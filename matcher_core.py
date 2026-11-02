# -*- coding: utf-8 -*-
"""
نواة المطابقة المشتركة بين نسخة Streamlit ونسخة سطح المكتب.

- دوال المطابقة الأصلية (match_names / match_names_dual) منسوخة كما هي حرفياً.
  الإضافة الوحيدة فيها: وسيط اختياري progress_cb لتحديث شريط التقدم، ولا يمس منطق المطابقة.
- إضافة قاعدة ثانية: كل قاعدة تُطابَق بنفس الدوال الأصلية منفصلة، ثم تُدمج النتائج:
    * وجد في قاعدة واحدة   -> تؤخذ نتيجتها كما هي + عمود "المصدر"
    * وجد في القاعدتين      -> ⚠️ تعارض بين القاعدتين، لا يُختار أي منهما تلقائياً
    * لم يوجد في أي منهما  -> ❌
"""
import bisect
import re
from io import BytesIO

import pandas as pd
from openpyxl.styles import PatternFill
from rapidfuzz import fuzz, process

# ============ دوال المساعدة (بدون أي تعديل عن الأصل) ============

def normalize_name(name):
    """تطبيع الاسم للمطابقة - الدالة الأصلية بدون تعديل"""
    if pd.isnull(name):
        return ""
    name = str(name).strip()
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()


def normalize_school_name(name):
    """تطبيع أسماء المدارس (بدون أي تعديل عن الأصل)"""
    if pd.isnull(name):
        return ""
    name = str(name).strip()
    name = re.sub(r'\bش\s*\.\s*', 'الشهيد ', name)
    name = re.sub(r'\b[اعمثبتر]\s*\.\s*', '', name)
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()


def get_first_three_words(name):
    if pd.isnull(name) or name == "":
        return ""
    words = str(name).split()
    return " ".join(words[:3]) if len(words) >= 3 else " ".join(words)


def is_first_three_words_match(name1, name2):
    words1 = name1.split()
    words2 = name2.split()
    length = min(len(words1), len(words2), 3)
    if length == 0:
        return False
    return all(words1[i] == words2[i] for i in range(length))


def school_similarity(a, b):
    """(بدون أي تعديل عن الأصل)"""
    if not a or not b:
        return 0
    return max(fuzz.ratio(a, b), fuzz.token_set_ratio(a, b))


SCHOOL_THRESHOLD = 70
PROGRESS_EVERY = 250


def _tick(progress_cb, done, total):
    if progress_cb is not None and (done % PROGRESS_EVERY == 0 or done == total):
        try:
            progress_cb(done, total)
        except Exception:
            pass  # تحديث الواجهة لا يجوز أن يوقف المطابقة


# ============ دوال المطابقة الأصلية (المنطق كما هو حرفياً) ============

def match_names(names_df, database_df, name_column_file, name_column_db,
                selected_columns=None, progress_cb=None):
    names_df = names_df.copy()
    database_df = database_df.copy()

    names_df["normalized_name"] = names_df[name_column_file].apply(normalize_name)
    database_df["normalized_name"] = database_df[name_column_db].apply(normalize_name)

    duplicate_names = set(
        database_df["normalized_name"][database_df["normalized_name"].duplicated(keep=False)]
    )

    database_df = database_df.drop_duplicates(subset=["normalized_name"])

    columns_to_get = [name_column_db]
    if selected_columns:
        columns_to_get.extend([col for col in selected_columns if col != name_column_db])

    database_map = database_df.set_index("normalized_name")[columns_to_get].to_dict(orient="index")
    db_names_list = list(database_map.keys())
    db_names_set = set(db_names_list)

    db_names_sorted = sorted(db_names_list)
    rank = {name: i for i, name in enumerate(db_names_list)}

    def find_prefix_match(normalized_name):
        candidates = []
        lo = bisect.bisect_left(db_names_sorted, normalized_name)
        hi = bisect.bisect_left(db_names_sorted, normalized_name + '￿')
        candidates.extend(db_names_sorted[lo:hi])
        for i in range(1, len(normalized_name) + 1):
            prefix = normalized_name[:i]
            if prefix in db_names_set:
                candidates.append(prefix)
        if not candidates:
            return None
        return min(candidates, key=lambda c: rank[c])

    matched_results = []
    total = len(names_df)

    for n, (original_name, normalized_name) in enumerate(
            zip(names_df[name_column_file], names_df["normalized_name"]), 1):
        best_match = None
        best_score = 0
        match_data = None

        if normalized_name in db_names_set:
            best_match = normalized_name
            best_score = 100
            match_data = database_map[best_match]
        else:
            result_extract = process.extractOne(
                normalized_name, db_names_list, scorer=fuzz.ratio
            )
            if result_extract:
                best_match, best_score, _ = result_extract

            if best_score >= 85 and best_match:
                if is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name):
                    match_data = database_map[best_match]

            if not match_data:
                db_name = find_prefix_match(normalized_name)
                if db_name is not None:
                    match_data = database_map[db_name]
                    best_match = db_name
                    best_score = fuzz.ratio(normalized_name, best_match)

        if match_data:
            note = "✅ تطابق"
            if best_match in duplicate_names:
                note = "⚠️ تطابق لكن يوجد أكثر من شخص بنفس الاسم - تحقق يدوياً"

            result = {
                "الاسم الأصلي": original_name,
                "الاسم المطابق": match_data[name_column_db],
                "نسبة التطابق": f"{round(best_score)}%",
                "ملاحظة": note
            }
            if selected_columns:
                for col in selected_columns:
                    if col in match_data:
                        result[col] = match_data[col]
            matched_results.append(result)
        else:
            result = {
                "الاسم الأصلي": original_name,
                "الاسم المطابق": "",
                "نسبة التطابق": "",
                "ملاحظة": "❌ لم يتم العثور على تطابق"
            }
            if selected_columns:
                for col in selected_columns:
                    result[col] = ""
            matched_results.append(result)

        _tick(progress_cb, n, total)

    return pd.DataFrame(matched_results)


def match_names_dual(names_df, database_df, name_column_file, name_column_db,
                     second_col_file, second_col_db, selected_columns=None, progress_cb=None):
    names_df = names_df.copy()
    database_df = database_df.copy()

    names_df["normalized_name"] = names_df[name_column_file].apply(normalize_name)
    database_df["normalized_name"] = database_df[name_column_db].apply(normalize_name)
    names_df["normalized_school"] = names_df[second_col_file].apply(normalize_school_name)
    database_df["normalized_school"] = database_df[second_col_db].apply(normalize_school_name)

    columns_to_keep = [name_column_db, second_col_db, "normalized_school"]
    if selected_columns:
        columns_to_keep.extend([c for c in selected_columns if c not in columns_to_keep])

    name_groups = {}
    cols_idx = {c: database_df.columns.get_loc(c) for c in columns_to_keep if c in database_df.columns}
    name_idx = database_df.columns.get_loc("normalized_name")
    for row in database_df.itertuples(index=False, name=None):
        key = row[name_idx]
        entry = {c: row[idx] for c, idx in cols_idx.items()}
        name_groups.setdefault(key, []).append(entry)

    db_unique_names = list(name_groups.keys())
    db_unique_names_set = set(db_unique_names)

    db_names_sorted = sorted(db_unique_names)
    rank = {name: i for i, name in enumerate(db_unique_names)}

    def find_prefix_match(normalized_name):
        candidates_names = []
        lo = bisect.bisect_left(db_names_sorted, normalized_name)
        hi = bisect.bisect_left(db_names_sorted, normalized_name + '￿')
        candidates_names.extend(db_names_sorted[lo:hi])
        for i in range(1, len(normalized_name) + 1):
            prefix = normalized_name[:i]
            if prefix in db_unique_names_set:
                candidates_names.append(prefix)
        if not candidates_names:
            return None
        return min(candidates_names, key=lambda c: rank[c])

    matched_results = []

    file_cols = [name_column_file, "normalized_name", second_col_file, "normalized_school"]
    file_idx = {c: names_df.columns.get_loc(c) for c in file_cols}
    total = len(names_df)

    for n, row in enumerate(names_df.itertuples(index=False, name=None), 1):
        original_name = row[file_idx[name_column_file]]
        normalized_name = row[file_idx["normalized_name"]]
        original_school = row[file_idx[second_col_file]]
        normalized_school = row[file_idx["normalized_school"]]

        best_match = None
        best_score = 0
        candidates = []

        if normalized_name in db_unique_names_set:
            best_match = normalized_name
            best_score = 100
            candidates = name_groups[best_match]
        else:
            result_extract = process.extractOne(
                normalized_name, db_unique_names, scorer=fuzz.ratio
            )
            if result_extract:
                best_match, best_score, _ = result_extract

            if best_score >= 85 and best_match:
                if is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name):
                    candidates = name_groups[best_match]

            if not candidates:
                found = find_prefix_match(normalized_name)
                if found is not None:
                    candidates = name_groups[found]
                    best_match = found
                    best_score = fuzz.ratio(normalized_name, best_match)

        chosen = None
        note_name = "❌ لم يتم العثور على تطابق"
        matched_school_display = ""
        school_score_display = ""
        note_school = "❌ لم يتم العثور على تطابق"

        if len(candidates) == 1:
            chosen = candidates[0]
            note_name = "✅ تطابق"
        elif len(candidates) > 1:
            scored = [
                (c, school_similarity(normalized_school, c["normalized_school"]))
                for c in candidates
            ]
            qualified = [c for c, s in scored if s >= SCHOOL_THRESHOLD]
            if len(qualified) == 1:
                chosen = qualified[0]
                note_name = "✅ تطابق (تم الفصل بواسطة المدرسة)"
            else:
                note_name = "⚠️ تطابق الاسم لكن يوجد أكثر من شخص بنفس الاسم - تحقق يدوياً"
                note_school = "⚠️ بانتظار تحقق يدوي"

        if chosen:
            sc = school_similarity(normalized_school, chosen["normalized_school"])
            matched_school_display = chosen[second_col_db]
            school_score_display = f"{round(sc)}%"
            note_school = "✅ تطابق" if sc >= SCHOOL_THRESHOLD else "❌ لم يتم العثور على تطابق"

        result = {
            "الاسم الأصلي": original_name,
            "الاسم المطابق": chosen[name_column_db] if chosen else "",
            "نسبة تطابق الاسم": f"{round(best_score)}%" if candidates else "",
            "ملاحظة الاسم": note_name,
            f"{second_col_file} (الأصلي)": original_school,
            f"{second_col_db} (المطابق)": matched_school_display,
            "نسبة تطابق المدرسة": school_score_display,
            "ملاحظة المدرسة": note_school,
        }

        if selected_columns:
            for col in selected_columns:
                result[col] = chosen.get(col, "") if chosen else ""

        matched_results.append(result)
        _tick(progress_cb, n, total)

    return pd.DataFrame(matched_results)


# ============ القاعدتان: تشغيل الأصل على كل قاعدة ثم الدمج ============

CONFLICT_MARK = "القاعدتين"
NOT_FOUND = "❌ لم يتم العثور على تطابق"


def _found(note):
    """هل وُجد الاسم في هذه القاعدة؟ (✅ تطابق أو ⚠️ مكرر داخل نفس القاعدة)"""
    note = str(note)
    return ("✅" in note) or ("⚠️" in note)


def _sub_progress(progress_cb, part, parts):
    """تحويل تقدم مطابقة قاعدة واحدة إلى تقدم كلي"""
    if progress_cb is None:
        return None

    def cb(done, total):
        progress_cb(part * total + done, parts * total)
    return cb


def cross_database_names(db1, db2, name_col1, name_col2, label1, label2,
                         school_col1=None, school_col2=None):
    """أسماء موجودة حرفياً (بعد التطبيع) في القاعدتين معاً - فحص للقاعدتين نفسيهما"""
    a = pd.DataFrame({"normalized_name": db1[name_col1].apply(normalize_name),
                      f"الاسم في {label1}": db1[name_col1]})
    b = pd.DataFrame({"normalized_name": db2[name_col2].apply(normalize_name),
                      f"الاسم في {label2}": db2[name_col2]})
    if school_col1:
        a[f"المدرسة في {label1}"] = db1[school_col1].values
    if school_col2:
        b[f"المدرسة في {label2}"] = db2[school_col2].values
    a = a[a["normalized_name"] != ""]
    b = b[b["normalized_name"] != ""]
    common = a.merge(b, on="normalized_name", how="inner")
    return common.drop(columns=["normalized_name"]).reset_index(drop=True)


def match_two_databases(names_df, db1, db2, name_col_file, name_col1, name_col2,
                        label1="الملاك", label2="العقود",
                        second_col_file=None, second_col1=None, second_col2=None,
                        selected_columns=None, progress_cb=None):
    """
    يعيد: (النتائج الموحدة, جدول التعارضات, جدول الأسماء المشتركة بين القاعدتين)
    كل قاعدة تُطابق بالدالة الأصلية نفسها دون أي تغيير في الشروط.
    """
    selected_columns = list(selected_columns or [])
    dual = bool(second_col_file and second_col1 and second_col2)

    excl1 = {name_col1} | ({second_col1} if dual else set())
    excl2 = {name_col2} | ({second_col2} if dual else set())
    sel1 = [c for c in selected_columns if c in db1.columns and c not in excl1]
    sel2 = [c for c in selected_columns if c in db2.columns and c not in excl2]

    if dual:
        r1 = match_names_dual(names_df, db1, name_col_file, name_col1, second_col_file,
                              second_col1, sel1, _sub_progress(progress_cb, 0, 2))
        r2 = match_names_dual(names_df, db2, name_col_file, name_col2, second_col_file,
                              second_col2, sel2, _sub_progress(progress_cb, 1, 2))
        note_col, score_col = "ملاحظة الاسم", "نسبة تطابق الاسم"
        r1 = r1.rename(columns={f"{second_col_file} (الأصلي)": "المدرسة (الأصلية)",
                                f"{second_col1} (المطابق)": "المدرسة (المطابقة)"})
        r2 = r2.rename(columns={f"{second_col_file} (الأصلي)": "المدرسة (الأصلية)",
                                f"{second_col2} (المطابق)": "المدرسة (المطابقة)"})
        base_cols = ["الاسم الأصلي", "المصدر", "الاسم المطابق", "نسبة تطابق الاسم", "ملاحظة الاسم",
                     "المدرسة (الأصلية)", "المدرسة (المطابقة)", "نسبة تطابق المدرسة", "ملاحظة المدرسة"]
    else:
        r1 = match_names(names_df, db1, name_col_file, name_col1, sel1,
                         _sub_progress(progress_cb, 0, 2))
        r2 = match_names(names_df, db2, name_col_file, name_col2, sel2,
                         _sub_progress(progress_cb, 1, 2))
        note_col, score_col = "ملاحظة", "نسبة التطابق"
        base_cols = ["الاسم الأصلي", "المصدر", "الاسم المطابق", "نسبة التطابق", "ملاحظة"]

    out_cols = base_cols + selected_columns + [f"مطابقة {label1}", f"مطابقة {label2}"]
    rows, conflicts = [], []
    recs1 = r1.to_dict("records")
    recs2 = r2.to_dict("records")

    def brief(rec):
        if not _found(rec[note_col]):
            return ""
        nm = rec["الاسم المطابق"] or "(عدة أشخاص بنفس الاسم)"
        return f"{nm} ({rec[score_col]})"

    for a, b in zip(recs1, recs2):
        f1, f2 = _found(a[note_col]), _found(b[note_col])
        out = {c: "" for c in out_cols}
        out["الاسم الأصلي"] = a["الاسم الأصلي"]
        out[f"مطابقة {label1}"] = brief(a)
        out[f"مطابقة {label2}"] = brief(b)
        if dual:
            out["المدرسة (الأصلية)"] = a["المدرسة (الأصلية)"]

        if f1 and f2:
            out["المصدر"] = f"{label1} + {label2}"
            out[note_col] = f"⚠️ الاسم موجود في القاعدتين ({label1} و{label2}) - تحقق يدوياً"
            if dual:
                s1 = "✅" in str(a["ملاحظة المدرسة"])
                s2 = "✅" in str(b["ملاحظة المدرسة"])
                where = [lbl for lbl, ok in ((label1, s1), (label2, s2)) if ok]
                out["ملاحظة المدرسة"] = ("⚠️ المدرسة تطابق في: " + " و".join(where)) if where \
                    else "⚠️ بانتظار تحقق يدوي"
            c = {"الاسم الأصلي": a["الاسم الأصلي"]}
            if dual:
                c["المدرسة (الأصلية)"] = a["المدرسة (الأصلية)"]
            for lbl, rec, sel in ((label1, a, sel1), (label2, b, sel2)):
                c[f"الاسم المطابق - {lbl}"] = rec["الاسم المطابق"]
                c[f"النسبة - {lbl}"] = rec[score_col]
                c[f"الملاحظة - {lbl}"] = rec[note_col]
                if dual:
                    c[f"المدرسة - {lbl}"] = rec["المدرسة (المطابقة)"]
                    c[f"نسبة المدرسة - {lbl}"] = rec["نسبة تطابق المدرسة"]
                for col in sel:
                    c[f"{col} - {lbl}"] = rec.get(col, "")
            conflicts.append(c)
        elif f1 or f2:
            rec, lbl = (a, label1) if f1 else (b, label2)
            for col in base_cols:
                if col in rec:
                    out[col] = rec[col]
            out["المصدر"] = lbl
            for col in selected_columns:
                if col in rec:
                    out[col] = rec[col]
        else:
            out[note_col] = NOT_FOUND
            if dual:
                out["ملاحظة المدرسة"] = NOT_FOUND
        rows.append(out)

    results = pd.DataFrame(rows, columns=out_cols)
    conflicts_df = pd.DataFrame(conflicts)
    common_df = cross_database_names(db1, db2, name_col1, name_col2, label1, label2,
                                     second_col1 if dual else None, second_col2 if dual else None)
    return results, conflicts_df, common_df


# ============ دوال التصدير ============

RED = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
GREEN = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
YELLOW = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
ORANGE = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
PURPLE = PatternFill(start_color="E1BEE7", end_color="E1BEE7", fill_type="solid")


def _autofit(ws):
    for column in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
        ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 60)


def _color_sheet(ws, dual):
    hdr = {cell.value: i for i, cell in enumerate(ws[1])}
    if dual:
        ni, si = hdr.get("ملاحظة الاسم"), hdr.get("ملاحظة المدرسة")
    else:
        ni, si = hdr.get("ملاحظة"), None
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        name_val = str(row[ni].value) if ni is not None else ""
        school_val = str(row[si].value) if si is not None else ""
        if CONFLICT_MARK in name_val:
            fill = PURPLE
        elif dual:
            # نفس ألوان to_excel_dual الأصلية
            if "⚠️" in name_val or "⚠️" in school_val:
                fill = ORANGE
            elif "✅" in name_val and "✅" in school_val:
                fill = GREEN
            elif "✅" in name_val or "✅" in school_val:
                fill = YELLOW
            else:
                fill = RED
        else:
            # نفس ألوان to_excel_basic الأصلية (✅ بدون لون)
            if "⚠️" in name_val:
                fill = ORANGE
            elif "❌" in name_val:
                fill = RED
            else:
                continue
        for cell in row:
            cell.fill = fill


def to_excel_basic(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.book.active
        _autofit(ws)
        _color_sheet(ws, dual=False)
    output.seek(0)
    return output


def to_excel_dual(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.book.active
        _autofit(ws)
        _color_sheet(ws, dual=True)
    output.seek(0)
    return output


def to_excel_two_db(results, conflicts_df, common_df, dual):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        results.to_excel(writer, index=False, sheet_name="النتائج")
        ws = writer.sheets["النتائج"]
        ws.sheet_view.rightToLeft = True
        _autofit(ws)
        _color_sheet(ws, dual)
        for df, name in ((conflicts_df, "تعارض بين القاعدتين"),
                         (common_df, "أسماء مشتركة بين القاعدتين")):
            if df is not None and not df.empty:
                df.to_excel(writer, index=False, sheet_name=name)
                w = writer.sheets[name]
                w.sheet_view.rightToLeft = True
                _autofit(w)
                for row in w.iter_rows(min_row=2, max_row=w.max_row):
                    for cell in row:
                        cell.fill = PURPLE
    output.seek(0)
    return output


# ============ ملخص الإحصائيات (مشترك بين النسختين) ============

def summary(results, dual, two_db):
    nc = "ملاحظة الاسم" if dual else "ملاحظة"
    notes = results[nc].astype(str)
    s = {"📊 الإجمالي": len(results)}
    conflict = notes.str.contains(CONFLICT_MARK, regex=False)
    warn = notes.str.contains("⚠️", regex=False) & ~conflict
    ok = notes.str.contains("✅", regex=False)
    s["✅ متطابق" if not dual else "✅ تطابق الاسم"] = int(ok.sum())
    if dual:
        sch = results["ملاحظة المدرسة"].astype(str).str.contains("✅", regex=False)
        s["✅ تطابق المدرسة"] = int(sch.sum())
        s["🎯 تطابق الاثنين"] = int((ok & sch).sum())
    s["⚠️ مكرر - تحقق يدوي"] = int(warn.sum())
    if two_db:
        s["🟣 موجود في القاعدتين"] = int(conflict.sum())
    s["❌ غير متطابق"] = int(notes.str.contains("❌", regex=False).sum())
    return s
