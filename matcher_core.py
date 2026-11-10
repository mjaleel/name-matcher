# -*- coding: utf-8 -*-
"""
نواة المطابقة المشتركة بين نسخة Streamlit ونسخة سطح المكتب — الإصدار 2.

نفس الدوال ونفس الوسائط ونفس شكل النتائج والألوان، فلا يحتاج app_streamlit.py أي تعديل.

ما تغيّر عن الإصدار الأول:
1. التطبيع: حذف الكشيدة (ـ) والتشكيل، وتوحيد الحروف الفارسية (چ گ ک پ ڤ ی) و(ؤ ← و: رؤى = روى).
2. لا يُقبل تطابق تقريبي إلا إذا تقابلت الكلمات كلمة بكلمة؛ الفرق المسموح فرق نقاط
   أو حرفين متبادلين (حصاد/حماد، هيلم/هليم). كان شرط "النسبة ≥ 85 وأول ثلاث كلمات"
   يقبل «علي نجم عبدالله جويج» على أنه «علي نجم عبدالله حسين».
3. المطابقة بالبادئة على حدود الكلمات فقط، ويُشترط أن يكون اسم القاعدة رباعياً على الأقل
   (الورقة فيها اللقب والقاعدة بدونه).
4. إذا تقارب أكثر من شخص بنفس الدرجة: ⚠️ تحقق يدوياً بدل اختيار أحدهم بصمت.
5. عمود جديد «تشابه قريب»: الكلمات التي اختلفت بين الاسمين، وأي شخص آخر قريب جداً،
   أو مرشح قريب إذا لم يوجد تطابق مؤكد (❔). العتبات صارمة: لا يظهر إلا القريب فعلاً.
6. القاعدتان: إذا وُجد الاسم في قاعدة ووُجد اسم قريب جداً في الأخرى يُعلَّم ⚠️ بنفس لون
   التعارض، ويُضاف لورقة التعارضات.
"""
import bisect
import re
from io import BytesIO

import pandas as pd
from openpyxl.styles import PatternFill
from rapidfuzz import fuzz, process

# ============ التطبيع ============

_NOISE = re.compile(r'[ـً-ٰٟ]')          # كشيدة وتشكيل
_PERSIAN = str.maketrans({'چ': 'ج', 'گ': 'ك', 'ک': 'ك', 'پ': 'ب', 'ڤ': 'ف', 'ی': 'ي', 'ؤ': 'و'})


def normalize_name(name):
    """تطبيع الاسم للمطابقة: قواعد الأصل + حذف الكشيدة والتشكيل + الحروف الفارسية"""
    if pd.isnull(name):
        return ""
    name = _NOISE.sub('', str(name).strip()).translate(_PERSIAN)
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()


def normalize_school_name(name):
    """تطبيع أسماء المدارس (الأصل + حذف الكشيدة والحروف الفارسية)"""
    if pd.isnull(name):
        return ""
    name = _NOISE.sub('', str(name).strip()).translate(_PERSIAN)
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

# عتبات التشابه القريب (صارمة عمداً)
CANDIDATE_THRESHOLD = 85     # مرشح قريب لاسم لم يُطابَق
OTHER_CLOSE_THRESHOLD = 92   # شخص آخر قريب جداً من اسم مطابَق
CROSS_DB_THRESHOLD = 88      # اسم قريب في القاعدة الأخرى
WORD_THRESHOLD = 70          # أدنى تشابه لكلمة تُعدّ مقابلة لكلمة (فرق نقاط أو حرفان)
STRONG_DIFF_COST = 0.6       # فرق كلمة لا يمنع التطابق التلقائي: نقطتان أو تبديل حرفين فقط

NOTE_CANDIDATE = "❔ مرشح قريب - تحقق"
NOTE_CLOSE_RIVALS = "⚠️ أكثر من اسم قريب بنفس الدرجة - تحقق يدوياً"


def _tick(progress_cb, done, total):
    if progress_cb is not None and (done % PROGRESS_EVERY == 0 or done == total):
        try:
            progress_cb(done, total)
        except Exception:
            pass  # تحديث الواجهة لا يجوز أن يوقف المطابقة


# ============ تشابه الكلمات ============

_CLS = {}
for _k, _v in (('ب', 'بتثنيى'), ('ح', 'جحخ'), ('د', 'دذ'), ('ر', 'رز'), ('س', 'سش'),
               ('ص', 'صض'), ('ط', 'طظ'), ('ع', 'عغ'), ('ف', 'فق'), ('ه', 'هة')):
    for _c in _v:
        _CLS[_c] = _k


def _word_cost(a, b):
    """تكلفة التحويل بين كلمتين: حرف متشابه بالنقاط 0.3، حرفان متبادلان 0.5، حرف مختلف فعلاً 1"""
    if a == b:
        return 0.0
    n, m = len(a), len(b)
    if not n or not m:
        return float(max(n, m))
    prev2 = None
    prev = list(range(m + 1))
    for i in range(1, n + 1):
        cur = [i] + [0] * m
        x = a[i - 1]
        for j in range(1, m + 1):
            y = b[j - 1]
            sub = 0 if x == y else (0.3 if _CLS.get(x, x) == _CLS.get(y, y) else 1)
            v = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + sub)
            if prev2 is not None and j > 1 and x == b[j - 2] and a[i - 2] == y and x != y:
                v = min(v, prev2[j - 2] + 0.5)
            cur[j] = v
        prev2, prev = prev, cur
    return prev[m]


def _word_sim(a, b):
    """تشابه كلمتين 0-100"""
    if a == b:
        return 100.0
    n, m = len(a), len(b)
    if not n or not m:
        return 0.0
    return 100.0 * (1 - _word_cost(a, b) / max(n, m))


def _words_agree(qw, cw):
    """تقابل كلمة بكلمة على الجزء المشترك. يعيد (عدد الكلمات المختلفة, عدد الكلمات المقارنة, الفروق) أو None"""
    k = min(len(qw), len(cw))
    diffs = []
    for i in range(k):
        if qw[i] == cw[i]:
            continue
        if _word_sim(qw[i], cw[i]) < WORD_THRESHOLD:
            return None
        diffs.append((qw[i], cw[i]))
    if len(diffs) > 2:
        return None
    return len(diffs), k, diffs


def _weak(diffs):
    """هل في الفروق حرف مختلف فعلاً (لا مجرد نقاط أو تبديل)؟ مثل جاسب/جاسم"""
    return any(_word_cost(a, b) > STRONG_DIFF_COST for a, b in diffs)


def _align_score(qw, cw):
    """تشابه اسمين 0-100 يتحمل كلمة ناقصة أو كلمتين ملتصقتين؛ الاسم الأول وزنه مضاعف"""
    n, m = len(qw), len(cw)
    INF = 1e9
    C = [[INF] * (m + 1) for _ in range(n + 1)]
    C[0][0] = 0.0
    for i in range(n + 1):
        for j in range(m + 1):
            c0 = C[i][j]
            if c0 >= INF:
                continue
            w = 2 if (i == 0 and j == 0) else 1
            if i < n and j < m:
                C[i + 1][j + 1] = min(C[i + 1][j + 1], c0 + w * (1 - _word_sim(qw[i], cw[j]) / 100))
            if i < n:   # كلمة في الورقة فقط
                C[i + 1][j] = min(C[i + 1][j], c0 + (0.15 if j == m else (2 if i == 0 else 0.6)))
            if j < m:   # كلمة في القاعدة فقط
                C[i][j + 1] = min(C[i][j + 1], c0 + (0.3 if i == n else (2 if j == 0 else 0.6)))
            if i + 1 < n and j < m:
                C[i + 2][j + 1] = min(C[i + 2][j + 1], c0 + 0.1 + (1 - _word_sim(qw[i] + qw[i + 1], cw[j]) / 100))
            if i < n and j + 1 < m:
                C[i + 1][j + 2] = min(C[i + 1][j + 2], c0 + 0.1 + (1 - _word_sim(qw[i], cw[j] + cw[j + 1]) / 100))
    return max(0.0, 100 * (1 - C[n][m] / max(3, min(n, m) + 1)))


# ============ فهرس القاعدة ============

class _Index:
    """يُبنى مرة واحدة لكل قاعدة: بحث سريع بدون مقارنة كل اسم بكل اسم"""

    def __init__(self, names_in_order):
        self.names = list(names_in_order)
        self.set = set(self.names)
        self.sorted = sorted(self.names)
        self.first3 = {}
        for nm in self.names:
            self.first3.setdefault(get_first_three_words(nm), []).append(nm)

    def _starting_with(self, n):
        key = n + " "
        lo = bisect.bisect_left(self.sorted, key)
        out = []
        for i in range(lo, min(lo + 50, len(self.sorted))):
            if not self.sorted[i].startswith(key):
                break
            out.append(self.sorted[i])
        return out

    def _pool(self, n, qw):
        pool = set(self.first3.get(get_first_three_words(n), []))
        pool.update(self._starting_with(n))
        for k in range(len(qw) - 1, 2, -1):          # الورقة فيها لقب والقاعدة بدونه
            p = " ".join(qw[:k])
            if p in self.set:
                pool.add(p)
        # أقرب الأسماء بالحروف، للاسم كاملاً وبدون آخر كلمة وآخر كلمتين
        for k in (len(qw), len(qw) - 1, len(qw) - 2):
            if k >= 3:
                for c, _, _ in process.extract(" ".join(qw[:k]), self.names, scorer=fuzz.ratio,
                                               limit=12, score_cutoff=70):
                    pool.add(c)
        return pool

    def find(self, n):
        """
        يعيد dict:
          best: الاسم المطابق (مطبّعاً) أو None
          score: النسبة
          rivals: أسماء بنفس الدرجة تقريباً (تحتاج تحقق يدوي)
          diffs: الكلمات المختلفة بين الورقة والقاعدة
          candidate: (اسم, نسبة) مرشح قريب حين لا يوجد تطابق مؤكد
          close: [(اسم, نسبة)] أشخاص آخرون قريبون جداً
        """
        res = dict(best=None, score=0, rivals=[], diffs=[], candidate=None, close=[], paper_longer=False)
        if not n:
            return res
        qw = n.split()
        if n in self.set:
            res.update(best=n, score=100)
            pool = self._pool(n, qw)
        else:
            pool = self._pool(n, qw)
            ok = []
            for c in pool:
                cw = c.split()
                if len(cw) < 3:
                    continue
                g = _words_agree(qw, cw)
                if not g or g[1] < 3:
                    continue
                if len(cw) < len(qw) and len(cw) < 4:     # القاعدة أقصر: يجب أن تكون رباعية على الأقل
                    continue
                shared = fuzz.ratio(" ".join(qw[:g[1]]), " ".join(cw[:g[1]]))
                # (فرق قوي؟, كلمات مقارنة, عدد الفروق, نسبة الجزء المشترك, الاسم, الفروق)
                ok.append((_weak(g[2]), g[1], g[0], shared, c, g[2]))
            # الفروق القوية (نقاط/تبديل) أولاً، ثم الأكثر كلمات مقارنة، ثم الأقل فروقاً
            ok.sort(key=lambda t: (t[0], -t[1], t[2], -t[3]))
            if ok:
                weak, k, soft, sc, c, diffs = ok[0]
                res.update(diffs=diffs, paper_longer=len(c.split()) < len(qw))
                if weak:
                    # حرف مختلف فعلاً (جاسب/جاسم): لا يُقبل تلقائياً، يُعرض مرشحاً قريباً فقط
                    if sc >= CANDIDATE_THRESHOLD:
                        res["candidate"] = (c, round(sc))
                else:
                    res.update(best=c, score=sc)
                    res["rivals"] = [o[4] for o in ok[1:] if not o[0] and o[1] == k
                                     and o[2] <= soft + 1 and sc - o[3] < 6]
        # تشابه قريب: أشخاص آخرون، أو مرشح لاسم لم يُطابَق
        scored = []
        for c in pool:
            if c == res["best"] or c in res["rivals"] or (res["candidate"] and c == res["candidate"][0]):
                continue
            cw = c.split()
            if _word_sim(qw[0], cw[0]) < 80:               # الاسم الأول مختلف = شخص آخر
                continue
            if len(qw) > 1 and len(cw) > 1 and _word_sim(qw[1], cw[1]) < WORD_THRESHOLD:
                continue                                    # اسم الأب مختلف = شخص آخر (لا مرشح بعيد)
            scored.append((c, round(_align_score(qw, cw))))
        scored.sort(key=lambda t: -t[1])
        if res["best"] is None:
            if res["candidate"]:
                res["close"] = [s for s in scored[:2] if s[1] >= CANDIDATE_THRESHOLD]
            elif scored and scored[0][1] >= CANDIDATE_THRESHOLD:
                res["candidate"] = scored[0]
                res["close"] = [s for s in scored[1:3] if s[1] >= CANDIDATE_THRESHOLD]
        else:
            res["close"] = [s for s in scored[:2] if s[1] >= OTHER_CLOSE_THRESHOLD]
        return res


def _show(w):
    """الكلمة المطبّعة للعرض: التطبيع يحوّل ه إلى ة، فنعيدها كما يكتبها الناس"""
    return w.replace("ة", "ه")


def _similarity_note(res, label_of):
    """نص عمود «تشابه قريب»"""
    parts = []
    if res["diffs"]:
        parts.append("اختلاف حرفي: " + "، ".join(f"{_show(a)} ← {_show(b)}" for a, b in res["diffs"]))
    if res["paper_longer"]:
        parts.append("في الورقة كلمات زائدة (اللقب)")
    if res["candidate"]:
        c, s = res["candidate"]
        parts.append(f"مرشح قريب: {label_of(c)} ({s}%)")
    for c, s in res["close"]:
        parts.append(f"اسم قريب آخر: {label_of(c)} ({s}%)")
    return " | ".join(parts)


# ============ المطابقة ============

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
    database_df = database_df[database_df["normalized_name"] != ""]

    columns_to_get = [name_column_db]
    if selected_columns:
        columns_to_get.extend([col for col in selected_columns if col != name_column_db])

    database_map = database_df.set_index("normalized_name")[columns_to_get].to_dict(orient="index")
    index = _Index(database_map.keys())
    label_of = lambda c: database_map[c][name_column_db]

    matched_results = []
    total = len(names_df)
    for n, (original_name, normalized_name) in enumerate(
            zip(names_df[name_column_file], names_df["normalized_name"]), 1):
        res = index.find(normalized_name)
        result = {"الاسم الأصلي": original_name, "الاسم المطابق": "", "نسبة التطابق": "", "ملاحظة": ""}
        data = None
        if res["best"] and not res["rivals"]:
            data = database_map[res["best"]]
            result["ملاحظة"] = ("⚠️ تطابق لكن يوجد أكثر من شخص بنفس الاسم - تحقق يدوياً"
                               if res["best"] in duplicate_names else "✅ تطابق")
            result["نسبة التطابق"] = f"{round(res['score'])}%"
        elif res["best"]:
            result["ملاحظة"] = NOTE_CLOSE_RIVALS
            res["close"] = [(c, round(_align_score(normalized_name.split(), c.split())))
                            for c in [res["best"]] + res["rivals"]][:3]
            res["diffs"] = []
        elif res["candidate"]:
            data = database_map[res["candidate"][0]]
            result["ملاحظة"] = NOTE_CANDIDATE
            result["نسبة التطابق"] = f"{res['candidate'][1]}%"
        else:
            result["ملاحظة"] = "❌ لم يتم العثور على تطابق"
        if data:
            result["الاسم المطابق"] = data[name_column_db]
        result["تشابه قريب"] = _similarity_note(res, label_of)
        if selected_columns:
            for col in selected_columns:
                result[col] = data.get(col, "") if data else ""
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
        if not key:
            continue
        entry = {c: row[idx] for c, idx in cols_idx.items()}
        name_groups.setdefault(key, []).append(entry)

    index = _Index(name_groups.keys())
    label_of = lambda c: name_groups[c][0][name_column_db]

    matched_results = []
    file_cols = [name_column_file, "normalized_name", second_col_file, "normalized_school"]
    file_idx = {c: names_df.columns.get_loc(c) for c in file_cols}
    total = len(names_df)

    for n, row in enumerate(names_df.itertuples(index=False, name=None), 1):
        original_name = row[file_idx[name_column_file]]
        normalized_name = row[file_idx["normalized_name"]]
        original_school = row[file_idx[second_col_file]]
        normalized_school = row[file_idx["normalized_school"]]

        res = index.find(normalized_name)
        chosen = None
        note_name = "❌ لم يتم العثور على تطابق"
        note_school = "❌ لم يتم العثور على تطابق"
        matched_school_display = school_score_display = ""
        score_display = ""

        target, candidate_mode = None, False
        if res["best"] and not res["rivals"]:
            target = res["best"]
        elif res["best"]:
            note_name = NOTE_CLOSE_RIVALS
            note_school = "⚠️ بانتظار تحقق يدوي"
            res["close"] = [(c, round(_align_score(normalized_name.split(), c.split())))
                            for c in [res["best"]] + res["rivals"]][:3]
            res["diffs"] = []
        elif res["candidate"]:
            target, candidate_mode = res["candidate"][0], True

        if target:
            candidates = name_groups[target]
            if len(candidates) == 1:
                chosen = candidates[0]
                note_name = NOTE_CANDIDATE if candidate_mode else "✅ تطابق"
            else:
                scored = [(c, school_similarity(normalized_school, c["normalized_school"])) for c in candidates]
                qualified = [c for c, s in scored if s >= SCHOOL_THRESHOLD]
                if len(qualified) == 1:
                    chosen = qualified[0]
                    note_name = NOTE_CANDIDATE if candidate_mode else "✅ تطابق (تم الفصل بواسطة المدرسة)"
                else:
                    note_name = "⚠️ تطابق الاسم لكن يوجد أكثر من شخص بنفس الاسم - تحقق يدوياً"
                    note_school = "⚠️ بانتظار تحقق يدوي"
            score_display = f"{res['candidate'][1]}%" if candidate_mode else f"{round(res['score'])}%"

        if chosen:
            sc = school_similarity(normalized_school, chosen["normalized_school"])
            matched_school_display = chosen[second_col_db]
            school_score_display = f"{round(sc)}%"
            note_school = "✅ تطابق" if sc >= SCHOOL_THRESHOLD else "❌ لم يتم العثور على تطابق"

        result = {
            "الاسم الأصلي": original_name,
            "الاسم المطابق": chosen[name_column_db] if chosen else "",
            "نسبة تطابق الاسم": score_display,
            "ملاحظة الاسم": note_name,
            "تشابه قريب": _similarity_note(res, label_of),
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


# ============ القاعدتان ============

CONFLICT_MARK = "القاعدتين"
NOT_FOUND = "❌ لم يتم العثور على تطابق"


def _found(note):
    """هل وُجد الاسم في هذه القاعدة؟ (✅ تطابق أو ⚠️ تحقق داخل نفس القاعدة)"""
    note = str(note)
    return ("✅" in note) or ("⚠️" in note)


def _is_candidate(note):
    return "❔" in str(note)


def _sub_progress(progress_cb, part, parts):
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


def _pct(v):
    try:
        return int(str(v).rstrip("%"))
    except ValueError:
        return 0


def match_two_databases(names_df, db1, db2, name_col_file, name_col1, name_col2,
                        label1="الملاك", label2="العقود",
                        second_col_file=None, second_col1=None, second_col2=None,
                        selected_columns=None, progress_cb=None):
    """
    يعيد: (النتائج الموحدة, جدول التعارضات, جدول الأسماء المشتركة بين القاعدتين)
    كل قاعدة تُطابق منفصلة ثم تُدمج:
      * في قاعدة واحدة فقط                     -> نتيجتها + عمود "المصدر"
      * في القاعدتين                           -> ⚠️ تعارض، لا يُختار أي منهما تلقائياً
      * في قاعدة، ومرشح قريب جداً في الأخرى     -> نتيجة القاعدة + ⚠️ تنبيه تشابه بين القاعدتين
      * لا تطابق مؤكد، ومرشح قريب في إحداهما     -> ❔ مرشح قريب
      * لا شيء                                 -> ❌
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
                     "تشابه قريب", "المدرسة (الأصلية)", "المدرسة (المطابقة)", "نسبة تطابق المدرسة",
                     "ملاحظة المدرسة"]
    else:
        r1 = match_names(names_df, db1, name_col_file, name_col1, sel1,
                         _sub_progress(progress_cb, 0, 2))
        r2 = match_names(names_df, db2, name_col_file, name_col2, sel2,
                         _sub_progress(progress_cb, 1, 2))
        note_col, score_col = "ملاحظة", "نسبة التطابق"
        base_cols = ["الاسم الأصلي", "المصدر", "الاسم المطابق", "نسبة التطابق", "ملاحظة", "تشابه قريب"]

    out_cols = base_cols + selected_columns + [f"مطابقة {label1}", f"مطابقة {label2}"]
    rows, conflicts = [], []
    recs1 = r1.to_dict("records")
    recs2 = r2.to_dict("records")

    def brief(rec):
        if not (_found(rec[note_col]) or _is_candidate(rec[note_col])):
            return ""
        nm = rec["الاسم المطابق"] or "(عدة أشخاص بنفس الاسم)"
        tag = " - مرشح" if _is_candidate(rec[note_col]) else ""
        return f"{nm} ({rec[score_col]}){tag}"

    def conflict_row(a, b, kind):
        c = {"نوع التعارض": kind, "الاسم الأصلي": a["الاسم الأصلي"]}
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
        return c

    def copy_rec(out, rec, lbl):
        for col in base_cols:
            if col in rec:
                out[col] = rec[col]
        out["المصدر"] = lbl
        for col in selected_columns:
            if col in rec:
                out[col] = rec[col]

    for a, b in zip(recs1, recs2):
        f1, f2 = _found(a[note_col]), _found(b[note_col])
        c1, c2 = _is_candidate(a[note_col]), _is_candidate(b[note_col])
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
            conflicts.append(conflict_row(a, b, "موجود في القاعدتين"))
        elif f1 or f2:
            rec, lbl = (a, label1) if f1 else (b, label2)
            other, olbl = (b, label2) if f1 else (a, label1)
            copy_rec(out, rec, lbl)
            # مرشح قريب جداً في القاعدة الأخرى: يجب الانتباه له
            if _is_candidate(other[note_col]) and _pct(other[score_col]) >= CROSS_DB_THRESHOLD:
                out[note_col] = (f"⚠️ موجود في {lbl}، ويوجد اسم قريب جداً في {olbl}: "
                                 f"{other['الاسم المطابق']} ({other[score_col]}) - تشابه بين {CONFLICT_MARK}")
                conflicts.append(conflict_row(a, b, f"موجود في {lbl} وقريب في {olbl}"))
            elif _is_candidate(other[note_col]):
                # مرشح في القاعدة الأخرى أبعد قليلاً: معلومة فقط في «تشابه قريب» بدون تغيير اللون
                extra = f"في {olbl} مرشح: {other['الاسم المطابق']} ({other[score_col]})"
                out["تشابه قريب"] = " | ".join(x for x in (out.get("تشابه قريب", ""), extra) if x)
        elif c1 or c2:
            if c1 and c2:
                rec, lbl = (a, label1) if _pct(a[score_col]) >= _pct(b[score_col]) else (b, label2)
                copy_rec(out, rec, lbl)
                out[note_col] = f"{NOTE_CANDIDATE} (مرشحان في {CONFLICT_MARK})"
                conflicts.append(conflict_row(a, b, "مرشح قريب في القاعدتين"))
            else:
                rec, lbl = (a, label1) if c1 else (b, label2)
                copy_rec(out, rec, lbl)
        else:
            out[note_col] = NOT_FOUND
            if dual:
                out["ملاحظة المدرسة"] = NOT_FOUND
            near = " | ".join(x for x in (a.get("تشابه قريب", ""), b.get("تشابه قريب", "")) if x)
            out["تشابه قريب"] = near
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
BLUE = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")   # ❔ مرشح قريب


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
        elif "❔" in name_val:
            fill = BLUE
        elif dual:
            if "⚠️" in name_val or "⚠️" in school_val:
                fill = ORANGE
            elif "✅" in name_val and "✅" in school_val:
                fill = GREEN
            elif "✅" in name_val or "✅" in school_val:
                fill = YELLOW
            else:
                fill = RED
        else:
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


# ============ ملخص الإحصائيات ============

def summary(results, dual, two_db):
    nc = "ملاحظة الاسم" if dual else "ملاحظة"
    notes = results[nc].astype(str)
    s = {"📊 الإجمالي": len(results)}
    conflict = notes.str.contains(CONFLICT_MARK, regex=False)
    warn = notes.str.contains("⚠️", regex=False) & ~conflict
    ok = notes.str.contains("✅", regex=False)
    cand = notes.str.contains("❔", regex=False) & ~conflict
    s["✅ متطابق" if not dual else "✅ تطابق الاسم"] = int(ok.sum())
    if dual:
        sch = results["ملاحظة المدرسة"].astype(str).str.contains("✅", regex=False)
        s["✅ تطابق المدرسة"] = int(sch.sum())
        s["🎯 تطابق الاثنين"] = int((ok & sch).sum())
    s["⚠️ تحقق يدوي"] = int(warn.sum())
    s["❔ مرشح قريب"] = int(cand.sum())
    if two_db:
        s["🟣 القاعدتين"] = int(conflict.sum())
    s["❌ غير متطابق"] = int(notes.str.contains("❌", regex=False).sum())
    return s
