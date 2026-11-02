# -*- coding: utf-8 -*-
# تشغيل:  streamlit run app_streamlit.py   (يجب أن يكون matcher_core.py بجانبه)
import pandas as pd
import streamlit as st

import matcher_core as mc

st.set_page_config(page_title="مطابقة الأسماء", layout="wide")
st.markdown("### 👨‍💻 برمجة: محمد عبدالجليل")
st.title("🔐 نظام مطابقة الأسماء")


@st.cache_data(show_spinner="جاري قراءة الملف...")
def load_excel(data: bytes, _name: str):
    from io import BytesIO
    return pd.read_excel(BytesIO(data))


def show_table(container, df):
    """عرض فقط: تحويل لنص لتجنب أخطاء Arrow مع الأعمدة المختلطة (لا يؤثر على ملف الإكسل)"""
    view = df.astype(str).replace({"nan": "", "None": "", "NaT": ""})
    try:
        container.dataframe(view, width="stretch")
    except TypeError:  # إصدارات Streamlit القديمة
        container.dataframe(view, use_container_width=True)


password = st.text_input("أدخل كلمة المرور:", type="password")

if password == "mjaleel":
    st.success("✅ تم التحقق من كلمة المرور.")
    st.subheader("📋 مطابقة الأسماء العامة")

    c1, c2, c3 = st.columns(3)
    with c1:
        file1 = st.file_uploader("📄 ملف الأسماء", type="xlsx", key="file1_name")
    with c2:
        file2 = st.file_uploader("📊 القاعدة الأولى (الملاك)", type="xlsx", key="file2_name")
    with c3:
        file3 = st.file_uploader("📊 القاعدة الثانية (العقود) - اختياري", type="xlsx", key="file3_name")

    if file1 and file2:
        try:
            names_df = load_excel(file1.getvalue(), file1.name)
            db1 = load_excel(file2.getvalue(), file2.name)
            db2 = load_excel(file3.getvalue(), file3.name) if file3 else None
        except Exception as e:
            st.error(f"❌ تعذّرت قراءة أحد الملفات: {e}")
            st.stop()

        two_db = db2 is not None
        st.markdown("---")
        st.subheader("⚙️ إعدادات المطابقة")

        if two_db:
            l1, l2 = st.columns(2)
            label1 = l1.text_input("اسم القاعدة الأولى:", "الملاك").strip() or "الملاك"
            label2 = l2.text_input("اسم القاعدة الثانية:", "العقود").strip() or "العقود"
            if label1 == label2:
                label2 += " 2"
        else:
            label1, label2 = "الملاك", "العقود"

        cols = st.columns(3 if two_db else 2)
        name_col_file = cols[0].selectbox("🔹 عمود الاسم في ملف الأسماء:",
                                          names_df.columns.tolist(), key="name_col_file_tab1")
        name_col_db = cols[1].selectbox(f"🔹 عمود الاسم في {label1}:",
                                        db1.columns.tolist(), key="name_col_db_tab1")
        name_col_db2 = cols[2].selectbox(f"🔹 عمود الاسم في {label2}:",
                                         db2.columns.tolist(), key="name_col_db2") if two_db else None

        st.markdown("---")
        use_second_column = st.checkbox("➕ إضافة عمود ثانٍ للمطابقة (مثل اسم المدرسة)",
                                        key="use_second_col_tab1")
        second_col_file = second_col_db = second_col_db2 = None
        if use_second_column:
            cols = st.columns(3 if two_db else 2)
            second_col_file = cols[0].selectbox(
                "🔸 العمود الثاني في ملف الأسماء:",
                [c for c in names_df.columns if c != name_col_file], key="second_col_file_tab1")
            second_col_db = cols[1].selectbox(
                f"🔸 العمود الثاني في {label1}:",
                [c for c in db1.columns if c != name_col_db], key="second_col_db_tab1")
            if two_db:
                second_col_db2 = cols[2].selectbox(
                    f"🔸 العمود الثاني في {label2}:",
                    [c for c in db2.columns if c != name_col_db2], key="second_col_db2")

        st.markdown("---")
        excl1 = {name_col_db, second_col_db}
        available = [c for c in db1.columns if c not in excl1]
        if two_db:
            excl2 = {name_col_db2, second_col_db2}
            available += [c for c in db2.columns if c not in excl2 and c not in available]
        selected_columns = st.multiselect("📌 اختر الأعمدة المراد جلبها من قاعدة البيانات:",
                                          available, key="cols_tab1")
        if two_db and selected_columns:
            only1 = [c for c in selected_columns if c not in db2.columns]
            only2 = [c for c in selected_columns if c not in db1.columns]
            if only1:
                st.caption(f"ℹ️ موجودة في {label1} فقط: {'، '.join(map(str, only1))}")
            if only2:
                st.caption(f"ℹ️ موجودة في {label2} فقط: {'، '.join(map(str, only2))}")

        st.markdown("---")
        dual = bool(use_second_column and second_col_file and second_col_db
                    and (second_col_db2 or not two_db))

        sig = (file1.file_id, file2.file_id, file3.file_id if file3 else None, label1, label2,
               name_col_file, name_col_db, name_col_db2, dual, second_col_file, second_col_db,
               second_col_db2, tuple(selected_columns))

        if st.button("🚀 بدء المطابقة", key="match_names_btn"):
            bar = st.progress(0.0, text="جاري المطابقة...")

            def cb(done, total):
                bar.progress(min(done / max(total, 1), 1.0), text=f"جاري المطابقة... {done:,} / {total:,}")

            try:
                if two_db:
                    res, conf, com = mc.match_two_databases(
                        names_df, db1, db2, name_col_file, name_col_db, name_col_db2,
                        label1, label2,
                        second_col_file if dual else None, second_col_db if dual else None,
                        second_col_db2 if dual else None, selected_columns, cb)
                    bar.progress(1.0, text="جاري تجهيز ملف الإكسل...")
                    xls = mc.to_excel_two_db(res, conf, com, dual).getvalue()
                elif dual:
                    res = mc.match_names_dual(names_df, db1, name_col_file, name_col_db,
                                              second_col_file, second_col_db, selected_columns, cb)
                    conf = com = None
                    xls = mc.to_excel_dual(res).getvalue()
                else:
                    res = mc.match_names(names_df, db1, name_col_file, name_col_db,
                                         selected_columns, cb)
                    conf = com = None
                    xls = mc.to_excel_basic(res).getvalue()
                st.session_state["result"] = dict(res=res, conf=conf, com=com, xls=xls,
                                                  dual=dual, two_db=two_db, sig=sig)
                bar.empty()
            except Exception as e:
                bar.empty()
                st.session_state.pop("result", None)
                st.error(f"❌ حدث خطأ أثناء المطابقة: {e}")

        # النتائج محفوظة في الجلسة: لا تختفي عند الضغط على زر التحميل
        r = st.session_state.get("result")
        if r and r["sig"] != sig:
            st.info("ℹ️ تغيّرت الإعدادات أو الملفات - اضغط «بدء المطابقة» لإعادة المطابقة.")
            r = None
        if r:
            st.success("✅ تمت المطابقة بنجاح!")
            stats = mc.summary(r["res"], r["dual"], r["two_db"])
            for col, (k, v) in zip(st.columns(len(stats)), stats.items()):
                col.metric(k, f"{v:,}")
            legend = "🟢 تطابق الاثنين &nbsp; 🟡 تطابق واحد &nbsp; " if r["dual"] else ""
            legend += "🟠 مكرر داخل القاعدة &nbsp; 🔴 لا تطابق"
            if r["two_db"]:
                legend += " &nbsp; 🟣 موجود في القاعدتين"
            st.markdown("💡 **ألوان الإكسل:** " + legend)

            if r["two_db"]:
                t1, t2, t3 = st.tabs([
                    "📋 النتائج",
                    f"🟣 تعارض بين القاعدتين ({len(r['conf']):,})",
                    f"👥 أسماء مشتركة بين القاعدتين ({len(r['com']):,})"])
                show_table(t1, r["res"])
                show_table(t2, r["conf"])
                show_table(t3, r["com"])
            else:
                show_table(st, r["res"])

            st.download_button("⬇️ تحميل النتائج", r["xls"],
                               file_name="نتائج_المطابقة.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif password:
    st.error("❌ كلمة المرور غير صحيحة.")
