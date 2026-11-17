import base64
import datetime as dt
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from processing import (
    COL_BUCKET,
    COL_STATUS,
    ST_DONE,
    ST_NOT_DONE,
    ST_PARTIAL,
    apply_rules,
    compute_new_values,
    default_old_columns,
    dpd_columns,
    guess_update_cols,
    new_col_names,
    overdue_columns,
    read_all_sheets,
    write_workbook,
)

LOGO_PATH = Path(__file__).parent / "assets" / "logo.png"
PURPLE, GREEN = "#4F008C", "#00C48C"

st.set_page_config(page_title="تحديث شيت Ejadah", page_icon=str(LOGO_PATH), layout="wide")

_logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
st.markdown(
    f"""
<style>
html, body, [data-testid='stAppViewContainer'] {{direction: rtl; text-align: right;}}
[data-testid="stHeader"] {{background: transparent;}}
.block-container {{padding-top: 1.5rem; max-width: 1200px;}}
.hero {{
  display: flex; align-items: center; gap: 20px; direction: rtl;
  background: linear-gradient(120deg, {PURPLE} 0%, #6A12AE 100%);
  border-bottom: 6px solid {GREEN}; border-radius: 18px;
  padding: 20px 28px; margin-bottom: 1.5rem; color: #fff;
  box-shadow: 0 6px 20px rgba(79,0,140,.25);
}}
.hero img {{width: 76px; height: 76px; border-radius: 16px; box-shadow: 0 2px 8px rgba(0,0,0,.3);}}
.hero h1 {{margin: 0; font-size: 1.9rem; color: #fff; padding: 0;}}
.hero p {{margin: 4px 0 0; color: #D9F8EE; font-size: 1rem;}}
/* بطاقات الأرقام */
[data-testid="stMetric"] {{
  background: #F4EEFA; border-right: 5px solid {GREEN}; border-radius: 12px; padding: 12px 16px;
}}
[data-testid="stMetricValue"] {{color: {PURPLE}; font-weight: 700;}}
/* التابات */
.stTabs [data-baseweb="tab-list"] {{gap: 8px;}}
.stTabs [data-baseweb="tab"] {{background: #F4EEFA; border-radius: 10px 10px 0 0; padding: 8px 20px; color: {PURPLE};}}
.stTabs [aria-selected="true"] {{background: {PURPLE}; color: #fff !important;}}
.stTabs [data-baseweb="tab-highlight"] {{background-color: {GREEN}; height: 4px;}}
/* الأزرار */
.stButton > button[kind="primary"], .stDownloadButton > button {{
  background: {GREEN}; color: #fff; border: none; border-radius: 12px; font-weight: 700; padding: .6rem 1.6rem;
}}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {{background: #00A878; color: #fff;}}
/* رفع الملفات */
[data-testid="stFileUploader"] section {{border: 2px dashed {PURPLE}; border-radius: 14px; background: #FAF6FD;}}
[data-testid="stFileUploader"] button {{border-color: {PURPLE}; color: {PURPLE};}}
h3, h4, strong {{color: {PURPLE};}}
hr {{border-color: {GREEN} !important; opacity: .4;}}
</style>
<div class="hero">
  <img src="data:image/png;base64,{_logo_b64}" alt="STC Bank">
  <div>
    <h1>تحديث شيت Ejadah اليومي</h1>
    <p>ارفع الملفين، راجع النتيجة، ونزّل الشيت المحدّث</p>
  </div>
</div>
""",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner="بقرا الملف...")
def _read(file_bytes: bytes):
    return read_all_sheets(file_bytes)


# ---------------------------------------------------------------- 1) الملفات
c1, c2 = st.columns(2)
master_file = c1.file_uploader("1) الملف الأساسي (آخر شيت متحدث)", type=["xlsx"], key="master")
update_file = c2.file_uploader("2) ملف التحديث الجديد", type=["xlsx"], key="update")
if not (master_file and update_file):
    st.info("ارفع الملفين عشان نبدأ.")
    st.stop()

master_bytes, update_bytes = master_file.getvalue(), update_file.getvalue()
M, U = _read(master_bytes), _read(update_bytes)
sheets = [s for s in M if s in U]
if not sheets:
    st.error("مفيش شيتات بنفس الاسم في الملفين (المفروض PF و Credit Cards).")
    st.stop()
skipped = [s for s in M if s not in U]
if skipped:
    st.warning(f"الشيتات دي مش موجودة في ملف التحديث وهتتساب زي ما هي: {', '.join(skipped)}")

# ---------------------------------------------------------------- 2) اليوم
today = dt.datetime.now(ZoneInfo("Africa/Cairo")).date()
run_date = st.date_input("تاريخ اليوم (بيتسمى بيه الأعمدة وبيتكتب في «تاريخ تحديث التقرير الشامل»)", value=today)
names = new_col_names(run_date)
st.caption(f"الأعمدة الجديدة: `{names.dpd}` · `{names.overdue}` · `{names.diff_dpd}` · `{names.diff_overdue}`")

# ---------------------------------------------------------------- 3) إعدادات كل شيت
per_sheet, summaries = {}, []
tabs = st.tabs(sheets)
for tab, sheet in zip(tabs, sheets):
    with tab:
        m, u = M[sheet], U[sheet]
        key_col, dpd_c, od_c, upd_cols = guess_update_cols(u)

        st.markdown("**أعمدة ملف التحديث**")
        a, b, c = st.columns(3)
        key_col = a.selectbox("عمود الحساب", upd_cols, index=upd_cols.index(key_col), key=f"k_{sheet}")
        dpd_c = b.selectbox("عمود DPD", upd_cols, index=upd_cols.index(dpd_c) if dpd_c in upd_cols else 0, key=f"d_{sheet}")
        od_c = c.selectbox("عمود Overdue", upd_cols, index=upd_cols.index(od_c) if od_c in upd_cols else 0, key=f"o_{sheet}")

        st.markdown("**الأعمدة القديمة اللي هنقارن بيها (من الملف الأساسي)**")
        dpd_opts, od_opts = dpd_columns(m), overdue_columns(m)
        old_dpd_def, old_od_def = default_old_columns(m, names)
        a, b = st.columns(2)
        old_dpd = a.selectbox("DPD القديم", dpd_opts, index=dpd_opts.index(old_dpd_def) if old_dpd_def in dpd_opts else 0, key=f"od_{sheet}")
        old_od = b.selectbox("Overdue القديم", od_opts, index=od_opts.index(old_od_def) if old_od_def in od_opts else 0, key=f"oo_{sheet}")

        vals, info = compute_new_values(m, u, key_col, dpd_c, od_c, old_dpd, old_od)

        st.markdown("**قيم «فرق DPD» اللي تتحول لـ «معالج»**")
        uniq = sorted(vals["diff_dpd"].dropna().unique())
        counts = vals["diff_dpd"].value_counts()
        fmt = lambda v: f"{v:g}  ({int(counts.get(v, 0))} حساب)"
        selected = st.multiselect(
            "اختار القيم (الافتراضي: صفر وأي قيمة سالبة)",
            options=uniq,
            default=[v for v in uniq if v <= 0],
            format_func=fmt,
            key=f"sel_{sheet}_{old_dpd}_{key_col}_{dpd_c}",
        )

        res = apply_rules(m, vals, selected, run_date)
        per_sheet[sheet] = dict(master=m, vals=vals, res=res)

        n_missing = int((~vals["in_update"]).sum())
        before = m[COL_STATUS].fillna(ST_NOT_DONE).value_counts()
        after = res[COL_STATUS].value_counts()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("حسابات في الماستر", len(m))
        k2.metric("مش موجودة في التحديث", n_missing)
        k3.metric("معالج (بعد)", int(after.get(ST_DONE, 0)), int(after.get(ST_DONE, 0) - before.get(ST_DONE, 0)))
        k4.metric("مسدد أقل من القسط (بعد)", int(after.get(ST_PARTIAL, 0)), int(after.get(ST_PARTIAL, 0) - before.get(ST_PARTIAL, 0)))
        if info["dup_in_update"]:
            st.warning(f"في {info['dup_in_update']} حساب مكرر في ملف التحديث، أخدت آخر ظهور ليهم.")
        if info["extra_in_update"]:
            st.warning(f"في {info['extra_in_update']} حساب في ملف التحديث مش موجود في الماستر (اتجاهلوا).")

        left, right = st.columns(2)
        left.dataframe(after.rename("العدد"), use_container_width=True)
        right.dataframe(res[COL_BUCKET].value_counts().rename("العدد"), use_container_width=True)

        with st.expander("معاينة أول 50 صف"):
            prev = pd.concat([m.iloc[:, :1], vals.rename(columns={"new_dpd": names.dpd, "new_od": names.overdue, "diff_dpd": names.diff_dpd, "diff_od": names.diff_overdue}).drop(columns="in_update"), res], axis=1).head(50)
            st.dataframe(prev, use_container_width=True)

# ---------------------------------------------------------------- 4) التحميل
st.divider()
if st.button("⚙️ جهّز الملف المحدّث", type="primary"):
    with st.spinner("بكتب الملف..."):
        st.session_state["out"] = write_workbook(master_bytes, per_sheet, names)
        st.session_state["out_name"] = f"Update_Ejadah_Assigment_{run_date.day}_{names.dpd[-3:].title()}_{run_date.year}.xlsx"
if "out" in st.session_state:
    st.download_button(
        "⬇️ تحميل الملف المحدّث",
        data=st.session_state["out"],
        file_name=st.session_state["out_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
