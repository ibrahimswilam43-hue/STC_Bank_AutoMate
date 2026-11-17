"""منطق تحديث شيت Ejadah — بدون أي اعتماد على Streamlit عشان نقدر نختبره لوحده."""
from __future__ import annotations

import io
import re
from copy import copy
from dataclasses import dataclass

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

# ---------- أسماء الأعمدة الثابتة ----------
COL_NEW_OVERDUE = "x_New_overdue"
COL_NEW_DPD = "x_New_DPD"
COL_STATUS = "حالة المعالجة - التمويل"
COL_BUCKET = "تحديث فئة العميل - التمويل"
COL_UPDATE_DATE = "تاريخ تحديث التقرير الشامل - التمويل"

ST_DONE = "معالج"
ST_NOT_DONE = "غير معالج"
ST_PARTIAL = "مسدد اقل من قيمة القسط"

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
MONTHS_TITLE = [m.title() for m in MONTHS]

DPD_COL_RE = re.compile(r"^DPD\d{1,2}[A-Za-z]{3}$")


# ---------- أسماء الأعمدة الجديدة حسب اليوم ----------
@dataclass
class NewColNames:
    dpd: str
    overdue: str
    diff_dpd: str
    diff_overdue: str


def new_col_names(d) -> NewColNames:
    day, mon, mon_t = d.day, MONTHS[d.month - 1], MONTHS_TITLE[d.month - 1]
    # نفس تسميات الشيت الحالي: DPD16NOV / OVERDUE16NOV / فرق dpd 16Nov / فرق overdue 16 Nov
    return NewColNames(
        dpd=f"DPD{day}{mon}",
        overdue=f"OVERDUE{day}{mon}",
        diff_dpd=f"فرق dpd {day}{mon_t}",
        diff_overdue=f"فرق overdue {day} {mon_t}",
    )


# ---------- أدوات ----------
def norm_key(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def bucket_from_dpd(v) -> str:
    """نفس معادلة الإكسل بالظبط."""
    if pd.isna(v) or v <= 0:
        return "Current"
    if v <= 29:
        return "1-29 DPD"
    if v <= 59:
        return "30-59 DPD"
    if v <= 89:
        return "60-89 DPD"
    if v <= 119:
        return "90-119 DPD"
    if v <= 149:
        return "120-149 DPD"
    if v <= 179:
        return "150-179 DPD"
    return "أكبر من 179"


def read_sheet(ws) -> pd.DataFrame:
    """يقرأ الشيت كـ DataFrame ويحتفظ برقم الصف الأصلي في العمود _row."""
    rows = list(ws.iter_rows(values_only=True))
    headers = [h if h is not None else f"_col{i}" for i, h in enumerate(rows[0])]
    df = pd.DataFrame(rows[1:], columns=headers)
    df["_row"] = np.arange(2, 2 + len(df))
    first = headers[0]
    df = df[df[first].notna()].reset_index(drop=True)
    return df


def load_wb(file_bytes: bytes):
    return openpyxl.load_workbook(io.BytesIO(file_bytes))


def read_all_sheets(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    wb = load_wb(file_bytes)
    return {ws.title: read_sheet(ws) for ws in wb.worksheets}


def guess_update_cols(upd: pd.DataFrame):
    cols = [c for c in upd.columns if c != "_row" and not str(c).startswith("_col")]
    key = cols[0]
    dpd = next((c for c in cols if str(c).strip().lower() == "dpd"), None) or next(
        (c for c in cols if "dpd" in str(c).lower()), None
    )
    od = next((c for c in cols if "overdue" in str(c).lower()), None)
    return key, dpd, od, cols


def dpd_columns(master: pd.DataFrame) -> list[str]:
    return [c for c in master.columns if "dpd" in str(c).lower() and not str(c).startswith("فرق") and c != COL_NEW_DPD]


def overdue_columns(master: pd.DataFrame) -> list[str]:
    return [c for c in master.columns if "overdue" in str(c).lower() and not str(c).startswith("فرق") and c != COL_NEW_OVERDUE]


def default_old_columns(master: pd.DataFrame, names: NewColNames):
    """آخر عمود DPDxxNOV قبل اليوم الجديد + الـ OVERDUE المقابل له."""
    cands = [c for c in master.columns if DPD_COL_RE.match(str(c)) and c != names.dpd]
    old_dpd = cands[-1] if cands else None
    # لو العمود الجديد موجود بالفعل (إعادة تشغيل) نختار اللي قبله في الترتيب
    if names.dpd in master.columns:
        idx = list(master.columns).index(names.dpd)
        before = [c for c in cands if list(master.columns).index(c) < idx]
        old_dpd = before[-1] if before else old_dpd
    old_od = None
    if old_dpd:
        guess = str(old_dpd).replace("DPD", "OVERDUE", 1)
        old_od = guess if guess in master.columns else None
    return old_dpd, old_od


# ---------- الحساب ----------
def compute_new_values(master, upd, upd_key, upd_dpd, upd_od, old_dpd, old_od):
    """يرجّع DataFrame بنفس ترتيب الماستر: new_dpd / new_od / diff_dpd / diff_od (NaN لو الحساب مش في التحديث)."""
    u = upd.copy()
    u["_k"] = u[upd_key].map(norm_key)
    u = u[u["_k"].notna()]
    dup = int(u["_k"].duplicated().sum())
    u = u.drop_duplicates("_k", keep="last").set_index("_k")
    u_dpd = pd.to_numeric(u[upd_dpd], errors="coerce")
    u_od = pd.to_numeric(u[upd_od], errors="coerce")

    mkey = master[master.columns[0]].map(norm_key)
    out = pd.DataFrame(index=master.index)
    out["new_dpd"] = mkey.map(u_dpd)
    out["new_od"] = mkey.map(u_od)
    out["in_update"] = mkey.isin(u.index)
    out["diff_dpd"] = out["new_dpd"] - pd.to_numeric(master[old_dpd], errors="coerce")
    out["diff_od"] = (out["new_od"] - pd.to_numeric(master[old_od], errors="coerce")).round(2)
    extra_in_update = int((~pd.Index(u.index).isin(set(mkey.dropna()))).sum())
    return out, {"dup_in_update": dup, "extra_in_update": extra_in_update}


def apply_rules(master, vals: pd.DataFrame, treated_dpd_diffs, run_date):
    """يرجّع الأعمدة الخمسة المحدّثة (x_New_overdue .. التاريخ)."""
    old_status = master[COL_STATUS].fillna(ST_NOT_DONE).astype(str).str.strip()
    status = old_status.copy()

    already_done = old_status == ST_DONE
    missing = ~vals["in_update"]
    dpd_hit = vals["diff_dpd"].isin(list(treated_dpd_diffs))
    od_changed = vals["diff_od"].notna() & (vals["diff_od"].abs() > 0.005)

    status[~already_done & (missing | dpd_hit)] = ST_DONE
    status[~already_done & ~(missing | dpd_hit) & od_changed] = ST_PARTIAL

    new_dpd0 = vals["new_dpd"].fillna(0)
    res = pd.DataFrame(index=master.index)
    res[COL_NEW_OVERDUE] = vals["new_od"].fillna(0)
    res[COL_NEW_DPD] = new_dpd0
    res[COL_STATUS] = status
    res[COL_BUCKET] = new_dpd0.map(bucket_from_dpd)
    res[COL_UPDATE_DATE] = pd.Timestamp(run_date)
    return res


# ---------- الكتابة في الإكسل ----------
def _clean(v):
    if v is None:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        f = float(v)
        return int(f) if f.is_integer() else f
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    return v


def write_workbook(master_bytes: bytes, per_sheet: dict, names: NewColNames) -> bytes:
    """per_sheet[sheet] = dict(master=df, vals=df, res=df)"""
    wb = load_wb(master_bytes)
    for sheet, pack in per_sheet.items():
        ws = wb[sheet]
        master, vals, res = pack["master"], pack["vals"], pack["res"]
        header = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value is not None}
        rows = master["_row"].tolist()

        # الأعمدة الجديدة: لو موجودة نحدّثها مكانها، لو لأ نضيفها في الآخر
        last_col = max(header.values())
        template_col = last_col  # ناخد التنسيق من آخر عمود
        new_cols = [
            (names.dpd, vals["new_dpd"]),
            (names.overdue, vals["new_od"]),
            (names.diff_dpd, vals["diff_dpd"]),
            (names.diff_overdue, vals["diff_od"]),
        ]
        for name, series in new_cols:
            if name not in header:
                last_col += 1
                header[name] = last_col
                ws.cell(1, last_col, name)
                src_h = ws.cell(1, template_col)
                dst_h = ws.cell(1, last_col)
                dst_h._style = copy(src_h._style)
                ws.column_dimensions[get_column_letter(last_col)].width = 14
            series_values = series.tolist()
            col = header[name]
            for r, v in zip(rows, series_values):
                v = _clean(v)
                # نفس عُرف الشيت: الحساب اللي مش في التحديث بيبقى #N/A
                ws.cell(r, col, "#N/A" if v is None else v)

        # الأعمدة الخمسة المحدّثة
        for name in [COL_NEW_OVERDUE, COL_NEW_DPD, COL_STATUS, COL_BUCKET, COL_UPDATE_DATE]:
            if name not in header:
                raise ValueError(f"العمود '{name}' مش موجود في شيت {sheet}")
            col = header[name]
            for r, v in zip(rows, res[name].tolist()):
                cell = ws.cell(r, col, _clean(v))
                if name == COL_UPDATE_DATE:
                    cell.number_format = "yyyy-mm-dd"

        # نسخ تنسيق الصفوف للأعمدة الجديدة + مد الفلتر
        for name, _ in new_cols:
            col = header[name]
            if col > template_col:
                for r in rows:
                    ws.cell(r, col)._style = copy(ws.cell(r, template_col)._style)
        if ws.auto_filter.ref:
            ws.auto_filter.ref = f"A1:{get_column_letter(max(header.values()))}{ws.max_row}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
