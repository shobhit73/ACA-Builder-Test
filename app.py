# app.py
# ============================
# ACA-1095 Builder — Streamlit App (login + defaults for L15 & Part III)
# ============================

import io
import os
import zipfile
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject, DictionaryObject
from reportlab.pdfgen import canvas

# ----------------------------
# Page config (must be first Streamlit call)
# ----------------------------
st.set_page_config(page_title="ACA-1095 Builder", layout="wide")

# =========================
# Simple Login (no external libs)
# =========================
USERS = {
    "admin": "admin123",
    "hr": "hrpass456",
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None

def login_screen():
    st.title("🔐 ACA-1095 Builder - Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login", type="primary", use_container_width=True):
        if u in USERS and USERS[u] == p:
            st.session_state.logged_in = True
            st.session_state.username = u
            st.success(f"Welcome {u}!")
            st.rerun()
        else:
            st.error("Invalid username or password")

def logout_button():
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.rerun()

# Gate the app
if not st.session_state.logged_in:
    login_screen()
    st.stop()

# After login
st.sidebar.success(f"Logged in as {st.session_state.username}")
logout_button()
st.title("ACA-1095 Builder")

# =========================
# Defaults & Controls (Sidebar)
# =========================
with st.sidebar.expander("Defaults & Inference"):
    # Line 15 defaults
    default_l15_amount = st.number_input(
        "Default Line 15 amount (if required but missing)",
        min_value=0.00, step=1.00, value=0.00, format="%.2f",
        help="Used when Line 14 requires Line 15 but no pay deduction amount is found."
    )
    collapse_l15_all12 = st.checkbox(
        "Collapse Line 15 to 'All 12' if uniform",
        value=True
    )

    # Part III defaults
    infer_part3_from_l16 = st.checkbox(
        "Infer Part III coverage from Line 16 = 2C when enrollment is missing",
        value=True
    )
    part3_include_emp_when_empty = st.checkbox(
        "If still nothing, include EMPLOYEE in Part III using default",
        value=False
    )
    part3_default_emp_all12 = st.checkbox(
        "Default Employee coverage = All 12 (when forced include above is ON)",
        value=True,
        help="If not checked, no months will be marked when forced include happens."
    )

# =========================
# Utility helpers
# =========================
TRUTHY = {"y","yes","true","t","1",1,True}
FALSY  = {"n","no","false","f","0",0,False,None,np.nan}
ACTIVE_STATUS = {"FT","FULL-TIME","FULL TIME","PT","PART-TIME","PART TIME","ACTIVE"}

EXPECTED_SHEETS = {
    "emp demographic": ["employeeid","firstname","lastname","ssn","addressline1","addressline2","city","state","zipcode"],
    "emp status": ["employeeid","employmentstatus","role","statusstartdate","statusenddate"],
    "emp eligibility": ["employeeid","iseligibleforcoverage","minimumvaluecoverage","eligibilitystartdate","eligibilityenddate"],
    "emp enrollment": ["employeeid","isenrolled","enrollmentstartdate","enrollmentenddate"],
    "dep enrollment": ["employeeid","dependentrelationship","eligible","enrolled","eligiblestartdate","eligibleenddate","firstname","lastname","mi","ssn","dob"],
    "pay deductions": ["employeeid","amount","startdate","enddate"]
}
CANON_ALIASES = {
    "mimimumvaluecoverage": "minimumvaluecoverage",
    "minimimvaluecoverage": "minimumvaluecoverage",
    "zip": "zipcode", "zip code": "zipcode",
    "ssn (digits only)": "ssn",
}

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (df.columns.str.strip().str.replace(r"\s+", " ", regex=True).str.lower())
    return df

def _coerce_str(x) -> str:
    if pd.isna(x): return ""
    return str(x).strip()

def to_bool(val) -> bool:
    if isinstance(val, str):
        v = val.strip().lower()
        if v in TRUTHY: return True
        if v in FALSY:  return False
    return bool(val) and val not in FALSY

def _last_day_of_month(y: int, m: int) -> date:
    return date(y,12,31) if m==12 else (date(y, m+1, 1) - timedelta(days=1))

def parse_date_safe(d, default_end: bool=False):
    if pd.isna(d): return None
    if isinstance(d, (datetime, np.datetime64)):
        dt = pd.to_datetime(d, errors="coerce");  return None if pd.isna(dt) else dt.date()
    s = str(d).strip()
    if not s: return None
    try:
        if len(s)==4 and s.isdigit():
            y = int(s); return date(y,12,31) if default_end else date(y,1,1)
        if len(s)==7 and s[4]=="-":
            y,m = map(int, s.split("-"));  return _last_day_of_month(y,m) if default_end else date(y,m,1)
    except: pass
    dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.isna(dt):
        try:
            y,m = map(int, s.split("-")[:2])
            return _last_day_of_month(y,m) if default_end else date(y,m,1)
        except: return None
    return dt.date()

def month_bounds(year:int, month:int):
    return date(year, month, 1), _last_day_of_month(year, month)

def _any_overlap(df, start_col, end_col, m_start, m_end, mask=None) -> bool:
    if df.empty: return False
    _m = mask if mask is not None else pd.Series(True, index=df.index)
    s = df.loc[_m, start_col].fillna(pd.Timestamp.min).dt.date
    e = df.loc[_m, end_col].fillna(pd.Timestamp.max).dt.date
    return bool(((e >= m_start) & (s <= m_end)).any())

def _all_month(df, start_col, end_col, m_start, m_end, mask=None) -> bool:
    if df.empty: return False
    _m = mask if mask is not None else pd.Series(True, index=df.index)
    s = df.loc[_m, start_col].fillna(pd.Timestamp.min).dt.date
    e = df.loc[_m, end_col].fillna(pd.Timestamp.max).dt.date
    return bool(((s <= m_start) & (e >= m_end)).any())

# =========================
# Excel ingestion & transforms
# =========================
def load_excel(file_bytes: bytes) -> dict:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    out = {}
    for raw in xls.sheet_names:
        df = pd.read_excel(xls, raw)
        df = normalize_columns(df)
        df = df.rename(columns={k:v for k,v in CANON_ALIASES.items() if k in df.columns})
        out[raw.strip().lower()] = df
    return out

def _pick_sheet(data: dict, key: str) -> pd.DataFrame:
    if key in data: return data[key]
    for k in data:
        if key in k: return data[k]
    return pd.DataFrame()

def _ensure_employeeid_str(df):
    if df.empty or "employeeid" not in df.columns: return df
    df = df.copy(); df["employeeid"] = df["employeeid"].map(_coerce_str); return df

def _parse_date_cols(df, cols, default_end_cols=()):
    if df.empty: return df
    df = df.copy(); endset = set(default_end_cols)
    for c in cols:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: parse_date_safe(x, default_end=c in endset))
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def _boolify(df, cols):
    if df.empty: return df
    df = df.copy()
    for c in cols:
        if c in df.columns: df[c] = df[c].apply(to_bool)
    return df

def prepare_inputs(data: dict):
    cleaned = {}
    for sheet, cols in EXPECTED_SHEETS.items():
        df = _pick_sheet(data, sheet)
        if df.empty:
            cleaned[sheet] = pd.DataFrame(columns=cols); continue
        for misspell, canon in CANON_ALIASES.items():
            if misspell in df.columns and canon not in df.columns:
                df = df.rename(columns={misspell: canon})
        df = _ensure_employeeid_str(df)
        if sheet == "emp status":
            if "employmentstatus" in df.columns:
                df["employmentstatus"] = df["employmentstatus"].astype(str).str.strip().str.upper()
            if "role" in df.columns:
                df["role"] = df["role"].astype(str).str.strip().str.upper()
            df = _parse_date_cols(df, ["statusstartdate","statusenddate"], default_end_cols=["statusenddate"])
        elif sheet == "emp eligibility":
            df = _boolify(df, ["iseligibleforcoverage","minimumvaluecoverage"])
            df = _parse_date_cols(df, ["eligibilitystartdate","eligibilityenddate"], default_end_cols=["eligibilityenddate"])
        elif sheet == "emp enrollment":
            df = _boolify(df, ["isenrolled"])
            df = _parse_date_cols(df, ["enrollmentstartdate","enrollmentenddate"], default_end_cols=["enrollmentenddate"])
        elif sheet == "dep enrollment":
            if "dependentrelationship" in df.columns:
                df["dependentrelationship"] = df["dependentrelationship"].astype(str).str.strip().str.title()
            df = _boolify(df, ["eligible","enrolled"])
            df = _parse_date_cols(df, ["eligiblestartdate","eligibleenddate","dob"], default_end_cols=["eligibleenddate"])
        elif sheet == "pay deductions":
            df = _parse_date_cols(df, ["startdate","enddate"], default_end_cols=["enddate"])
            if "amount" in df.columns:
                df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
        cleaned[sheet] = df
    return (cleaned["emp demographic"], cleaned["emp status"], cleaned["emp eligibility"],
            cleaned["emp enrollment"], cleaned["dep enrollment"], cleaned["pay deductions"])

def choose_report_year(emp_elig: pd.DataFrame, fallback_to_current=True) -> int:
    if emp_elig.empty or not {"eligibilitystartdate","eligibilityenddate"} <= set(emp_elig.columns):
        return datetime.now().year if fallback_to_current else 2024
    counts={}
    for _,r in emp_elig.iterrows():
        s = pd.to_datetime(r.get("eligibilitystartdate"), errors="coerce")
        e = pd.to_datetime(r.get("eligibilityenddate"), errors="coerce")
        if pd.isna(s) and pd.isna(e): continue
        s = s or pd.Timestamp.min; e = e or pd.Timestamp.max
        for y in range(s.year, e.year+1): counts[y]=counts.get(y,0)+1
    return max(sorted(counts), key=lambda y:(counts[y], y)) if counts else (datetime.now().year if fallback_to_current else 2024)

def _collect_employee_ids(*dfs):
    ids=set()
    for df in dfs:
        if df is None or df.empty: continue
        if "employeeid" in df.columns:
            ids.update(map(_coerce_str, df["employeeid"].dropna().tolist()))
    return sorted(ids)

def _grid_for_year(employee_ids, year:int) -> pd.DataFrame:
    recs=[]
    for emp in employee_ids:
        for m in range(1,13):
            ms,me = month_bounds(year,m)
            recs.append({"employeeid":emp,"year":year,"monthnum":m,"month":ms.strftime("%b"),
                         "monthstart":ms,"monthend":me})
    g = pd.DataFrame.from_records(recs)
    g["monthstart"]=pd.to_datetime(g["monthstart"]); g["monthend"]=pd.to_datetime(g["monthend"])
    return g

def build_interim(emp_demo, emp_status, emp_elig, emp_enroll, dep_enroll, year=None) -> pd.DataFrame:
    if year is None: year = choose_report_year(emp_elig)
    employee_ids = _collect_employee_ids(emp_demo, emp_status, emp_elig, emp_enroll, dep_enroll)
    grid = _grid_for_year(employee_ids, year)
    demo = emp_demo[["employeeid","firstname","lastname"]].drop_duplicates("employeeid") if not emp_demo.empty else pd.DataFrame(columns=["employeeid","firstname","lastname"])
    out = grid.merge(demo, on="employeeid", how="left")
    stt, elg, enr, dep = emp_status.copy(), emp_elig.copy(), emp_enroll.copy(), dep_enroll.copy()
    for df in (stt,elg,enr,dep):
        if not df.empty:
            for c in df.columns:
                if c.endswith("date") and not np.issubdtype(df[c].dtype, np.datetime64):
                    df[c] = pd.to_datetime(df[c], errors="coerce")
    flags=[]
    for _,row in out.iterrows():
        emp = row["employeeid"]; ms=row["monthstart"].date(); me=row["monthend"].date()
        st_emp = stt[stt["employeeid"]==emp] if not stt.empty else stt
        el_emp = elg[elg["employeeid"]==emp] if not elg.empty else elg
        en_emp = enr[enr["employeeid"]==emp] if not enr.empty else enr
        de_emp = dep[dep["employeeid"]==emp] if not dep.empty else dep

        employed=False
        if not st_emp.empty and {"employmentstatus","statusstartdate","statusenddate"} <= set(st_emp.columns):
            employed = _any_overlap(st_emp, "statusstartdate","statusenddate", ms,me, mask=st_emp["employmentstatus"].isin(ACTIVE_STATUS))
        ft=False
        if not st_emp.empty and {"role","statusstartdate","statusenddate"} <= set(st_emp.columns):
            ft = _any_overlap(st_emp, "statusstartdate","statusenddate", ms,me, mask=st_emp["role"].eq("FT"))
        eligible_any=False; eligible_allmonth=False; eligible_mv_any=False
        if not el_emp.empty and {"eligibilitystartdate","eligibilityenddate"} <= set(el_emp.columns):
            eligible_any = _any_overlap(el_emp, "eligibilitystartdate","eligibilityenddate", ms,me)
            eligible_allmonth = _all_month(el_emp, "eligibilitystartdate","eligibilityenddate", ms,me)
            if "minimumvaluecoverage" in el_emp.columns:
                eligible_mv_any = _any_overlap(el_emp, "eligibilitystartdate","eligibilityenddate", ms,me, mask=el_emp["minimumvaluecoverage"].fillna(False))
        enrolled_allmonth=False
        if not en_emp.empty and {"enrollmentstartdate","enrollmentenddate"} <= set(en_emp.columns):
            en_mask = en_emp["isenrolled"].fillna(False) if "isenrolled" in en_emp.columns else pd.Series(True,index=en_emp.index)
            enrolled_allmonth = _all_month(en_emp, "enrollmentstartdate","enrollmentenddate", ms,me, mask=en_mask)
        offer_spouse=False; offer_dependents=False
        if not de_emp.empty and {"dependentrelationship","eligiblestartdate","eligibleenddate"} <= set(de_emp.columns):
            offer_spouse = _any_overlap(de_emp, "eligiblestartdate","eligibleenddate", ms,me, mask=de_emp["dependentrelationship"].eq("Spouse"))
            offer_dependents = _any_overlap(de_emp, "eligiblestartdate","eligibleenddate", ms,me, mask=de_emp["dependentrelationship"].eq("Child"))
        waitingperiod_month = bool(employed and ft and not eligible_any)

        # Line 14
        if eligible_allmonth and eligible_mv_any:
            l14 = "1E" if (offer_spouse and offer_dependents) else ("1C" if (offer_dependents and not offer_spouse) else ("1D" if (offer_spouse and not offer_dependents) else "1B"))
        elif eligible_allmonth and not eligible_mv_any:
            l14 = "1F"
        else:
            l14 = "1H"

        # Line 16
        if enrolled_allmonth: l16 = "2C"
        elif waitingperiod_month: l16 = "2D"
        elif not employed: l16 = "2A"
        elif employed and not ft: l16 = "2B"
        else: l16 = ""

        flags.append({
            "employed": employed, "ft": ft,
            "eligibleforcoverage": eligible_any, "eligible_allmonth": eligible_allmonth, "eligible_mv": eligible_mv_any,
            "offer_ee_allmonth": eligible_allmonth, "enrolled_allmonth": enrolled_allmonth,
            "offer_spouse": offer_spouse, "offer_dependents": offer_dependents, "waitingperiod_month": waitingperiod_month,
            "line14_final": l14, "line16_final": l16,
        })
    interim = pd.concat([out.reset_index(drop=True), pd.DataFrame(flags)], axis=1)
    base_cols = ["employeeid","firstname","lastname","year","monthnum","month","monthstart","monthend"]
    flag_cols = ["employed","ft","eligibleforcoverage","eligible_allmonth","eligible_mv","offer_ee_allmonth",
                 "enrolled_allmonth","offer_spouse","offer_dependents","waitingperiod_month","line14_final","line16_final"]
    keep = [c for c in base_cols if c in interim.columns] + [c for c in flag_cols if c in interim.columns]
    interim = interim[keep].sort_values(["employeeid","year","monthnum"]).reset_index(drop=True)
    return interim

def build_final(interim: pd.DataFrame) -> pd.DataFrame:
    df = interim.copy()
    out = df.loc[:, ["employeeid","month","line14_final","line16_final"]].rename(columns={
        "employeeid":"EmployeeID","month":"Month","line14_final":"Line14_Final","line16_final":"Line16_Final"
    })
    if "monthnum" in df.columns:
        out = out.join(df["monthnum"]).sort_values(["EmployeeID","monthnum"]).drop(columns=["monthnum"])
    else:
        order = {m:i for i,m in enumerate(MONTHS, start=1)}
        out["_ord"]=out["Month"].map(order); out=out.sort_values(["EmployeeID","_ord"]).drop(columns=["_ord"])
    return out.reset_index(drop=True)

# =========================
# PDF helpers (Part I + II + III)
# =========================
def normalize_ssn_digits(ssn: str) -> str:
    d = "".join(ch for ch in str(ssn) if str(ch).isdigit())
    return f"{d[0:3]}-{d[3:5]}-{d[5:9]}" if len(d)>=9 else d

# IRS 1095-C 2024 field names (page 1):
F_PART1 = ["f1_1[0]","f1_2[0]","f1_3[0]","f1_4[0]","f1_5[0]","f1_6[0]","f1_7[0]","f1_8[0]"]
F_L14 = ["f1_17[0]","f1_18[0]","f1_19[0]","f1_20[0]","f1_21[0]","f1_22[0]","f1_23[0]",
         "f1_24[0]","f1_25[0]","f1_26[0]","f1_27[0]","f1_28[0]","f1_29[0]"]
F_L16 = ["f1_43[0]","f1_44[0]","f1_45[0]","f1_46[0]","f1_47[0]","f1_48[0]","f1_49[0]",
         "f1_50[0]","f1_51[0]","f1_52[0]","f1_53[0]","f1_54[0]","f1_55[0]"]
# NEW: Line 15 (All 12 + Jan..Dec) — adjust if your PDF uses different ids
F_L15 = [f"f1_{i}[0]" for i in range(30, 43)]  # f1_30 .. f1_42

# Part III (page 3): names + checkboxes
P3_NAME_ROWS = [
    ["f3_61[0]","f3_62[0]","f3_63[0]","f3_64[0]","f3_65[0]"],   # row 18
    ["f3_66[0]","f3_67[0]","f3_68[0]","f3_69[0]","f3_70[0]"],   # row 19
    ["f3_71[0]","f3_72[0]","f3_73[0]","f3_74[0]","f3_75[0]"],   # row 20
    ["f3_76[0]","f3_77[0]","f3_78[0]","f3_79[0]","f3_80[0]"],   # row 21
    ["f3_81[0]","f3_82[0]","f3_83[0]","f3_84[0]","f3_85[0]"],   # row 22
    ["f3_91[0]","f3_92[0]","f3_93[0]","f3_95[0]","f3_96[0]"],   # row 23
    ["f3_97[0]","f3_98[0]","f3_99[0]","f3_100[0]","f3_101[0]"], # row 24
    ["f3_102[0]","f3_103[0]","f3_104[0]","f3_105[0]","f3_106[0]"], # row 25
    ["f3_107[0]","f3_108[0]","f3_109[0]","f3_110[0]","f3_111[0]"], # row 26
    ["f3_113[0]","f3_114[0]","f3_115[0]","f3_116[0]","f3_117[0]"], # row 27
    ["f3_118[0]","f3_119[0]","f3_120[0]","f3_121[0]","f3_122[0]"], # row 28
]
def _seq(a,b): return [f"c3_{i}" for i in range(a,b+1)]
P3_CHECK_ROWS = [
    _seq(16, 28), _seq(29, 41), _seq(42, 54), _seq(55, 67), _seq(68, 80),
    _seq(81, 93), _seq(94, 106), _seq(107,119), _seq(120,132), _seq(133,145), _seq(146,158),
]

def set_need_appearances(writer: PdfWriter):
    root = writer._root_object
    if "/AcroForm" not in root:
        root.update({NameObject("/AcroForm"): DictionaryObject()})
    root["/AcroForm"].update({NameObject("/NeedAppearances"): BooleanObject(True)})

def find_rects(reader: PdfReader, target_names, page_index=0):
    rects = {}
    pg = reader.pages[page_index]
    annots = pg.get("/Annots")
    if not annots: return rects
    try: arr = annots.get_object()
    except Exception: arr = annots
    for a in arr:
        obj = a.get_object()
        if obj.get("/Subtype")!="/Widget": continue
        nm = obj.get("/T")
        ft = obj.get("/FT")
        if ft != "/Tx" or nm not in target_names: continue
        r = obj.get("/Rect")
        if r and len(r)==4:
            rects[nm] = tuple(float(r[i]) for i in range(4))
    return rects

def build_overlay(page_w, page_h, rects_and_values, font="Helvetica", font_size=10.5, inset=2.0):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_w, page_h))
    c.setFont(font, font_size)
    for rect, val in rects_and_values:
        if not val: continue
        x0,y0,x1,y1 = rect
        text_x = x0 + inset
        text_y = y1 - font_size - inset
        if text_y < y0 + inset: text_y = y0 + inset
        c.drawString(text_x, text_y, val)
    c.save()
    packet.seek(0)
    return PdfReader(packet)

def flatten_pdf(reader: PdfReader):
    out = PdfWriter()
    for i, page in enumerate(reader.pages):
        annots = page.get("/Annots")
        if annots:
            try: arr = annots.get_object()
            except Exception: arr = annots
            keep=[]
            for a in arr:
                try:
                    if a.get_object().get("/Subtype") != "/Widget":
                        keep.append(a)
                except Exception:
                    keep.append(a)
            if keep:
                page[NameObject("/Annots")] = keep
            else:
                if "/Annots" in page:
                    del page[NameObject("/Annots")]
        out.add_page(page)
    if "/AcroForm" in out._root_object:
        del out._root_object[NameObject("/AcroForm")]
    return out

# ======== Line 15 + Part III helpers with defaults ========
L15_REQUIRED_CODES = {"1B","1C","1D","1E","1J","1K","1L","1M","1N","1O","1P","1Q","1T","1U"}

def _fmt_money(x):
    if x in (None, "", np.nan): return ""
    try: return f"{float(x):.2f}"
    except: return ""

def _overlaps_month(s, e, y, m):
    ms, me = month_bounds(y, m)
    s = pd.to_datetime(s, errors="coerce"); e = pd.to_datetime(e, errors="coerce")
    if pd.isna(s) and pd.isna(e): return False
    s = s if not pd.isna(s) else pd.Timestamp.min
    e = e if not pd.isna(e) else pd.Timestamp.max
    return (e.date() >= ms) and (s.date() <= me)

def compute_l15_for_employee(emp_id: str, year_used: int, final_emp_df: pd.DataFrame, pay_ded_df: pd.DataFrame,
                             default_amount: float = 0.00, collapse_all12: bool = True) -> dict:
    """
    Returns {"Jan":"12.34", ..., "Dec":"", "ALL": ""}.
    Uses pay_deductions.amount as EE self-only cost for overlapping months.
    If Line 14 needs an amount but none exists, uses provided default_amount.
    """
    l14_map = {r["Month"]: str(r["Line14_Final"]).strip() for _, r in final_emp_df.iterrows()}
    out = {}
    for idx, mname in enumerate(MONTHS, start=1):
        code = l14_map.get(mname, "")
        if code not in L15_REQUIRED_CODES:
            out[mname] = ""
            continue
        amt = None
        if pay_ded_df is not None and not pay_ded_df.empty:
            rows = pay_ded_df[pay_ded_df["employeeid"].astype(str)==str(emp_id)]
            for _, rr in rows.iterrows():
                if _overlaps_month(rr.get("startdate"), rr.get("enddate"), year_used, idx):
                    amt = rr.get("amount")
                    if pd.notna(amt): break
        use_amt = default_amount if (amt is None or pd.isna(amt)) else amt
        out[mname] = _fmt_money(use_amt)

    uniq = {v for v in out.values() if v != ""}
    out["ALL"] = (list(uniq)[0] if len(uniq)==1 else "") if collapse_all12 else ""
    if out["ALL"]:
        for mname in MONTHS: out[mname] = ""
    return out

def build_part3_people(emp_row: pd.Series,
                       year_used: int,
                       final_emp_df: pd.DataFrame,
                       emp_enroll_df: pd.DataFrame,
                       dep_enroll_df: pd.DataFrame,
                       infer_from_l16_2c: bool = True,
                       force_include_emp: bool = False,
                       default_emp_all12: bool = True) -> list:
    """
    Returns [{"first","mi","last","ssn","dob","all12", "months"}...]
    Priority:
      1) Use enrollment overlap (employee & dependents).
      2) If none and infer_from_l16_2c: mark months where Line16 = '2C'.
      3) If still none and force_include_emp: include employee with default_emp_all12.
    """
    people = []

    def months_from_periods(df, mask_col="isenrolled"):
        months_cov = set()
        if df is None or df.empty: return months_cov
        mask = df[mask_col].fillna(True) if mask_col in df.columns else pd.Series(True, index=df.index)
        for _, r in df[mask].iterrows():
            s = r.get("enrollmentstartdate") or r.get("eligiblestartdate")
            e = r.get("enrollmentenddate") or r.get("eligibleenddate")
            s = pd.to_datetime(s, errors="coerce")
            e = pd.to_datetime(e, errors="coerce")
            if pd.isna(s) and pd.isna(e): continue
            s = s if not pd.isna(s) else pd.Timestamp(year_used,1,1)
            e = e if not pd.isna(e) else pd.Timestamp(year_used,12,31)
            for m in range(1,13):
                ms, me = month_bounds(year_used, m)
                if (e.date() >= ms) and (s.date() <= me):
                    months_cov.add(m)
        return months_cov

    # 1) Employee months from enrollment
    emp_months = set()
    if emp_enroll_df is not None and not emp_enroll_df.empty:
        subset = emp_enroll_df[emp_enroll_df["employeeid"].astype(str)==str(emp_row.get("employeeid"))]
        emp_months = months_from_periods(subset)

    # 2) If none, infer from Line 16 = 2C
    if not emp_months and infer_from_l16_2c and final_emp_df is not None and not final_emp_df.empty:
        for idx, mname in enumerate(MONTHS, start=1):
            v = str(final_emp_df[final_emp_df["Month"]==mname]["Line16_Final"].iloc[0]) if not final_emp_df[final_emp_df["Month"]==mname].empty else ""
            if v == "2C":
                emp_months.add(idx)

    # 3) If still none, possibly force include with default
    if not emp_months and force_include_emp:
        emp_months = set(range(1,13)) if default_emp_all12 else set()

    # Add employee if any months or forced include triggered
    if emp_months or force_include_emp:
        people.append({
            "first": _coerce_str(emp_row.get("firstname")),
            "mi":    "",
            "last":  _coerce_str(emp_row.get("lastname")),
            "ssn":   normalize_ssn_digits(_coerce_str(emp_row.get("ssn"))),
            "dob":   "",
            "all12": len(emp_months)==12,
            "months": emp_months
        })

    # Dependents (names must exist)
    if dep_enroll_df is not None and not dep_enroll_df.empty:
        fn_col = next((c for c in dep_enroll_df.columns if c.lower() in {"firstname","first"}), None)
        ln_col = next((c for c in dep_enroll_df.columns if c.lower() in {"lastname","last"}), None)
        mi_col = next((c for c in dep_enroll_df.columns if c.lower() in {"mi","middle","middleinitial"}), None)
        ssn_col= next((c for c in dep_enroll_df.columns if c.lower() in {"ssn","tin"}), None)
        dob_col= next((c for c in dep_enroll_df.columns if c.lower() in {"dob","dateofbirth"}), None)

        deps = dep_enroll_df[dep_enroll_df["employeeid"].astype(str)==str(emp_row.get("employeeid"))]
        for _, r in deps.iterrows():
            # months by any elig/enrolled overlap
            mset = months_from_periods(pd.DataFrame([r]))
            first = _coerce_str(r.get(fn_col)) if fn_col else ""
            last  = _coerce_str(r.get(ln_col)) if ln_col else ""
            if not first and not last:
                continue
            people.append({
                "first": first,
                "mi":    _coerce_str(r.get(mi_col)) if mi_col else "",
                "last":  last,
                "ssn":   normalize_ssn_digits(_coerce_str(r.get(ssn_col))) if ssn_col else "",
                "dob":   _coerce_str(pd.to_datetime(r.get(dob_col), errors="coerce").date()) if dob_col else "",
                "all12": len(mset)==12,
                "months": mset
            })

    return people

def fill_pdf_for_employee(pdf_bytes: bytes,
                          emp_row: pd.Series,
                          final_df_emp: pd.DataFrame,
                          year_used: int,
                          pay_ded_df: pd.DataFrame = None,
                          emp_enroll_df: pd.DataFrame = None,
                          dep_enroll_df: pd.DataFrame = None):
    """Return (editable_filename, editable_bytes, flattened_filename, flattened_bytes)"""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    page0 = reader.pages[0]
    W = float(page0.mediabox.width); H = float(page0.mediabox.height)

    # ---- Part I
    first  = _coerce_str(emp_row.get("firstname"))
    mi     = ""
    last   = _coerce_str(emp_row.get("lastname"))
    ssn    = normalize_ssn_digits(_coerce_str(emp_row.get("ssn")))
    addr1  = _coerce_str(emp_row.get("addressline1"))
    addr2  = _coerce_str(emp_row.get("addressline2"))
    city   = _coerce_str(emp_row.get("city"))
    state  = _coerce_str(emp_row.get("state"))
    zipcode= _coerce_str(emp_row.get("zipcode"))
    street = addr1 if not addr2 else f"{addr1} {addr2}"

    part1_map = {
        "f1_1[0]": first,
        "f1_2[0]": mi,
        "f1_3[0]": last,
        "f1_4[0]": ssn,
        "f1_5[0]": street,
        "f1_6[0]": city,
        "f1_7[0]": state,
        "f1_8[0]": zipcode,
    }

    # ---- Line 14 & 16
    l14_by_m = {row["Month"]: _coerce_str(row["Line14_Final"]) for _,row in final_df_emp.iterrows()}
    l16_by_m = {row["Month"]: _coerce_str(row["Line16_Final"]) for _,row in final_df_emp.iterrows()}

    def all12_value(d):
        vals = [d.get(m, "") for m in MONTHS]
        uniq = {v for v in vals if v}
        return list(uniq)[0] if len(uniq)==1 else ""

    l14_all = all12_value(l14_by_m)
    l16_all = all12_value(l16_by_m)
    l14_values = [l14_all] + [l14_by_m.get(m,"") for m in MONTHS]
    l16_values = [l16_all] + [l16_by_m.get(m,"") for m in MONTHS]

    part2_map = {}
    for name,val in zip(F_L14, l14_values): part2_map[name]=val
    for name,val in zip(F_L16, l16_values): part2_map[name]=val

    # ---- Line 15 with defaults
    try:
        l15_map = compute_l15_for_employee(
            emp_id=_coerce_str(emp_row.get("employeeid")),
            year_used=year_used,
            final_emp_df=final_df_emp,
            pay_ded_df=pay_ded_df if pay_ded_df is not None else pd.DataFrame(),
            default_amount=default_l15_amount,
            collapse_all12=collapse_l15_all12
        )
        l15_values = [l15_map.get("ALL","")] + [l15_map.get(m,"") for m in MONTHS]
        for name, val in zip(F_L15, l15_values): part2_map[name] = val
    except Exception:
        pass  # skip if IDs don't match this PDF

    # ---- Part III with defaults/inference
    part3_map = {}
    try:
        people = build_part3_people(
            emp_row, year_used, final_df_emp,
            emp_enroll_df, dep_enroll_df,
            infer_from_l16_2c=infer_part3_from_l16,
            force_include_emp=part3_include_emp_when_empty,
            default_emp_all12=part3_default_emp_all12
        )
        for idx, person in enumerate(people[:len(P3_NAME_ROWS)]):
            nfields = P3_NAME_ROWS[idx]
            part3_map.update({
                nfields[0]: person["first"],
                nfields[1]: person["mi"],
                nfields[2]: person["last"],
                nfields[3]: person["ssn"],
                nfields[4]: person["dob"],
            })
            checks = P3_CHECK_ROWS[idx]
            cvals = [person["all12"]] + [ (m in person["months"]) for m in range(1,13) ]
            for nm, val in zip(checks, cvals):
                part3_map[nm] = bool(val)
    except Exception:
        pass

    # ---- EDITABLE output
    writer_edit = PdfWriter()
    for i in range(len(reader.pages)):
        writer_edit.add_page(reader.pages[i])

    for i in range(len(writer_edit.pages)):
        try:
            writer_edit.update_page_form_field_values(writer_edit.pages[i], {**part1_map, **part2_map, **part3_map})
        except Exception:
            pass

    root = writer_edit._root_object
    if "/AcroForm" not in root:
        root.update({NameObject("/AcroForm"): DictionaryObject()})
    root["/AcroForm"].update({NameObject("/NeedAppearances"): BooleanObject(True)})

    rects = find_rects(reader, list({**part1_map, **part2_map}.keys()), page_index=0)
    overlay_pairs = [(rects[nm], (part1_map|part2_map)[nm]) for nm in (part1_map|part2_map) if nm in rects and (part1_map|part2_map)[nm]]
    if overlay_pairs:
        overlay_pdf = build_overlay(float(page0.mediabox.width), float(page0.mediabox.height), overlay_pairs)
        writer_edit.pages[0].merge_page(overlay_pdf.pages[0])

    first_last = f"{first}_{last}".strip().replace(" ","_") or (_coerce_str(emp_row.get("employeeid")) or "employee")
    editable_name = f"1095c_filled_fields_{first_last}_{year_used}.pdf"
    editable_bytes = io.BytesIO()
    writer_edit.write(editable_bytes)
    editable_bytes.seek(0)

    # ---- FLATTENED output
    reader_after = PdfReader(io.BytesIO(editable_bytes.getvalue()))
    writer_flat = flatten_pdf(reader_after)
    flattened_name = f"1095c_filled_flattened_{first_last}_{year_used}.pdf"
    flattened_bytes = io.BytesIO()
    writer_flat.write(flattened_bytes)
    flattened_bytes.seek(0)

    return editable_name, editable_bytes, flattened_name, flattened_bytes

def save_excel_outputs(interim: pd.DataFrame, final: pd.DataFrame, year:int) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as xw:
        final.to_excel(xw, index=False, sheet_name=f"Final {year}")
        interim.to_excel(xw, index=False, sheet_name=f"Interim {year}")
    buf.seek(0)
    return buf.getvalue()

# =========================
# UI — Step 1: Upload Excel
# =========================
st.subheader("1) Enter your Excel file")
excel_file = st.file_uploader("Upload ACA input workbook (.xlsx)", type=["xlsx"], key="excel")

interim_df = None
final_df = None
emp_demo_df = None
emp_status_df = None
emp_elig_df = None
emp_enroll_df = None
dep_enroll_df = None
pay_ded_df = None
year_used = None

if excel_file is not None:
    try:
        data = load_excel(excel_file.getvalue())
        emp_demo_df, emp_status_df, emp_elig_df, emp_enroll_df, dep_enroll_df, pay_ded_df = prepare_inputs(data)
        year_used = choose_report_year(emp_elig_df)
        interim_df = build_interim(emp_demo_df, emp_status_df, emp_elig_df, emp_enroll_df, dep_enroll_df, year=year_used)
        final_df = build_final(interim_df)

        st.success(f"Processed successfully for YEAR = {year_used}")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Interim Table (preview)**")
            st.dataframe(interim_df.head(200), use_container_width=True)
        with c2:
            st.markdown("**Final Table (preview)**")
            st.dataframe(final_df.head(200), use_container_width=True)

        excel_bytes = save_excel_outputs(interim_df, final_df, year_used)
        st.download_button(
            label=f"Download Final & Interim Excel ({year_used})",
            data=excel_bytes,
            file_name=f"final_and_interim_{year_used}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Failed to process Excel: {e}")

# =========================
# UI — Step 2: Upload PDF
# =========================
st.subheader("2) Upload blank 1095-C PDF (employee copy, 2024 layout)")
pdf_file = st.file_uploader("Upload 1095-C PDF", type=["pdf"], key="pdf")

# =========================
# UI — Step 3: Generate PDFs
# =========================
st.subheader("3) Generate PDFs")

if excel_file is None:
    st.info("Upload the Excel workbook to continue.")
elif pdf_file is None:
    st.info("Upload the blank 1095-C PDF to enable PDF generation.")
else:
    if emp_demo_df is not None and not emp_demo_df.empty:
        all_ids = sorted(emp_demo_df["employeeid"].astype(str).unique().tolist())
        selected_emp = st.selectbox("Choose EmployeeID to generate a single 1095-C PDF:", all_ids, index=0 if all_ids else None)
        colA, colB = st.columns(2)

        with colA:
            if st.button("Generate Single (Part I + II + L15 + Part III)", type="primary", use_container_width=True):
                try:
                    emp_row = emp_demo_df[emp_demo_df["employeeid"].astype(str)==selected_emp].iloc[0]
                    final_emp = final_df[final_df["EmployeeID"].astype(str)==selected_emp].copy()
                    if final_emp.empty:
                        final_emp = pd.DataFrame({"Month":MONTHS, "Line14_Final":[""]*12, "Line16_Final":[""]*12})
                    else:
                        ord_map = {m:i for i,m in enumerate(MONTHS, start=1)}
                        final_emp["_ord"] = final_emp["Month"].map(ord_map)
                        final_emp = final_emp.sort_values("_ord").drop(columns=["_ord"])

                    editable_name, editable_bytes, flat_name, flat_bytes = fill_pdf_for_employee(
                        pdf_file.getvalue(), emp_row, final_emp, year_used,
                        pay_ded_df=pay_ded_df,
                        emp_enroll_df=emp_enroll_df,
                        dep_enroll_df=dep_enroll_df
                    )
                    st.success("PDFs generated for selected employee.")
                    st.download_button("Download Editable PDF", editable_bytes.getvalue(), file_name=editable_name, mime="application/pdf")
                    st.download_button("Download Flattened PDF", flat_bytes.getvalue(), file_name=flat_name, mime="application/pdf")
                except Exception as e:
                    st.error(f"Failed to generate PDFs: {e}")

        with colB:
            st.write("Bulk generate for multiple employees")
            default_selection = all_ids if len(all_ids) <= 10 else all_ids[:10]
            bulk_ids = st.multiselect("Select employees (leave empty to generate for ALL)", all_ids, default=default_selection, key="bulk_ids")
            if st.button("Bulk Generate ZIP (Editable + Flattened for each)", use_container_width=True):
                try:
                    ids_to_run = bulk_ids if bulk_ids else all_ids
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
                        for emp_id in ids_to_run:
                            emp_row = emp_demo_df[emp_demo_df["employeeid"].astype(str)==emp_id]
                            if emp_row.empty:
                                continue
                            emp_row = emp_row.iloc[0]
                            final_emp = final_df[final_df["EmployeeID"].astype(str)==emp_id].copy()
                            if final_emp.empty:
                                final_emp = pd.DataFrame({"Month":MONTHS, "Line14_Final":[""]*12, "Line16_Final":[""]*12})
                            else:
                                ord_map = {m:i for i,m in enumerate(MONTHS, start=1)}
                                final_emp["_ord"] = final_emp["Month"].map(ord_map)
                                final_emp = final_emp.sort_values("_ord").drop(columns=["_ord"])

                            editable_name, editable_bytes, flat_name, flat_bytes = fill_pdf_for_employee(
                                pdf_file.getvalue(), emp_row, final_emp, year_used,
                                pay_ded_df=pay_ded_df,
                                emp_enroll_df=emp_enroll_df,
                                dep_enroll_df=dep_enroll_df
                            )
                            z.writestr(editable_name, editable_bytes.getvalue())
                            z.writestr(flat_name, flat_bytes.getvalue())

                    zip_buf.seek(0)
                    st.success(f"Bulk ZIP generated ({len(ids_to_run)} employees).")
                    st.download_button(
                        "Download ZIP",
                        zip_buf.getvalue(),
                        file_name=f"1095c_bulk_{year_used}.zip",
                        mime="application/zip"
                    )
                except Exception as e:
                    st.error(f"Bulk generation failed: {e}")
    else:
        st.warning("No employees found in the uploaded Excel.")
