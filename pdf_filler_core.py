# pdf_filler_core.py
from __future__ import annotations
import os, io, json
from typing import Dict, Any, Optional, List, Tuple
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject

MONTHS = ["Jan","Feb","Mar","Apr","May","June","July","Aug","Sept","Oct","Nov","Dec"]
MONTH_ALIASES = {
    "Jan": "Jan", "January": "Jan",
    "Feb": "Feb", "February": "Feb",
    "Mar": "Mar", "March": "Mar",
    "Apr": "Apr", "April": "Apr",
    "May": "May",
    "Jun": "June", "June": "June",
    "Jul": "July", "July": "July",
    "Aug": "Aug", "August": "Aug",
    "Sep": "Sept", "Sept": "Sept", "September": "Sept",
    "Oct": "Oct", "October": "Oct",
    "Nov": "Nov", "November": "Nov",
    "Dec": "Dec", "December": "Dec",
}

def month_to_canonical(m: str) -> str:
    m = str(m).strip()
    return MONTH_ALIASES.get(m, m)

def set_need_appearances(reader: PdfReader, writer: PdfWriter):
    try:
        root = reader.trailer.get("/Root")
        if hasattr(root, "get_object"): root = root.get_object()
        acro = None
        if root: acro = root.get("/AcroForm")
        if hasattr(acro, "get_object"): acro = acro.get_object()
        if acro is None: return
        acro.update({NameObject("/NeedAppearances"): BooleanObject(True)})
        writer._root_object.update({NameObject("/AcroForm"): acro})
    except Exception:
        pass

def set_field(writer: PdfWriter, field_name: Optional[str], value: Optional[str]):
    if not field_name or value is None: return
    try:
        writer.update_page_form_field_values(writer.pages[0], {field_name: str(value)})
    except Exception:
        try:
            writer.update_page_form_field_values(writer.pages[0], {field_name.strip(): str(value)})
        except Exception:
            pass

def derive_part1_values(emp_id: int, demo_df: Optional[pd.DataFrame]) -> Dict[str, Any]:
    out = {
        "employee_first": None, "employee_middle": None, "employee_last": None, "employee_ssn": None,
        "employee_addr1": None, "employee_city": None, "employee_state": None, "employee_zip": None, "employee_country": None,
        "employer_name": None, "employer_ein": None, "employer_addr1": None, "employer_city": None, "employer_state": None,
        "employer_zip": None, "employer_country": None, "employer_phone": None,
    }
    if demo_df is None: return out
    g = demo_df.loc[demo_df["EmployeeID"] == emp_id]
    if g.empty: return out

    def pick(s): 
        s = s.dropna().astype(str).str.strip()
        return s.iloc[0] if len(s) else None

    colmap = {
        "FirstName": "employee_first",
        "MiddleInitial": "employee_middle",
        "LastName": "employee_last",
        "SSN": "employee_ssn",
        "AddressLine1": "employee_addr1",
        "City": "employee_city",
        "State": "employee_state",
        "ZipCode": "employee_zip",
        "Country": "employee_country",
        "EmployerName": "employer_name",
        "EIN": "employer_ein",
        "EmployerAddress": "employer_addr1",
        "EmployerCity": "employer_city",
        "EmployerState": "employer_state",
        "EmployerZipCode": "employer_zip",
        "EmployerCountry": "employer_country",
        "ContactTelephone": "employer_phone",
    }
    for src, tgt in colmap.items():
        if src in g.columns:
            out[tgt] = pick(g[src])

    if out["employee_ssn"]:
        out["employee_ssn"] = str(out["employee_ssn"]).replace("-", "").replace(" ", "")
    for key in ("employee_zip", "employer_zip"):
        if out[key] is not None:
            out[key] = str(out[key]).split(".")[0]
    return out

def load_field_map(json_bytes: Optional[bytes]) -> Dict[str, Any]:
    if not json_bytes:
        return {"line14": {}, "line16": {}, "part1": {}, "part2": {}}
    m = json.loads(json_bytes.decode("utf-8"))
    m.setdefault("line14", {}); m.setdefault("line16", {})
    m.setdefault("part1", {});  m.setdefault("part2", {})
    return m

def all_12_same(codes_by_month: Dict[str, Optional[str]]) -> Optional[str]:
    vals = [codes_by_month.get(m) for m in MONTHS]
    vals = [v for v in vals if v and str(v).strip()]
    if len(vals) == 12 and len(set(vals)) == 1:
        return list(set(vals))[0]
    return None

def line_dict_from_block(block: pd.DataFrame, col_name: str) -> Dict[str, Optional[str]]:
    out = {m: None for m in MONTHS}
    for _, r in block.iterrows():
        mo = month_to_canonical(r["Month"])
        if mo in out:
            out[mo] = r.get(col_name)
    return out

def fill_line_codes(writer: PdfWriter, fieldmap: Dict[str, Any], codes_by_month: Dict[str, Optional[str]], which: str):
    section = fieldmap.get(which, {}) or {}
    if not section: return
    same = all_12_same(codes_by_month)
    if same and "all" in section:
        set_field(writer, section["all"], same)
        for m in MONTHS:
            if m in section: set_field(writer, section[m], "")
    else:
        for m in MONTHS:
            if m in section: set_field(writer, section[m], codes_by_month.get(m) or "")

def fill_one_employee(
    emp_id: int,
    pdf_template_bytes: bytes,
    field_map: Dict[str, Any],
    demo_df: Optional[pd.DataFrame],
    monthly_summary_df: Optional[pd.DataFrame],  # must have Employee_ID, Month, line_14, line_16 if provided
    plan_start_month: Optional[str] = None
) -> bytes:
    reader = PdfReader(io.BytesIO(pdf_template_bytes))
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    set_need_appearances(reader, writer)

    # Part I
    p1_vals = derive_part1_values(emp_id, demo_df)
    for k, v in p1_vals.items():
        f = field_map.get("part1", {}).get(k)
        if f: set_field(writer, f, v)

    # Part II (optional if monthly summary present)
    if monthly_summary_df is not None:
        df = monthly_summary_df.copy()
        req = {"Employee_ID","Month","line_14","line_16"}
        missing = req - set(df.columns)
        if missing:
            raise ValueError(f"Monthly summary missing required columns: {missing}")
        df["Employee_ID"] = pd.to_numeric(df["Employee_ID"], errors="coerce").astype("Int64")

        block = df.loc[df["Employee_ID"] == emp_id].copy()
        if not block.empty:
            block["Month"] = block["Month"].apply(month_to_canonical)
            block["__midx"] = block["Month"].apply(lambda x: MONTHS.index(str(x)) if str(x) in MONTHS else 0)
            block = block.sort_values(["__midx"]).drop(columns="__midx", errors="ignore")
            l14 = line_dict_from_block(block, "line_14")
            l16 = line_dict_from_block(block, "line_16")
            fill_line_codes(writer, field_map, l14, "line14")
            fill_line_codes(writer, field_map, l16, "line16")

    if field_map.get("part2", {}).get("plan_start_month") and plan_start_month:
        set_field(writer, field_map["part2"]["plan_start_month"], plan_start_month)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()
