# app.py
from __future__ import annotations
import io, os, zipfile, tempfile
from typing import Optional, Dict, Any, List, Tuple
import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject

# =========================
# FIXED RESOURCES (no upload)
# =========================
PDF_TEMPLATE_PATH = "assets/f1095c.pdf"  # keep your blank template here

# PASTE YOUR FIELD MAP HERE (example shape below)
FIELD_MAP: Dict[str, Any] = {
    # If the same code is used for all months, "all" takes priority; otherwise month keys.
    "line14": {
        "all": "l14_all",    # or omit "all" and use month keys
        "Jan": "l14_jan", "Feb": "l14_feb", "Mar": "l14_mar", "Apr": "l14_apr",
        "May": "l14_may", "June": "l14_jun", "July": "l14_jul", "Aug": "l14_aug",
        "Sept": "l14_sep", "Oct": "l14_oct", "Nov": "l14_nov", "Dec": "l14_dec"
    },
    "line16": {
        "all": "l16_all",
        "Jan": "l16_jan", "Feb": "l16_feb", "Mar": "l16_mar", "Apr": "l16_apr",
        "May": "l16_may", "June": "l16_jun", "July": "l16_jul", "Aug": "l16_aug",
        "Sept": "l16_sep", "Oct": "l16_oct", "Nov": "l16_nov", "Dec": "l16_dec"
    },
    # Optional sections — include only if your PDF has these text fields
    "part1": {
        "employee_first": "p1_emp_first",
        "employee_middle": "p1_emp_middle",
        "employee_last": "p1_emp_last",
        "employee_ssn": "p1_emp_ssn",
        "employee_addr1": "p1_emp_addr1",
        "employee_city": "p1_emp_city",
        "employee_state": "p1_emp_state",
        "employee_zip": "p1_emp_zip",
        "employee_country": "p1_emp_country",
        "employer_name": "p1_er_name",
        "employer_ein": "p1_er_ein",
        "employer_addr1": "p1_er_addr1",
        "employer_city": "p1_er_city",
        "employer_state": "p1_er_state",
        "employer_zip": "p1_er_zip",
        "employer_country": "p1_er_country",
        "employer_phone": "p1_er_phone",
    },
    "part2": {
        # If you want to fill plan start month like "01"
        "plan_start_month": "p2_plan_start_month"
    },
}
# Replace the above field names with your real PDF field names.


# =========================
# INTERIM BUILDER (your logic, lightly guarded)
# =========================
def _safe_to_datetime(s):
    return pd.to_datetime(s, errors="coerce")

def build_interim_from_excel(
    excel_bytes: bytes,
    year: int = 2025,
    demo_sheet: str = "Emp Demographic",
    elig_sheet: str = "Emp Eligibility",
    enr_sheet: str = "Emp Enrollment",
    exclude_waived: bool = True,
) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    demographic_df = pd.read_excel(xls, sheet_name=demo_sheet)
    eligibility_df = pd.read_excel(xls, sheet_name=elig_sheet)
    enrollment_df = pd.read_excel(xls, sheet_name=enr_sheet)

    for df in [demographic_df, eligibility_df, enrollment_df]:
        df.columns = df.columns.str.strip().str.replace(" ", "_")

    # Date coercion
    for col in ["StatusStartDate", "StatusEndDate"]:
        if col in demographic_df.columns:
            demographic_df[col] = _safe_to_datetime(demographic_df[col])
    for col in ["EligibilityStartDate", "EligibilityEndDate"]:
        if col in eligibility_df.columns:
            eligibility_df[col] = _safe_to_datetime(eligibility_df[col])
    for col in ["EnrollmentStartDate", "EnrollmentEndDate"]:
        if col in enrollment_df.columns:
            enrollment_df[col] = _safe_to_datetime(enrollment_df[col])

    far_future = pd.Timestamp("2262-04-11")
    if "StatusEndDate" in demographic_df.columns:
        demographic_df["StatusEndDate"] = demographic_df["StatusEndDate"].fillna(far_future)
    if "EligibilityEndDate" in eligibility_df.columns:
        eligibility_df["EligibilityEndDate"] = eligibility_df["EligibilityEndDate"].fillna(far_future)
    if "EnrollmentEndDate" in enrollment_df.columns:
        enrollment_df["EnrollmentEndDate"] = enrollment_df["EnrollmentEndDate"].fillna(far_future)

    if exclude_waived:
        if "EligiblePlan" in eligibility_df.columns:
            eligibility_df = eligibility_df[
                ~eligibility_df["EligiblePlan"].astype(str).str.contains("Waive", case=False, na=False)
            ]
        if "PlanCode" in enrollment_df.columns:
            enrollment_df = enrollment_df[
                ~enrollment_df["PlanCode"].astype(str).str.contains("Waive", case=False, na=False)
            ]

    emp_df = demographic_df[[
        "EmployeeID", "FirstName", "MiddleInitial", "LastName",
        "Role", "StatusStartDate", "StatusEndDate"
    ]].copy()

    emp_df["Name"] = (
        emp_df["FirstName"].fillna("") + " " +
        emp_df["MiddleInitial"].fillna("") + " " +
        emp_df["LastName"].fillna("")
    ).str.replace("  ", " ").str.strip()

    months = pd.date_range(start=f"{year}-01-01", end=f"{year}-12-31", freq="MS")

    records = []
    for _, emp in emp_df.iterrows():
        for month_start in months:
            month_end = month_start + pd.offsets.MonthEnd(1)

            employed_full_month = (emp["StatusStartDate"] <= month_start) and (emp["StatusEndDate"] >= month_end)
            is_full_time = emp["Role"] == "FT" and employed_full_month
            is_part_time = emp["Role"] == "PT" and employed_full_month

            elig_rows = eligibility_df[
                (eligibility_df["EmployeeID"] == emp["EmployeeID"]) &
                (eligibility_df["EligibilityStartDate"] <= month_end) &
                (eligibility_df["EligibilityEndDate"] >= month_start)
            ]
            enroll_rows = enrollment_df[
                (enrollment_df["EmployeeID"] == emp["EmployeeID"]) &
                (enrollment_df["EnrollmentStartDate"] <= month_end) &
                (enrollment_df["EnrollmentEndDate"] >= month_start)
            ]

            # Eligibility flags
            eligible_mv = False
            employee_eligible = False
            spouse_eligible = False
            child_eligible = False
            if not elig_rows.empty:
                all_plans = set(elig_rows["EligiblePlan"].astype(str))
                all_tiers = set(elig_rows["EligibleTier"].astype(str))
                if all_plans == {"PlanA"} and all_tiers.intersection({"EMPFAM", "EMP", "EMPCHILD", "EMPSPOUSE"}):
                    eligible_mv = True
                if all_tiers.intersection({"EMP", "EMPFAM", "EMPSPOUSE"}):
                    employee_eligible = True
                if all_tiers.intersection({"EMPFAM", "EMPSPOUSE"}):
                    spouse_eligible = True
                if "EMPFAM" in all_tiers:
                    child_eligible = True

            # Enrollment flags
            employee_enrolled = False
            spouse_enrolled = False
            child_enrolled = False
            if not enroll_rows.empty:
                all_tiers_enr = set(enroll_rows["Tier"].dropna().astype(str))
                if all_tiers_enr.intersection({"EMP", "EMPFAM", "EMPSPOUSE"}):
                    employee_enrolled = True
                if all_tiers_enr.intersection({"EMPFAM", "EMPSPOUSE"}):
                    spouse_enrolled = True
                if "EMPFAM" in all_tiers_enr:
                    child_enrolled = True

            records.append({
                "Employee_ID": emp["EmployeeID"],
                "Name": emp["Name"],
                "Month": month_start.strftime("%b"),
                "Is_Employed_full_month": "Yes" if employed_full_month else "No",
                "Is_full_time_full_month": "Yes" if is_full_time else "No",
                "Is_Part_time_full_month": "Yes" if is_part_time else "No",
                "eligible_mv": eligible_mv,
                "employee_eligible": employee_eligible,
                "spouse_eligible": spouse_eligible,
                "child_eligible": child_eligible,
                "employee_enrolled": employee_enrolled,
                "spouse_enrolled": spouse_enrolled,
                "child_enrolled": child_enrolled
            })

    return pd.DataFrame(records)


# =========================
# PDF FILLER (in-memory)
# =========================
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

def enable_need_appearances(reader: PdfReader, writer: PdfWriter):
    try:
        root = reader.trailer.get("/Root")
        if hasattr(root, "get_object"):
            root = root.get_object()
        acro = None
        if root:
            acro = root.get("/AcroForm")
        if hasattr(acro, "get_object"):
            acro = acro.get_object()
        if acro is None:
            return
        acro.update({NameObject("/NeedAppearances"): BooleanObject(True)})
        writer._root_object.update({NameObject("/AcroForm"): acro})
    except Exception as e:
        print("ℹ️ Could not set NeedAppearances:", e)

def set_form_text(writer: PdfWriter, field_name: str, value: str):
    if not field_name:
        return
    try:
        writer.update_page_form_field_values(writer.pages[0], {field_name: value})
    except Exception:
        try:
            writer.update_page_form_field_values(writer.pages[0], {field_name.strip(): value})
        except Exception as e:
            print(f"⚠️ Unable to set field '{field_name}': {e}")

def fill_part1(writer: PdfWriter, fieldmap: Dict[str, Any], vals: Dict[str, Any]):
    p1 = fieldmap.get("part1", {}) or {}
    if not p1:
        return
    for k, v in vals.items():
        if k in p1 and v is not None:
            set_form_text(writer, p1[k], str(v))

def all_12_same(codes_by_month: Dict[str, Optional[str]]) -> Optional[str]:
    vals = [codes_by_month.get(m) for m in MONTHS]
    vals = [v for v in vals if v and str(v).strip()]
    if len(vals) == 12 and len(set(vals)) == 1:
        return list(set(vals))[0]
    return None

def fill_line_codes(writer: PdfWriter, fieldmap: Dict[str, Any],
                    codes_by_month: Dict[str, Optional[str]], which: str):
    section = fieldmap.get(which, {}) or {}
    if not section:
        return
    same = all_12_same(codes_by_month)
    if same and "all" in section:
        set_form_text(writer, section["all"], same)
        for m in MONTHS:
            if m in section:
                set_form_text(writer, section[m], "")
    else:
        for m in MONTHS:
            if m in section:
                set_form_text(writer, section[m], codes_by_month.get(m) or "")

def line_dict_from_block(block: pd.DataFrame, col_name: str) -> Dict[str, Optional[str]]:
    out = {m: None for m in MONTHS}
    for _, r in block.iterrows():
        mo = month_to_canonical(r["Month"])
        if mo in out:
            out[mo] = r.get(col_name)
    return out

def derive_part1_values_from_demo_row(row: Optional[pd.Series]) -> Dict[str, Any]:
    # You can wire this to demographics if you extend the app; here we keep blank/defaults.
    return {
        "employee_first": None, "employee_middle": None, "employee_last": None, "employee_ssn": None,
        "employee_addr1": None, "employee_city": None, "employee_state": None, "employee_zip": None, "employee_country": None,
        "employer_name": None, "employer_ein": None, "employer_addr1": None, "employer_city": None,
        "employer_state": None, "employer_zip": None, "employer_country": None, "employer_phone": None,
    }

def fill_one_employee_pdf_bytes(
    emp_id: int,
    monthly_summary: pd.DataFrame,  # must include Employee_ID, Month, line_14, line_16
    fieldmap: Dict[str, Any],
    pdf_template_path: str,
    plan_start_month: Optional[str] = None
) -> bytes:
    block = monthly_summary.loc[monthly_summary["Employee_ID"] == emp_id].copy()
    if block.empty:
        raise ValueError("No monthly rows found for this employee.")

    block["Month"] = block["Month"].apply(month_to_canonical)
    block["__midx"] = block["Month"].apply(lambda x: MONTHS.index(str(x)) if str(x) in MONTHS else 0)
    block = block.sort_values(["__midx"]).drop(columns="__midx", errors="ignore")

    l14 = line_dict_from_block(block, "line_14")
    l16 = line_dict_from_block(block, "line_16")
    p1_vals = derive_part1_values_from_demo_row(None)

    reader = PdfReader(pdf_template_path)
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    enable_need_appearances(reader, writer)

    fill_part1(writer, fieldmap, p1_vals)
    if "part2" in fieldmap and "plan_start_month" in fieldmap["part2"] and plan_start_month:
        set_form_text(writer, fieldmap["part2"]["plan_start_month"], plan_start_month)

    fill_line_codes(writer, fieldmap, l14, "line14")
    fill_line_codes(writer, fieldmap, l16, "line16")

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="ACA One-Flow: Interim + PDFs", page_icon="🧾", layout="wide")
st.title("🧾 ACA 1095-C — Build Interim → Generate PDFs")

with st.sidebar:
    st.header("Step 0 — Settings")
    year = st.number_input("Year for Interim", min_value=2000, max_value=2100, value=2025, step=1)
    exclude_waived = st.checkbox("Exclude Waived plans", value=True)
    plan_start = st.text_input("Plan Start Month (optional, e.g., 01)", value="")

# Step 1: upload Excel
uploaded = st.file_uploader("Upload Input Data Excel (single file)", type=["xlsx"])

# Internal state to gate the PDF section
if "interim_ready" not in st.session_state:
    st.session_state.interim_ready = False
if "interim_df" not in st.session_state:
    st.session_state.interim_df = None

colA, colB = st.columns([1, 1])
with colA:
    build_btn = st.button("🚀 Build Interim", use_container_width=True)

if build_btn:
    if not uploaded:
        st.error("Please upload the Excel file first.")
        st.stop()
    try:
        interim_df = build_interim_from_excel(
            excel_bytes=uploaded.getvalue(),
            year=year,
            exclude_waived=exclude_waived,
        )
        # For PDF filler, we also need monthly summary with line_14/line_16.
        # If your pipeline produces that separately, you can merge it here.
        # For demo, we create empty placeholders so UI stays consistent.
        if "line_14" not in interim_df.columns:
            interim_df["line_14"] = ""
        if "line_16" not in interim_df.columns:
            interim_df["line_16"] = ""

        st.session_state.interim_df = interim_df
        st.session_state.interim_ready = True
        st.success(f"Interim built with {len(interim_df):,} rows.")
    except Exception as e:
        st.exception(e)
        st.stop()

# Show interim + download
if st.session_state.interim_ready and st.session_state.interim_df is not None:
    st.subheader("Interim Preview")
    st.dataframe(st.session_state.interim_df.head(500), use_container_width=True, height=420)

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        csv_bytes = st.session_state.interim_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Interim (CSV)", data=csv_bytes,
                           file_name=f"Interim_{year}.csv", mime="text/csv", use_container_width=True)
    with dl_col2:
        try:
            import xlsxwriter
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                st.session_state.interim_df.to_excel(writer, index=False, sheet_name=f"Interim_{year}")
            st.download_button("⬇️ Download Interim (XLSX)", data=buf.getvalue(),
                               file_name=f"Interim_{year}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        except Exception:
            st.info("Install xlsxwriter to enable Excel download (`pip install xlsxwriter`).")

    st.markdown("---")
    st.subheader("Generate PDFs (unlocked after building Interim)")

    if not os.path.exists(PDF_TEMPLATE_PATH):
        st.error(f"Missing PDF template at {PDF_TEMPLATE_PATH}. Put your blank 1095-C there.")
    else:
        mode = st.radio("Choose generation mode", ["Single employee", "Bulk (all employees)"], horizontal=True)
        plan_start_clean = plan_start.strip() or None

        if mode == "Single employee":
            emp_id = st.number_input("Employee_ID to generate", min_value=0, step=1, value=1001)
            go_single = st.button("🧾 Generate Single PDF", type="primary")
            if go_single:
                # Validate required columns
                df = st.session_state.interim_df.copy()
                for col in ["Employee_ID", "Month", "line_14", "line_16"]:
                    if col not in df.columns:
                        st.error(f"Interim is missing required column: {col}")
                        st.stop()
                try:
                    pdf_bytes = fill_one_employee_pdf_bytes(
                        emp_id=int(emp_id),
                        monthly_summary=df,
                        fieldmap=FIELD_MAP,
                        pdf_template_path=PDF_TEMPLATE_PATH,
                        plan_start_month=plan_start_clean
                    )
                    st.success(f"Generated PDF for Employee {int(emp_id)}")
                    st.download_button(
                        "⬇️ Download 1095-C PDF",
                        data=pdf_bytes,
                        file_name=f"1095C_{int(emp_id)}.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(str(e))

        else:
            go_bulk = st.button("🧾 Generate Bulk PDFs (ZIP)", type="primary")
            if go_bulk:
                df = st.session_state.interim_df.copy()
                for col in ["Employee_ID", "Month", "line_14", "line_16"]:
                    if col not in df.columns:
                        st.error(f"Interim is missing required column: {col}")
                        st.stop()
                # Unique employees
                emp_ids = [int(x) for x in pd.to_numeric(df["Employee_ID"], errors="coerce").dropna().unique().tolist()]
                if not emp_ids:
                    st.error("No valid Employee_ID values found.")
                    st.stop()

                mem_zip = io.BytesIO()
                errors: List[str] = []
                with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for eid in emp_ids:
                        try:
                            pdf_bytes = fill_one_employee_pdf_bytes(
                                emp_id=eid,
                                monthly_summary=df,
                                fieldmap=FIELD_MAP,
                                pdf_template_path=PDF_TEMPLATE_PATH,
                                plan_start_month=plan_start_clean
                            )
                            zf.writestr(f"1095C_{eid}.pdf", pdf_bytes)
                        except Exception as e:
                            errors.append(f"Employee {eid}: {e}")
                mem_zip.seek(0)

                if errors:
                    st.warning("Some employees failed:")
                    for m in errors[:20]:
                        st.write(f"• {m}")
                    if len(errors) > 20:
                        st.write(f"… and {len(errors)-20} more.")

                st.success(f"Prepared {len(emp_ids)-len(errors)} PDF(s).")
                st.download_button(
                    "⬇️ Download ZIP of PDFs",
                    data=mem_zip,
                    file_name="1095C_PDFs.zip",
                    mime="application/zip"
                )
