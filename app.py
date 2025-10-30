# app.py
import io, zipfile
import pandas as pd
import streamlit as st
from interim_core import build_interim_from_excel
from pdf_filler_core import load_field_map, fill_one_employee

st.set_page_config(page_title="ACA Interim + 1095-C PDFs", page_icon="🧾", layout="wide")
st.title("🧾 ACA Interim Builder & 1095-C PDF Generator")

with st.sidebar:
    st.header("Settings")
    year = st.number_input("Year", min_value=2000, max_value=2100, value=2025, step=1)
    exclude_waived = st.checkbox("Exclude Waived plans", value=True)
    plan_start = st.text_input("Plan Start Month (optional, e.g., 01)", value="")
    st.caption("Plan start fills Part II if your mapping includes it.")

st.subheader("1) Upload required files")
colA, colB, colC = st.columns([1.2,1,1])
with colA:
    f_input = st.file_uploader("Input Data Excel (Emp Demographic / Eligibility / Enrollment)", type=["xlsx"])
with colB:
    f_pdf   = st.file_uploader("Blank 1095-C PDF template", type=["pdf"])
with colC:
    f_map   = st.file_uploader("Field Mapping JSON (optional but recommended)", type=["json"])

st.subheader("2) (Optional) Upload Monthly Summary (for Line 14/16)")
f_summary = st.file_uploader("Monthly Summary Excel (Employee_ID, Month, line_14, line_16)", type=["xlsx"])

st.markdown("---")
build_btn = st.button("🚀 Build Interim & Enable PDF Generation")

if build_btn:
    if not f_input or not f_pdf:
        st.error("Please upload at least the Input Excel and the blank PDF template.")
        st.stop()

    try:
        # Build interim from main input
        interim_df = build_interim_from_excel(
            excel_bytes=f_input.getvalue(),
            year=year,
            exclude_waived=exclude_waived,
        )
        st.success(f"Interim built with {len(interim_df):,} rows.")
        st.dataframe(interim_df, use_container_width=True, height=420)

        # Downloads for Interim
        csv_bytes = interim_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Interim CSV", data=csv_bytes, file_name=f"Interim_{year}.csv", mime="text/csv")

        try:
            import xlsxwriter
            xio = io.BytesIO()
            with pd.ExcelWriter(xio, engine="xlsxwriter") as wr:
                interim_df.to_excel(wr, sheet_name=f"Interim_{year}", index=False)
            st.download_button(
                "⬇️ Download Interim Excel (XLSX)",
                data=xio.getvalue(),
                file_name=f"Interim_{year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception:
            st.info("Install xlsxwriter for XLSX downloads (`pip install xlsxwriter`).")

        # Prepare data for PDF generation
        demo_df = None
        try:
            # Re-open demographic sheet for Part I (from the same input file)
            xls = pd.ExcelFile(io.BytesIO(f_input.getvalue()))
            demo_df = pd.read_excel(xls, sheet_name="Emp Demographic")
            demo_df.columns = demo_df.columns.str.strip().str.replace(" ", "_")
            if "EmployeeID" in demo_df.columns:
                demo_df["EmployeeID"] = pd.to_numeric(demo_df["EmployeeID"], errors="coerce").astype("Int64")
        except Exception:
            pass

        field_map = load_field_map(f_map.read() if f_map else None)

        summary_df = None
        if f_summary:
            try:
                summary_df = pd.read_excel(io.BytesIO(f_summary.getvalue()))
            except Exception as e:
                st.warning(f"Could not read Monthly Summary: {e}")

        st.markdown("### 3) Generate PDFs")
        st.caption("Choose single employee or generate for all employees (bulk). Part I always fills. Part II fills only if Monthly Summary + mapping for Line14/16 are provided.")

        # SINGLE
        emp_ids = sorted(pd.to_numeric(interim_df["Employee_ID"].dropna(), errors="coerce").dropna().astype(int).unique().tolist())
        c1, c2 = st.columns([1, 2])
        with c1:
            selected_emp = st.selectbox("Employee for Single PDF", emp_ids if emp_ids else [0])
            single_btn = st.button("Generate Single PDF")

        if single_btn:
            try:
                pdf_bytes = fill_one_employee(
                    emp_id=int(selected_emp),
                    pdf_template_bytes=f_pdf.getvalue(),
                    field_map=field_map,
                    demo_df=demo_df,
                    monthly_summary_df=summary_df,  # may be None → Part II skipped
                    plan_start_month=plan_start.strip() or None
                )
                st.download_button(
                    "⬇️ Download Single 1095-C PDF",
                    data=pdf_bytes,
                    file_name=f"1095C_{int(selected_emp)}.pdf",
                    mime="application/pdf",
                )
                st.success(f"Generated PDF for Employee {int(selected_emp)}.")
            except Exception as e:
                st.error(f"Single PDF error: {e}")

        # BULK
        with c2:
            bulk_btn = st.button("Generate Bulk PDFs (all employees in Interim)")
        if bulk_btn:
            try:
                ids = emp_ids
                mem_zip = io.BytesIO()
                with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for eid in ids:
                        try:
                            pdf_data = fill_one_employee(
                                emp_id=int(eid),
                                pdf_template_bytes=f_pdf.getvalue(),
                                field_map=field_map,
                                demo_df=demo_df,
                                monthly_summary_df=summary_df,
                                plan_start_month=plan_start.strip() or None
                            )
                            zf.writestr(f"1095C_{int(eid)}.pdf", pdf_data)
                        except Exception as ee:
                            # write a small txt marker for failures
                            zf.writestr(f"ERROR_{int(eid)}.txt", f"{ee}")
                mem_zip.seek(0)
                st.download_button(
                    "⬇️ Download ZIP (All PDFs)",
                    data=mem_zip,
                    file_name=f"1095C_ALL_{year}.zip",
                    mime="application/zip",
                )
                st.success(f"Bulk ZIP ready for {len(emp_ids)} employees.")
            except Exception as e:
                st.error(f"Bulk PDF error: {e}")

    except Exception as e:
        st.exception(e)
