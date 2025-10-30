# app.py
import os, io, zipfile, tempfile, json
import streamlit as st
from pdf_filler_core import generate_pdfs

st.set_page_config(page_title="1095-C PDF Filler", page_icon="🧾", layout="centered")

st.title("🧾 ACA 1095-C PDF Filler")
st.caption("Fill Part II (and optional Part I) from your monthly summary & mapping JSON")

with st.sidebar:
    st.subheader("Mode")
    mode = st.radio("Choose run mode:", ["bulk", "single"], index=0)
    emp_id = None
    if mode == "single":
        emp_id = st.number_input("Employee_ID to generate", min_value=0, step=1, value=1001)

    plan_start = st.text_input("Plan Start Month (optional, e.g., 01)", value="")
    max_emps = st.number_input("Max employees (optional, for testing)", min_value=0, step=1, value=0)
    max_emps = None if max_emps == 0 else int(max_emps)

st.markdown("### 1) Upload files")

col1, col2 = st.columns(2)
with col1:
    f_summary = st.file_uploader("Monthly summary Excel (.xlsx)", type=["xlsx"])
    f_demo    = st.file_uploader("Demographics Excel (.xlsx) (optional)", type=["xlsx"])
with col2:
    f_pdf     = st.file_uploader("Blank 1095-C PDF template (.pdf)", type=["pdf"])
    f_map     = st.file_uploader("Field mapping JSON (.json)", type=["json"])

st.markdown("---")
run = st.button("🚀 Generate PDFs")

if run:
    # Basic validation
    if not f_summary or not f_pdf or not f_map:
        st.error("Please upload at least: Monthly summary, PDF template, and Field map JSON.")
        st.stop()

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Persist uploads to disk so core can use file paths
            sum_path = os.path.join(tmpdir, "summary.xlsx")
            pdf_path = os.path.join(tmpdir, "template.pdf")
            map_path = os.path.join(tmpdir, "fieldmap.json")
            out_dir  = os.path.join(tmpdir, "out_pdfs")
            with open(sum_path, "wb") as f: f.write(f_summary.read())
            with open(pdf_path, "wb") as f: f.write(f_pdf.read())
            with open(map_path, "wb") as f: f.write(f_map.read())
            demo_path = None
            if f_demo:
                demo_path = os.path.join(tmpdir, "demo.xlsx")
                with open(demo_path, "wb") as f: f.write(f_demo.read())

            with st.status("Working…", expanded=True) as status:
                st.write("Validating files and field map…")
                made, errors, out_dir = generate_pdfs(
                    summary_xlsx_path=sum_path,
                    pdf_template_path=pdf_path,
                    field_map_json_path=map_path,
                    demographics_xlsx_path=demo_path,
                    mode=mode,
                    employee_id=int(emp_id) if emp_id is not None else None,
                    max_employees=max_emps,
                    plan_start_month=plan_start.strip() or None,
                    out_dir=out_dir
                )

                # Zip outputs for download
                st.write("Packaging results…")
                mem_zip = io.BytesIO()
                with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name in sorted(os.listdir(out_dir)):
                        if name.lower().endswith(".pdf"):
                            zf.write(os.path.join(out_dir, name), arcname=name)
                mem_zip.seek(0)

                status.update(label="Done!", state="complete")

            st.success(f"Created {made} PDF(s).")
            if errors:
                st.warning("Some employees failed:")
                for m in errors[:20]:
                    st.write(f"• {m}")
                if len(errors) > 20:
                    st.write(f"… and {len(errors)-20} more.")

            st.download_button(
                "⬇️ Download ZIP of PDFs",
                data=mem_zip,
                file_name="1095C_PDFs.zip",
                mime="application/zip",
            )

    except Exception as e:
        st.error(f"Error: {e}")
