# interim_core.py
from __future__ import annotations
import io
import pandas as pd
from typing import Optional

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
    """
    Build the interim table exactly as in your snippet, with small guardrails.
    """
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    demographic_df = pd.read_excel(xls, sheet_name=demo_sheet)
    eligibility_df = pd.read_excel(xls, sheet_name=elig_sheet)
    enrollment_df = pd.read_excel(xls, sheet_name=enr_sheet)

    for df in [demographic_df, eligibility_df, enrollment_df]:
        df.columns = df.columns.str.strip().str.replace(" ", "_")

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
            eligibility_df = eligibility_df[~eligibility_df["EligiblePlan"].astype(str).str.contains("Waive", case=False, na=False)]
        if "PlanCode" in enrollment_df.columns:
            enrollment_df = enrollment_df[~enrollment_df["PlanCode"].astype(str).str.contains("Waive", case=False, na=False)]

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
