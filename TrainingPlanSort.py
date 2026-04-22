import re
from pathlib import Path

import pandas as pd

# =========================
# Paths
# =========================

input_path = Path(
    r"C:\Users\e.muraj\Downloads\training records by TP-for conversion.xlsx"
)

# Read Excel (default first sheet; set sheet_name=... if needed)
df = pd.read_excel(input_path)

# =========================
# Helper functions
# =========================

def clean_legacy_document_number(value):
    """
    Legacy Document Number logic (from Training Number):

    1) If it begins with 'CORP' (case-insensitive), return as-is.
    2) Otherwise, if it ends with '-######' (dash + EXACTLY 6 digits), remove that suffix.
    3) After that, if the result starts with a digit, prepend 'SOP-'.
    """
    if pd.isna(value):
        return value
    if not isinstance(value, str):
        value = str(value)

    s = value.strip()

    # Step 1: CORP exception
    if s.upper().startswith("CORP"):
        return s

    # Step 2: remove trailing dash + 6 digits
    s = re.sub(r"-\d{6}$", "", s).strip()

    # Step 3: prepend SOP- if starts with digit
    if re.match(r"^\d", s):
        s = "SOP-" + s

    return s


def clean_plan_name(value):
    """
    Clean 'Training Plan Name' to just the descriptive part:
    'TP-0001_ Data Integrity' -> 'Data Integrity'
    If no underscore, just strip whitespace.
    """
    if pd.isna(value):
        return ""
    if not isinstance(value, str):
        value = str(value)

    s = value.strip()
    if "_" in s:
        return s.split("_", 1)[1].strip()
    return s


def build_required_roles(sub_df):
    """
    For a group of rows with the same Legacy Document Number:
    Build the 'Required for Role(s)' string as:

        TP-0001_Data Integrity|TP-0011_Quality Core|...

    using:
        - 'Training Plan: Training Plan ID'
        - cleaned 'Training Plan Name'
    """
    seen = []
    for _, row in sub_df.iterrows():
        plan_id = row.get("Training Plan: Training Plan ID", None)
        plan_name_raw = row.get("Training Plan Name", None)

        if pd.isna(plan_id):
            continue

        plan_id_str = str(plan_id).strip()
        if not plan_id_str:
            continue

        clean_name = clean_plan_name(plan_name_raw)

        # If we have a clean name, combine; else just use ID
        if clean_name:
            role_str = f"{plan_id_str}_{clean_name}"
        else:
            role_str = plan_id_str

        if role_str not in seen:
            seen.append(role_str)

    return "|".join(seen)


def pick_course_name(sub_df):
    """
    For a group of rows with the same Legacy Document Number:
    Pick a representative Course Name from 'Training Name'.
    """
    for v in sub_df["Training Name"]:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            return s
    return ""


# =========================
# Apply transformations
# =========================

# 1) Create Legacy Document Number from Training Number
df["Legacy Document Number"] = df["Training Number"].apply(clean_legacy_document_number)

# 2) Group by Legacy Document Number
groups = df.groupby("Legacy Document Number", dropna=True)

rows = []
for legacy_doc_num, sub_df in groups:
    if pd.isna(legacy_doc_num):
        continue

    course_name = pick_course_name(sub_df)
    required_roles = build_required_roles(sub_df)

    rows.append({
        "Legacy Document Number": legacy_doc_num,
        "Course Name": course_name,
        "Required for Role(s)": required_roles
    })

result = pd.DataFrame(rows)

# Ensure column order
result = result[["Legacy Document Number", "Course Name", "Required for Role(s)"]]

# =========================
# Save output in same folder
# =========================

output_path = input_path.with_name(input_path.stem + "_legacy_courses" + input_path.suffix)
result.to_excel(output_path, index=False)

print(f"Legacy course mapping saved to: {output_path}")
