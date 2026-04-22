import pandas as pd

# Path to your Excel file (fixed with raw string)
excel_path = r"C:\Users\e.muraj\Downloads\Training Records by Plan (STK19NOV2025) - As exported 08DEC2025.xlsx"

column_name = "Training Plan Name"

# Read the Excel file
df = pd.read_excel(excel_path)

# Make sure the column exists
if column_name not in df.columns:
    raise KeyError(f"Column '{column_name}' not found in the Excel file. Available columns: {list(df.columns)}")

# Get the column, drop NaNs, convert to string, and strip whitespace
col_series = df[column_name].dropna().astype(str).str.strip()

# Filter values that do NOT start with "TP-"
invalid_values = col_series[~col_series.str.startswith("TP-")]

# Get unique values only
unique_invalid_values = invalid_values.unique()

print("Values in 'Training Plan Name' that do NOT start with 'TP-':")
for value in unique_invalid_values:
    print(value)
