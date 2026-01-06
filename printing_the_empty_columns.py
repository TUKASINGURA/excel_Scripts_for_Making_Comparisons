import argparse
import pandas as pd

# --- ARGPARSER / INPUTS ---
parser = argparse.ArgumentParser(description="Find and save rows with empty target column.")
parser.add_argument("--input", "-i", default="VOD_descriptions.xlsx", help="Input Excel file")
parser.add_argument("--sheet", "-s", default="Sheet1", help="Sheet name")
parser.add_argument("--column", "-c", default="vods_description", help="Target column to check")
parser.add_argument("--print-full", action="store_true", dest="print_full", help="Print full file to console (may be very large)")
args = parser.parse_args()

input_file = args.input
sheet = args.sheet
target_column = args.column

# --- LOAD THE EXCEL FILE ---
df = pd.read_excel(input_file, sheet_name=sheet, engine="openpyxl")

# --- TREAT WHITESPACE-ONLY CELLS AS EMPTY ---
df = df.replace(r'^\s*$', pd.NA, regex=True)

# --- CHECK IF COLUMN EXISTS ---
if target_column not in df.columns:
    raise ValueError(f"Column '{target_column}' not found in the sheet")

# --- FIND ROWS WHERE DESCRIPTION IS EMPTY ---
rows_with_empty_description = df[df[target_column].isna()]

# --- PRINT RESULTS ---
print(f"Rows where '{target_column}' is empty:\n")

if not rows_with_empty_description.empty:
    print(rows_with_empty_description.to_string(index=True))
else:
    print(f"No rows found with empty '{target_column}'.")

print("\nSummary:")
print(f"Total rows: {len(df)}")
print(f"Rows with empty '{target_column}': {len(rows_with_empty_description)}")

# --- OPTIONAL: print entire file ---
if args.print_full:
    print("\n--- Full file contents ---\n")
    print(df.to_string(index=True))

# --- SAVE rows with empty target column to files ---
out_base = f"rows_with_empty_{target_column}"
xlsx_path = f"{out_base}.xlsx"
csv_path = f"{out_base}.csv"
rows_with_empty_description.to_excel(xlsx_path, index=False)
rows_with_empty_description.to_csv(csv_path, index=False)

# write a concise summary log so background runs can be checked quickly
log_path = f"{out_base}.log"
with open(log_path, "w", encoding="utf-8") as fh:
    fh.write(f"Input file: {input_file}\n")
    fh.write(f"Sheet: {sheet}\n")
    fh.write(f"Target column: {target_column}\n")
    fh.write(f"Total rows: {len(df)}\n")
    fh.write(f"Rows with empty '{target_column}': {len(rows_with_empty_description)}\n")
    fh.write(f"Saved Excel: {xlsx_path}\n")
    fh.write(f"Saved CSV: {csv_path}\n")

print(f"\nSaved rows with empty '{target_column}' to: {xlsx_path} and {csv_path}")
print(f"Summary written to: {log_path}")
