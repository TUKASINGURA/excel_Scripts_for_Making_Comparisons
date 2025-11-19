import pandas as pd

# --- INPUT FILE ---
input_file = "sample.xlsx"
sheet = "Sheet1"   # change if needed

# --- OUTPUT FILE ---
output_file = "cleaned_no_empty_rows.xlsx"

# --- LOAD THE EXCEL FILE ---
df = pd.read_excel(input_file, sheet_name=sheet, engine="openpyxl")

# --- REMOVE ROWS THAT ARE ENTIRELY EMPTY ---
df_clean = df.dropna(how="all")

# If you ALSO want to remove rows where all cells are empty strings or whitespace:
df_clean = df_clean.replace(r'^\s*$', pd.NA, regex=True).dropna(how="all")

# --- SAVE THE CLEANED FILE ---
df_clean.to_excel(output_file, index=False)

print(f"Original rows: {len(df)}")
print(f"Rows after removing empty rows: {len(df_clean)}")
print(f"Saved cleaned file as: {output_file}")
