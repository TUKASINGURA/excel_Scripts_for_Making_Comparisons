import pandas as pd

# --- INPUT FILE ---
input_file = "non_matched_rows_not_containing_UG.xlsx"
sheet = "Sheet1"   # change if needed

# --- OUTPUT FILE ---
output_file = "cleaned_no_empty_rows.xlsx"

# --- LOAD THE EXCEL FILE ---
df = pd.read_excel(input_file, sheet_name=sheet, engine="openpyxl")

# Normalize empty strings/whitespace to NA for accurate detection
df_norm = df.replace(r'^\s*$', pd.NA, regex=True)

# --- FOCUS COLUMN ---
# Set `target_column` to the header name you want to check.
# If None, the script checks full rows as before.
target_column = "country"  # e.g. "subscriptions_packages_id"

if target_column:
	if target_column not in df_norm.columns:
		raise KeyError(f"Column '{target_column}' not found in input. Available: {list(df_norm.columns)}")

	# Rows where the target column is empty (NA) and where it has a value
	rows_without_value = df_norm[df_norm[target_column].isna()].copy()
	rows_with_value = df_norm[df_norm[target_column].notna()].copy()

	if len(rows_without_value) > 0:
		print(f"Rows where column '{target_column}' is empty:")
		print(rows_without_value.to_string())
	else:
		print(f"No empty cells found in column '{target_column}'.")

	# --- SAVE BOTH SETS TO FILES ---
	file_with = f"{target_column}_with_value.xlsx"
	file_without = f"{target_column}_without_value.xlsx"
	rows_with_value.to_excel(file_with, index=False)
	rows_without_value.to_excel(file_without, index=False)
	print(f"Saved {len(rows_with_value)} rows where '{target_column}' has a value: {file_with}")
	print(f"Saved {len(rows_without_value)} rows where '{target_column}' is empty: {file_without}")

	# Keep cleaned file (rows where target column has a value)
	df_clean = rows_with_value.copy()
	removed_count = len(df) - len(df_clean)
	print(f"Removed {removed_count} rows where '{target_column}' was empty.")
else:
	# Rows that are completely empty (all columns NA)
	rows_completely_empty = df_norm[df_norm.isna().all(axis=1)]

	# Rows that have at least one empty cell
	rows_with_any_empty = df_norm[df_norm.isna().any(axis=1)]

	if len(rows_completely_empty) > 0:
		print("Rows completely empty:")
		print(rows_completely_empty.to_string(index=False))
	else:
		print("No completely empty rows found.")

	if len(rows_with_any_empty) > 0:
		print("\nRows with at least one empty cell:")
		print(rows_with_any_empty.to_string(index=False))
	else:
		print("No rows with empty cells found.")

	# --- REMOVE ROWS THAT ARE ENTIRELY EMPTY ---
	df_clean = df_norm.dropna(how="all")

# --- SAVE THE CLEANED FILE ---
df_clean.to_excel(output_file, index=False)

print(f"Original rows: {len(df)}")
print(f"Rows after removing empty rows: {len(df_clean)}")
print(f"Saved cleaned file as: {output_file}")
