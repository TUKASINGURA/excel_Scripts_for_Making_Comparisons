import pandas as pd
import csv

# --- SCRIPT PURPOSE ---
# Count how many times each similar value appears in a specified column of a file.

# --- CONFIGURATION: MODIFY THESE VALUES ---
input_file = "filtered_by_date_original_text.xlsx"  # Change this to your input file (can be .csv or .xlsx)
sheet_name = "Sheet1"  # Used only for Excel files
column_to_search = "devices_type"  # Change this to the column you want to count
output_file = "value_counts_report.xlsx"  # Output Excel file name

# --- LOAD THE DATA ---
def detect_sep(path):
    """
    Try to detect CSV delimiter using csv.Sniffer on a sample of the file.
    If detection fails, return comma as a sensible default.
    """
    try:
        with open(path, "r", encoding="utf-8", newline="") as f:
            sample = f.read(4096)
            return csv.Sniffer().sniff(sample).delimiter
    except Exception:
        return ","

def read_table(path, sheet=None):
    """
    Read a file as Excel if it ends with an Excel extension,
    otherwise treat it as CSV and try to detect the separator.
    """
    lower = path.lower()
    if lower.endswith(('.xls', '.xlsx', '.xlsm')):
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    else:
        sep = detect_sep(path)
        df = pd.read_csv(path, sep=sep, engine="python", encoding="utf-8", skipinitialspace=True)
    # Normalize column names: remove BOM and trim whitespace
    df.columns = df.columns.astype(str).str.replace("\ufeff", "").str.strip()
    return df

# Load the input file
df = read_table(input_file, sheet=sheet_name)

# --- VERIFY COLUMN EXISTS ---
if column_to_search not in df.columns:
    print(f"\n❌ ERROR: Column '{column_to_search}' not found in the file!")
    print(f"\nAvailable columns are:")
    for i, col in enumerate(df.columns, 1):
        print(f"  {i}. {col}")
else:
    # Get total rows
    total_rows = len(df)
    
    # --- COUNT ALL UNIQUE VALUES IN THE COLUMN ---
    # Convert values to string and strip whitespace for accurate matching
    value_counts = df[column_to_search].astype(str).str.strip().value_counts()
    
    # --- CREATE RESULT DATAFRAME ---
    result_data = []
    for value, count_val in value_counts.items():
        percentage = (count_val / total_rows * 100)
        # Handle empty values
        display_value = "[EMPTY]" if value == "" else value
        result_data.append({
            'Value': display_value,
            'Count': count_val,
            'Percentage': f"{percentage:.2f}%"
        })
    
    result_df = pd.DataFrame(result_data)
    
    # --- PRINT RESULTS TO CONSOLE ---
    print("\n" + "="*70)
    print("VALUE COUNT REPORT")
    print("="*70)
    print(f"File:              {input_file}")
    print(f"Column:            {column_to_search}")
    print(f"Total rows:        {total_rows}")
    print(f"Unique values:     {len(value_counts)}")
    print("="*70)
    print(f"\nCount of each value in '{column_to_search}' column:\n")
    
    for idx, row in result_df.iterrows():
        print(f"  {row['Value']:30} : {row['Count']:5} occurrences ({row['Percentage']:>7})")
    
    print("\n" + "="*70)
    
    # --- SAVE TO EXCEL ---
    result_df.to_excel(output_file, index=False, sheet_name='Value Counts')
    print(f"\n✓ Results saved to: {output_file}\n")
