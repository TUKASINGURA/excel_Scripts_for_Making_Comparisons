
import pandas as pd
import os

# --- INPUT FILE ---
input_file = "most.csv"  # CSV file with semicolon separator
output_file = "customer_data_organized.xlsx"

# --- HELPER FUNCTION TO SAFELY CONVERT TO STRING ---
def safe_str(value):
    """Convert any value to string safely, handling NaN and empty values"""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()

# --- READ THE DATA ---
# Read CSV with semicolon separator and use first row as header
df = pd.read_csv(input_file, sep=";", dtype=str)

# Fill NaN values with empty strings
df = df.fillna("")

# Get the header and data info
print("Original columns:", df.columns.tolist())
print("Original shape:", df.shape)
print(f"\nHeaders detected: {list(df.columns)}")

# --- STRIP WHITESPACE FROM ALL VALUES ---
# Remove leading/trailing whitespace and quotes from all columns
for col in df.columns:
    df[col] = df[col].apply(lambda x: safe_str(x).strip('"'))

# --- SAVE TO EXCEL ---
df.to_excel(output_file, index=False, sheet_name='Sheet1')

print(f"\nData successfully organized and saved to: {output_file}")
print(f"Output shape: {df.shape}")
print("\nColumn names:")
print(df.columns.tolist())
print("\nFirst few rows:")
print(df.head())
