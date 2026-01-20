import pandas as pd
import re
import os

# --- INPUT FILE ---
input_file = "paid.csv"  # Can be .csv or .xlsx
output_file = "customer_data_organized.xlsx"

# --- PACKAGE TYPES TO SPLIT ---
package_types = [
    "AfroMobile Offer Package",
    "No Registration",
    "Month (Uganda)",
    "Disapora Monthly Package",
    "Next Media Staff"
]

# --- HELPER FUNCTION TO SAFELY CONVERT TO STRING ---
def safe_str(value):
    """Convert any value to string safely, handling NaN and empty values"""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()

# --- READ THE DATA ---
# Detect file type and read accordingly
if input_file.endswith('.xlsx'):
    df = pd.read_excel(input_file)
else:  # CSV
    df = pd.read_csv(input_file, sep=";")

# Get the header
print("Original columns:", df.columns.tolist())
print("Original shape:", df.shape)

# --- CREATE NEW DATAFRAME WITH REORGANIZED DATA ---
new_data = []

for _, row in df.iterrows():
    # Extract main fields - use safe_str to handle NaN values
    customer_id = safe_str(row.iloc[0])
    customer_login = safe_str(row.iloc[1])
    packages = safe_str(row.iloc[2])
    city = safe_str(row.iloc[3])
    country = safe_str(row.iloc[4])
    total_sessions = safe_str(row.iloc[5])
    latest_subscription = safe_str(row.iloc[6])
    
    # Remove quotes if present
    customer_id = customer_id.strip('"')
    customer_login = customer_login.strip('"')
    packages = packages.strip('"')
    city = city.strip('"')
    country = country.strip('"')
    latest_subscription = latest_subscription.strip('"')
    
    # Clean up packages field - split by tabs or multiple spaces
    package_list = re.split(r'\s{2,}|\t+', packages)
    package_list = [p.strip() for p in package_list if p.strip()]
    
    # Create a new row with split package columns
    new_row = {
        'customer_id': customer_id,
        'customer_login': customer_login,
        'city': city,
        'country': country,
        'total_sessions': total_sessions,
        'latest_subscription_started': latest_subscription,
    }
    
    # Add package columns
    for pkg_type in package_types:
        new_row[pkg_type] = 'Yes' if any(pkg_type in p for p in package_list) else ''
    
    new_data.append(new_row)

# --- CREATE NEW DATAFRAME ---
new_df = pd.DataFrame(new_data)

# --- SAVE TO EXCEL ---
new_df.to_excel(output_file, index=False, sheet_name='Sheet1')
print(f"\nData successfully organized and saved to: {output_file}")
print(f"Output shape: {new_df.shape}")
print("\nColumn names:")
print(new_df.columns.tolist())
print("\nFirst few rows:")
print(new_df.head())
