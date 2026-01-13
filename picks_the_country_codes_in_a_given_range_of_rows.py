import pandas as pd
import pycountry
import re
from collections import defaultdict

# --- INPUT FILE ---
input_file = "countries.xlsx"
sheet = "Sheet1"

# --- OUTPUT FILE ---
output_file = "available_country_codes_with_true_counts.xlsx"

# --- LOAD EXCEL FILE ---
df = pd.read_excel(input_file, sheet_name=sheet, engine="openpyxl")

# --- CLEAN EMPTY ROWS ---
df = df.dropna(how="all")
df = df.replace(r'^\s*$', pd.NA, regex=True).dropna(how="all")

# --- BUILD COUNTRY LOOKUP ---
country_lookup = {}

for c in pycountry.countries:
    country_lookup[c.name.lower()] = c.alpha_2
    country_lookup[c.alpha_2.lower()] = c.alpha_2

# Common aliases
country_lookup.update({
    "usa": "US",
    "u.s.a": "US",
    "uk": "GB",
    "u.k": "GB",
})

# --- STORAGE ---
country_counts = defaultdict(int)

# --- SCAN CELL BY CELL ---
for _, row in df.iterrows():
    for cell in row:
        if pd.isna(cell):
            continue

        text = str(cell).lower()

        for country_key, country_code in country_lookup.items():
            matches = re.findall(rf"\b{re.escape(country_key)}\b", text)
            if matches:
                country_name = pycountry.countries.get(alpha_2=country_code).name
                country_counts[(country_name, country_code)] += len(matches)

# --- OUTPUT DATAFRAME ---
output_df = pd.DataFrame([
    {
        "Country Name": name,
        "Country Code": code,
        "Count": count
    }
    for (name, code), count in country_counts.items()
]).sort_values("Count", ascending=False)

# --- SAVE ---
output_df.to_excel(output_file, index=False)

# --- SUMMARY ---
print(f"Total rows scanned: {len(df)}")
print(f"Unique countries found: {len(output_df)}")
print(f"Saved file: {output_file}")
