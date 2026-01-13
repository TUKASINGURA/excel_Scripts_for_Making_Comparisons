import pandas as pd
import pycountry
import re
from collections import defaultdict

# --- INPUT FILE ---
input_file = "sample.xlsx"
sheet = "Sheet1"   # change if needed

# --- OUTPUT FILE ---
output_file = "available_country_codes_with_counts.xlsx"

# --- LOAD THE EXCEL FILE (SAME FORMAT) ---
df = pd.read_excel(input_file, sheet_name=sheet, engine="openpyxl")

# --- REMOVE ROWS THAT ARE ENTIRELY EMPTY ---
df_clean = df.dropna(how="all")
df_clean = df_clean.replace(r'^\s*$', pd.NA, regex=True).dropna(how="all")

# --- PREPARE COUNTRY LOOKUP ---
countries = {c.name.lower(): c.alpha_2 for c in pycountry.countries}

# Common variations
countries.update({
    "usa": "US",
    "u.s.a": "US",
    "uk": "GB",
    "u.k": "GB",
})

# --- STORAGE FOR COUNTS ---
country_counts = defaultdict(int)

# --- SCAN ALL ROWS & COLUMNS ---
for _, row in df_clean.iterrows():
    row_text = " ".join(row.astype(str)).lower()

    for country_name, country_code in countries.items():
        if re.search(rf"\b{re.escape(country_name)}\b", row_text):
            country_counts[(country_name.title(), country_code)] += 1

# --- CREATE OUTPUT DATAFRAME ---
output_df = pd.DataFrame(
    [
        {
            "Country Name": name,
            "Country Code": code,
            "Count": count
        }
        for (name, code), count in country_counts.items()
    ]
).sort_values("Count", ascending=False)

# --- SAVE RESULT ---
output_df.to_excel(output_file, index=False)

# --- SUMMARY ---
print(f"Rows scanned: {len(df_clean)}")
print(f"Unique countries found: {len(output_df)}")
print(f"Saved country count file as: {output_file}")
