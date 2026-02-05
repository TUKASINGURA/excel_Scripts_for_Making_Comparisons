import pandas as pd
import pycountry

# Function to get country name from country code
def get_country_name(code):
    try:
        return pycountry.countries.get(alpha_2=code).name
    except (AttributeError, LookupError):
        return code  # Return code if not found

# Configuration
input_file = 'channels.xlsx'  # Change this to your actual file
country_code_column = 'country'  # Specify which column contains country codes
output_file = 'output_with_country_names.xlsx'  # Change this to your desired output file

# Read the file
df = pd.read_excel(input_file)  # Using Excel reader

# Create a new column with country names
df['country_name'] = df[country_code_column].apply(get_country_name)

# Save to a new file
output_file = 'output_with_country_names.xlsx'  # Change this to your desired output file
df.to_excel(output_file, index=False)

print(f"Done! File saved as '{output_file}'")
print("\nFirst few rows:")
print(df.head(10))
