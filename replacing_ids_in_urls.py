import pandas as pd
import os

def generate_urls_from_excel(excel_file, id_column, output_file=None):
    """
    Generate URLs by replacing IDs in a URL template with IDs from an Excel file.
    
    Parameters:
    - excel_file: Path to the Excel file containing IDs
    - id_column: Column name or index containing the IDs
    - output_file: Optional file to save the generated URLs
    """
    
    # Base URL template - modify this with your URL pattern
    url_template = "https://website/application.com/front/vods/edit-vod/{id}?language=en"
 
    try:
        # Read the Excel fil
        df = pd.read_excel(excel_file)
        
        # Get the IDs from the specified column
        ids = df[id_column].astype(str).tolist()
        
        # Generate URLs
        urls = [url_template.format(id=id_value.strip()) for id_value in ids]
        
        print(f"Generated {len(urls)} URLs from {excel_file}\n")
        
        # Print URLs to console
        for i, url in enumerate(urls, 1):
            print(f"{i}. {url}")
        
        # Save to output file if specified
        if output_file:
            with open(output_file, 'w') as f:
                for url in urls:
                    f.write(url + '\n')
            print(f"\nURLs saved to: {output_file}")
        
        return urls
    
    except FileNotFoundError:
        print(f"Error: Excel file not found: {excel_file}")
        return []
    except KeyError:
        print(f"Error: Column '{id_column}' not found in Excel file")
        print(f"Available columns: {list(df.columns)}")
        return []
    except Exception as e:
        print(f"Error: {e}")
        return []


def main():
    """Main function - Modify these variables according to your needs"""
    
    # Configuration
    excel_file = "mw.xlsx"  # Change this to your Excel file name
    id_column = "MiddleWare"  # Change this to your column name containing IDs
    output_file = "generated_urls.txt"  # Output file for URLs
    
    # Check if the Excel file exists
    if not os.path.exists(excel_file):
        print(f"Please provide the Excel file: {excel_file}")
        print("\nExample usage:")
        print("1. Place your Excel file in the same directory as this script")
        print("2. Update the 'excel_file' variable with your file name")
        print("3. Update the 'id_column' variable with the column name containing IDs")
        print("4. Update the url_template in the function if needed")
        return
    
    # Generate URLs
    generate_urls_from_excel(excel_file, id_column, output_file)


if __name__ == "__main__":
    main()
