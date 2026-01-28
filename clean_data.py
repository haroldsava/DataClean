import pandas as pd
import re

# --- CONFIGURATION ---
INPUT_FILE = 'registry.xlsx'       # I've corrected the extension here to .xlsx
OUTPUT_FILE = 'registry_cleaned.xlsx'
COLUMN_NAME = 'NUMAR CONTRACT'     # Make sure this matches EXACTLY (Case sensitive)
# ---------------------

def clean_contract_number(value):
    text = str(value)
    # Regex: Finds Separator -> Date -> Everything after
    pattern = r'[\s\-/]*\d{2}\.\d{2}\.\d{4}.*'
    cleaned_text = re.sub(pattern, '', text)
    return cleaned_text.strip()

def main():
    try:
        print(f"Loading {INPUT_FILE}...")
        
        # Read all sheets
        all_sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
        sheets_processed = 0

        # Create a writer to save the results
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            
            for sheet_name, df in all_sheets.items():
                # Clean column names (remove hidden spaces around the names)
                df.columns = df.columns.str.strip()

                if COLUMN_NAME in df.columns:
                    print(f"Processing sheet: '{sheet_name}'...")
                    df[COLUMN_NAME] = df[COLUMN_NAME].apply(clean_contract_number)
                    sheets_processed += 1
                else:
                    print(f"Copying sheet: '{sheet_name}' (Target column not found)")
                
                # Save the sheet (whether modified or not) to the new file
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        if sheets_processed > 0:
            print(f"\nSuccess! Cleaned data saved to: {OUTPUT_FILE}")
        else:
            print(f"\nWarning: The column '{COLUMN_NAME}' was not found in any sheet.")
            print("Check that the column header in Excel looks exactly like the CONFIGURATION line.")

    except FileNotFoundError:
        print(f"Error: Could not find '{INPUT_FILE}'. Check the spelling and file extension.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()