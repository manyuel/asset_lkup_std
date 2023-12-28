import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
import datetime

# Path to the folder containing the Excel files
excel_files_path = Path(r"C:\filepath.xlsx")

# Load the new master table from CSV
master_table = pd.read_csv(r"filepath.csv")

# Trimming whitespaces in master table
master_table['Legacy Attr Value'] = master_table['Legacy Attr Value'].str.strip()
master_table['Attribute Value'] = master_table['Attribute Value'].str.strip()

# Create a dictionary for easy lookup from the master table
master_dict = master_table.set_index('Legacy Attr Value')[['Attribute Value', 'Description']].to_dict('index')


# Function to replace Asset Codes worksheet with new master table
def replace_asset_codes_sheet(wb, master_df):
    if 'Asset Codes' in wb.sheetnames:
        ws = wb['Asset Codes']
        wb.remove(ws)
    ws = wb.create_sheet('Asset Codes')
    for r in dataframe_to_rows(master_df, index=False, header=True):
        ws.append(r)

# Function to update 'Master in Loc' tab
def update_master_in_loc_sheet(sheet, master_dict, skip_codes):
    for row in sheet.iter_rows(min_row=3, max_col=8): # min_row=3: Avoid headers because early templates have merged cells which can't be modified
        old_asset_code = (row[6].value or "").strip()
        if old_asset_code in master_dict and not any(code in old_asset_code for code in skip_codes):
            new_values = master_dict[old_asset_code] # Get the new values from the master dictionary
            row[6].value = new_values['Attribute Value']
            row[7].value = new_values['Description']
        else:
            row[6].value = (old_asset_code or "").replace('-', '_')

# Counters
success = 0
fail = 0

# Process each Excel file
for file in excel_files_path.glob('*.xlsx'):
    try:
        wb = load_workbook(file)
        sheet = wb['Master in Loc']
        # Set and modify skip_codes based on file name conditions
        
        skip_codes = ['CO', 'PA'] if 'MEX' in file.stem else []

        # Update 'Master in Loc' sheet
        update_master_in_loc_sheet(sheet, master_dict, skip_codes)

        # Replace 'Asset Codes' sheet with new master table
        replace_asset_codes_sheet(wb, master_table)

        # Save the workbook
        wb.save(file)

        success += 1
        print(f"{file.name} updated successfully.")
    except Exception as e:
        fail += 1
        print(f"Failed to process {file.name}: {e}.")

print(f"Process finished at {datetime.datetime.now().strftime('%H:%M - %d/%m/%Y')}")
print(f"{success} files were updated correctly, {fail} files failed to update.")
