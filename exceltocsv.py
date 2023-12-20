import pandas as pd
import os
from openpyxl import load_workbook

# Path to the folder containing Excel files
folder_path = "C:\\Users\\yahmadia\\OneDrive\\Desktop\\CEN100\\Shell03"

# List all files in the folder
files = os.listdir(folder_path)

# Filter Excel files
excel_files = [f for f in files if f.endswith('.xlsx')]

for file in excel_files:
    # Read Excel file
    wb = load_workbook(filename=os.path.join(folder_path, file))

    # Select the first sheet
    sheet = wb.active

    # Convert sheet to DataFrame
    data = sheet.values
    columns = next(data)
    df = pd.DataFrame(data, columns=columns)

    # Convert and save as CSV
    csv_file = os.path.splitext(file)[0] + '.csv'
    df.to_csv(os.path.join(folder_path, csv_file), index=False)
