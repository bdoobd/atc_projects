import re
from pathlib import Path

import pandas as pd
import openpyxl as xl

def parse_list(data_list):
    for i in range(0, len(data_list), 3):

        # breakpoint()

        if i + 2 >= len(data_list):
            break 

        # line with quantity and amount values
        row1 = data_list[i]
        # messy string for parsing net weight, gross weíght and HS code
        row2 = data_list[i + 2]

        qty = int(row1[14]) if row1[14] else 0
        amount = float(row1[30]) if row1[30] else 0.0

        print(row2)
        # if isinstance(data_list[i], str):
        #     # Remove leading and trailing whitespace
        #     data_list[i] = data_list[i].strip()
        #     # Remove any non-alphanumeric characters (except spaces)
        #     data_list[i] = re.sub(r'[^a-zA-Z0-9\s]', '', data_list[i])

if __name__ == "__main__":
    try:
        table_file = input("Enter the path to the Excel file: ")
        if Path(table_file).is_file():
            print(f"File found: {table_file}")
        else:
            print("File not found.")
            exit()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        exit()

    print(f"Loading Excel file from: {table_file}")
    # # Load the Excel file
    # file_path = "path/to/your/excel/file.xlsx"
    workbook = xl.load_workbook(table_file)
    
    # # Select the active sheet
    sheet = workbook.active

    # # Define a list to store the parsed data
    parsed_data = []

    # # Get rows with the data starting from row 27 (index 26)
    rows = list(sheet.iter_rows(values_only=True))
    rows_with_data = rows[26:]

    parse_list(rows_with_data)

    # Call the parse_list function
    # # Read the data into a pandas DataFrame
    # data = sheet.values
    # columns = next(data)[0:]  # Get the first row as column names
    # df = pd.DataFrame(data, columns=columns)
    
    # # Example: Print the DataFrame
    # print(df)