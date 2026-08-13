import re
from pathlib import Path

import pandas as pd
import openpyxl as xl

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
    # Load the Excel file
    # file_path = "path/to/your/excel/file.xlsx"
    # workbook = xl.load_workbook(file_path)
    
    # # Select the active sheet
    # sheet = workbook.active
    
    # # Read the data into a pandas DataFrame
    # data = sheet.values
    # columns = next(data)[0:]  # Get the first row as column names
    # df = pd.DataFrame(data, columns=columns)
    
    # # Example: Print the DataFrame
    # print(df)