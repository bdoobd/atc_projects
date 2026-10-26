import re
from pathlib import Path

import pandas as pd
import openpyxl as xl

def parse_list(data_list):

    parse_data = []

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

        messy_string = row2[0]

        if not messy_string.startswith("CN Code:"):
            break

        # HS code, net and gross weight are extracted using regex patterns
        CN_regex = re.search(r'CN Code:\s*(\d+)', messy_string)
        Net_regex = re.search(r'Net weight\s\(g\):([\d\s]+)', messy_string)
        Gross_regex = re.search(r'Gross weight\s\(g\):([\d\s]+)', messy_string)

        code = CN_regex.group(1) if CN_regex else None
        net_str = Net_regex.group(1) if Net_regex else None
        gross_str = Gross_regex.group(1) if Gross_regex else None

        net_weight = float(re.sub(r'\s+', '', net_str))/1000 if net_str else 0
        gross_weight = float(re.sub(r'\s+', '', gross_str))/1000 if gross_str else 0

        parse_data.append({
            'HS Code': code,
            'Quantity': qty,
            'Net Weight (kg)': net_weight,
            'Gross Weight (kg)': gross_weight,
            'Amount': amount
        })

    return parse_data

def write_exel_talbe(data, output_file):
    df = pd.DataFrame(data)

    tmp_index =df.set_index('HS Code')
    grouped_df = tmp_index.groupby(tmp_index.index)
    compact_df = grouped_df.sum()

    compact_df.loc['Total'] = compact_df.sum(numeric_only=True)

    df.to_excel(f'table_{output_file}.xlsx', index=False)
    compact_df.to_excel(f"compact_{output_file}.xlsx")

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
    # parsed_data = []

    # # Get rows with the data starting from row 27 (index 26)
    rows = list(sheet.iter_rows(values_only=True))
    rows_with_data = rows[26:]

    ready_data = parse_list(rows_with_data)

    write_exel_talbe(ready_data, 'output')

    # Call the parse_list function
    # # Read the data into a pandas DataFrame
    # data = sheet.values
    # columns = next(data)[0:]  # Get the first row as column names
    # df = pd.DataFrame(data, columns=columns)
    
    # # Example: Print the DataFrame
    # print(df)