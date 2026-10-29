from pathlib import Path
import sys
import re

import pdfplumber
import pandas as pd
import openpyxl

def get_doc_folder():
    user_input = input("Enter the path to the folder containing PDF files: ")

    try:
        if not user_input:
            raise ValueError("No folder path provided.")
        
        folder_path = Path(user_input)
        if not folder_path.is_dir():
            raise ValueError(f"The provided path '{folder_path}' is not a valid directory.")
        return folder_path
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(100)

def extract_weight_and_dimensions(pdf_file):
    WEIGHT_HEADER_PATTERN = re.compile(r"Gross\s+wt\.\s+kg\s+Net\s+weight", re.IGNORECASE)
    DIMS_HEADER_PATTERN = re.compile(r"Length\s+Width\s+Height", re.IGNORECASE)

    WEIGHT_PATTERN = re.compile(r"([\d.]+)\s*kg\s+([\d.]+)\s*kg\s+([\d.]+)\s*m3", re.IGNORECASE)
    DIMS_PATTERN = re.compile(r"([\d.]+)\s*cm\s+([\d.]+)\s*cm\s+([\d.]+)\s*cm", re.IGNORECASE)

    output = {
        "length": 0.0,
        "width": 0.0,
        "height": 0.0,
        "net_weight": 0.0,
        "gross_weight": 0.0,
        "volume": 0.0
    }

    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            
            lines = text.split('\n');

            for i, line in enumerate(lines):
                if WEIGHT_HEADER_PATTERN.search(line):
                    if i + 1 < len(lines):
                        weight_line = lines[i + 1]
                        match = WEIGHT_PATTERN.search(weight_line)
                        if match:
                            output['gross_weight'] = float(match.group(1))
                            output['net_weight'] = float(match.group(2))
                            output['volume'] = float(match.group(3))
                        else:
                            print(f"No weight data found in '{pdf_file.name}' after the header line.")

                if DIMS_HEADER_PATTERN.search(line):
                    if i + 1 < len(lines):
                        dims_line = lines[i + 1]
                        match = DIMS_PATTERN.search(dims_line)
                        if match:
                            output['length'] = float(match.group(1))
                            output['width'] = float(match.group(2))
                            output['height'] = float(match.group(3))
                        else:
                            print(f"No dimension data found in '{pdf_file.name}' after the header line.")

    except Exception as e:
        print(f"Failed to process '{pdf_file.name}': {e}")

    return output

def write_to_excel(output, filename='output.xlsx'):
    df = pd.DataFrame(output)
    
    total_rows = df.shape[0]
    
    total_net = df['net_weight'].sum()
    total_gross = df['gross_weight'].sum()
    total_volume = df['volume'].sum()
    
    total_index = f'Total ({total_rows} rows)'
    
    df.loc[total_index] = {
            'net_weight': total_net,
            'gross_weight': total_gross,
            'volume': total_volume
    }
    
    cleaned_totals = df.fillna('')
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        cleaned_totals.to_excel(writer, sheet_name='Sheet1')
    
        worksheet = writer.sheets['Sheet1']
    
        for column in worksheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in column)
            letter = column[0].column_letter
            worksheet.column_dimensions[letter].width = max(max_length + 3, 10)
    
if __name__ == "__main__":
    folder = get_doc_folder()
    output = []

    for pdf_file in folder.glob("*.pdf"):
        pdf_output = extract_weight_and_dimensions(pdf_file)
        output.append(pdf_output)

    write_to_excel(output)

    