from pathlib import Path
import os
import re

import pandas as pd
import pdfplumber

def get_file_folder():
    """
    Запрашивает у пользователя путь к папке, где находится PDF-файл, и проверяет его существование.
    Ввести название папки без слешей, например: "files" если папка находится в текущей директории, или полный путь к папке, например: "C:/Users/Username/Documents/files".

    Returns:
    
        Путь к папке в виде объекта Path, если путь существует и является директорией, иначе None.
    """
    try:
        user_input = input("Enter the folder path where the PDF file is located: ")
        if not user_input:
            raise ValueError("Folder path cannot be empty.")

        folder = Path(user_input)
        if not folder.exists() or not folder.is_dir():
            raise ValueError(f"The provided path does not exist or is not a directory {folder}.")

        return folder
        
    except Exception as e:
        print(f"Error: {e}")
        return None

def handle_invoice(pdf_path):
    """
    Обрабатывает коммерческий инвойс (Invoice) и извлекает данные.
    
    
    Args:
        pdf_path (Path): Путь к PDF-файлу инвойса.
        
        Returns:
            dict: Словарь с извлеченными данными.
    
    """
    INVOICE_NUMBER_PATT = re.compile(r'(CD2026[A-Z0-9-]+)', re.IGNORECASE)
    TOTAL_AMOUNT_PATT = re.compile(r'(?:Invoice\s*amount)\s*([\d\s.,-]*)', re.IGNORECASE)
    HS_CODE_PATT = re.compile(r'\b([\d]{8,10})\b', re.IGNORECASE)
    WEIGHT_PATT = re.compile(r'([\d.,]+)\s(?:kg)', re.IGNORECASE)

    page_text = ""

    data = {
            "Invoice number": None,
            "Total amount": None,
            "Items": []
        }
    # TODO: Инвойс с двумя таоженными кодами дублирует сумму инвойса, нужно убрать дублирование 
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text += page.extract_text(keep_blank_chars = True) or '-- TEXT NOT FOUND --'
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")

    invoice_found = INVOICE_NUMBER_PATT.search(page_text)
    total_found = TOTAL_AMOUNT_PATT.search(page_text)
    
    if invoice_found: data['Invoice number'] = invoice_found.group(1).strip()
    if total_found:
        total_string = total_found.group(1).strip()
        clean_string = re.sub(r'[^\d.-]','', total_string)
        data['Total amount'] = float(clean_string) if clean_string else 0.00
    
    for line in page_text.split('\n'):
        if re.match(r'\b[\d]{8,10}\b', line):
            hs = HS_CODE_PATT.search(line)
            weight = WEIGHT_PATT.search(line)
            weight_str = weight.group(1).strip() if weight else '0'
            clean_weight = re.sub(r'[^\d.]', '', weight_str)
            data['Items'].append({'HS': hs.group(1).strip() if hs else '-- HS code not found --', 'Weight': float(clean_weight)})
    
    return data

def handle_customs_invoice(file):
    """
    Обрабатывает таможенный инвойс (Customs Invoice) и извлекает данные.
    
    Args:
        file (Path): Путь к PDF-файлу таможенного инвойса.
        
    Returns:
        dict: Словарь с извлеченными данными.
    """
    ORDER_PATT = re.compile(r'([A-Z]\d{6})', re.IGNORECASE)
    TOTAL_PATT = re.compile(r'(?:TotalExclusiveTax)\s*([\d.,-]+)', re.IGNORECASE)
    HS_PATT = re.compile(r'\b([\d]{8,10})\b-(?:[A-Z,a-z ]+),', re.IGNORECASE)
    WEIGHT_PATT = re.compile(r',(\d+(?:[.,]\d+)*)\s*(?:KG|kg)', re.IGNORECASE)

    full_text = ''

    data = {
        "Invoice number": None,
        "Total amount": None,
        "Items": []
    }

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            full_text += page.extract_text(keep_blank_chars=True) or '-- Customs Invoice text not found --'

    order_found = ORDER_PATT.search(full_text)
    total_amount_found = TOTAL_PATT.search(full_text)

    if order_found: data['Invoice number'] = order_found.group(1).split()

    if total_amount_found:
        total_amount_string = total_amount_found.group(1).strip() if total_amount_found else '0'
        total_amount_clear = re.sub(r'[^[\d.-]]', '', total_amount_string)
        data['Total amount'] = float(total_amount_clear)

    for line in full_text.split('\n'):
        if re.match(HS_PATT, line):
            hs_code_found = HS_PATT.search(line)
            hs_code = hs_code_found.group(1).strip() if hs_code_found else '-- HS code not found --'
            
            weight_found = WEIGHT_PATT.search(line)
            weight_string = weight_found.group(1).strip() if weight_found else '0'
            weight_clean = re.sub(r'[^\d.]+', '', weight_string)

            data['Items'].append({
                'HS': hs_code, 'Weight': float(weight_clean)})
        
    return data

def write_to_excel(data, output_file):
    """
    Записывает данные в Excel-файл.
    
    Args:
        data (list): Список словарей с данными для записи.
        output_file (str): Имя выходного Excel-файла.
    """
    df = pd.DataFrame(data)

    columns_order = ['Invoice number', 'HS Code', 'Weight', 'Total amount']
    df = df[columns_order]

    df.loc['Total'] = df[['Weight', 'Total amount']].sum(numeric_only=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=True)
        worksheet = writer.sheets['Sheet1']

        number_format = '#,##0.00'

        total_amount_column = 4
        weight_column = 5

        for row in range(2, len(df) + 2):
            worksheet.cell(row=row, column=total_amount_column).number_format = number_format
            worksheet.cell(row=row, column=weight_column).number_format = number_format

    print(f'Data written to {output_file}')

if __name__ == "__main__":

    working_path = get_file_folder()
    result = []

    if working_path is None:
        print("No valid folder path provided. Exiting.")
        exit(1) 

    for file in os.listdir(working_path):
        found = working_path / file
        if found.exists() and found.is_file() and file.endswith('.pdf'):

            if file.startswith('CD2026'):
                fetched_data = handle_invoice(found)

            if file.startswith('Customs invoice'):
                fetched_data = handle_customs_invoice(found)
               
            for i, item in enumerate(fetched_data['Items']):
                result.append({
                    "Invoice number": fetched_data['Invoice number'],
                    "Total amount": fetched_data['Total amount'] if i == 0 else None,
                    "HS Code": item['HS'],
                    "Weight": item['Weight']
                })

            print(f"Data extracted from {found}: {fetched_data}")


    write_to_excel(result, 'result_data_script.xlsx')
    # df = pd.DataFrame(result)
    # df.loc['Total'] = df[['Weight', 'Total amount']].sum(numeric_only=True)
    # df.to_excel('result_data_script.xlsx', index=True)
    # print('Processing completed. Resulting DataFrame:')
    # print(df)