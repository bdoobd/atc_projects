from pathlib import Path
import os
import re

import pandas as pd
import pdfplumber

def get_file_folder():
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
    INVOICE_NUMBER_PATT = re.compile(r'(CD2026[A-Z0-9-]+)', re.IGNORECASE)
    TOTAL_AMOUNT_PATT = re.compile(r'(?:Invoice\s*amount)\s*([\d\s.,]*)', re.IGNORECASE)
    HS_CODE_PATT = re.compile(r'\b([\d]{8,10})\b', re.IGNORECASE)
    WEIGHT_PATT = re.compile(r'([\d.,]+)\s(?:kg)')

    page_text = ""

    data = {
            "Invoice number": None,
            "Total amount": None,
            "Items": []
        }

    invoice_found = INVOICE_NUMBER_PATT.search(page_text)
    total_found = TOTAL_AMOUNT_PATT.search(page_text)

    print(f"Invoice number found: {invoice_found.group(1).strip() if invoice_found else 'Not found'}")
    print(f"Total amount found: {total_found.group(1).strip() if total_found else 'Not found'}")
    
    if invoice_found: data['Invoice number'] = invoice_found.group(1).strip()
    if total_found:
        total_string = total_found.group(1).strip()
        clean_string = re.sub(r'[^\d.]','', total_string)
        data['Total amount'] = float(clean_string)
    
    for line in page_text.split('\n'):
        if re.match(r'\b[\d]{8,10}\b', line):
            hs = HS_CODE_PATT.search(line)
            weight = WEIGHT_PATT.search(line)
            weight_str = weight.group(1).strip()
            clean_weight = re.sub(r'[^\d.]', '', weight_str)
            data['Items'].append({'HS': hs.group(1).strip(), 'weight': float(clean_weight)})
    
    return data

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text += page.extract_text(keep_blank_chars=True) or ' -- text not found -- '

    # try:
    #     with pdfplumber.open(pdf_path) as pdf:
    #         first_page = pdf.pages[0]
    #         text = first_page.extract_text()
    #         print(f"Extracted text from {pdf_path}:\n{text}")
    # except Exception as e:
    #     print(f"Error processing {pdf_path}: {e}")

def handle_customs_invoice(pdf_path):
    pass
    # try:
    #     with pdfplumber.open(pdf_path) as pdf:
    #         first_page = pdf.pages[0]
    #         text = first_page.extract_text()
    #         print(f"Extracted text from {pdf_path}:\n{text}")
    # except Exception as e:
    #     print(f"Error processing {pdf_path}: {e}")

if __name__ == "__main__":

    file = 'CD202606800.pdf'

    working_path = get_file_folder()

    file = working_path / file

    if file.exists() and file.suffix.lower() == '.pdf':
        print(f"Processing file: {file}")
    else:
        print(f"File does not exist or is not a PDF: {file}")

        # pdf_path = working_path / file
        # result = handle_invoice(pdf_path)
        # print(f"Invoice data extracted from {file}: {result}")

    # result = handle_invoice(file)
    # print(f"Invoice data extracted from {file}: {result}")

    # for file in os.listdir(working_path):
    #     if file.endswith('.pdf'):
    #         if file.startswith('CD2026'):
    #             result =handle_invoice(file)
    #             print(f"Invoice data extracted from {file}: {result}")

    #         if file.startswith('Customs invoice'):
    #             print(f"Running Customs Invoice handler for file: {file}")
    #         # pdf_path = working_path / file
           

