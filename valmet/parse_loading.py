from pathlib import Path
import sys
import re
import os
from collections import defaultdict

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
            raise ValueError(
                f"The provided path '{folder_path}' is not a valid directory."
            )
        return folder_path
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(100)


def extract_weight_and_dimensions(pdf_file):
    # Шаблон для поиска строки с информацией о каждой упаковке
    # 1 PCS  BOX  12.5 kg  50 x 40 x 30 cm
    package_pattern = re.compile(
        r"\d+\s+"
        r"(?P<unit>[A-Z]{3})\s+"
        r"(?P<weight>\d+[.,]\d+)\s*kg\s+"
        r"(?P<length>\d+)\s*x\s*(?P<width>\d+)\s*x\s*(?P<height>\d+)\s*cm",
        re.IGNORECASE,
    )
    # Шаблон для поиска общей информации о количестве упаковок
    total_qty_pattern = re.compile(r"Total\s+no\s+of\s+packages\s+(\d+)", re.IGNORECASE)
    # Шаблон для поиска общей информации о весе
    total_wt_pattern = re.compile(r"Gross\s+wt\.\s*kg\s+(\d+[.,]\d+)", re.IGNORECASE)

    items = []
    # Сгруппированные данные по уникальным комбинациям (unit, length, width, height)
    grouped = defaultdict(lambda: {"quantity": 0, "weight": 0.0})
    doc_total_qty = None
    doc_total_wt = None

    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    for match in package_pattern.finditer(full_text):
        unit = match.group("unit").upper()
        weight = float(match.group("weight").replace(",", "."))
        length = int(match.group("length"))
        width = int(match.group("width"))
        height = int(match.group("height"))

        items.append({
            "quantity": 1,
            "unit": unit,
            "weight": weight,
            "length": length,
            "width": width,
            "height": height,
        })

       
        # Группировка по уникальным комбинациям (unit, length, width, height) для подсчета количества и суммарного веса
        key = (unit, length, width, height)
        grouped[key]["quantity"] += 1
        grouped[key]["weight"] = round(grouped[key]["weight"] + weight, 2)

    # Итоговые значения количества и веса из документа
    qty_match = total_qty_pattern.search(full_text)
    wt_match = total_wt_pattern.search(full_text)
    if qty_match:
        doc_total_qty = int(qty_match.group(1))
    if wt_match:
        doc_total_wt = float(wt_match.group(1).replace(",", "."))

    summary_by_dimensions = [
        {
            "quantity": data["quantity"],
            "unit": unit,
            "length": length,
            "width": width,
            "height": height,
            "weight": data["weight"],
        }
        for (unit, length, width, height), data in grouped.items()
    ]

    parsed_qty = len(items)
    parsed_wt = round(sum(x["weight"] for x in items), 2)

    return {
        "items": items,
        "grouped": summary_by_dimensions,
        "validation": {
            "parsed_qty": parsed_qty,
            "doc_qty": doc_total_qty,
            "parsed_weight": parsed_wt,
            "doc_weight": doc_total_wt,
            "is_valid": (parsed_qty == doc_total_qty and parsed_wt == doc_total_wt),
        },
    }


def write_to_excel(output, filename="output.xlsx"):
    df = pd.DataFrame(output)

    total_rows = df.shape[0]

    total_quantity = df["quantity"].sum()
    total_weight = df["weight"].sum()

    df.index = range(1, len(df) + 1)
    df.index.name = "No."

    total_index = f"Total ({total_rows} rows)"

    df.loc[total_index] = {
        "quantity": total_quantity,
        "weight": total_weight,
    }

    cleaned_totals = df.fillna("")

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        cleaned_totals.to_excel(writer, sheet_name="Sheet1")

        worksheet = writer.sheets["Sheet1"]

        for column in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            letter = column[0].column_letter
            worksheet.column_dimensions[letter].width = max(max_length + 3, 10)


if __name__ == "__main__":
    folder = get_doc_folder()

    files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")
    ]

    print(f"Found {len(files)} PDF files in {folder}")

    pdf_file = folder / files[0]

    items = extract_weight_and_dimensions(pdf_file)['items']
    grouped = extract_weight_and_dimensions(pdf_file)['grouped']

    if not pdf_file:
        print(f"File not found in {folder}")
        sys.exit(0)

    write_to_excel(items, 'items.xlsx')
    write_to_excel(grouped, 'grouped.xlsx')

    print("Ready, please find xlsx files in folder")
