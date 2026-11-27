from pathlib import Path
import re

import pandas as pd
import pdfplumber

def get_doc_folder():
    """
    Запрашивает путь к папке с PDF-документами и возвращает список PDF-файлов.

    Путь может быть абсолютным или относительно текущей рабочей директории.
    """
    folder_path = Path(input("Введите путь к папке с PDF-документами: ").strip())

    if not folder_path.exists():
        print(f"Ошибка: папка не найдена: {folder_path}")
        return []
    if not folder_path.is_dir():
        print(f"Ошибка: указанный путь не является папкой: {folder_path}")
        return []

    pdf_files = sorted(
        (file_path for file_path in folder_path.iterdir()
         if file_path.is_file() and file_path.suffix.lower() == ".pdf"),
        key=lambda file_path: file_path.name.lower()
    )
    if not pdf_files:
        print(f"В папке нет PDF-документов: {folder_path}")

    return pdf_files

def extract_data(file_path):
    """
    Извлекает данные о грузовых местах и итогах из PDF packing list.

    Возвращает словарь с Shipment ID, списком Handling Units, Board Meter LDM
    (если он указан в документе) и общим Manual Gross Weight.
    """
    file_path = Path(file_path)
    with pdfplumber.open(file_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    shipment_match = re.search(r"\bShipment\s*ID\s*:?\s*(\d+)\b", text, re.IGNORECASE)
    if not shipment_match:
        raise ValueError(f"Shipment ID не найден в файле: {file_path}")

    number_pattern = r"(?:\d+(?:[.,]\d*)?|[.,]\d+)"
    row_pattern = re.compile(
        rf"^\s*(?P<unit_id>\d+)\s+(?P<type>.+?)\s+(?P=unit_id)\s+"
        rf"(?P<values>(?:{number_pattern}\s+){{3,}}{number_pattern})\s*$"
    )
    handling_units = []
    for line in text.splitlines():
        row_match = row_pattern.match(line)
        if not row_match:
            continue

        values = [
            float(value.replace(",", "."))
            for value in re.findall(number_pattern, row_match.group("values"))
        ]
        handling_units.append({
            "Handling Unit ID": row_match.group("unit_id"),
            "Type Description": row_match.group("type").strip(),
            "Width": values[0],
            "Depth": values[1],
            "Height": values[2],
            "Manual Gross Weight": values[-1],
        })

    if not handling_units:
        raise ValueError(f"Handling Unit данные не найдены в файле: {file_path}")

    totals_match = re.search(
        rf"Board\s+Meter\s+LDM\s+Manual\s+Gross\s+Weight\s+"
        rf"({number_pattern})\s+({number_pattern})",
        text,
        re.IGNORECASE,
    )
    if totals_match:
        board_meter_ldm = float(totals_match.group(1).replace(",", "."))
        total_gross_weight = float(totals_match.group(2).replace(",", "."))
    else:
        volume_totals_match = re.search(
            rf"Manual\s+Volume\s+Manual\s+Gross\s+Weight\s+"
            rf"({number_pattern})(?:\s+({number_pattern}))?",
            text,
            re.IGNORECASE,
        )
        if not volume_totals_match:
            raise ValueError(f"Итоги Manual Gross Weight не найдены: {file_path}")

        board_meter_ldm = None
        gross_weight_value = volume_totals_match.group(2) or volume_totals_match.group(1)
        total_gross_weight = float(gross_weight_value.replace(",", "."))

    return {
        "Shipment ID": shipment_match.group(1),
        "Handling Units": handling_units,
        "Board Meter LDM": board_meter_ldm,
        "Total Manual Gross Weight": total_gross_weight,
    }

def write_to_excel(data, output_path="output.xlsx"):
    """
    Записывает данные об отправках и грузовых местах в Excel-файл.

    Создаёт листы "Shipments" со сводными данными и "Handling Units" с
    отдельной строкой для каждого грузового места.
    """
    if isinstance(data, dict):
        data = [data]
    if not data:
        raise ValueError("Нет данных для записи в Excel.")

    shipments = []
    handling_units = []
    for shipment in data:
        shipments.append({
            "Shipment ID": shipment["Shipment ID"],
            "Board Meter LDM": shipment["Board Meter LDM"],
            "Total Manual Gross Weight": shipment["Total Manual Gross Weight"],
        })
        for unit in shipment["Handling Units"]:
            handling_units.append({
                "Shipment ID": shipment["Shipment ID"],
                "Handling Unit ID": unit["Handling Unit ID"],
                "Type Description": unit["Type Description"],
                "Width": unit["Width"],
                "Depth": unit["Depth"],
                "Height": unit["Height"],
                "Volume (m3)": unit["Width"] * unit["Depth"] * unit["Height"],
                "Manual Gross Weight": unit["Manual Gross Weight"],
            })

    output_path = Path(output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(shipments).to_excel(writer, sheet_name="Shipments", index=False)
        handling_units_df = pd.DataFrame(handling_units)
        handling_units_df.to_excel(writer, sheet_name="Handling Units", index=False)

        worksheet = writer.sheets["Handling Units"]
        summary_row = len(handling_units_df) + 3
        summary = [
            ("Summary", "Value"),
            ("Handling Unit Count", int(handling_units_df["Handling Unit ID"].count())),
            ("Total Volume (m3)", handling_units_df["Volume (m3)"].sum()),
            ("Total Manual Gross Weight", handling_units_df["Manual Gross Weight"].sum()),
        ]
        for row_offset, values in enumerate(summary):
            for column, value in enumerate(values, start=1):
                worksheet.cell(row=summary_row + row_offset, column=column, value=value)

    print(f"Excel file created: {output_path}")
    return output_path

if __name__ == "__main__":
    pdf_files = get_doc_folder()

    if not pdf_files:
        raise SystemExit("Нет PDF-файлов для обработки.")

    output = [extract_data(pdf_file) for pdf_file in pdf_files]
    write_to_excel(output, pdf_files[0].parent / "output.xlsx")