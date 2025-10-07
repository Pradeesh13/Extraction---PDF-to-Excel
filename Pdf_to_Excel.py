import pdfplumber
import pandas as pd
import configparser
import os

# Read input path from path.ini (full PDF file path)
path_ini_file = os.path.join("Info", "Config", "path.ini")
config = configparser.ConfigParser()
config.read(path_ini_file)

input_path = config.get('input', 'path')

input_pdf = input_path  # use as full PDF filepath, no join

output_excel = os.path.join("Info", "Data_Extracted", "Extracted.xlsx")

with pdfplumber.open(input_pdf) as pdf:
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for i, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if not tables:
                text = page.extract_text().split("\n")
                df = pd.DataFrame(text, columns=["Text"])
                df.to_excel(writer, sheet_name=f"Page_{i}", index=False)
            else:
                for j, table in enumerate(tables, start=1):
                    df = pd.DataFrame(table)
                    sheet_name = f"Page_{i}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\033[1mExtracted tables and data from PDF\033[0m")
