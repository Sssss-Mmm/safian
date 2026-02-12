import pandas as pd
import os

file_path = '2026통합발주서_영업_연습.xlsb'

try:
    # Read all sheet names
    xls = pd.ExcelFile(file_path, engine='pyxlsb')
    print(f"Sheet Names: {xls.sheet_names}")

    # Read the first few rows of each sheet to understand columns
    for sheet_name in xls.sheet_names:
        print(f"\n--- Sheet: {sheet_name} ---")
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', nrows=5)
        print("Columns:", df.columns.tolist())
        print(df.head())

except Exception as e:
    print(f"Error reading file: {e}")
