import pandas as pd
import openpyxl

filepath = r"c:\Users\Administrator\Desktop\proyecto_tosi\excel\piezas.xlsm"
wb = openpyxl.load_workbook(filepath, read_only=True)
sheets = wb.sheetnames

for sheet in sheets:
    print(f"\n--- Sheet: {sheet} ---")
    try:
        # Read first 10 rows
        df = pd.read_excel(filepath, sheet_name=sheet, header=None, nrows=10)
        # Scan each row to see if "Ref" is in the values
        for idx, row in df.iterrows():
            valores = [str(v).strip() for v in row.values if pd.notna(v)]
            if any("Ref" in str(v) for v in valores):
                print(f"Row {idx} contains 'Ref': {valores}")
            if any("Designa" in str(v) for v in valores):
                print(f"Row {idx} contains 'Designa': {valores}")
    except Exception as e:
        print(f"Error reading sheet {sheet}: {e}")
