import pandas as pd
import openpyxl

filepath = r"c:\Users\Administrator\Desktop\proyecto_tosi\excel\piezas.xlsm"
wb = openpyxl.load_workbook(filepath, read_only=True)
sheets = wb.sheetnames

with open("sheet_contents.txt", "w", encoding="utf-8") as f:
    for sheet in sheets:
        f.write(f"\n=========================================\n")
        f.write(f"SHEET: {sheet}\n")
        f.write(f"=========================================\n")
        try:
            df = pd.read_excel(filepath, sheet_name=sheet, header=None, nrows=15)
            f.write(df.to_string())
            f.write("\n")
        except Exception as e:
            f.write(f"Error: {e}\n")

print("Done. Saved to sheet_contents.txt")
