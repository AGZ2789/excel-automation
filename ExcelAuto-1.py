# ExcelAuto.py
from pathlib import Path
from openpyxl import Workbook
import platform

# 1) Create a workbook and write some cells
wb = Workbook()
sheet = wb.active
sheet["A1"] = "Hello"
sheet["B1"] = "World"

out_path = Path(__file__).with_name("hello_world.xlsx")
wb.save(out_path)

if platform.system() == "Windows":
    try:
        from win32com.client import Dispatch  # requires pywin32
        xl = Dispatch("Excel.Application")
        xl.Visible = True
        xl.Workbooks.Open(str(out_path.resolve()))
        # Tip: leave Excel open so you can see it; close manually when done.
    except Exception as e:
        print("Excel auto-open skipped:", e)

print(f"Created: {out_path.resolve()}")