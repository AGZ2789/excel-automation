# Excel Automation (Python)

Creates an Excel file, updates cells, and (on Windows) opens it automatically using COM.

## Setup
To execute the scripts, open a terminal (or Command Prompt on Windows) in the project folder.

```bash
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
```

## Run

Step 1: Run ExcelAuto-1.py to create a new Excel file named hello_world.xlsx with “Hello” in cell A1 and “World” in cell B1.
```bash
python ExcelAuto1.py
```

Step 2: Run ExcelAuto-2.py to open the existing Excel file and append or update additional cell values.
```bash
python ExcelAuto2.py
```

## Notes

* `pywin32` is Windows-only. On macOS/Linux, the Excel file is created but not auto-opened.
* Generated `.xlsx` files are ignored in `.gitignore` and will not be uploaded to GitHub.
