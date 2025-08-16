# Excel Automation (Python)

A beginner-friendly Python automation project demonstrating how to **create, update, and read Excel spreadsheets** programmatically.
This three-step mini-project uses the `openpyxl` library for cross-platform Excel file handling and `pywin32` for optional auto-opening on Windows.

## Project Overview

1. **ExcelAuto-1.py** → Create a new Excel file with two cells (`A1` = "Hello", `B1` = "World").
2. **ExcelAuto-2.py** → Update or append additional values to the existing file.
3. **ExcelAuto-3.py** → Read the Excel file and display its contents in the terminal.

## Setup

To execute the scripts, open a terminal (or Command Prompt on Windows) in the project folder.

```bash
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
```

## Run

**Step 1:** Run `ExcelAuto-1.py` to create a new Excel file named `hello_world.xlsx` with “Hello” in cell A1 and “World” in cell B1.

```bash
python ExcelAuto-1.py
```

**Step 2:** Run `ExcelAuto-2.py` to open the existing Excel file and append or update additional cell values.

```bash
python ExcelAuto-2.py
```

**Step 3:** Run `ExcelAuto-3.py` to read the Excel file and print cell values to the terminal. This completes the **create → update → read** cycle.

```bash
python ExcelAuto-3.py
```

## Notes

* `pywin32` is Windows-only. On macOS/Linux, the Excel file is created but not auto-opened.
* Generated `.xlsx` files are ignored in `.gitignore` and will not be uploaded to GitHub.
