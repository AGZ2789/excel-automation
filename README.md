# Excel Automation (Python)

Creates an Excel file, updates cells, and (on Windows) opens it automatically using COM.

## Setup

```bash
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
```

## Run

```bash
python ExcelAuto.py
```

## Notes

* `pywin32` is Windows-only. On macOS/Linux, the Excel file is created but not auto-opened.
* Generated `.xlsx` files are ignored in `.gitignore` and will not be uploaded to GitHub.
