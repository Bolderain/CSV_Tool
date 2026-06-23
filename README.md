# CSV Tool

A Windows desktop tool for converting device lists (CSV / XLSX) into the import format required by the EniM device management platform.

Supports **Repeater**, **Headend**, and **Proxie** device types.

---

## Features

- Drag & drop CSV or Excel files onto the file list
- Auto-detects Serial and MAC columns
- Multi-sheet Excel support with optional "no header row" mode
- Live output preview before export
- Saves firmware presets per customer (template, MD5, size)
- Saves column mappings per customer
- Duplicate serial number detection
- Validation summary with row counts and warnings

---

## Running from source

**Requirements:** Python 3.10+

All other dependencies (PySide6, openpyxl) are installed automatically on first launch into a private environment at `C:\ct\.venv`.

```
python csv_tool_modern.py
```

---

## Running the .exe (no Python needed)

Download the `CSV_Tool` folder from the latest release and double-click `CSV_Tool.exe`.  
The `_internal` folder must stay next to the `.exe`.

---

## Building the .exe yourself

Install PyInstaller into your Python environment, then run:

```
pyinstaller CSV_Tool.spec
```

The output is placed in `dist\CSV_Tool\`.

---

## Presets

Firmware configuration presets are stored in `presets\presets.json`.  
Column mappings are stored in `presets\mappings.json`.  
Both files are created automatically on first run if they don't exist.
