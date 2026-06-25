# CSV Tool

A Windows desktop tool for converting device lists (CSV / XLSX) into the import format required by the EniM device management platform.

Supports **Repeater**, **Headend**, and **Proxie** device types.

---

## Features

- Drag & drop CSV or Excel files onto the file list
- Four device modes: **Repeater**, **Headend**, **Proxie**, **1T**
- Auto-detects Serial and MAC columns; MAC address always normalized to colon-separated format (`AA:BB:CC:DD:EE:FF`)
- Multi-sheet Excel support with optional "no header row" mode
- Live output preview before export
- Saves firmware presets per customer (template, MD5, size)
- Saves column mappings per customer
- Duplicate serial number detection
- Validation summary with row counts and warnings
- **Headend: Calculate from Proxie** — toggle to derive headend values from a Proxie file; shown with an orange warning and reflected in the live preview immediately

---

## Getting Started

1. **Add files** — click *Add files* and pick one or more `.csv` or `.xlsx` device exports
2. **Check column mapping** — verify Serial and MAC columns in the *Input preview*; adjust via the dropdowns if needed
3. **Save mapping** *(recommended)* — enter a customer name (e.g. `Bayernwerk`) and click *Save mapping for customer*
4. **Pick a preset** — select a firmware preset for the target device type or fill in the fields directly
5. **Run** — click **Repeater**, **Headend**, **Proxie**, or **1T**

Output filename format:
```
Device_import_YYYYMMDD_CUSTOMERNAME_ROWCOUNT_TYPE.csv
```

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
