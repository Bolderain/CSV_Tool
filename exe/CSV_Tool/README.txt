CSV Tool (CSV_Tool.exe) – README
================================

Purpose
-------
This tool converts device lists (CSV or XLSX) into a strict “Device_import_…” CSV format for:
- Repeater (R310)
- Headend (M300)
- Proxie (P300)

Input can be:
- .csv (comma OR semicolon separated, auto-detected)
- .xlsx (first sheet is used)

Output is always:
- .csv comma-separated (delimiter = ,)
- strict columns depending on script mode


Installation / Distribution (Windows)
------------------------------------
You will receive a folder called:

  CSV_Tool\
    CSV_Tool.exe
    (many additional files/folders)

IMPORTANT:
- Keep ALL files in the CSV_Tool folder.
- Do NOT move only CSV_Tool.exe out of the folder.

How to run:
1) Copy the entire CSV_Tool folder to a writable location, e.g.
   - Desktop
   - Documents
   - a shared network folder (if you have write permissions)
2) Double-click CSV_Tool.exe

First start will create:
- CSV_Tool\presets\presets.json
- CSV_Tool\presets\mappings.json

These files store your saved configuration presets and customer column mappings.


Main Features
-------------
1) Input file list
   - Add files (.csv / .xlsx)
   - Remove selected
   - Clear list
   - Selecting a file automatically shows the Input Preview

2) Input Preview (first 5 rows)
   Shows:
   - detected file type (CSV/XLSX)
   - detected delimiter (CSV only: ',' or ';')
   - headers
   - first 5 rows (for quick validation)

3) Column Mapping (Serial / MAC)
   - The tool tries to auto-detect the serial and MAC column.
   - If it is not detected correctly, pick the correct columns using:
     - Serial column dropdown
     - MAC column dropdown
   - “Save mapping for customer” stores the chosen columns under the
     current “Filename input name (customer)”.

   Result:
   - Next time you run files for the same customer name, the tool will
     automatically reuse the saved mapping.

4) Desired Configuration (presets + editable)
   - You can select “Edit for”: Repeater / Headend / Proxie
   - Choose a preset (“DEFAULT” or custom preset)
   - Fields are editable:
     - desiredConfigurationTemplate
     - desiredConfigurationMd5
     - desiredConfigurationSize
   - If you change any field, the preset switches to “Custom”.
   - You can save the current values as a new preset:
     - enter a unique “New preset name”
     - click “Save preset”
     - the tool checks that the name does not already exist

5) Scripts (strict output generators)
   Buttons:
   - Repeater
   - Headend
   - Proxie

   Each script:
   - reads every selected input file
   - ignores fully empty rows (with a WARNING in the log)
   - generates a new comma-separated output CSV
   - output filename format:
     Device_import_YYYYMMDD_INPUTNAME_ROWCOUNT_TYPE.csv

   Notes about ROWCOUNT:
   - ROWCOUNT equals “number of data rows written” (header not included)

   Device-specific logic:
   A) Repeater (type R310)
      - serialNumber: copied from input Serial column
      - macAddress: copied from input MAC column
      - type: R310
      - registrationStatus: ACTIVATED
      - desiredConfiguration*: taken from the selected/editable preset

   B) Headend (type M300)
      - serialNumber: input serial with C… changed to B…
        (example: C1008... -> B1008...)
      - macAddress: input MAC minus 1
        (example: 00:0B:C2:19:98:EF -> 00:0B:C2:19:98:EE)
      - type: M300
      - registrationStatus: ACTIVATED
      - desiredConfiguration*: taken from the selected/editable preset

   C) Proxie (type P300)
      - serialNumber: copied from input Serial column
      - macAddress: copied from input MAC column
      - accessToken:
          00185803 + mac (no separators, lower-case)
        example:
          MAC input: 00:0B:C2:19:98:EF
          accessToken: 00185803000bc21998ef
      - type: P300
      - registrationStatus: ACTIVATED
      - desiredConfiguration*: taken from the selected/editable preset


Step-by-step Usage
------------------
1) Add input files
   - Click “Add files” and select one or more .csv or .xlsx files.
   - The newly added file is auto-selected and previewed.

2) Verify input preview + mapping
   - Check headers and the first 5 rows.
   - Verify Serial column and MAC column mapping.
   - If incorrect: select the correct columns from the dropdowns.
   - If you want to reuse it next time:
     - fill “Filename input name (customer)”
     - click “Save mapping for customer”

3) Set output filename inputs
   - “Filename date (YYYYMMDD)”
   - “Filename input name (customer)”

4) Select desired configuration preset
   - In “Desired configuration”, choose:
     - “Edit for” = the mode you plan to run
     - preset (DEFAULT or custom)
   - Optionally edit Template/MD5/Size
   - Optional: save a new preset name using “Save preset”

5) Run a script
   - Click “Repeater”, “Headend”, or “Proxie”
   - Output files are written to the selected output folder
   - Check the log for warnings/errors


Troubleshooting
---------------
- “Output already exists …”
  Delete/rename the existing output file or change date/customer name.

- “Could not find required columns (serial and mac)”
  Use the mapping dropdowns and/or save mapping for the customer.

- Empty rows in input CSV
  The tool skips fully empty rows and logs a WARNING with the row number.

- Presets not saving
  Ensure the CSV_Tool folder is in a writable location (not Program Files).


Files created by the tool
-------------------------
- presets\presets.json
  Stores saved desired configuration presets per mode.

- presets\mappings.json
  Stores saved Serial/MAC column mappings per customer name.