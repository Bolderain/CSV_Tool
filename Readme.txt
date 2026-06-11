Start the tool

	Open a terminal in the folder and run:

	python csv_tool.py


	On first start, the tool auto-installs required dependencies (e.g., openpyxl) into a local virtual environment and restarts itself.

Add an input file

	Click Add files

	Select one or more .csv or .xlsx files

	Input can be comma , or semicolon ; separated (auto-detected)

	Output is always comma-separated ,

Check / fix column detection

	In Input preview (first 5 rows) and mapping:

	Verify Serial column and MAC column

	If the tool did not detect them correctly, pick them manually in the dropdowns

Save mapping for a customer (recommended)

	Enter a stable name into Filename input name (customer) (e.g. ESWE, Bayernwerk)

	Click Save mapping for customer

	Result:

	The tool remembers which input columns are Serial/MAC for that customer and reuses them next time automatically.

Set the output filename parts

	Filename date (YYYYMMDD): used in output name

	Filename input name (customer): used in output name and for saved mapping

	Output naming format:
	Device_import_YYYYMMDD_INPUTNAME_ROWCOUNT_TYPE.csv

Choose desired configuration (presets)

	In Desired configuration (presets + editable):

	Select Edit for: Repeater / Headend / Proxie

	Select a Preset (or edit fields directly)

	Optional: save your edited values as a new preset with Save preset

Run a transformation

	Click one of:

	Repeater

	Headend

	Proxie

The tool writes a new output CSV (comma-separated) into the output folder (default: same folder as csv_tool.py).