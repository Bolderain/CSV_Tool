# csv_tool.py
# PySide6 GUI version (replaces tkinter)
# Features:
# - Input: .csv (comma or semicolon) and .xlsx
# - Output: always comma-separated CSV
# - Modes: Repeater / Headend / Proxie
# - Presets stored in ./presets/presets.json (next to this script)
# - Customer mappings stored in ./presets/mappings.json
# - Auto preview on selection/add (first 5 rows), shows delimiter/sheet, headers, and detected mapping
# - Manual mapping dropdowns + "Save mapping for customer"
# - Skips empty rows and logs warnings (does not fail)
# - Self-bootstraps dependencies (PySide6, openpyxl) into a private venv in %LOCALAPPDATA%\csv_tool\.venv

# -----------------------------
# Dependency bootstrap
# -----------------------------
import os
import sys
import subprocess
import shutil
from pathlib import Path

REQUIRED_PACKAGES = [
    "openpyxl",
    "PySide6-Essentials==6.10.2",
]



def _run(cmd: list[str]) -> None:
    subprocess.check_call(cmd)


def _real(p: Path) -> Path:
    return Path(os.path.realpath(str(p)))


def _venv_dir() -> Path:
    # very short path to avoid MAX_PATH with PySide6 files
    # avoid Windows Store redirected LocalCache paths
    root = os.environ.get("SystemDrive", "C:")
    return _real(Path(root + r"\ct") / ".venv")


def _venv_python(venv_path: Path) -> Path:
    return venv_path / "Scripts" / "python.exe"


def _ensure_deps():
    try:
        import openpyxl  # noqa
        from PySide6 import QtWidgets  # noqa

        return
    except ModuleNotFoundError:
        pass

    venv_path = _real(_venv_dir())
    venv_path.parent.mkdir(parents=True, exist_ok=True)

    py = _venv_python(venv_path)
    cfg = venv_path / "pyvenv.cfg"

    # Recreate broken venv
    if not py.exists() or not cfg.exists():
        if venv_path.exists():
            shutil.rmtree(venv_path, ignore_errors=True)
        _run([sys.executable, "-m", "venv", str(venv_path)])

        venv_path = _real(venv_path)
        py = _venv_python(venv_path)
        cfg = venv_path / "pyvenv.cfg"

        if not py.exists() or not cfg.exists():
            raise RuntimeError(f"Venv creation failed: missing pyvenv.cfg at {cfg}")

    _run([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    _run([str(py), "-m", "pip", "install", *REQUIRED_PACKAGES])

    os.execv(str(py), [str(py)] + sys.argv)


_ensure_deps()

# -----------------------------
# Imports (after deps)
# -----------------------------
import csv
import json
import datetime as dt
import re

from openpyxl import load_workbook

from PySide6.QtCore import Qt, QSignalBlocker
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QGroupBox,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QLabel,
    QLineEdit,
    QComboBox,
    QPlainTextEdit,
    QScrollArea,
    QSizePolicy,
)


# -----------------------------
# Output schemas
# -----------------------------
REPEATER_OUTPUT_FIELDS = [
    "serialNumber",
    "macAddress",
    "type",
    "registrationStatus",
    "desiredConfigurationTemplate",
    "desiredConfigurationMd5",
    "desiredConfigurationSize",
]

HEADEND_OUTPUT_FIELDS = [
    "serialNumber",
    "macAddress",
    "type",
    "registrationStatus",
    "desiredConfigurationTemplate",
    "desiredConfigurationMd5",
    "desiredConfigurationSize",
]

PROXIE_OUTPUT_FIELDS = [
    "serialNumber",
    "macAddress",
    "type",
    "registrationStatus",
    "accessToken",
    "desiredConfigurationTemplate",
    "desiredConfigurationMd5",
    "desiredConfigurationSize",
]

MODE_FIXED = {
    "Repeater": {"type": "R310", "registrationStatus": "ACTIVATED"},
    "Headend": {"type": "M300", "registrationStatus": "ACTIVATED"},
    "Proxie": {"type": "P300", "registrationStatus": "ACTIVATED"},
}

ACCESS_TOKEN_PREFIX = "00185803"

SERIAL_HEADER_ALIASES = [
    "serialnumber",
    "serialno",
    "serialnr",
    "serialnum",
    "serial",
    "sn",
    "sno",
    "deviceserialnumber",
    "devserialnumber",
    "deviceid",
    "devicenumber",
    "seriennummer",
    "seriennr",
    "sernr",
    "idserial",
    "meterserial",
]

MAC_HEADER_ALIASES = [
    "macaddress",
    "macadress",
    "macaddr",
    "mac",
    "macid",
    "maceth",
    "hwaddress",
    "hardwareaddress",
    "ethernetaddress",
    "ethmac",
    "lanmac",
    "wlanmac",
    "eui",
    "eui64",
    "mac_eth",
    "macethernet",
]


# -----------------------------
# Paths: presets and mappings stored next to script
# -----------------------------
def _script_dir() -> Path:
    return Path(__file__).resolve().parent


def _preset_folder() -> Path:
    p = _script_dir() / "presets"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _presets_json_path() -> Path:
    return _preset_folder() / "presets.json"


def _mappings_json_path() -> Path:
    return _preset_folder() / "mappings.json"


# -----------------------------
# Helpers
# -----------------------------
def _norm_header(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").strip().lower())


def _pick_column(fieldnames: list[str], aliases: list[str], must_contain: list[str]) -> str | None:
    norm_map = {_norm_header(fn): fn for fn in fieldnames}
    for a in aliases:
        if a in norm_map:
            return norm_map[a]
    for fn in fieldnames:
        n = _norm_header(fn)
        if all(token in n for token in must_contain):
            return fn
    return None


def _sanitize_filename_part(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", s)
    s = s.strip().strip(".")
    return s or "input"


def _validate_yyyymmdd(s: str) -> bool:
    return bool(re.fullmatch(r"\d{8}", (s or "").strip()))


def _final_output_name(date_yyyymmdd: str, inputname: str, data_rows: int, device_type: str) -> str:
    date_part = _sanitize_filename_part(date_yyyymmdd)
    name_part = _sanitize_filename_part(inputname)
    type_part = _sanitize_filename_part(device_type)
    return f"Device_import_{date_part}_{name_part}_{data_rows}_{type_part}.csv"


def _validate_md5_hex32(s: str) -> bool:
    return bool(re.fullmatch(r"[0-9a-fA-F]{32}", (s or "").strip()))


def _validate_size_numeric(s: str) -> bool:
    return bool(re.fullmatch(r"\d+", (s or "").strip()))


def _detect_input_delimiter(file_obj) -> str:
    sample = file_obj.read(8192)
    file_obj.seek(0)

    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        if getattr(dialect, "delimiter", None) in (",", ";"):
            return dialect.delimiter
    except csv.Error:
        pass

    lines = sample.splitlines()
    lines = [l for l in lines if l.strip()][:10]
    text = "\n".join(lines)
    return ";" if text.count(";") > text.count(",") else ","


def _is_empty_row(row: dict) -> bool:
    return all(((v or "").strip() == "") for v in row.values())


def _serial_c_to_b(serial: str) -> str:
    s = (serial or "").strip()
    if not s:
        raise ValueError("Headend mode: serialNumber is empty.")
    if s[0] in ("C", "c"):
        return "B" + s[1:]
    if s[0] in ("B", "b"):
        return "B" + s[1:]
    raise ValueError(f"Headend mode: serialNumber does not start with C or B: '{s}'")


def _mac_minus_one(mac: str) -> str:
    raw = re.sub(r"[^0-9a-fA-F]", "", (mac or "").strip())
    if len(raw) != 12:
        raise ValueError(f"Headend mode: invalid MAC (expected 12 hex digits): '{mac}'")
    value = int(raw, 16)
    if value == 0:
        raise ValueError(f"Headend mode: MAC underflow on decrement: '{mac}'")
    value -= 1
    h = f"{value:012X}"
    return ":".join(h[i : i + 2] for i in range(0, 12, 2))


def _normalize_mac_12hex_lower(mac: str) -> str:
    raw = re.sub(r"[^0-9a-fA-F]", "", (mac or "").strip())
    if len(raw) != 12:
        raise ValueError(f"Proxie mode: invalid MAC (expected 12 hex digits): '{mac}'")
    return raw.lower()


def _access_token_from_mac(mac: str) -> str:
    return ACCESS_TOKEN_PREFIX + _normalize_mac_12hex_lower(mac)


def _build_run_cfg(mode: str, template: str, md5: str, size: str) -> dict:
    if mode not in MODE_FIXED:
        raise ValueError(f"Unknown mode: {mode}")

    template = (template or "").strip()
    md5 = (md5 or "").strip()
    size = (size or "").strip()

    if not template:
        raise ValueError("desiredConfigurationTemplate is empty.")
    if not md5:
        raise ValueError("desiredConfigurationMd5 is empty.")
    if not size:
        raise ValueError("desiredConfigurationSize is empty.")
    if not _validate_md5_hex32(md5):
        raise ValueError("desiredConfigurationMd5 must be 32 hex characters.")
    if not _validate_size_numeric(size):
        raise ValueError("desiredConfigurationSize must be numeric.")

    return {
        "type": MODE_FIXED[mode]["type"],
        "registrationStatus": MODE_FIXED[mode]["registrationStatus"],
        "desiredConfigurationTemplate": template,
        "desiredConfigurationMd5": md5.lower(),
        "desiredConfigurationSize": size,
    }


def _read_xlsx_headers_and_rows(path: Path, max_rows: int | None = None):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        header = next(it, None)
        if header is None:
            raise ValueError(f"No header found in file: {path}")
        fieldnames = [str(h).strip() if h is not None else "" for h in header]
        if all(h == "" for h in fieldnames):
            raise ValueError(f"Header row is empty in file: {path}")

        rows = []
        for excel_row_num, values in enumerate(it, start=2):
            row = {}
            for i, key in enumerate(fieldnames):
                v = values[i] if i < len(values) else None
                row[key] = "" if v is None else str(v).strip()
            rows.append((excel_row_num, row))
            if max_rows is not None and len(rows) >= max_rows:
                break

        return ws.title, fieldnames, rows
    finally:
        wb.close()


def _read_input_preview(infile: Path, max_rows: int = 5) -> dict:
    suffix = infile.suffix.lower()

    if suffix == ".csv":
        with infile.open("r", encoding="utf-8-sig", newline="") as fin:
            delim = _detect_input_delimiter(fin)
            reader = csv.DictReader(fin, delimiter=delim)
            if reader.fieldnames is None:
                raise ValueError(f"No header found in file: {infile}")
            headers = list(reader.fieldnames)
            rows = []
            for row_index, row in enumerate(reader, start=2):
                rows.append((row_index, row))
                if len(rows) >= max_rows:
                    break
            return {"type": "csv", "delimiter": delim, "sheet": "", "headers": headers, "rows": rows}

    if suffix == ".xlsx":
        sheet, headers, rows = _read_xlsx_headers_and_rows(infile, max_rows=max_rows)
        return {"type": "xlsx", "delimiter": "", "sheet": sheet, "headers": headers, "rows": rows}

    raise ValueError(f"Unsupported input type: {infile.suffix}")


def _iter_input_rows(infile: Path, log_fn):
    suffix = infile.suffix.lower()

    if suffix == ".csv":
        with infile.open("r", encoding="utf-8-sig", newline="") as fin:
            in_delim = _detect_input_delimiter(fin)
            log_fn(f"Detected input delimiter: '{in_delim}' for {infile.name}")
            reader = csv.DictReader(fin, delimiter=in_delim)
            if reader.fieldnames is None:
                raise ValueError(f"No header found in file: {infile}")
            fieldnames = list(reader.fieldnames)
            for row_index, row in enumerate(reader, start=2):
                yield row_index, fieldnames, row
        return

    if suffix == ".xlsx":
        wb = load_workbook(infile, read_only=True, data_only=True)
        try:
            ws = wb.active
            log_fn(f"Detected input type: xlsx (sheet '{ws.title}') for {infile.name}")
            it = ws.iter_rows(values_only=True)
            header = next(it, None)
            if header is None:
                raise ValueError(f"No header found in file: {infile}")
            fieldnames = [str(h).strip() if h is not None else "" for h in header]
            if all(h == "" for h in fieldnames):
                raise ValueError(f"Header row is empty in file: {infile}")

            for excel_row_num, values in enumerate(it, start=2):
                row = {}
                for i, key in enumerate(fieldnames):
                    v = values[i] if i < len(values) else None
                    row[key] = "" if v is None else str(v).strip()
                yield excel_row_num, fieldnames, row
        finally:
            wb.close()
        return

    raise ValueError(f"Unsupported input type: {infile.suffix}")


# -----------------------------
# Presets / Mappings storage
# -----------------------------
def _default_presets_payload() -> dict:
    return {
        "Repeater": {
            "DEFAULT": {
                "desiredConfigurationTemplate": "CFG-0001-0100-EBRO-20230223-FTPS-v1.tar.gz",
                "desiredConfigurationMd5": "f09ffd00e02bf9ad2108ea5199d46d2c",
                "desiredConfigurationSize": "195",
            }
        },
        "Headend": {
            "DEFAULT": {
                "desiredConfigurationTemplate": "CFG-0001-0100-HE-20230223-FTPS_AGC-v1.tar.gz",
                "desiredConfigurationMd5": "648be952a640f0d00f5da9f12854477f",
                "desiredConfigurationSize": "215",
            }
        },
        "Proxie": {
            "DEFAULT": {
                "desiredConfigurationTemplate": "CFG-0001-0100-PROXY-20230131-FTPS-v1.tar.gz",
                "desiredConfigurationMd5": "9768cbe43fa48cec894443e8c03ed399",
                "desiredConfigurationSize": "630",
            }
        },
    }


def _load_presets() -> dict:
    path = _presets_json_path()
    if not path.exists():
        payload = _default_presets_payload()
        with path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        return payload

    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    out = {"Repeater": {}, "Headend": {}, "Proxie": {}}
    for mode in out.keys():
        v = data.get(mode, {})
        if isinstance(v, dict):
            out[mode] = v
    return out


def _save_presets(presets: dict) -> None:
    path = _presets_json_path()
    payload = {
        "Repeater": presets.get("Repeater", {}),
        "Headend": presets.get("Headend", {}),
        "Proxie": presets.get("Proxie", {}),
    }
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)


def _load_mappings() -> dict:
    path = _mappings_json_path()
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_mappings(mappings: dict) -> None:
    path = _mappings_json_path()
    with path.open("w", encoding="utf-8") as f:
        json.dump(mappings, f, indent=2, ensure_ascii=False)


# -----------------------------
# Transform factories with mapping override
# -----------------------------
def _resolve_serial_mac_keys(input_fieldnames: list[str], mapping_override: dict | None) -> tuple[str | None, str | None, str]:
    if mapping_override:
        s = (mapping_override.get("serial") or "").strip()
        m = (mapping_override.get("mac") or "").strip()
        if s and m and s in input_fieldnames and m in input_fieldnames:
            return s, m, "manual"

    serial_key = _pick_column(input_fieldnames, SERIAL_HEADER_ALIASES, must_contain=["serial"])
    mac_key = _pick_column(input_fieldnames, MAC_HEADER_ALIASES, must_contain=["mac"])

    if serial_key and mac_key:
        return serial_key, mac_key, "auto"

    return serial_key, mac_key, "auto"


def _repeater_transform_factory(input_fieldnames: list[str], run_cfg: dict, mapping_override: dict | None):
    serial_key, mac_key, _src = _resolve_serial_mac_keys(input_fieldnames, mapping_override)
    if not serial_key or not mac_key:
        raise ValueError(
            "Repeater mode: could not find required columns (serial and mac). "
            "Use Mapping dropdowns and save mapping for the customer."
        )

    def transform(row: dict, row_index: int) -> dict:
        serial = (row.get(serial_key) or "").strip()
        mac = (row.get(mac_key) or "").strip()
        if not serial:
            raise ValueError(f"Repeater mode: empty serialNumber at row {row_index} (column: '{serial_key}')")
        if not mac:
            raise ValueError(f"Repeater mode: empty macAddress at row {row_index} (column: '{mac_key}')")

        return {
            "serialNumber": serial,
            "macAddress": mac,
            "type": run_cfg["type"],
            "registrationStatus": run_cfg["registrationStatus"],
            "desiredConfigurationTemplate": run_cfg["desiredConfigurationTemplate"],
            "desiredConfigurationMd5": run_cfg["desiredConfigurationMd5"],
            "desiredConfigurationSize": run_cfg["desiredConfigurationSize"],
        }

    return transform


def _headend_transform_factory(input_fieldnames: list[str], run_cfg: dict, mapping_override: dict | None):
    serial_key, mac_key, _src = _resolve_serial_mac_keys(input_fieldnames, mapping_override)
    if not serial_key or not mac_key:
        raise ValueError(
            "Headend mode: could not find required columns (serial and mac). "
            "Use Mapping dropdowns and save mapping for the customer."
        )

    def transform(row: dict, row_index: int) -> dict:
        in_serial = (row.get(serial_key) or "").strip()
        in_mac = (row.get(mac_key) or "").strip()
        if not in_serial:
            raise ValueError(f"Headend mode: empty serialNumber at row {row_index} (column: '{serial_key}')")
        if not in_mac:
            raise ValueError(f"Headend mode: empty macAddress at row {row_index} (column: '{mac_key}')")

        out_serial = _serial_c_to_b(in_serial)
        out_mac = _mac_minus_one(in_mac)

        return {
            "serialNumber": out_serial,
            "macAddress": out_mac,
            "type": run_cfg["type"],
            "registrationStatus": run_cfg["registrationStatus"],
            "desiredConfigurationTemplate": run_cfg["desiredConfigurationTemplate"],
            "desiredConfigurationMd5": run_cfg["desiredConfigurationMd5"],
            "desiredConfigurationSize": run_cfg["desiredConfigurationSize"],
        }

    return transform


def _proxie_transform_factory(input_fieldnames: list[str], run_cfg: dict, mapping_override: dict | None):
    serial_key, mac_key, _src = _resolve_serial_mac_keys(input_fieldnames, mapping_override)
    if not serial_key or not mac_key:
        raise ValueError(
            "Proxie mode: could not find required columns (serial and mac). "
            "Use Mapping dropdowns and save mapping for the customer."
        )

    def transform(row: dict, row_index: int) -> dict:
        serial = (row.get(serial_key) or "").strip()
        mac_in = (row.get(mac_key) or "").strip()
        if not serial:
            raise ValueError(f"Proxie mode: empty serialNumber at row {row_index} (column: '{serial_key}')")
        if not mac_in:
            raise ValueError(f"Proxie mode: empty macAddress at row {row_index} (column: '{mac_key}')")

        token = _access_token_from_mac(mac_in)

        return {
            "serialNumber": serial,
            "macAddress": mac_in.strip(),
            "type": run_cfg["type"],
            "registrationStatus": run_cfg["registrationStatus"],
            "accessToken": token,
            "desiredConfigurationTemplate": run_cfg["desiredConfigurationTemplate"],
            "desiredConfigurationMd5": run_cfg["desiredConfigurationMd5"],
            "desiredConfigurationSize": run_cfg["desiredConfigurationSize"],
        }

    return transform


# -----------------------------
# GUI
# -----------------------------
class CsvToolWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV Tool")

        # Start size: shows all features without resizing (still scrollable if user shrinks)
        self.resize(1120, 880)
        self.setMinimumSize(980, 780)


        self.script_dir = _script_dir()
        self.output_dir = self.script_dir

        self.files: list[Path] = []

        self.presets = _load_presets()
        self.mappings = _load_mappings()

        self.mode_edit = "Repeater"

        self._building = False
        self._applying_preset = False
        self._per_mode_selected_preset = {
            "Repeater": self._default_preset_name("Repeater"),
            "Headend": self._default_preset_name("Headend"),
            "Proxie": self._default_preset_name("Proxie"),
        }

        self._build_ui()
        self._log(f"Presets file: {_presets_json_path()}")
        self._log(f"Mappings file: {_mappings_json_path()}")

        # apply defaults
        self._set_mode_edit("Repeater")

    def _default_preset_name(self, mode: str) -> str:
        block = self.presets.get(mode, {})
        if isinstance(block, dict) and block:
            if "DEFAULT" in block:
                return "DEFAULT"
            return sorted(block.keys())[0]
        return "DEFAULT"

    def _build_ui(self):
        self._building = True

        central = QWidget()
        root_layout = QVBoxLayout(central)
        root_layout.setSpacing(10)
        root_layout.setContentsMargins(10, 10, 10, 10)

        # Scroll container (lets the whole form scroll if window is smaller)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        self.form_layout = QVBoxLayout(container)
        self.form_layout.setSpacing(10)
        self.form_layout.setContentsMargins(10, 10, 10, 10)
        scroll.setWidget(container)

        root_layout.addWidget(scroll)
        self.setCentralWidget(central)

        # -----------------------------
        # Input files group
        # -----------------------------
        gb_files = QGroupBox("Input files (.csv or .xlsx)")
        files_layout = QGridLayout(gb_files)
        files_layout.setColumnStretch(0, 1)

        self.file_list = QListWidget()
        self.file_list.setToolTip("Input files list. Selecting a file auto-previews it.")
        self.file_list.setFixedHeight(90)  # small, no huge empty area
        self.file_list.currentRowChanged.connect(self._on_file_selected)

        files_layout.addWidget(self.file_list, 0, 0, 4, 1)

        btn_add = QPushButton("Add files")
        btn_add.setToolTip("Add one or more input files (.csv or .xlsx). Newly added file is auto-selected and previewed.")
        btn_add.clicked.connect(self._add_files)

        btn_remove = QPushButton("Remove selected")
        btn_remove.setToolTip("Remove the selected file from the list.")
        btn_remove.clicked.connect(self._remove_selected)

        btn_clear = QPushButton("Clear list")
        btn_clear.setToolTip("Remove all files from the list.")
        btn_clear.clicked.connect(self._clear_list)

        btn_preview = QPushButton("Preview selected")
        btn_preview.setToolTip("Show the first 5 rows, detected delimiter/sheet, headers, and mapping.")
        btn_preview.clicked.connect(self._preview_selected)

        files_layout.addWidget(btn_add, 0, 1)
        files_layout.addWidget(btn_remove, 1, 1)
        files_layout.addWidget(btn_clear, 2, 1)
        files_layout.addWidget(btn_preview, 3, 1)

        self.form_layout.addWidget(gb_files)

        # -----------------------------
        # Output and filename group
        # -----------------------------
        gb_out = QGroupBox("Output")
        out_layout = QGridLayout(gb_out)
        out_layout.setColumnStretch(1, 1)

        out_layout.addWidget(QLabel("Output folder"), 0, 0)
        self.out_dir = QLineEdit(str(self.output_dir))
        self.out_dir.setReadOnly(True)
        out_layout.addWidget(self.out_dir, 0, 1)

        btn_choose = QPushButton("Choose")
        btn_choose.setToolTip("Select the output folder. Output is always comma-separated CSV.")
        btn_choose.clicked.connect(self._choose_output_dir)
        out_layout.addWidget(btn_choose, 0, 2)

        out_layout.addWidget(QLabel("Filename date (YYYYMMDD)"), 1, 0)
        self.date_edit = QLineEdit(dt.datetime.now().strftime("%Y%m%d"))
        self.date_edit.setToolTip("Date used in output filename.")
        self.date_edit.setMaximumWidth(120)
        out_layout.addWidget(self.date_edit, 1, 1, alignment=Qt.AlignLeft)

        out_layout.addWidget(QLabel("INPUTNAME_CUSTOMER"), 1, 1, alignment=Qt.AlignCenter)
        self.customer_edit = QLineEdit("")
        self.customer_edit.setToolTip("Customer name used in output filename and customer mapping key.")
        self.customer_edit.setMaximumWidth(260)
        out_layout.addWidget(self.customer_edit, 1, 2, alignment=Qt.AlignLeft)

        hint = QLabel("Output filename: Device_import_YYYYMMDD_INPUTNAME_CUSTOMER_ROWCOUNT_TYPE.csv")
        hint.setStyleSheet("color: gray;")
        out_layout.addWidget(hint, 2, 0, 1, 3)

        self.form_layout.addWidget(gb_out)

        # -----------------------------
        # Desired configuration group
        # -----------------------------
        gb_cfg = QGroupBox("Desired configuration (presets + editable)")
        cfg_layout = QGridLayout(gb_cfg)
        cfg_layout.setColumnStretch(5, 1)

        cfg_layout.addWidget(QLabel("Edit for"), 0, 0)

        self.edit_for = QComboBox()
        self.edit_for.addItems(["Repeater", "Headend", "Proxie"])
        self.edit_for.setToolTip("Select which device type you want to edit desired configuration for.")
        self.edit_for.currentTextChanged.connect(self._set_mode_edit)
        cfg_layout.addWidget(self.edit_for, 0, 1)

        cfg_layout.addWidget(QLabel("Preset"), 0, 2)
        self.preset_combo = QComboBox()
        self.preset_combo.setToolTip("Select a preset to fill Template/MD5/Size. Editing fields switches to Custom.")
        self.preset_combo.currentTextChanged.connect(self._on_preset_changed)
        cfg_layout.addWidget(self.preset_combo, 0, 3)

        cfg_layout.addWidget(QLabel("New preset name"), 0, 4)
        self.new_preset_name = QLineEdit("")
        self.new_preset_name.setToolTip("Name for a new preset. Must be unique within the selected mode.")
        cfg_layout.addWidget(self.new_preset_name, 0, 5)

        btn_save_preset = QPushButton("Save preset")
        btn_save_preset.setToolTip("Save the current Template/MD5/Size as a new preset for this mode.")
        btn_save_preset.clicked.connect(self._save_preset)
        cfg_layout.addWidget(btn_save_preset, 0, 6)

        cfg_layout.addWidget(QLabel("desiredConfigurationTemplate"), 1, 0)
        self.template_edit = QLineEdit("")
        self.template_edit.textChanged.connect(self._mark_custom_if_user_edit)
        cfg_layout.addWidget(self.template_edit, 1, 1, 1, 6)

        cfg_layout.addWidget(QLabel("desiredConfigurationMd5"), 2, 0)
        self.md5_edit = QLineEdit("")
        self.md5_edit.textChanged.connect(self._mark_custom_if_user_edit)
        cfg_layout.addWidget(self.md5_edit, 2, 1, 1, 6)

        cfg_layout.addWidget(QLabel("desiredConfigurationSize"), 3, 0)
        self.size_edit = QLineEdit("")
        self.size_edit.setMaximumWidth(120)
        self.size_edit.textChanged.connect(self._mark_custom_if_user_edit)
        cfg_layout.addWidget(self.size_edit, 3, 1, alignment=Qt.AlignLeft)

        self.form_layout.addWidget(gb_cfg)

        # -----------------------------
        # Preview + mapping group
        # -----------------------------
        gb_preview = QGroupBox("Input preview (first 5 rows) and mapping")
        pv_layout = QVBoxLayout(gb_preview)

        self.preview_info = QLabel("No file previewed.")
        self.preview_info.setStyleSheet("color: gray;")
        self.mapping_info = QLabel("Mapping: none")
        self.mapping_info.setStyleSheet("color: gray;")

        pv_layout.addWidget(self.preview_info)
        pv_layout.addWidget(self.mapping_info)

        row = QHBoxLayout()
        row.addWidget(QLabel("Serial column"))
        self.serial_combo = QComboBox()
        self.serial_combo.setToolTip("Select which input column contains the serial number.")
        self.serial_combo.setMinimumWidth(240)
        row.addWidget(self.serial_combo)

        row.addSpacing(10)
        row.addWidget(QLabel("MAC column"))
        self.mac_combo = QComboBox()
        self.mac_combo.setToolTip("Select which input column contains the MAC address.")
        self.mac_combo.setMinimumWidth(240)
        row.addWidget(self.mac_combo)

        row.addSpacing(10)
        btn_save_map = QPushButton("Save mapping for customer")
        btn_save_map.setToolTip("Save selected Serial/MAC columns for the current customer (Filename input name).")
        btn_save_map.clicked.connect(self._save_mapping_for_customer)
        row.addWidget(btn_save_map)

        row.addStretch(1)
        pv_layout.addLayout(row)

        self.preview_text = QPlainTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.preview_text.setFixedHeight(140)
        pv_layout.addWidget(self.preview_text)

        self.form_layout.addWidget(gb_preview)

        # -----------------------------
        # Scripts group
        # -----------------------------
        gb_scripts = QGroupBox("Scripts")
        sc_layout = QHBoxLayout(gb_scripts)

        btn_rep = QPushButton("Repeater")
        btn_rep.setToolTip("Generate output CSV for Repeaters (R310).")
        btn_rep.clicked.connect(lambda: self._run_mode("Repeater"))

        btn_he = QPushButton("Headend")
        btn_he.setToolTip("Generate output CSV for Headends (M300). Serial C->B, MAC minus 1.")
        btn_he.clicked.connect(lambda: self._run_mode("Headend"))

        btn_px = QPushButton("Proxie")
        btn_px.setToolTip("Generate output CSV for Proxies (P300). Adds accessToken = 00185803 + MAC (no separators, lower-case).")
        btn_px.clicked.connect(lambda: self._run_mode("Proxie"))

        self.status = QLabel("Ready.")
        self.status.setStyleSheet("color: gray;")

        sc_layout.addWidget(btn_rep)
        sc_layout.addWidget(btn_he)
        sc_layout.addWidget(btn_px)
        sc_layout.addSpacing(16)
        sc_layout.addWidget(self.status)
        sc_layout.addStretch(1)

        self.form_layout.addWidget(gb_scripts)

        # -----------------------------
        # Log group
        # -----------------------------
        gb_log = QGroupBox("Log")
        log_layout = QVBoxLayout(gb_log)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.log_text.setFixedHeight(160)
        log_layout.addWidget(self.log_text)

        self.form_layout.addWidget(gb_log)

        self._building = False

    # -----------------------------
    # Logging
    # -----------------------------
    def _log(self, msg: str) -> None:
        ts = dt.datetime.now().strftime("%H:%M:%S")
        self.log_text.appendPlainText(f"[{ts}] {msg}")

    # -----------------------------
    # File list actions
    # -----------------------------
    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select input files",
            "",
            "CSV or Excel (*.csv *.xlsx);;CSV files (*.csv);;Excel files (*.xlsx);;All files (*.*)",
        )
        if not paths:
            return

        first_new_index = None
        added = 0

        for p in paths:
            fp = Path(p)
            if fp not in self.files:
                self.files.append(fp)
                added += 1
                if first_new_index is None:
                    first_new_index = len(self.files) - 1

        self._refresh_file_list()
        self._log(f"Added {added} file(s). Total: {len(self.files)}.")

        if self.files and first_new_index is not None:
            self.file_list.setCurrentRow(first_new_index)
            self._preview_selected()

    def _remove_selected(self):
        idx = self.file_list.currentRow()
        if idx < 0 or idx >= len(self.files):
            return
        removed = self.files.pop(idx)
        self._refresh_file_list()
        self._log(f"Removed: {removed.name}. Total: {len(self.files)}.")

        if self.files:
            self.file_list.setCurrentRow(min(idx, len(self.files) - 1))
            self._preview_selected()
        else:
            self._clear_preview()

    def _clear_list(self):
        self.files.clear()
        self._refresh_file_list()
        self._log("Cleared file list.")
        self._clear_preview()

    def _refresh_file_list(self):
        with QSignalBlocker(self.file_list):
            self.file_list.clear()
            for f in self.files:
                self.file_list.addItem(QListWidgetItem(str(f)))

    def _on_file_selected(self, _row: int):
        # auto preview on selection
        if not self._building:
            self._preview_selected()

    def _selected_file(self) -> Path | None:
        idx = self.file_list.currentRow()
        if idx < 0 or idx >= len(self.files):
            return None
        return self.files[idx]

    def _clear_preview(self):
        self.preview_info.setText("No file previewed.")
        self.mapping_info.setText("Mapping: none")
        self.preview_text.setPlainText("")
        with QSignalBlocker(self.serial_combo):
            self.serial_combo.clear()
        with QSignalBlocker(self.mac_combo):
            self.mac_combo.clear()

    # -----------------------------
    # Output folder
    # -----------------------------
    def _choose_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select output folder", str(self.output_dir))
        if not d:
            return
        self.output_dir = Path(d)
        self.out_dir.setText(str(self.output_dir))
        self._log(f"Output folder set: {self.output_dir}")

    # -----------------------------
    # Desired config preset logic
    # -----------------------------
    def _set_mode_edit(self, mode: str):
        self.mode_edit = mode
        # populate presets list for this mode
        self.presets = _load_presets()
        names = sorted(self.presets.get(mode, {}).keys())
        if "Custom" not in names:
            names.append("Custom")

        with QSignalBlocker(self.preset_combo):
            self.preset_combo.clear()
            self.preset_combo.addItems(names)

            selected = self._per_mode_selected_preset.get(mode, "DEFAULT")
            if selected not in names:
                selected = "DEFAULT" if "DEFAULT" in names else (names[0] if names else "Custom")
            self.preset_combo.setCurrentText(selected)

        # apply selected preset to fields
        self._apply_preset(mode, self.preset_combo.currentText())

    def _apply_preset(self, mode: str, preset_name: str):
        if preset_name == "Custom":
            return
        block = self.presets.get(mode, {})
        values = block.get(preset_name, {})
        if not isinstance(values, dict):
            return

        self._applying_preset = True
        try:
            with QSignalBlocker(self.template_edit):
                self.template_edit.setText(str(values.get("desiredConfigurationTemplate", "")).strip())
            with QSignalBlocker(self.md5_edit):
                self.md5_edit.setText(str(values.get("desiredConfigurationMd5", "")).strip())
            with QSignalBlocker(self.size_edit):
                self.size_edit.setText(str(values.get("desiredConfigurationSize", "")).strip())
        finally:
            self._applying_preset = False

    def _on_preset_changed(self, preset: str):
        if self._building:
            return
        self._per_mode_selected_preset[self.mode_edit] = preset
        self._apply_preset(self.mode_edit, preset)

    def _mark_custom_if_user_edit(self, _text: str):
        if self._building or self._applying_preset:
            return
        # Switch current mode to Custom if user edits any config field
        if self.preset_combo.currentText() != "Custom":
            with QSignalBlocker(self.preset_combo):
                if self.preset_combo.findText("Custom") >= 0:
                    self.preset_combo.setCurrentText("Custom")
            self._per_mode_selected_preset[self.mode_edit] = "Custom"

    def _save_preset(self):
        mode = self.mode_edit
        name = (self.new_preset_name.text() or "").strip()

        if not name:
            QMessageBox.critical(self, "Preset name", "Preset name is empty.")
            return
        if name in ("DEFAULT", "Custom"):
            QMessageBox.critical(self, "Preset name", "Preset name cannot be DEFAULT or Custom.")
            return

        self.presets = _load_presets()
        if name in self.presets.get(mode, {}):
            QMessageBox.critical(self, "Preset name", f"Preset name already exists for {mode}. Choose another name.")
            return

        template = (self.template_edit.text() or "").strip()
        md5 = (self.md5_edit.text() or "").strip()
        size = (self.size_edit.text() or "").strip()

        if not template:
            QMessageBox.critical(self, "Preset values", "desiredConfigurationTemplate is empty.")
            return
        if not _validate_md5_hex32(md5):
            QMessageBox.critical(self, "Preset values", "desiredConfigurationMd5 must be 32 hex characters.")
            return
        if not _validate_size_numeric(size):
            QMessageBox.critical(self, "Preset values", "desiredConfigurationSize must be numeric.")
            return

        if mode not in self.presets:
            self.presets[mode] = {}
        self.presets[mode][name] = {
            "desiredConfigurationTemplate": template,
            "desiredConfigurationMd5": md5.lower(),
            "desiredConfigurationSize": size,
        }
        _save_presets(self.presets)

        self._log(f"Saved preset '{name}' for {mode}.")
        self.new_preset_name.setText("")
        self._set_mode_edit(mode)
        with QSignalBlocker(self.preset_combo):
            if self.preset_combo.findText(name) >= 0:
                self.preset_combo.setCurrentText(name)
        self._per_mode_selected_preset[mode] = name
        QMessageBox.information(self, "Preset saved", f"Saved preset '{name}' for {mode}.")

    # -----------------------------
    # Preview + mapping
    # -----------------------------
    def _customer_key(self) -> str:
        return (self.customer_edit.text() or "").strip()

    def _load_customer_mapping(self) -> dict | None:
        customer = self._customer_key()
        if not customer:
            return None
        self.mappings = _load_mappings()
        v = self.mappings.get(customer)
        if isinstance(v, dict):
            return {"serial": str(v.get("serial", "")).strip(), "mac": str(v.get("mac", "")).strip()}
        return None

    def _preview_selected(self):
        f = self._selected_file()
        if not f:
            self._clear_preview()
            return

        info = _read_input_preview(f, max_rows=5)
        headers = info["headers"]

        with QSignalBlocker(self.serial_combo):
            self.serial_combo.clear()
            self.serial_combo.addItems(headers)
        with QSignalBlocker(self.mac_combo):
            self.mac_combo.clear()
            self.mac_combo.addItems(headers)

        saved = self._load_customer_mapping()
        serial_auto = _pick_column(headers, SERIAL_HEADER_ALIASES, must_contain=["serial"])
        mac_auto = _pick_column(headers, MAC_HEADER_ALIASES, must_contain=["mac"])

        mapping_src = "none"
        serial_eff = ""
        mac_eff = ""

        if saved and saved.get("serial") in headers and saved.get("mac") in headers:
            mapping_src = "saved"
            serial_eff = saved.get("serial", "")
            mac_eff = saved.get("mac", "")
        elif serial_auto and mac_auto:
            mapping_src = "auto"
            serial_eff = serial_auto
            mac_eff = mac_auto

        if serial_eff:
            with QSignalBlocker(self.serial_combo):
                self.serial_combo.setCurrentText(serial_eff)
        if mac_eff:
            with QSignalBlocker(self.mac_combo):
                self.mac_combo.setCurrentText(mac_eff)

        if info["type"] == "csv":
            self.preview_info.setText(
                f"Preview: {f.name} (csv, delimiter '{info['delimiter']}'), headers: {len(headers)}"
            )
        else:
            self.preview_info.setText(
                f"Preview: {f.name} (xlsx, sheet '{info['sheet']}'), headers: {len(headers)}"
            )

        if mapping_src == "auto":
            self.mapping_info.setText(f"Mapping: auto (serial '{serial_eff}', mac '{mac_eff}')")
        elif mapping_src == "saved":
            self.mapping_info.setText(
                f"Mapping: saved for customer '{self._customer_key()}' (serial '{serial_eff}', mac '{mac_eff}')"
            )
        else:
            self.mapping_info.setText("Mapping: not detected (select columns manually)")

        # render preview text
        def trunc(v: str, n: int = 60) -> str:
            v = (v or "").strip()
            return v if len(v) <= n else v[: n - 3] + "..."

        lines = []
        lines.append("Headers:")
        lines.append("  " + " | ".join(headers))
        lines.append("")
        lines.append("Rows (first 5):")

        rows = info["rows"]
        if not rows:
            lines.append("  (no data rows)")
        else:
            for row_index, row in rows:
                parts = [trunc(str(row.get(h, ""))) for h in headers]
                lines.append(f"{row_index}: " + " | ".join(parts))

        self.preview_text.setPlainText("\n".join(lines))

    def _save_mapping_for_customer(self):
        customer = self._customer_key()
        if not customer:
            QMessageBox.critical(self, "Customer mapping", "Filename input name (customer) is empty.")
            return

        f = self._selected_file()
        if not f:
            QMessageBox.critical(self, "Customer mapping", "No file selected.")
            return

        info = _read_input_preview(f, max_rows=1)
        headers = info["headers"]

        serial_col = (self.serial_combo.currentText() or "").strip()
        mac_col = (self.mac_combo.currentText() or "").strip()

        if not serial_col or not mac_col:
            QMessageBox.critical(self, "Customer mapping", "Select both Serial column and MAC column.")
            return
        if serial_col not in headers or mac_col not in headers:
            QMessageBox.critical(self, "Customer mapping", "Selected columns are not in the current file headers.")
            return

        self.mappings = _load_mappings()
        if customer in self.mappings:
            res = QMessageBox.question(
                self,
                "Customer mapping",
                f"Mapping for '{customer}' already exists. Overwrite?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if res != QMessageBox.Yes:
                return

        self.mappings[customer] = {"serial": serial_col, "mac": mac_col}
        _save_mappings(self.mappings)

        self.mapping_info.setText(f"Mapping: saved for customer '{customer}' (serial '{serial_col}', mac '{mac_col}')")
        self._log(f"Saved mapping for customer '{customer}': serial='{serial_col}', mac='{mac_col}'")
        QMessageBox.information(self, "Customer mapping", f"Saved mapping for '{customer}'.")

    # -----------------------------
    # Run scripts
    # -----------------------------
    def _validate_common_inputs(self) -> tuple[str, str]:
        if not self.files:
            raise ValueError("No input files selected.")

        date_str = (self.date_edit.text() or "").strip()
        customer = (self.customer_edit.text() or "").strip()

        if not _validate_yyyymmdd(date_str):
            raise ValueError("Filename date must be 8 digits in YYYYMMDD format.")
        if not customer:
            raise ValueError("Filename input name (customer) is empty.")
        return date_str, customer

    def _mapping_override_for_run(self) -> dict | None:
        # If user selected something in the dropdowns, prefer that (even if not saved yet)
        serial_col = (self.serial_combo.currentText() or "").strip()
        mac_col = (self.mac_combo.currentText() or "").strip()
        if serial_col and mac_col:
            return {"serial": serial_col, "mac": mac_col}

        saved = self._load_customer_mapping()
        if saved and saved.get("serial") and saved.get("mac"):
            return saved
        return None

    def _run_mode(self, mode: str):
        try:
            date_str, customer = self._validate_common_inputs()

            run_cfg = _build_run_cfg(
                mode=mode,
                template=self.template_edit.text(),
                md5=self.md5_edit.text(),
                size=self.size_edit.text(),
            )

            mapping_override = self._mapping_override_for_run()

            self.status.setText(f"Running: {mode} ...")
            self._log(f"Run started. Mode: {mode}")

            time_tag = dt.datetime.now().strftime("%H%M%S")

            for i, infile in enumerate(self.files, start=1):
                tmp_name = f"__tmp__{date_str}_{time_tag}_{i}.csv"
                tmp_path = self.output_dir / tmp_name

                data_rows = self._process_one_file(infile, tmp_path, mode, run_cfg, mapping_override)

                final_name = _final_output_name(date_str, customer, data_rows, run_cfg["type"])
                final_path = self.output_dir / final_name

                if final_path.exists():
                    tmp_path.unlink(missing_ok=True)
                    raise FileExistsError(f"Output already exists: {final_path}")

                tmp_path.replace(final_path)
                self._log(f"Saved: {final_path.name}")

            self.status.setText("Done.")
            self._log(f"Run finished successfully. Mode: {mode}")
        except Exception as e:
            self.status.setText("Error.")
            self._log(f"ERROR: {e}")
            QMessageBox.critical(self, f"Error [{mode}]", str(e))

    def _process_one_file(self, infile: Path, out_path: Path, mode: str, run_cfg: dict, mapping_override: dict | None) -> int:
        self._log(f"Processing: {infile.name}")

        it = _iter_input_rows(infile, self._log)
        first = next(it, None)
        if first is None:
            raise ValueError(f"No data in file: {infile}")

        first_row_index, input_fields, first_row = first

        if mode == "Repeater":
            row_transform = _repeater_transform_factory(input_fields, run_cfg, mapping_override)
            out_fields = REPEATER_OUTPUT_FIELDS
        elif mode == "Headend":
            row_transform = _headend_transform_factory(input_fields, run_cfg, mapping_override)
            out_fields = HEADEND_OUTPUT_FIELDS
        elif mode == "Proxie":
            row_transform = _proxie_transform_factory(input_fields, run_cfg, mapping_override)
            out_fields = PROXIE_OUTPUT_FIELDS
        else:
            raise ValueError(f"Unknown mode: {mode}")

        with out_path.open("w", encoding="utf-8", newline="") as fout:
            writer = csv.DictWriter(fout, fieldnames=out_fields, delimiter=",")
            writer.writeheader()

            data_rows = 0
            skipped_empty = 0

            def handle(row_index: int, row: dict):
                nonlocal data_rows, skipped_empty
                if _is_empty_row(row):
                    skipped_empty += 1
                    self._log(f"WARNING: skipped empty row {row_index} in {infile.name}")
                    return
                out = row_transform(row, row_index)
                writer.writerow(out)
                data_rows += 1

            handle(first_row_index, first_row)
            for row_index, _fields, row in it:
                handle(row_index, row)

        self._log(f"OK: {infile.name} data rows written: {data_rows} (skipped empty: {skipped_empty})")
        return data_rows


def main():
    app = QApplication(sys.argv)
    # Force light mode (ignore system dark theme)
    # Force light mode via stylesheet (avoids QtGui imports)
    app.setStyle("Fusion")
    app.setStyleSheet("""
        QWidget { background: #f5f5f5; color: #000000; }
        QLineEdit, QPlainTextEdit, QListWidget, QComboBox {
            background: #ffffff;
            color: #000000;
            selection-background-color: #0078d7;
            selection-color: #ffffff;
        }
        QPushButton {
            background: #f0f0f0;
            color: #000000;
            border: 1px solid #c8c8c8;
            padding: 4px 10px;
            border-radius: 4px;
        }
        QPushButton:hover { background: #e8e8e8; }
        QGroupBox { border: 1px solid #c8c8c8; margin-top: 8px; }
        QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 3px; }
    """)


    w = CsvToolWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
