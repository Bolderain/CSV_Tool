"""
Unit tests for csv_tool_modern.py backend logic.
All test data is synthetic — no real customer data.

Run with:  C:\\ct\\.venv\\Scripts\\pytest.exe tests\\
"""

import sys
import os
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
import csv_tool_modern as ct

DATA = Path(__file__).parent / "data"

# ---------------------------------------------------------------------------
# Shared run configs (synthetic firmware values)
# ---------------------------------------------------------------------------

CFG_REP = {
    "type": "R310",
    "registrationStatus": "ACTIVATED",
    "desiredConfigurationTemplate": "CFG-REP.tar.gz",
    "desiredConfigurationMd5": "a" * 32,
    "desiredConfigurationSize": "195",
}

CFG_HE = {
    "type": "M300",
    "registrationStatus": "ACTIVATED",
    "desiredConfigurationTemplate": "CFG-HE.tar.gz",
    "desiredConfigurationMd5": "b" * 32,
    "desiredConfigurationSize": "215",
}

CFG_PROXIE = {
    "type": "P300",
    "registrationStatus": "ACTIVATED",
    "desiredConfigurationTemplate": "CFG-PROXY.tar.gz",
    "desiredConfigurationMd5": "c" * 32,
    "desiredConfigurationSize": "630",
}


# ---------------------------------------------------------------------------
# MAC helpers
# ---------------------------------------------------------------------------

class TestNormalizeMacColonsep:
    def test_plain_12hex(self):
        assert ct._normalize_mac_colonsep("000BC2100001") == "00:0B:C2:10:00:01"

    def test_already_colon_separated(self):
        assert ct._normalize_mac_colonsep("00:0B:C2:10:00:01") == "00:0B:C2:10:00:01"

    def test_dash_separated(self):
        assert ct._normalize_mac_colonsep("00-0B-C2-10-00-01") == "00:0B:C2:10:00:01"

    def test_lowercase_input(self):
        assert ct._normalize_mac_colonsep("000bc2100001") == "00:0B:C2:10:00:01"

    def test_too_short_raises(self):
        with pytest.raises(ValueError):
            ct._normalize_mac_colonsep("000BC2")

    def test_empty_raises(self):
        with pytest.raises(ValueError):
            ct._normalize_mac_colonsep("")


class TestAccessToken:
    def test_format(self):
        assert ct._access_token_from_mac("000BC2100001") == "00185803000bc2100001"

    def test_strips_separators(self):
        assert ct._access_token_from_mac("00:0B:C2:10:00:01") == "00185803000bc2100001"

    def test_prefix(self):
        token = ct._access_token_from_mac("000BC2300001")
        assert token.startswith("00185803")


# ---------------------------------------------------------------------------
# Headend transform — serial and MAC taken as-is (Jakub's workflow)
# ---------------------------------------------------------------------------

class TestHeadendTransform:
    def _factory(self, fields):
        return ct._headend_transform_factory(fields, CFG_HE, None)

    def test_serial_kept_as_is(self):
        f = self._factory(["SN", "MAC"])
        result = f({"SN": "C10625001001M", "MAC": "000BC2100001"}, 1)
        assert result["serialNumber"] == "C10625001001M"

    def test_mac_normalized_not_decremented(self):
        f = self._factory(["SN", "MAC"])
        result = f({"SN": "C10625001001M", "MAC": "000BC2100002"}, 1)
        # Must be :02, NOT :01 (old MAC-1 behaviour)
        assert result["macAddress"] == "00:0B:C2:10:00:02"

    def test_serial_not_transformed_c_to_b(self):
        f = self._factory(["SN", "MAC"])
        result = f({"SN": "C10625001001M", "MAC": "000BC2100001"}, 1)
        # Must start with C, NOT B (old C->B behaviour)
        assert result["serialNumber"].startswith("C")

    def test_non_c_serial_accepted(self):
        """Headend serials from Jakub can start with any character (e.g. ECRX...)."""
        f = self._factory(["SN", "MAC"])
        result = f({"SN": "ECRX3623000001", "MAC": "000BC2170001"}, 1)
        assert result["serialNumber"] == "ECRX3623000001"

    def test_device_type_and_status(self):
        f = self._factory(["SN", "MAC"])
        result = f({"SN": "C10625001001M", "MAC": "000BC2100001"}, 1)
        assert result["type"] == "M300"
        assert result["registrationStatus"] == "ACTIVATED"

    def test_empty_serial_raises(self):
        f = self._factory(["SN", "MAC"])
        with pytest.raises(ValueError, match="empty serialNumber"):
            f({"SN": "", "MAC": "000BC2100001"}, 1)

    def test_empty_mac_raises(self):
        f = self._factory(["SN", "MAC"])
        with pytest.raises(ValueError, match="empty macAddress"):
            f({"SN": "C10625001001M", "MAC": ""}, 1)


# ---------------------------------------------------------------------------
# Repeater transform
# ---------------------------------------------------------------------------

class TestRepeaterTransform:
    def _factory(self, fields):
        return ct._repeater_transform_factory(fields, CFG_REP, None)

    def test_basic(self):
        f = self._factory(["serialNumber", "macAddress"])
        result = f({"serialNumber": "C10625002001", "macAddress": "000BC2200001"}, 1)
        assert result["serialNumber"] == "C10625002001"
        assert result["macAddress"] == "000BC2200001"  # Repeater passes MAC through as-is
        assert result["type"] == "R310"

    def test_mac_passed_through_unchanged(self):
        f = self._factory(["serialNumber", "macAddress"])
        for mac_in in ["000BC2200001", "00:0B:C2:20:00:01", "00-0B-C2-20-00-01"]:
            result = f({"serialNumber": "C10625002001", "macAddress": mac_in}, 1)
            assert result["macAddress"] == mac_in


# ---------------------------------------------------------------------------
# Proxie transform
# ---------------------------------------------------------------------------

class TestProxieTransform:
    def _factory(self, fields):
        return ct._proxie_transform_factory(fields, CFG_PROXIE, None)

    def test_access_token_generated(self):
        f = self._factory(["SN", "bpl mac"])
        result = f({"SN": "C10625003001", "bpl mac": "000BC2300001"}, 1)
        assert result["serialNumber"] == "C10625003001"
        assert result["accessToken"] == "00185803000bc2300001"
        assert result["macAddress"] == "00:0B:C2:30:00:01"
        assert result["type"] == "P300"


# ---------------------------------------------------------------------------
# File-based: validate CSV files end-to-end
# ---------------------------------------------------------------------------

class TestValidateExport:
    def _run(self, filename, mode, cfg, sheet_name=None, no_header=False):
        result = ct._validate_export(
            infile=DATA / filename,
            mode=mode,
            run_cfg=cfg,
            mapping_override=None,
            sheet_name=sheet_name,
            no_header=no_header,
            log_fn=lambda msg: None,
        )
        return result

    # --- CSV ---

    def test_repeater_comma_csv(self):
        r = self._run("repeater_comma.csv", "Repeater", CFG_REP)
        assert r["ok_count"] == 3
        assert r["errors"] == []

    def test_repeater_semicolon_csv(self):
        r = self._run("repeater_semicolon.csv", "Repeater", CFG_REP)
        assert r["ok_count"] == 2
        assert r["errors"] == []

    def test_headend_direct_csv(self):
        r = self._run("headend_direct.csv", "Headend", CFG_HE)
        assert r["ok_count"] == 3
        assert r["errors"] == []

    def test_proxie_bpl_csv(self):
        r = self._run("proxie_bpl.csv", "Proxie", CFG_PROXIE)
        assert r["ok_count"] == 3
        assert r["errors"] == []

    def test_duplicate_serial_detected(self):
        r = self._run("duplicates.csv", "Repeater", CFG_REP)
        assert len(r["errors"]) == 1
        assert "Duplicate" in r["errors"][0][1]
        assert "C10625002001" in r["errors"][0][1]

    # --- Excel: single sheet ---

    def test_repeater_single_sheet_xlsx(self):
        r = self._run("repeater_single_sheet.xlsx", "Repeater", CFG_REP, sheet_name="Repeaters")
        assert r["ok_count"] == 3
        assert r["errors"] == []

    def test_repeater_empty_rows_skipped(self):
        r = self._run("repeater_empty_rows.xlsx", "Repeater", CFG_REP, sheet_name="Repeaters")
        assert r["ok_count"] == 3          # 3 data rows
        assert r["skipped_empty"] == 2     # 2 empty rows skipped
        assert r["errors"] == []

    def test_repeater_no_header_xlsx(self):
        # Columns are auto-named Col_A, Col_B when no header row exists
        r = ct._validate_export(
            infile=DATA / "repeater_no_header.xlsx",
            mode="Repeater",
            run_cfg=CFG_REP,
            mapping_override={"serial": "Col_A", "mac": "Col_B"},
            sheet_name="Sheet1",
            no_header=True,
            log_fn=lambda msg: None,
        )
        assert r["ok_count"] == 2
        assert r["errors"] == []

    # --- Excel: Jakub's multi-sheet format ---

    def test_jakub_headend_sheet(self):
        """Headend sheet: SN + MAC, taken as-is — no transformation."""
        r = self._run("jakub_multisheet.xlsx", "Headend", CFG_HE, sheet_name="Headend")
        assert r["ok_count"] == 2
        assert r["errors"] == []

    def test_jakub_proxie_sheet(self):
        """Proxie sheet: SN + bpl mac (+ eth mac ignored) — accessToken generated."""
        r = self._run("jakub_multisheet.xlsx", "Proxie", CFG_PROXIE, sheet_name="Proxies")
        assert r["ok_count"] == 3
        assert r["errors"] == []

    def test_jakub_wrong_sheet_falls_back_to_first(self):
        """If the requested sheet doesn't exist the tool falls back to the first sheet."""
        r = self._run("jakub_multisheet.xlsx", "Headend", CFG_HE, sheet_name="DoesNotExist")
        # Falls back to first sheet ("Headend") which has 2 rows
        assert r["ok_count"] == 2
