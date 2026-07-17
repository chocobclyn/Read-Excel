"""Sanity checks that run outside OBS: python test_sanity.py"""
import importlib.util
import os
import sys
import tempfile
import unittest.mock

sys.modules["obspython"] = unittest.mock.MagicMock()

_spec = importlib.util.spec_from_file_location(
    "ore", os.path.join(os.path.dirname(os.path.abspath(__file__)), "read-excel.py"))
ore = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ore)

# dispatch map
assert ore.KEY_FOR_TYPE == {"text": "text", "image": "file", "media": "local_file", "browser": "url"}

# color: #FF8000 -> ABGR 0xFF0080FF
assert ore.parse_color("#FF8000") == 0xFF0080FF
assert ore.parse_color("ff8000") == 0xFF0080FF

# gsheet export URL builder
url = "https://docs.google.com/spreadsheets/d/ABC123-_x/edit?gid=55#gid=55"
assert ore.gsheet_export_url(url) == "https://docs.google.com/spreadsheets/d/ABC123-_x/export?format=csv&gid=55"
assert ore.gsheet_export_url("https://docs.google.com/spreadsheets/d/ABC/edit").endswith("gid=0")
assert ore.gsheet_export_url("not a sheet url") is None

# visibility truthiness
assert "TRUE".strip().lower() in ore.TRUTHY
assert "show" in ore.TRUTHY and "false" not in ore.TRUTHY

# load_data: missing file returns None, never raises
ore.db_format = "excel"
ore.excel_file = os.path.join(tempfile.gettempdir(), "does_not_exist_12345.xlsx")
assert ore.load_data() is None

# load_data + update_obs_sources end to end on a real CSV (obs is a MagicMock)
with tempfile.TemporaryDirectory() as tmp:
    csv_path = os.path.join(tmp, "control.csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        f.write("Source Type,Source Name,Value\n"
                "text,Title,Hello\n"
                ",,\n"
                "mystery,X,1\n")
    ore.db_format = "csv"
    ore.excel_file = csv_path
    ore.source_type_column = "Source Type"
    ore.source_name_column = "Source Name"
    ore.value_column = "Value"
    ore.is_running = True
    ore.last_file_modified = 0
    ore.data_cache = None

    headers, rows = ore.load_data()
    assert headers == ["Source Type", "Source Name", "Value"]
    assert rows[0]["Value"] == "Hello"

    # picker probe: csv reports one pseudo-sheet and the real headers
    assert ore.probe_file(csv_path, "csv", "") == (["CSV"], ["Source Type", "Source Name", "Value"])
    assert ore.probe_file("no_such_file.csv", "csv", "") == ([], [])

    ore.update_obs_sources()
    assert "updated 1, skipped 2, errors 0" in ore.status_text, ore.status_text

    # bad column name is reported, not silent
    ore.value_column = "Nope"
    ore.update_obs_sources()
    assert "Column 'Nope' not found" in ore.status_text
    ore.value_column = "Value"

    # xlsx round-trip (only if openpyxl is available in this environment)
    try:
        from openpyxl import Workbook
    except ImportError:
        print("openpyxl not installed here — skipped xlsx check")
    else:
        xlsx_path = os.path.join(tmp, "control.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["Source Type", "Source Name", "Value"])
        ws.append(["text", "Title", "Hi"])
        ws.append([None, None, None])
        wb.save(xlsx_path)
        ore.db_format = "excel"
        ore.excel_file = xlsx_path
        ore.worksheet = ""
        headers, rows = ore.load_data()
        assert headers == ["Source Type", "Source Name", "Value"]
        assert rows[0]["Value"] == "Hi"

        # picker probe: xlsx reports real sheet names and headers
        sheets, cols = ore.probe_file(xlsx_path, "excel", "")
        assert sheets == wb.sheetnames and cols == ["Source Type", "Source Name", "Value"]
        ore.data_cache = None
        ore.last_file_modified = 0
        ore.update_obs_sources()
        assert "updated 1, skipped 1, errors 0" in ore.status_text, ore.status_text

print("all sanity checks passed")
