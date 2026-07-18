"""Sanity checks that run outside OBS: python test_sanity.py"""
import importlib.util
import os
import sys
import tempfile
import unittest.mock

obs_mock = unittest.mock.MagicMock()
sys.modules["obspython"] = obs_mock

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

# truthiness + setting-cell parsing
assert "TRUE".strip().lower() in ore.TRUTHY
assert "show" in ore.TRUTHY and "false" not in ore.TRUTHY
assert ore.parse_setting("Visible", "TRUE") is True
assert ore.parse_setting("Visible", "hide") is False
assert ore.parse_setting("Bounds Type", "1") == 1
assert ore.parse_setting("Scale", "0.5") == 0.5

# load_data: missing file returns None, never raises
ore.db_format = "excel"
ore.excel_file = os.path.join(tempfile.gettempdir(), "does_not_exist_12345.xlsx")
assert ore.load_data() is None

# load_data + update_obs_sources end to end on a real CSV (obs is a MagicMock)
with tempfile.TemporaryDirectory() as tmp:
    csv_path = os.path.join(tmp, "control.csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        f.write("Source Type,Source Name,Value,Visible,Pos X,Pos Y,Scale,Rotation,Bounds Type\n"
                "text,Title,Hello,,,,,,\n"
                ",,,,,,,,\n"
                "mystery,X,1,,,,,,\n"
                ",Logo,,FALSE,100,200,0.5,45,1\n"
                ",Bad,,,,oops,,,\n")
    ore.db_format = "csv"
    ore.excel_file = csv_path
    ore.source_type_column = "Source Type"
    ore.source_name_column = "Source Name"
    ore.value_column = "Value"
    ore.is_running = True
    ore.last_file_modified = 0
    ore.data_cache = None

    all_headers = ["Source Type", "Source Name", "Value",
                   "Visible", "Pos X", "Pos Y", "Scale", "Rotation", "Bounds Type"]
    headers, rows = ore.load_data()
    assert headers == all_headers
    assert rows[0]["Value"] == "Hello"

    # picker probe: csv reports one pseudo-sheet and the real headers
    assert ore.probe_file(csv_path, "csv", "") == (["CSV"], all_headers)
    assert ore.probe_file("no_such_file.csv", "csv", "") == ([], [])

    # one fake scene so apply_item_settings has something to iterate
    obs_mock.obs_frontend_get_scenes.return_value = [unittest.mock.MagicMock()]
    ore.update_obs_sources()
    # Title=content, Logo=settings-only -> updated; blank + mystery -> skipped; Bad Pos Y -> error
    assert "updated 2, skipped 2, errors 1" in ore.status_text, ore.status_text
    assert obs_mock.obs_sceneitem_set_visible.call_args[0][1] is False
    assert obs_mock.obs_sceneitem_set_rot.call_args[0][1] == 45.0
    assert obs_mock.obs_sceneitem_set_bounds_type.call_args[0][1] == 1

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
