import pandas as pd
import obspython as obs
from openpyxl import load_workbook
import time
import os
import threading
from obs_template_generator import generate_template_button_clicked

excel_file = ""
worksheet = ""
source_type_column = ""
source_name_column = ""
value_column = ""
db_format = "excel"
is_running = False
reload_interval = 3
last_update_time = 0
last_file_modified = 0
worksheets = ["N/A"]
columns = []
data_cache = None
update_thread = None
stop_event = threading.Event()


def script_description():
    """Returns the description shown in OBS"""
    return (
        "<b>Enhanced Excel/CSV Source Controller</b>"
        "<hr>"
        "Automatically updates OBS sources based on Excel/CSV file content. "
        "Supports text, image, media, and browser sources"
        "<br><br>"
        "Columns needed:"
        "<ul>"
        "<li>Source Type (text/image/media/browser)</li>"
        "<li>Source Name (as named in OBS)</li>"
        "<li>Value (What you want to set the source to)"
        "</ul>"
        "By John Paolo 'CHOCO!' Baclayon <br>"
        "This program is free to use and free to modify under the MIT License"
    )

def load_data():
    global last_file_modified
    try:
        if db_format == "excel" and os.path.isfile(excel_file):
            last_file_modified = os.path.getmtime(excel_file)
            return pd.read_excel(excel_file, sheet_name=worksheet, header=0, engine='openpyxl')
        elif db_format == "csv" and os.path.isfile(excel_file):
            last_file_modified = os.path.getmtime(excel_file)
            return pd.read_csv(excel_file)
        else:
            return pd.DataFrame()
    except Exception as e:
        print(f"Failed to load data: {e}")
        return pd.DataFrame()


def update_obs_sources():
    global last_update_time, last_file_modified, data_cache

    if not is_running:
        return

    try:
        if db_format in ["excel", "csv"] and os.path.isfile(excel_file):
            current_modified = os.path.getmtime(excel_file)
            if current_modified == last_file_modified and data_cache is not None:
                df = data_cache
            else:
                df = load_data()
                data_cache = df
        else:
            df = load_data()
            data_cache = df

        for index, row in df.iterrows():
            try:
                source_type = row[source_type_column]
                source_name = row[source_name_column]
                value = row[value_column]
                if pd.notna(source_type) and isinstance(source_type, str):
                    type_lower = source_type.lower()
                    if type_lower == "text":
                        update_text_source(source_name, value)
                    elif type_lower in ["image", "media"]:
                        update_media_source(source_name, value, type_lower)
                    elif type_lower == "browser":
                        update_browser_source(source_name, value)
            except Exception as e:
                print(f"Row {index}: Error - {e}")
                continue
        last_update_time = time.time()
    except Exception as e:
        print(f"Error reading data: {e}")


def threaded_updater():
    while not stop_event.wait(reload_interval):
        update_obs_sources()


def update_text_source(source_name, new_text):
    source = obs.obs_get_source_by_name(source_name)
    if source:
        settings = obs.obs_data_create()
        obs.obs_data_set_string(settings, "text", str(new_text))
        obs.obs_source_update(source, settings)
        obs.obs_data_release(settings)
        obs.obs_source_release(source)


def update_media_source(source_name, file_path, source_type):
    source = obs.obs_get_source_by_name(source_name)
    if source:
        settings = obs.obs_data_create()
        key = "file" if source_type == "image" else "local_file"
        obs.obs_data_set_string(settings, key, file_path)
        obs.obs_source_update(source, settings)
        obs.obs_data_release(settings)
        obs.obs_source_release(source)


def update_browser_source(source_name, url):
    source = obs.obs_get_source_by_name(source_name)
    if source:
        settings = obs.obs_data_create()
        obs.obs_data_set_string(settings, "url", url)
        obs.obs_source_update(source, settings)
        obs.obs_data_release(settings)
        obs.obs_source_release(source)


def script_properties():
    props = obs.obs_properties_create()

    db_list = obs.obs_properties_add_list(props, "db_format", "Data Source", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
    obs.obs_property_list_add_string(db_list, "Excel (.xlsm/.xlsx)", "excel")
    obs.obs_property_list_add_string(db_list, "CSV", "csv")

    obs.obs_properties_add_path(props, "excel_file", "File Path", obs.OBS_PATH_FILE, "*.xlsm;*.xlsx;*.csv", None)
    obs.obs_properties_add_text(props, "worksheet", "Sheet Name/Label", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_text(props, "source_type_column", "Source Type Column", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_text(props, "source_name_column", "Source Name Column", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_text(props, "value_column", "Value Column", obs.OBS_TEXT_DEFAULT)

    reload_interval_list = obs.obs_properties_add_list(props, "reload_interval", "Reload Interval (seconds)", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_INT)
    obs.obs_property_list_add_int(reload_interval_list, "3 seconds", 3)
    obs.obs_property_list_add_int(reload_interval_list, "5 seconds", 5)
    obs.obs_property_list_add_int(reload_interval_list, "10 seconds", 10)

    obs.obs_properties_add_button(props, "toggle_script", "Start/Stop Script", toggle_script)
    obs.obs_properties_add_button(props, "generate_template", "Generate Source Template", generate_template_button_clicked)

    return props


def script_update(settings):
    global excel_file, worksheet, source_type_column, source_name_column, value_column, reload_interval, db_format, google_sheet_id, google_service_account_json
    db_format = obs.obs_data_get_string(settings, "db_format")
    excel_file = obs.obs_data_get_string(settings, "excel_file")
    worksheet = obs.obs_data_get_string(settings, "worksheet")
    source_type_column = obs.obs_data_get_string(settings, "source_type_column")
    source_name_column = obs.obs_data_get_string(settings, "source_name_column")
    value_column = obs.obs_data_get_string(settings, "value_column")
    reload_interval = obs.obs_data_get_int(settings, "reload_interval")


def toggle_script(props, prop):
    global is_running, update_thread, stop_event
    is_running = not is_running
    if is_running:
        stop_event.clear()
        update_thread = threading.Thread(target=threaded_updater)
        update_thread.daemon = True
        update_thread.start()
        print("Script is now: Running")
    else:
        stop_event.set()
        if update_thread:
            update_thread.join()
        print("Script is now: Stopped")
