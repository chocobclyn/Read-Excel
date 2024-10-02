import pandas as pd
import obspython as obs
from openpyxl import load_workbook
import time

excel_file = ""
worksheet = ""
source_type_column = ""
source_name_column = ""
value_column = ""
is_running = False  
reload_interval = 3  
last_update_time = 0  
worksheets = []
columns = []

def update_obs_sources():
    global last_update_time

    
    if excel_file and worksheet and is_running:
        
        if time.time() - last_update_time >= reload_interval:
            try:
                
                df = pd.read_excel(excel_file, sheet_name=worksheet, header=0, engine='openpyxl')

               
                for index, row in df.iterrows():
                    try:
                        
                        source_type = row[source_type_column]  
                        source_name = row[source_name_column]  
                        value = row[value_column]  

                        
                        if pd.notna(source_type) and isinstance(source_type, str):
                            if source_type.lower() == "text":
                                update_text_source(source_name, value)
                            elif source_type.lower() == "image":
                                update_media_source(source_name, value)
                    except KeyError:
                        
                        continue
                last_update_time = time.time()  
            except Exception as e:
                
                print(f"Error reading Excel file: {e}")


def update_text_source(source_name, new_text):
    source = obs.obs_get_source_by_name(source_name)
    if source:
        settings = obs.obs_data_create()
        obs.obs_data_set_string(settings, "text", str(new_text))
        obs.obs_source_update(source, settings)
        obs.obs_data_release(settings)
        obs.obs_source_release(source)


def update_media_source(source_name, file_path):
    source = obs.obs_get_source_by_name(source_name)
    if source:
        settings = obs.obs_data_create()
        obs.obs_data_set_string(settings, "file", file_path)  
        obs.obs_source_update(source, settings)
        obs.obs_data_release(settings)
        obs.obs_source_release(source)


def load_worksheets_and_columns():
    global worksheets, columns
    if excel_file:
        try:
            
            workbook = load_workbook(excel_file, read_only=True)
            worksheets = workbook.sheetnames

            
            if worksheet in worksheets:
                df = pd.read_excel(excel_file, sheet_name=worksheet, header=0, engine='openpyxl')
                columns = list(df.columns)  
        except Exception as e:
            print(f"Error loading workbook or worksheet: {e}")


def script_tick(seconds):
    update_obs_sources()  


def script_properties():
    props = obs.obs_properties_create()

    
    obs.obs_properties_add_path(props, "excel_file", "Excel File", obs.OBS_PATH_FILE, "*.xlsm", None)

    
    worksheet_list = obs.obs_properties_add_list(props, "worksheet", "Worksheet", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
    for sheet in worksheets:
        obs.obs_property_list_add_string(worksheet_list, sheet, sheet)

    
    source_type_column_list = obs.obs_properties_add_list(props, "source_type_column", "Source Type Column", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
    source_name_column_list = obs.obs_properties_add_list(props, "source_name_column", "Source Name Column", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
    value_column_list = obs.obs_properties_add_list(props, "value_column", "Value Column", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)

    for col in columns:
        obs.obs_property_list_add_string(source_type_column_list, col, col)
        obs.obs_property_list_add_string(source_name_column_list, col, col)
        obs.obs_property_list_add_string(value_column_list, col, col)

    
    reload_interval_list = obs.obs_properties_add_list(props, "reload_interval", "Reload Interval (seconds)", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_INT)
    obs.obs_property_list_add_int(reload_interval_list, "3 seconds", 3)
    obs.obs_property_list_add_int(reload_interval_list, "5 seconds", 5)
    obs.obs_property_list_add_int(reload_interval_list, "10 seconds", 10)

    
    obs.obs_properties_add_button(props, "toggle_script", "Start/Stop Script", toggle_script)

    return props


def script_update(settings):
    global excel_file, worksheet, source_type_column, source_name_column, value_column, reload_interval
    excel_file = obs.obs_data_get_string(settings, "excel_file")
    worksheet = obs.obs_data_get_string(settings, "worksheet")
    source_type_column = obs.obs_data_get_string(settings, "source_type_column")
    source_name_column = obs.obs_data_get_string(settings, "source_name_column")
    value_column = obs.obs_data_get_string(settings, "value_column")
    reload_interval = obs.obs_data_get_int(settings, "reload_interval")

    
    load_worksheets_and_columns()


def toggle_script(props, prop):
    global is_running
    is_running = not is_running 
    state = "Running" if is_running else "Stopped"
    print(f"Script is now: {state}")

