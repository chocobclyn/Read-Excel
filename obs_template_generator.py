import obspython as obs
import xlsxwriter
import os
import ctypes

def categorize_source_type(source_id):
    id_map = {
        "wasapi_process_output_capture": "",
        "wasapi_input_capture": "",
        "wasapi_output_capture": "",
        "browser_source": "Browser",
        "color_source_v3": "",
        "monitor_capture": "",
        "game_capture": "",
        "image_source": "Image",
        "slideshow_v2": "",
        "ffmpeg_source":"Media",
        "ndi_source": "",
        "source-clone": "",
        "text_gdiplus_v3": "Text",
        "dshow_input": "",
        "vlc_source": "",
        "window_capture": ""

    }
    return id_map.get(source_id, "")

def get_downloads_folder():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders") as key:
            downloads, _ = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')
            return downloads
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Downloads")

def show_message_box(message):
    ctypes.windll.user32.MessageBoxW(0, message, "OBS Read Excel - Generated Source List", 0x40)

def generate_template_from_scenes():
    downloads_dir = get_downloads_folder()
    output_file = os.path.join(downloads_dir, "obs_source_template.xlsx")

    scenes = obs.obs_frontend_get_scenes()
    data = []

    for scene_item in scenes:
        scene = obs.obs_scene_from_source(scene_item)
        if not scene:
            continue
        items = obs.obs_scene_enum_items(scene)
        for item in items:
            source = obs.obs_sceneitem_get_source(item)
            if source:
                source_id = obs.obs_source_get_id(source)
                source_name = obs.obs_source_get_name(source)
                source_type = categorize_source_type(source_id)
                data.append([source_type, source_name, ""])
        obs.sceneitem_list_release(items)
        obs.obs_source_release(scene_item)

    try:
        workbook = xlsxwriter.Workbook(output_file)
        worksheet = workbook.add_worksheet()
        worksheet.write_row(0, 0, ["Source Type", "Source Name", "Value"])
        for row_idx, row_data in enumerate(data, 1):
            worksheet.write_row(row_idx, 0, row_data)
        workbook.close()
        show_message_box(f"Source list saved to your Downloads folder:\n\n{output_file}\n\nPlease make sure that you have the appropriate\nsoftware (Excel) to view/modify the file")
    except Exception as e:
        show_message_box(f"Failed to save template: {e}")

def generate_template_button_clicked(props, prop):
    generate_template_from_scenes()
    return True
