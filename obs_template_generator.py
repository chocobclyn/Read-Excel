import obspython as obs
import os


def categorize_source_type(source_id):
    # prefix match so version suffixes (text_gdiplus_v2/v3, color_source_v3...) all resolve
    for prefix, label in (
        ("browser_source", "Browser"),
        ("image_source", "Image"),
        ("ffmpeg_source", "Media"),
        ("text_gdiplus", "Text"),
        ("text_ft2_source", "Text"),
        ("color_source", "Color"),
    ):
        if source_id.startswith(prefix):
            return label
    return ""


def get_downloads_folder():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
            downloads, _ = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')
            return downloads
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Downloads")


def collect_items(items, data, seen):
    for item in items:
        source = obs.obs_sceneitem_get_source(item)
        if not source:
            continue
        if obs.obs_sceneitem_is_group(item):
            group_items = obs.obs_sceneitem_group_enum_items(item)
            collect_items(group_items, data, seen)
            obs.sceneitem_list_release(group_items)
            continue
        source_id = obs.obs_source_get_id(source)
        source_name = obs.obs_source_get_name(source)
        print(f"[template] found '{source_name}' (id: {source_id})")
        if source_name not in seen:
            seen.add(source_name)
            # every source is listed (like older builds); unknown types get a blank
            # Type cell, which the main script skips safely
            data.append([categorize_source_type(source_id), source_name, "CHANGE_ME"])


def generate_template_from_scenes():
    """Scans the scene collection and writes a starter control sheet. Returns a status message."""
    from openpyxl import Workbook

    output_file = os.path.join(get_downloads_folder(), "obs_source_template.xlsx")

    scenes = obs.obs_frontend_get_scenes()
    data = []
    seen = set()
    scene_count = 0
    for scene_source in scenes:
        scene = obs.obs_scene_from_source(scene_source)
        if not scene:
            continue
        scene_count += 1
        items = obs.obs_scene_enum_items(scene)
        collect_items(items, data, seen)
        obs.sceneitem_list_release(items)
    obs.source_list_release(scenes)

    try:
        wb = Workbook()
        ws = wb.active
        ws.append(["Source Type", "Source Name", "Value"])
        for row in data:
            ws.append(row)
        wb.save(output_file)
        return f"Template saved ({len(data)} sources from {scene_count} scenes): {output_file}"
    except Exception as e:
        return f"Failed to save template: {e}"
