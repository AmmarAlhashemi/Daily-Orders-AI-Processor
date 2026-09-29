import os
import subprocess

def hex_to_excel_color(hex_str: str) -> int:
    """Converts a standard hex color string (e.g. 'F4B084') to Excel RGB integer format."""
    hex_str = hex_str.lstrip('#')
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return r + (g * 256) + (b * 65536)

def open_in_explorer(file_path: str) -> None:
    """Opens File Explorer and selects the specified output file safely."""
    try:
        if os.path.isfile(file_path):
            norm_path = os.path.normpath(file_path)
            subprocess.Popen(["explorer", "/select,", norm_path])
    except Exception as ex:
        print(f"Failed to open File Explorer: {ex}")