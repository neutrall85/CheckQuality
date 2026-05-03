import os

def get_app_data_dir() -> str:
    return os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "ExcelReporter")

def ensure_dir(dir_path: str):
    os.makedirs(dir_path, exist_ok=True)