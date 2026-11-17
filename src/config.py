import os
import json
import sys

DEFAULT_EXCEL_PATH = r"\\172.30.0.10\Oficina_tecnica\99-PEDIDOS EN CURSO AKTIS\LISTADO PEDIDOS EN CURSO AKTIS.xlsm"
DEFAULT_SCAN_FOLDER = r"\\Mes-verificacio\escaneos"
DEFAULT_DEST_FOLDER = r"\\Mes-verificacio\escaneos"

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CONFIG_FILE_PATH = os.path.join(get_base_dir(), "config.json")

def load_config():
    config = {
        "excel_path": DEFAULT_EXCEL_PATH,
        "scan_folder": DEFAULT_SCAN_FOLDER,
        "dest_folder": DEFAULT_DEST_FOLDER,
        "include_existing_scans": False
    }
    if os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                user_cfg = json.load(f)
                config.update(user_cfg)
        except Exception as e:
            print(f"Error cargando config.json: {e}")
    else:
        save_config(config)
    return config

def save_config(config_dict):
    try:
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardando config.json: {e}")
