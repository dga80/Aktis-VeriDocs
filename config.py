import os
import sys
import json

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(get_base_dir(), "config.json")

DEFAULT_CONFIG = {
    "excel_path": r"\\172.30.0.10\Oficina_tecnica\99-PEDIDOS EN CURSO AKTIS\LISTADO PEDIDOS EN CURSO AKTIS.xlsm",
    "excel_sheet": "Listado",
    "escaneos_folder": r"\\Mes-verificacio\escaneos",
    "destination_base_folder": r"\\Mes-verificacio\escaneos",
    "oficina_tecnica_base": r"\\172.30.0.10\Oficina_tecnica",
    "copy_to_oficina_tecnica": True,
    "ocr_catalog_path": "codigosOCR.xlsx",
    "web_port": 5055,
    "web_host": "0.0.0.0",
    "ocr_enabled": True,
    "ocr_header_height_percent": 30,
    "ocr_footer_height_percent": 15,
    "ocr_lang": "es-ES",
    "auto_listen_on_startup": True,
    "duplex_mode": True
}

def load_config():
    target_file = CONFIG_FILE
    if not os.path.exists(target_file):
        internal_file = os.path.join(get_base_dir(), "_internal", "config.json")
        if os.path.exists(internal_file):
            target_file = internal_file

    if os.path.exists(target_file):
        try:
            with open(target_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                config = DEFAULT_CONFIG.copy()
                config.update(data)
                return config
        except Exception as e:
            print(f"Error loading config: {e}")
    return DEFAULT_CONFIG.copy()

def save_config(config_dict):
    try:
        temp_file = CONFIG_FILE + ".tmp"
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, indent=4, ensure_ascii=False)
        # Atomic replace on Windows
        os.replace(temp_file, CONFIG_FILE)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False
