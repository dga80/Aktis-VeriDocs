import os
import json

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULT_CONFIG = {
    "excel_path": r"\\172.30.0.10\Oficina_tecnica\99-PEDIDOS EN CURSO AKTIS\LISTADO PEDIDOS EN CURSO AKTIS.xlsm",
    "excel_sheet": "Listado",
    "escaneos_folder": r"\\Mes-verificacio\escaneos",
    "destination_base_folder": r"\\Mes-verificacio\escaneos",
    "ocr_catalog_path": "codigosOCR.xlsx",
    "web_port": 5055,
    "web_host": "0.0.0.0",
    "ocr_enabled": True,
    "ocr_header_height_percent": 30,
    "ocr_footer_height_percent": 15,
    "ocr_lang": "es-ES",
    "auto_listen_on_startup": True
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                config = DEFAULT_CONFIG.copy()
                config.update(data)
                return config
        except Exception as e:
            print(f"Error loading config: {e}")
    return DEFAULT_CONFIG.copy()

def save_config(config_dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False
