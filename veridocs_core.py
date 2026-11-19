import os
import sys
import re
import time
import json
import shutil
import threading
import warnings
import unicodedata
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any, Callable
import openpyxl
import pypdfium2 as pdfium

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from config import load_config
from ocr_service import (
    is_ocr_available,
    is_page_blank,
    extract_of_from_header,
    extract_footer_text,
    match_footer_code,
    extract_footer_and_detect_code
)

# Catálogo predeterminado si no se encuentra codigosOCR.xlsx
DEFAULT_FALLBACK_CATALOG = [
    {"codigo": "DC-0018", "documento": "Carros Mat.", "is_incidencia": False},
    {"codigo": "DC-0019", "documento": "Orden de Fabricación", "is_incidencia": False},
    {"codigo": "DC-0020", "documento": "Control Estado (incidencias-Materiales)", "is_incidencia": True},
    {"codigo": "DC-0021", "documento": "Control de pendientes", "is_incidencia": False},
    {"codigo": "DC-0022", "documento": "Control Estado Entrada Verifica", "is_incidencia": True},
    {"codigo": "DC-0023", "documento": "Pesajes", "is_incidencia": False},
    {"codigo": "DC-0024", "documento": "Expediciones-Submer", "is_incidencia": False},
    {"codigo": "DC-0025", "documento": "D.Conformidad UE", "is_incidencia": False},
    {"codigo": "DC-0026", "documento": "Decl.Conformidad UL508A", "is_incidencia": False},
    {"codigo": "DC-0027", "documento": "Certificado Garantía MCD", "is_incidencia": False},
    {"codigo": "DC-0028", "documento": "H.Verificación IEC - (ES)", "is_incidencia": False},
    {"codigo": "DC-0029", "documento": "Controles Internos General", "is_incidencia": False},
    {"codigo": "DC-0030", "documento": "H.Verificación IEC - (EN)", "is_incidencia": False},
    {"codigo": "DC-0031", "documento": "H.Verificación UL - (EN)", "is_incidencia": False},
    {"codigo": "DC-0032", "documento": "Informe Técnico Ingeniero", "is_incidencia": False},
    {"codigo": "DC-0033", "documento": "Informe Técnico OBRA", "is_incidencia": False},
    {"codigo": "DC-0034", "documento": "Pruebas FAT_(ES)", "is_incidencia": False},
    {"codigo": "DC-0035", "documento": "Pruebas FAT_(EN)", "is_incidencia": False},
    {"codigo": "DC-0036", "documento": "Pruebas FAT _Cliente_(ES)", "is_incidencia": False},
    {"codigo": "DC-0037", "documento": "Pruebas FAT _Cliente_(EN)", "is_incidencia": False},
    {"codigo": "DC-0038", "documento": "Hoja Almacén", "is_incidencia": False},
    {"codigo": "DC-0039", "documento": "Hoja Revisión Cubículos", "is_incidencia": False},
    {"codigo": "ANEXO 1_IT0001", "documento": "H.Tareas de Verificación", "is_incidencia": False},
    {"codigo": "ANEXO 2_IT0001", "documento": "H.Check MCD", "is_incidencia": False}
]

# Compatibilidad hacia atrás
DOCUMENTS_DEF = DEFAULT_FALLBACK_CATALOG


def get_effective_documents_def():
    return load_ocr_catalog()


_cached_catalog = None
_catalog_mtime = 0
_catalog_lock = threading.Lock()


def sanitize_filename(text: Any) -> str:
    """Limpia caracteres inválidos para nombres de archivos y carpetas en Windows."""
    if text is None:
        return ""
    text = str(text).strip()
    text = re.sub(r'[\\/:*?"<>|]', '-', text)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'-+', '-', text)
    return text.strip(' -.')


def load_ocr_catalog(catalog_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Carga todos los tipos de documentos posibles desde codigosOCR.xlsx.
    Identifica de forma dinámica el código, el nombre y si es tipo incidencia.
    Recarga automáticamente si el archivo fue modificado.
    """
    global _cached_catalog, _catalog_mtime

    cfg = load_config()
    target_path = catalog_path or cfg.get("ocr_catalog_path", "codigosOCR.xlsx")

    if not os.path.isabs(target_path):
        base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        cand_path = os.path.join(base_dir, target_path)
        if os.path.exists(cand_path):
            target_path = cand_path
        else:
            target_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), target_path)

    if not os.path.exists(target_path):
        return DEFAULT_FALLBACK_CATALOG.copy()

    with _catalog_lock:
        try:
            current_mtime = os.path.getmtime(target_path)
            if _cached_catalog is not None and current_mtime == _catalog_mtime:
                return _cached_catalog

            wb = openpyxl.load_workbook(target_path, data_only=True, read_only=True)
            ws = wb.active

            catalog = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 2:
                    continue
                code_raw = row[0]
                doc_raw = row[1]
                if not code_raw or not doc_raw:
                    continue

                code_str = str(code_raw).strip()
                doc_str = str(doc_raw).strip()
                is_inc = ("incidencia" in doc_str.lower()) or (code_str.upper() in ("DC-0020", "DC-0022"))

                catalog.append({
                    "codigo": code_str,
                    "documento": doc_str,
                    "is_incidencia": is_inc
                })

            wb.close()
            if catalog:
                _cached_catalog = catalog
                _catalog_mtime = current_mtime
                return _cached_catalog
        except Exception as e:
            print(f"[load_ocr_catalog] Error leyendo {target_path}: {e}")

    return DEFAULT_FALLBACK_CATALOG.copy()


def is_file_ready(filepath: str, check_interval: float = 0.5, max_checks: int = 30) -> bool:
    """Verifica que el escáner haya terminado de escribir el archivo y esté libre."""
    if not os.path.exists(filepath):
        return False

    last_size = -1
    stable_count = 0

    for _ in range(max_checks):
        try:
            if not os.path.exists(filepath):
                return False
            size = os.path.getsize(filepath)
            if size > 0 and size == last_size:
                stable_count += 1
            else:
                stable_count = 0
            last_size = size

            if stable_count >= 2:
                try:
                    with open(filepath, "r+b") as f:
                        pass
                    return True
                except (IOError, PermissionError):
                    stable_count = 0
        except Exception:
            pass
        time.sleep(check_interval)

    return False


class ExcelManager:
    def __init__(self):
        self.cached_orders = []
        self.last_loaded = 0
        self.cache_ttl = 60  # Segundos de caché para agilizar búsquedas
        self.lock = threading.RLock()

    def get_orders(self, force_reload: bool = False) -> List[Dict[str, Any]]:
        now = time.time()
        with self.lock:
            if not force_reload and self.cached_orders and (now - self.last_loaded < self.cache_ttl):
                return self.cached_orders

            cfg = load_config()
            excel_path = cfg.get("excel_path")
            sheet_name = cfg.get("excel_sheet", "Listado")

            if not excel_path or not os.path.exists(excel_path):
                print(f"[ExcelManager] Archivo no accesible: {excel_path}")
                return self.cached_orders

            orders = []
            try:
                wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                else:
                    ws = wb.active

                for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                    if not row or len(row) < 8:
                        continue

                    of_raw = row[7] # Columna H
                    if not of_raw:
                        continue

                    of_str = str(of_raw).strip()
                    if of_str.upper() in ("PTE", "PENDIENTE", "NONE", "-", ""):
                        continue

                    pedido_odoo = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
                    sb = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
                    desc = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
                    cliente = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                    ped_cliente = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""

                    orders.append({
                        "row": row_idx,
                        "of": of_str,
                        "of_sanitized": sanitize_filename(of_str),
                        "pedido_odoo": pedido_odoo,
                        "sb": sb,
                        "descripcion": desc,
                        "cliente": cliente,
                        "pedido_cliente": ped_cliente,
                        "display_label": f"{of_str} | {pedido_odoo} | {desc} ({cliente})"
                    })

                wb.close()
                self.cached_orders = orders
                self.last_loaded = now
                print(f"[ExcelManager] Cargadas {len(orders)} órdenes de fabricación exitosamente.")
            except Exception as e:
                print(f"[ExcelManager] Error leyendo Excel: {e}")

            return self.cached_orders

    def search(self, query: str) -> List[Dict[str, Any]]:
        orders = self.get_orders()
        if not query:
            return orders[:50]

        q = query.lower().strip()
        results = []
        for o in orders:
            if (q in o["of"].lower() or
                q in o["of_sanitized"].lower() or
                q in o["pedido_odoo"].lower() or
                q in o["descripcion"].lower() or
                q in o["cliente"].lower()):
                results.append(o)
                if len(results) >= 50:
                    break
        return results

    def find_by_of(self, of_query: str) -> Optional[Dict[str, Any]]:
        if not of_query:
            return None

        orders = self.get_orders()
        q_raw = of_query.strip().upper()
        q_norm = q_raw.replace('\\', '/').replace('-', '/')
        q_sanitized = sanitize_filename(q_raw).upper()

        # 1. Coincidencia exacta
        for o in orders:
            o_of = o["of"].upper()
            o_norm = o_of.replace('\\', '/').replace('-', '/')
            if o_of == q_raw or o_norm == q_norm or o["of_sanitized"].upper() == q_sanitized:
                return o

        # 2. Coincidencia por número final
        match_digits = re.search(r'(\d{3,6})$', q_raw)
        if match_digits:
            q_num_str = match_digits.group(1)
            q_num_int = int(q_num_str)
            for o in orders:
                parts = o["of"].replace('\\', '/').replace('-', '/').split('/')
                last_part = parts[-1].strip()
                if last_part.isdigit() and int(last_part) == q_num_int:
                    prefix_match = re.match(r'^([A-Z0-9]+)', q_raw)
                    if prefix_match:
                        pref = prefix_match.group(1)
                        if pref in o["of"].upper():
                            return o
                    else:
                        return o

        # 3. Coincidencia por terminación
        for o in orders:
            if o["of"].upper().endswith(q_raw) or o["of_sanitized"].upper().endswith(q_sanitized):
                return o

        # 4. Subcadena general
        for o in orders:
            if q_raw in o["of"].upper() or q_sanitized in o["of_sanitized"].upper():
                return o

        return None


# =========================================================================
# REPLICACIÓN AUTOMÁTICA EN OFICINA TÉCNICA (05-DOCUMENTACION FINAL)
# =========================================================================

COMPANY_PREFIX_MAP = {
    "AYC": "02-AyC",
    "MES": "01-MES",
    "LOESS": "03-LOESS",
    "CESA": "04-CESA",
    "CTI": "05-CTI"
}

KNOWN_CLIENT_ALIASES = {
    "LINEAS Y MONTAJES": "LYMET",
    "LINEAS Y MONTAJES ELECTRICOS": "LYMET",
    "CELULOSAS DE ASTURIAS": "ENCE",
    "CEASA": "ENCE",
    "SCHNEIDER ELECTRIC SYSTEM": "SCHNEIDER SESI",
}

SHORT_DOC_NAMES = {
    "DC-0018": "Carros Mat",
    "DC-0019": "Orden de Fabricacion",
    "DC-0020": "Incidencias",
    "DC-0021": "Control de pendientes",
    "DC-0022": "Control Estado Entrada Verifica",
    "DC-0023": "Pesajes",
    "DC-0024": "Expediciones-Submer",
    "DC-0025": "Decl. Conformidad",
    "DC-0026": "Decl. Conformidad UL508A",
    "DC-0027": "Certificado Garantia MCD",
    "DC-0028": "Listado verif.",
    "DC-0029": "Control Interno General",
    "DC-0030": "Listado verif. (EN)",
    "DC-0031": "Verificacion UL",
    "DC-0032": "Informe Tecnico Ingeniero",
    "DC-0033": "Informe Tecnico Obra",
    "DC-0034": "Pruebas FAT (ES)",
    "DC-0035": "Pruebas FAT (EN)",
    "DC-0036": "Pruebas FAT Cliente (ES)",
    "DC-0037": "Pruebas FAT Cliente (EN)",
    "DC-0038": "Hoja Almacen",
    "DC-0039": "Hoja Revision Cubiculos",
    "ANEXO 1_IT0001": "Tareas Verificacion MCD",
    "ANEXO 2_IT0001": "Check MCD"
}


def format_of_for_filename(of_str: str) -> str:
    """Formatea la OF con guiones bajos para el nombre de archivo (ej: MES_MO_05266)."""
    cleaned = of_str.replace('/', '_').replace('\\', '_').replace('-', '_')
    cleaned = re.sub(r'_+', '_', cleaned)
    return cleaned.strip('_')


def get_short_document_name(code: str, default_name: str, is_mcd: bool = False) -> str:
    """Retorna la nomenclatura resumida oficial del documento según el tipo de proyecto."""
    code_up = str(code).upper().strip()
    if code_up in SHORT_DOC_NAMES:
        name = SHORT_DOC_NAMES[code_up]
        if not is_mcd:
            if code_up == "ANEXO 1_IT0001":
                return "Tareas Verificacion"
            if code_up == "DC-0027":
                return "Certificado Garantia"
            if code_up == "DC-0025":
                return "Decl. Conformidad UE"
        return name
    return default_name


def norm_search_token(s: Any) -> str:
    """Normaliza un texto eliminando tildes, signos y caracteres no alfanuméricos."""
    if not s:
        return ""
    txt = unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'[^A-Za-z0-9]', '', txt).upper()


def to_windows_long_path(p: str) -> str:
    """Convierte una ruta a formato extendido de Windows (\\\\?\\) para soportar >260 caracteres."""
    if not p or not isinstance(p, str):
        return p
    if os.name != 'nt':
        return p
    if p.startswith('\\\\?\\'):
        return p
    abs_p = os.path.abspath(str(p).strip())
    if abs_p.startswith('\\\\'):
        return '\\\\?\\UNC\\' + abs_p[2:]
    return '\\\\?\\' + abs_p


def sort_year_folders(folders: List[str]) -> List[str]:
    """Ordena carpetas con referencias a años con prioridad dinámica."""
    cur_year = datetime.now().year
    next_year = cur_year + 1

    def get_priority(folder_name: str):
        m = re.search(r'(?<!\d)20\d{2}(?!\d)', folder_name)
        if not m:
            return (999, 0)
        y = int(m.group(0))
        if y == next_year:
            return (0, -y)
        elif y == cur_year:
            return (1, -y)
        elif y < cur_year:
            return (2, -y)
        else:
            return (3, -y)

    valid_folders = [f for f in folders if re.search(r'(?<!\d)20\d{2}(?!\d)', f)]
    return sorted(valid_folders, key=get_priority)


def find_oficina_tecnica_project_path(
    of_str: str,
    order_info: Optional[Dict[str, Any]] = None,
    base_ot_path: Optional[str] = None
) -> Tuple[Optional[str], str, bool]:
    """Localiza la ruta del proyecto en Oficina Técnica en base a la OF y datos de Aktis."""
    if not base_ot_path:
        cfg = load_config()
        base_ot_path = cfg.get("oficina_tecnica_base", r"\\172.30.0.10\Oficina_tecnica")

    if not os.path.isdir(base_ot_path):
        return None, f"Ruta base de Oficina Técnica no accesible: '{base_ot_path}'", False

    m_pref = re.match(r'^([A-Za-z0-9_-]{2,10})[/\\-]', of_str.strip())
    prefix = m_pref.group(1).upper() if m_pref else of_str.split('/')[0].split('-')[0].upper().strip()

    company_folder = COMPANY_PREFIX_MAP.get(prefix)
    if not company_folder:
        try:
            for cand in os.listdir(base_ot_path):
                if prefix.lower() in cand.lower():
                    company_folder = cand
                    break
        except Exception:
            pass

    if not company_folder:
        return None, f"No se encontró carpeta para el prefijo de empresa '{prefix}'", False

    proyectos_base = os.path.join(base_ot_path, company_folder, "00-Proyectos")
    if not os.path.isdir(proyectos_base):
        return None, f"No existe la carpeta 00-Proyectos: '{proyectos_base}'", False

    raw_client = (order_info.get("cliente") or "") if order_info else ""
    raw_desc = (order_info.get("descripcion") or "") if order_info else ""
    sb_val = (order_info.get("sb") or "").strip() if order_info else ""

    # Caso especial McDonald's
    is_mcd = bool("MCD" in raw_desc.upper() or "MCDONALD" in raw_client.upper() or "MC DONALD" in raw_client.upper())
    if is_mcd and (prefix == "MES" or "01-MES" in company_folder):
        mcd_client_folder = None
        for cand in os.listdir(proyectos_base):
            cand_norm = norm_search_token(cand)
            if "MCDONALD" in cand_norm and "FRIDECOR" not in cand.upper():
                mcd_client_folder = cand
                break
        if not mcd_client_folder:
            mcd_client_folder = "Mc Donald´s"

        mcd_client_path = os.path.join(proyectos_base, mcd_client_folder)
        if not os.path.isdir(mcd_client_path):
            return None, f"No existe la carpeta de McDonald's: '{mcd_client_path}'", False

        m_macro = re.search(r'[A-Z]\d{5}', sb_val)
        macro_code = m_macro.group(0).upper() if m_macro else ""

        macro_folders = [d for d in os.listdir(mcd_client_path) if os.path.isdir(os.path.join(mcd_client_path, d))]
        sorted_macros = sort_year_folders(macro_folders)

        if not sorted_macros:
            return None, f"No se encontró carpeta anual en '{mcd_client_path}'", False

        ordered_macro_cands = []
        if macro_code:
            ordered_macro_cands.extend([m for m in sorted_macros if macro_code in m.upper()])
        for m in sorted_macros:
            if m not in ordered_macro_cands:
                ordered_macro_cands.append(m)

        clean_sb = re.sub(r'[^A-Za-z0-9]', '', sb_val).upper()
        is_subcuadro = clean_sb.endswith('A') or ("SUBCUADRO" in raw_desc.upper() and "A" in sb_val.upper())
        base_project_code = clean_sb[:-1] if clean_sb.endswith('A') else clean_sb

        desc_toks = [norm_search_token(t) for t in re.split(r'[\s_\-/,.;:]+', raw_desc) if len(t) >= 4 and norm_search_token(t) not in ('REST', 'MCD', 'INSTORE', 'SUBCUADRO', 'VARIADOR')]

        matched_rest = None
        target_proy_01_path = None

        for macro_cand in ordered_macro_cands:
            cand_p01 = os.path.join(mcd_client_path, macro_cand, "01 - Proyectos")
            if not os.path.isdir(cand_p01):
                cand_p01 = os.path.join(mcd_client_path, macro_cand)

            try:
                restaurant_dirs = [d for d in os.listdir(to_windows_long_path(cand_p01)) if os.path.isdir(to_windows_long_path(os.path.join(cand_p01, d)))]
            except Exception:
                continue

            if base_project_code:
                for r_cand in restaurant_dirs:
                    if re.sub(r'[^A-Za-z0-9]', '', r_cand).upper().startswith(base_project_code):
                        matched_rest = r_cand
                        target_proy_01_path = cand_p01
                        break

            if not matched_rest and desc_toks:
                best_r, best_m = None, 0
                for r_cand in restaurant_dirs:
                    r_norm = norm_search_token(r_cand)
                    m = sum(1 for t in desc_toks if t in r_norm)
                    if m > best_m:
                        best_m = m
                        best_r = r_cand
                if best_m >= 1:
                    matched_rest = best_r
                    target_proy_01_path = cand_p01

            if matched_rest:
                break

        if not matched_rest:
            return None, f"Restaurante no encontrado para SB '{sb_val}' en carpetas de McDonald's", False

        rest_full_path = os.path.join(target_proy_01_path, matched_rest)
        doc_final_dir = None
        for sub in os.listdir(to_windows_long_path(rest_full_path)):
            if "05" in sub and "FINAL" in norm_search_token(sub):
                doc_final_dir = os.path.join(rest_full_path, sub)
                break
        if not doc_final_dir:
            doc_final_dir = os.path.join(rest_full_path, "05 - Doc. Final (As Built)")
            try:
                os.makedirs(to_windows_long_path(doc_final_dir), exist_ok=True)
            except Exception as e:
                return None, f"No se pudo crear carpeta 05 final en McDonald's: {e}", False

        sub_folder_name = "SubCuadro" if is_subcuadro else "Cuadro General"
        target_dump = os.path.join(doc_final_dir, sub_folder_name)
        try:
            os.makedirs(to_windows_long_path(target_dump), exist_ok=True)
        except Exception as e:
            return None, f"No se pudo crear carpeta '{sub_folder_name}': {e}", False

        return target_dump, "OK", True

    # Proyectos generales
    if not raw_client:
        return None, "No hay cliente informado en los datos de Aktis", False

    clean_client = re.sub(r'[,.]?\s*\b(S\.?L\.?U?|S\.?A\.?U?|S\.?A\.?|S\.?L\.?|SOCIEDAD LIMITADA|SOCIEDAD ANONIMA)\b', '', raw_client, flags=re.I).strip(' .,')
    clean_compact = norm_search_token(clean_client)

    alias_target = None
    for k_alias, v_alias in KNOWN_CLIENT_ALIASES.items():
        if k_alias in clean_client.upper():
            alias_target = v_alias
            break

    client_tokens = [norm_search_token(w) for w in re.split(r'[\s_,\-\.]+', clean_client) if len(w) >= 3 and norm_search_token(w) not in ("DEL", "LOS", "LAS", "POR", "PARA", "CON", "SIN", "GRUPO")]

    try:
        available_clients = [d for d in os.listdir(proyectos_base) if os.path.isdir(os.path.join(proyectos_base, d)) and not d.startswith("00_") and not d.startswith("0_") and not d.startswith("Thumbs")]
    except Exception as e:
        return None, f"Error listando proyectos en '{proyectos_base}': {e}", False

    scored_clients = []
    for cand in available_clients:
        cand_norm = norm_search_token(cand)
        score = 0
        if alias_target and norm_search_token(alias_target) in cand_norm:
            score = 120
        elif cand_norm == clean_compact:
            score = 100
        elif cand_norm in clean_compact or clean_compact in cand_norm:
            score = 85
        else:
            matched_toks = sum(1 for t in client_tokens if t in cand_norm)
            if matched_toks > 0:
                score = 50 + (matched_toks * 10)
        if score >= 40:
            scored_clients.append((score, cand))

    scored_clients.sort(key=lambda x: x[0], reverse=True)

    if not scored_clients:
        return None, f"Cliente '{raw_client}' no localizado en '{proyectos_base}'", False

    odoo_code = (order_info.get("pedido_odoo") or "").strip().upper() if order_info else ""
    sb_code = (order_info.get("sb") or "").strip().upper().split('.')[0] if order_info else ""
    desc_tokens = [norm_search_token(t) for t in re.split(r'[\s_\-/,.;:]+', raw_desc) if len(t) >= 3 and norm_search_token(t) not in ("DEL", "LOS", "LAS", "PARA", "CON", "CUADRO", "CUADROS")]

    matched_project_path = None

    for _, client_folder in scored_clients[:3]:
        client_path = os.path.join(proyectos_base, client_folder)
        try:
            entries = os.listdir(to_windows_long_path(client_path))
        except Exception:
            continue

        year_subdirs = sort_year_folders([d for d in entries if os.path.isdir(to_windows_long_path(os.path.join(client_path, d)))])
        search_paths = [os.path.join(client_path, y) for y in year_subdirs] if year_subdirs else [client_path]

        for sp in search_paths:
            try:
                subdirs = [d for d in os.listdir(to_windows_long_path(sp)) if os.path.isdir(to_windows_long_path(os.path.join(sp, d)))]
            except Exception:
                continue

            cands = []
            if odoo_code:
                cands = [d for d in subdirs if odoo_code in d.upper()]
            if not cands and sb_code:
                cands = [d for d in subdirs if sb_code in d.upper()]

            if cands:
                if len(cands) == 1:
                    matched_project_path = os.path.join(sp, cands[0])
                    break
                else:
                    best_c = None
                    best_m = -1
                    for c in cands:
                        c_norm = norm_search_token(c)
                        m = sum(1 for t in desc_tokens if t in c_norm)
                        if m > best_m:
                            best_m = m
                            best_c = c
                    matched_project_path = os.path.join(sp, best_c or cands[0])
                    break

        if matched_project_path:
            break

    if not matched_project_path:
        return None, f"Proyecto '{odoo_code}' no encontrado en las carpetas de '{scored_clients[0][1]}'", False

    doc_final_path = None
    try:
        for sub in os.listdir(to_windows_long_path(matched_project_path)):
            sub_norm = norm_search_token(sub)
            if "05" in sub and "FINAL" in sub_norm:
                doc_final_path = os.path.join(matched_project_path, sub)
                break
    except Exception:
        pass

    if not doc_final_path:
        doc_final_path = os.path.join(matched_project_path, "05-Documentacion final")
        try:
            os.makedirs(to_windows_long_path(doc_final_path), exist_ok=True)
        except Exception as e:
            return None, f"No se pudo crear carpeta de documentación final: {e}", False

    return doc_final_path, "OK", False


def replicate_to_oficina_tecnica(
    of_str: str,
    order_info: Dict[str, Any],
    dest_dir: str,
    log_callback: Optional[Callable[[str, str], None]] = None
) -> Optional[str]:
    """Replica los PDFs generados en la carpeta correspondiente de Oficina Técnica."""
    def log(msg: str, level: str = "info"):
        if log_callback:
            log_callback(msg, level)
        else:
            print(f"[{level.upper()}] {msg}")

    cfg = load_config()
    if not cfg.get("copy_to_oficina_tecnica", True):
        return None

    base_ot = cfg.get("oficina_tecnica_base", r"\\172.30.0.10\Oficina_tecnica")

    doc_final_path, status_msg, is_direct_dump = find_oficina_tecnica_project_path(of_str, order_info, base_ot)
    if not doc_final_path:
        log(f"  ℹ️ Oficina Técnica: {status_msg}. (No se pudo replicar copia automática)", "info")
        return None

    try:
        count = 0
        long_dest_dir = to_windows_long_path(dest_dir)
        long_doc_final = to_windows_long_path(doc_final_path)

        if is_direct_dump:
            os.makedirs(long_doc_final, exist_ok=True)
            for item in os.listdir(long_dest_dir):
                if item.startswith("_ORIGINAL"):
                    continue
                s_file = os.path.join(dest_dir, item)
                if os.path.isfile(to_windows_long_path(s_file)):
                    d_file = os.path.join(doc_final_path, item)
                    shutil.copy2(to_windows_long_path(s_file), to_windows_long_path(d_file))
                    count += 1
            log(f"  📁 Copia replicada en Oficina Técnica [{os.path.basename(doc_final_path)}]: {count} archivos -> '{doc_final_path}'", "success")
            return doc_final_path
        else:
            folder_name = os.path.basename(dest_dir)
            ot_target = os.path.join(doc_final_path, folder_name)
            long_ot_target = to_windows_long_path(ot_target)
            os.makedirs(long_ot_target, exist_ok=True)
            for item in os.listdir(long_dest_dir):
                if item.startswith("_ORIGINAL"):
                    continue
                s_file = os.path.join(dest_dir, item)
                if os.path.isfile(to_windows_long_path(s_file)):
                    d_file = os.path.join(ot_target, item)
                    shutil.copy2(to_windows_long_path(s_file), to_windows_long_path(d_file))
                    count += 1
            log(f"  📁 Copia replicada en Oficina Técnica: '{folder_name}' ({count} archivos) -> '{doc_final_path}'", "success")
            return ot_target
    except Exception as e:
        log(f"  ⚠️ Error replicando en Oficina Técnica: {e}", "warning")
        return None


def process_scanned_pdf(
    pdf_path: str,
    excel_manager: ExcelManager,
    catalog: Optional[List[Dict[str, Any]]] = None,
    log_callback: Optional[Callable[[str, str], None]] = None,
    progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None
) -> Dict[str, Any]:
    """
    Motor central de VeriDocs:
    1. Lee la cabecera (pág 1) para detectar la OF.
    2. Conecta con Aktis para obtener Pedido ODOO, Descripción, etc.
    3. Analiza el pie de cada hoja contra codigosOCR.xlsx.
    4. Gestión DÚPLEX segura:
       - Si duplex_mode es False: procesa cada página por separado (sin descartar reverso).
       - Si duplex_mode es True: mantiene ambas caras para Pruebas FAT (DC-0034/DC-0035).
         Para el resto, si la segunda cara tiene código OCR reconocido válido la conserva;
         si está en blanco o sin código, la descarta.
    5. Agrupa páginas por código y gestiona incidencias (1-N, copia DC-0022).
    6. Guarda en destino, archiva original y replica en Oficina Técnica.
    """
    created_files = []

    def log(msg: str, level: str = "info"):
        if log_callback:
            log_callback(msg, level)
        else:
            print(f"[{level.upper()}] {msg}")

    def notify_prog(phase: str, cur_p: int = 0, tot_p: int = 0, of_val: str = ""):
        if progress_callback:
            try:
                progress_callback({
                    "phase": phase,
                    "current_page": cur_p,
                    "total_pages": tot_p,
                    "of": of_val,
                    "created_files": list(created_files)
                })
            except Exception:
                pass

    cfg = load_config()
    header_pct = int(cfg.get("ocr_header_height_percent", 30))
    footer_pct = int(cfg.get("ocr_footer_height_percent", 15))
    ocr_lang = cfg.get("ocr_lang", "es-ES")
    ocr_enabled = cfg.get("ocr_enabled", True)
    duplex_mode = cfg.get("duplex_mode", True)
    dest_base = cfg.get("destination_base_folder") or cfg.get("escaneos_folder")

    if catalog is None:
        catalog = load_ocr_catalog()

    fname = os.path.basename(pdf_path)
    log(f"Iniciando procesamiento de archivo escaneado: '{fname}'...")

    if not os.path.exists(to_windows_long_path(pdf_path)):
        return {"success": False, "error": f"El archivo no existe: {pdf_path}"}

    try:
        src_doc = pdfium.PdfDocument(pdf_path)
    except Exception as e:
        log(f"Error al abrir PDF con pypdfium2: {e}", "error")
        return {"success": False, "error": f"No se pudo abrir el PDF: {e}"}

    total_pages = len(src_doc)
    if total_pages == 0:
        src_doc.close()
        return {"success": False, "error": "El documento PDF está vacío (0 páginas)."}

    log(f"Documento cargado correctamente con {total_pages} páginas.")
    notify_prog("Leyendo cabecera (pág 1) para detectar OF...", 0, total_pages)

    if not ocr_enabled:
        log("⚠️ Reconocimiento OCR desactivado en configuración.", "warning")

    # 1. OCR de la Cabecera de la Página 1 para detectar la OF
    log(f"🔍 Leyendo cabecera (primeros {header_pct}%) de la página 1...")
    of_res = extract_of_from_header(src_doc, page_index=0, header_height_percent=header_pct, lang=ocr_lang)

    detected_ofs = of_res.get("detected_ofs", [])
    best_of = of_res.get("best_of")

    if not best_of and total_pages > 1:
        log("  ℹ️ No se detectó OF en página 1, probando cabecera de página 2...")
        of_res = extract_of_from_header(src_doc, page_index=1, header_height_percent=header_pct, lang=ocr_lang)
        detected_ofs = of_res.get("detected_ofs", [])
        best_of = of_res.get("best_of")

    if not best_of:
        src_doc.close()
        err_msg = f"No se pudo detectar el código de la OF en la cabecera de '{fname}'."
        log(f"❌ {err_msg}", "error")
        return {
            "success": False,
            "error": err_msg,
            "raw_header_text": of_res.get("raw_text", ""),
            "total_pages": total_pages
        }

    log(f"✅ OF detectada en cabecera: '{best_of}'", "success")
    notify_prog(f"OF: {best_of}. Consultando pedidos en Aktis...", 1, total_pages, of_val=best_of)

    # 2. Buscar la OF en el Excel de Aktis
    order_info = None
    for cand in detected_ofs:
        order_info = excel_manager.find_by_of(cand)
        if order_info:
            break

    if not order_info:
        order_info = excel_manager.find_by_of(best_of)

    alert_warning = None
    if order_info:
        of_str = order_info["of"]
        raw_desc = order_info.get("descripcion", "")
        raw_client = order_info.get("cliente", "")
        sb_val = order_info.get("sb", "")
        is_mcd = bool("MCD" in raw_desc.upper() or "MCDONALD" in raw_client.upper() or "MC DONALD" in raw_client.upper())

        of_formatted = format_of_for_filename(of_str)
        col_h = of_formatted

        m_macro = re.search(r'[A-Z]\d{5}', sb_val)
        if is_mcd and m_macro:
            col_e = m_macro.group(0)
        else:
            col_e = sanitize_filename(order_info.get("pedido_odoo", ""))

        col_i = sanitize_filename(raw_desc)
        log(f"📋 Encontrada en Aktis: Pedido Odoo='{col_e}', Descripción='{col_i}', Cliente='{raw_client}'")
    else:
        of_str = best_of
        raw_desc = ""
        raw_client = "NO REGISTRADO EN AKTIS"
        sb_val = ""
        is_mcd = False
        of_formatted = format_of_for_filename(best_of)
        col_e = ""
        col_i = ""
        col_h = of_formatted
        alert_warning = f"Aviso: La OF '{best_of}' no se encontró en el listado de pedidos en curso Aktis."
        log(f"⚠️ {alert_warning}", "warning")
        order_info = {
            "of": of_str,
            "of_sanitized": col_h,
            "pedido_odoo": "",
            "descripcion": "",
            "cliente": raw_client,
            "sb": ""
        }

    # 3. Carpeta de destino de la OF
    dest_dir = os.path.join(dest_base, col_h)
    try:
        os.makedirs(to_windows_long_path(dest_dir), exist_ok=True)
    except Exception as e_mkdir:
        src_doc.close()
        log(f"Error creando carpeta destino {dest_dir}: {e_mkdir}", "error")
        return {"success": False, "error": f"No se pudo crear la carpeta {dest_dir}: {e_mkdir}"}

    log(f"Carpeta destino preparada: '{dest_dir}'")

    # 4. Análisis del pie de página de cada hoja del PDF
    log(f"🔍 Analizando pies de página de las {total_pages} páginas según 'codigosOCR.xlsx'...")
    page_classifications = []
    page_rotations = {}

    for p_idx in range(total_pages):
        notify_prog(f"Analizando pie pág {p_idx + 1} de {total_pages}...", p_idx + 1, total_pages, of_val=of_str)
        detect_res = extract_footer_and_detect_code(
            src_doc,
            catalog=catalog,
            page_index=p_idx,
            footer_height_percent=footer_pct,
            lang=ocr_lang
        )

        matched_doc = detect_res.get("matched_doc")
        matched_code = detect_res.get("matched_code")
        rot_needed = detect_res.get("rotation_needed", 0)
        page_rotations[p_idx] = rot_needed

        page_classifications.append({
            "page_idx": p_idx,
            "doc_def": matched_doc,
            "code": matched_code,
            "raw_text": detect_res.get("text", ""),
            "rotation_needed": rot_needed
        })

    # 5. Agrupación por tipo de documento con soporte DÚPLEX Seguro
    DUPLEX_PRESERVE_CODES = ("DC-0034", "DC-0035")
    groups = {}
    unrecognized_pages = []

    if not duplex_mode:
        # Modo 1 CARA: Cada página es independiente, no se descarta ningún reverso
        for item in page_classifications:
            p_idx = item["page_idx"]
            doc_def = item["doc_def"]
            code = (item.get("code") or "").strip().upper()
            rot_needed = item.get("rotation_needed", 0)
            is_blank = (code == "BLANK") or is_page_blank(src_doc, page_index=p_idx)

            if is_blank:
                log(f"  Pág {p_idx + 1}/{total_pages}: Hoja en blanco descartada.")
                notify_prog(f"Pág {p_idx + 1}/{total_pages}: Hoja en blanco (descartada)", p_idx + 1, total_pages, of_val=of_str)
            elif doc_def and code and code != "BLANK":
                doc_name = doc_def.get("documento", code)
                rot_tag = f" [Apaisado corregido {rot_needed}°]" if rot_needed else ""
                log(f"  Pág {p_idx + 1}/{total_pages}: Código '{code}' -> '{doc_name}'{rot_tag}")
                notify_prog(f"Pág {p_idx + 1}/{total_pages}: {code} ({doc_name}){rot_tag}", p_idx + 1, total_pages, of_val=of_str)

                if code not in groups:
                    groups[code] = {"doc_def": doc_def, "pages": []}
                groups[code]["pages"].append(p_idx)
            else:
                log(f"  Pág {p_idx + 1}/{total_pages}: ⚠️ Sin código reconocido.", "warning")
                notify_prog(f"Pág {p_idx + 1}/{total_pages}: Sin código reconocido", p_idx + 1, total_pages, of_val=of_str)
                unrecognized_pages.append(p_idx)
    else:
        # Modo DÚPLEX: Procesa por parejas de anverso y reverso
        s = 0
        while s * 2 < total_pages:
            p_front = s * 2
            p_back = p_front + 1 if p_front + 1 < total_pages else None

            item_f = page_classifications[p_front]
            item_b = page_classifications[p_back] if p_back is not None else None

            code_f = (item_f.get("code") or "").strip().upper()
            code_b = (item_b.get("code") or "").strip().upper() if item_b else ""

            is_fat_sheet = (code_f in DUPLEX_PRESERVE_CODES) or (code_b in DUPLEX_PRESERVE_CODES)

            if is_fat_sheet:
                fat_code = code_f if code_f in DUPLEX_PRESERVE_CODES else code_b
                fat_doc_def = item_f.get("doc_def") if code_f in DUPLEX_PRESERVE_CODES else (item_b.get("doc_def") if item_b else None)
                if not fat_doc_def:
                    for d in catalog:
                        if d.get("codigo", "").strip().upper() == fat_code:
                            fat_doc_def = d
                            break

                doc_name = fat_doc_def.get("documento", fat_code) if fat_doc_def else fat_code

                if fat_code not in groups:
                    groups[fat_code] = {
                        "doc_def": fat_doc_def or {"codigo": fat_code, "documento": doc_name},
                        "pages": []
                    }

                rot_f = page_rotations.get(p_front, 0)
                rot_tag_f = f" [Apaisado corregido {rot_f}°]" if rot_f else ""
                log(f"  Pág {p_front + 1}/{total_pages}: Código '{fat_code}' -> '{doc_name}' (Anverso FAT){rot_tag_f}")
                notify_prog(f"Pág {p_front + 1}/{total_pages}: {fat_code} (Anverso)", p_front + 1, total_pages, of_val=of_str)
                groups[fat_code]["pages"].append(p_front)

                if p_back is not None:
                    rot_b = page_rotations.get(p_back, 0)
                    rot_tag_b = f" [Apaisado corregido {rot_b}°]" if rot_b else ""
                    log(f"  Pág {p_back + 1}/{total_pages}: Doble cara mantenida para '{fat_code}' (Reverso FAT){rot_tag_b}")
                    notify_prog(f"Pág {p_back + 1}/{total_pages}: {fat_code} (Reverso)", p_back + 1, total_pages, of_val=of_str)
                    groups[fat_code]["pages"].append(p_back)

            else:
                # Comprobar si el reverso tiene su propio código válido independiente (no es un descarte)
                b_is_distinct_doc = (code_b and code_b != "BLANK" and item_b.get("doc_def") is not None)
                f_is_distinct_doc = (code_f and code_f != "BLANK" and item_f.get("doc_def") is not None)

                if f_is_distinct_doc and b_is_distinct_doc:
                    # Ambas caras contienen documentos reconocidos independientes: conservar ambos
                    for target_item, code_val in [(item_f, code_f), (item_b, code_b)]:
                        p_idx = target_item["page_idx"]
                        doc_def = target_item["doc_def"]
                        rot_needed = target_item.get("rotation_needed", 0)
                        doc_name = doc_def.get("documento", code_val)
                        rot_tag = f" [Apaisado corregido {rot_needed}°]" if rot_needed else ""
                        log(f"  Pág {p_idx + 1}/{total_pages}: Código '{code_val}' -> '{doc_name}' (Cara independiente){rot_tag}")
                        notify_prog(f"Pág {p_idx + 1}/{total_pages}: {code_val} ({doc_name}){rot_tag}", p_idx + 1, total_pages, of_val=of_str)

                        if code_val not in groups:
                            groups[code_val] = {"doc_def": doc_def, "pages": []}
                        groups[code_val]["pages"].append(p_idx)
                else:
                    # Documento normal dúplex con reverso descartable
                    if not code_f and code_b and is_page_blank(src_doc, page_index=p_front):
                        target_item = item_b
                        discard_idx = p_front
                    else:
                        target_item = item_f
                        discard_idx = p_back

                    if discard_idx is not None:
                        log(f"  Pág {discard_idx + 1}/{total_pages}: Segunda cara descartada por descarte (dúplex).")
                        notify_prog(f"Pág {discard_idx + 1}/{total_pages}: Segunda cara (descartada)", discard_idx + 1, total_pages, of_val=of_str)

                    p_idx = target_item["page_idx"]
                    doc_def = target_item["doc_def"]
                    code = (target_item.get("code") or "").strip().upper()
                    rot_needed = target_item.get("rotation_needed", 0)

                    is_blank = (code == "BLANK") or is_page_blank(src_doc, page_index=p_idx)

                    if is_blank:
                        log(f"  Pág {p_idx + 1}/{total_pages}: Hoja en blanco descartada.")
                        notify_prog(f"Pág {p_idx + 1}/{total_pages}: Hoja en blanco (descartada)", p_idx + 1, total_pages, of_val=of_str)
                    elif doc_def and code and code != "BLANK":
                        doc_name = doc_def.get("documento", code)
                        rot_tag = f" [Apaisado corregido {rot_needed}°]" if rot_needed else ""
                        log(f"  Pág {p_idx + 1}/{total_pages}: Código '{code}' -> '{doc_name}'{rot_tag}")
                        notify_prog(f"Pág {p_idx + 1}/{total_pages}: {code} ({doc_name}){rot_tag}", p_idx + 1, total_pages, of_val=of_str)

                        if code not in groups:
                            groups[code] = {"doc_def": doc_def, "pages": []}
                        groups[code]["pages"].append(p_idx)
                    else:
                        log(f"  Pág {p_idx + 1}/{total_pages}: ⚠️ Sin código reconocido.", "warning")
                        notify_prog(f"Pág {p_idx + 1}/{total_pages}: Sin código reconocido", p_idx + 1, total_pages, of_val=of_str)
                        unrecognized_pages.append(p_idx)

            s += 1

    # 6. Generación de los archivos PDF divididos
    escaneos_root = cfg.get("escaneos_folder") or dest_base
    inc_folder = os.path.join(escaneos_root, "incidencias")
    if not os.path.isdir(inc_folder) and os.path.isdir(os.path.join(escaneos_root, "Incidencias")):
        inc_folder = os.path.join(escaneos_root, "Incidencias")

    for code, g in groups.items():
        doc_def = g["doc_def"]
        pages = g["pages"]
        default_doc_name = sanitize_filename(doc_def.get("documento", code))
        doc_name = get_short_document_name(code, default_doc_name, is_mcd=is_mcd)
        is_inc = doc_def.get("is_incidencia", False) or (str(code).upper() in ("DC-0020", "DC-0022"))

        if is_inc and len(pages) > 1:
            total_inc = len(pages)
            for k, p in enumerate(pages, start=1):
                parts = [doc_name, f"{k}-{total_inc}"]
                if col_e:
                    parts.append(col_e)
                if col_i:
                    parts.append(col_i)
                parts.append(col_h)
                out_name = " - ".join(parts) + ".pdf"
                out_path = os.path.join(dest_dir, out_name)

                if os.path.exists(to_windows_long_path(out_path)):
                    base, ext = os.path.splitext(out_name)
                    out_path = os.path.join(dest_dir, f"{base}_{int(time.time())}{ext}")

                sub_doc = pdfium.PdfDocument.new()
                sub_doc.import_pages(src_doc, [p])
                rot_p = page_rotations.get(p, 0)
                if rot_p != 0:
                    sub_doc[0].set_rotation((sub_doc[0].get_rotation() + rot_p) % 360)
                sub_doc.save(to_windows_long_path(out_path))
                sub_doc.close()

                if str(code).strip().upper() == "DC-0022":
                    try:
                        os.makedirs(to_windows_long_path(inc_folder), exist_ok=True)
                        inc_dest = os.path.join(inc_folder, os.path.basename(out_path))
                        shutil.copy2(to_windows_long_path(out_path), to_windows_long_path(inc_dest))
                        log(f"  📋 Copia duplicada en carpeta 'Incidencias': '{os.path.basename(inc_dest)}'")
                    except Exception as e_copy_inc:
                        log(f"  ⚠️ No se pudo copiar en Incidencias: {e_copy_inc}", "warning")

                created_files.append({
                    "code": code,
                    "doc_name": f"{doc_name} ({k}-{total_inc})",
                    "filename": os.path.basename(out_path),
                    "path": out_path,
                    "pages_count": 1,
                    "pages_list": [p + 1]
                })
                log(f"  📄 Creado: '{os.path.basename(out_path)}' (Pág {p + 1})")
                notify_prog(f"Generado: {os.path.basename(out_path)}", total_pages, total_pages, of_val=of_str)
        else:
            parts = [doc_name]
            if col_e:
                parts.append(col_e)
            if col_i:
                parts.append(col_i)
            parts.append(col_h)
            out_name = " - ".join(parts) + ".pdf"
            out_path = os.path.join(dest_dir, out_name)

            if os.path.exists(to_windows_long_path(out_path)):
                base, ext = os.path.splitext(out_name)
                out_path = os.path.join(dest_dir, f"{base}_{int(time.time())}{ext}")

            sub_doc = pdfium.PdfDocument.new()
            sub_doc.import_pages(src_doc, pages)
            for idx_in_sub, orig_p in enumerate(pages):
                rot_p = page_rotations.get(orig_p, 0)
                if rot_p != 0:
                    sub_doc[idx_in_sub].set_rotation((sub_doc[idx_in_sub].get_rotation() + rot_p) % 360)
            sub_doc.save(to_windows_long_path(out_path))
            sub_doc.close()

            if str(code).strip().upper() == "DC-0022":
                try:
                    os.makedirs(to_windows_long_path(inc_folder), exist_ok=True)
                    inc_dest = os.path.join(inc_folder, os.path.basename(out_path))
                    shutil.copy2(to_windows_long_path(out_path), to_windows_long_path(inc_dest))
                    log(f"  📋 Copia duplicada en carpeta 'Incidencias': '{os.path.basename(inc_dest)}'")
                except Exception as e_copy_inc:
                    log(f"  ⚠️ No se pudo copiar en Incidencias: {e_copy_inc}", "warning")

            pages_human = [p + 1 for p in pages]
            created_files.append({
                "code": code,
                "doc_name": doc_name,
                "filename": os.path.basename(out_path),
                "path": out_path,
                "pages_count": len(pages),
                "pages_list": pages_human
            })
            log(f"  📄 Creado: '{os.path.basename(out_path)}' ({len(pages)} págs: {pages_human})")
            notify_prog(f"Generado: {os.path.basename(out_path)}", total_pages, total_pages, of_val=of_str)

    # 7. Páginas no reconocidas
    alert_unrecognized = None
    if unrecognized_pages:
        unrec_pages_human = [p + 1 for p in unrecognized_pages]
        unrec_name = f"REVISAR_PAGINAS_NO_IDENTIFICADAS - {col_h}.pdf"
        unrec_path = os.path.join(dest_dir, unrec_name)

        if os.path.exists(to_windows_long_path(unrec_path)):
            base, ext = os.path.splitext(unrec_name)
            unrec_path = os.path.join(dest_dir, f"{base}_{int(time.time())}{ext}")

        sub_doc = pdfium.PdfDocument.new()
        sub_doc.import_pages(src_doc, unrecognized_pages)
        for idx_in_sub, orig_p in enumerate(unrecognized_pages):
            rot_p = page_rotations.get(orig_p, 0)
            if rot_p != 0:
                sub_doc[idx_in_sub].set_rotation((sub_doc[idx_in_sub].get_rotation() + rot_p) % 360)
        sub_doc.save(to_windows_long_path(unrec_path))
        sub_doc.close()

        alert_unrecognized = f"Se detectaron {len(unrecognized_pages)} páginas sin código OCR reconocido (Páginas {unrec_pages_human}). Guardadas en '{os.path.basename(unrec_path)}'."
        log(f"⚠️ {alert_unrecognized}", "warning")

        created_files.append({
            "code": "UNRECOGNIZED",
            "doc_name": "Páginas no identificadas",
            "filename": os.path.basename(unrec_path),
            "path": unrec_path,
            "pages_count": len(unrecognized_pages),
            "pages_list": unrec_pages_human,
            "is_warning": True
        })

    # Cerrar el PDF fuente antes de moverlo
    src_doc.close()

    # 8. Archivar el PDF original completo en la carpeta de la OF
    archive_dest = os.path.join(dest_dir, f"_ORIGINAL_{fname}")
    if os.path.exists(to_windows_long_path(archive_dest)):
        base_orig, ext_orig = os.path.splitext(fname)
        archive_dest = os.path.join(dest_dir, f"_ORIGINAL_{int(time.time())}_{fname}")

    try:
        shutil.move(to_windows_long_path(pdf_path), to_windows_long_path(archive_dest))
        log(f"Archivado escaneo original en: '{os.path.basename(archive_dest)}'")
    except Exception as e_move:
        try:
            shutil.copy2(to_windows_long_path(pdf_path), to_windows_long_path(archive_dest))
            os.remove(to_windows_long_path(pdf_path))
            log(f"Archivado escaneo original en: '{os.path.basename(archive_dest)}'")
        except Exception as e_copy:
            log(f"No se pudo mover el PDF original: {e_copy}", "warning")

    # 9. Replicar copia de la carpeta en Oficina Técnica (05-Documentacion final)
    ot_replica_path = replicate_to_oficina_tecnica(
        of_str=of_str,
        order_info=order_info,
        dest_dir=dest_dir,
        log_callback=log
    )

    # Combinar alertas de forma segura sin perder ninguna
    alerts_list = []
    if alert_warning:
        alerts_list.append(alert_warning)
    if alert_unrecognized:
        alerts_list.append(alert_unrecognized)
    full_alert = "\n\n".join(alerts_list) if alerts_list else None

    log(f"🎉 Lote de la OF '{of_str}' completado con éxito. Generados {len(created_files)} archivos PDF.", "success")

    return {
        "success": True,
        "of": of_str,
        "order_info": order_info,
        "destination_dir": dest_dir,
        "dest_dir": dest_dir,  # Compatibilidad total entre interfaces
        "total_pages": total_pages,
        "created_files": created_files,
        "unrecognized_pages": [p + 1 for p in unrecognized_pages],
        "alert": full_alert,
        "archive_path": archive_dest,
        "ot_replica_path": ot_replica_path
    }


class ScanningSession:
    """
    Gestor de escucha automática continua para la carpeta de escaneos.
    Soporta cola protegida contra concurrencia, IDs monotónicos para eventos y lotes,
    recarga dinámica de códigos y gestión de alertas múltiples.
    """
    def __init__(self, excel_manager: Optional[ExcelManager] = None):
        self.excel_manager = excel_manager
        self.active = False
        self.worker_thread = None
        self.stop_event = threading.Event()
        self.lock = threading.RLock()
        self.processing_lock = threading.Lock()
        
        self.events_log = []
        self._next_event_id = 1
        self.batches_history = []
        self._next_batch_id = 1
        self.active_alerts = []
        self.current_processing_file = None
        self.current_progress = {
            "is_processing": False,
            "filename": "",
            "of": "",
            "phase": "",
            "current_page": 0,
            "total_pages": 0,
            "created_files": []
        }
        self.last_batch = None
        self.processed_files_set = set()  # set of (path, mtime, size)

    def update_progress(self, is_processing: bool, filename: str = "", of: str = "", phase: str = "", current_page: int = 0, total_pages: int = 0, created_files: list = None):
        with self.lock:
            self.current_progress = {
                "is_processing": is_processing,
                "filename": filename,
                "of": of,
                "phase": phase,
                "current_page": current_page,
                "total_pages": total_pages,
                "created_files": list(created_files) if created_files else []
            }

    def log(self, message: str, level: str = "info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        with self.lock:
            entry = {
                "id": self._next_event_id,
                "time": timestamp,
                "message": message,
                "level": level
            }
            self._next_event_id += 1
            self.events_log.append(entry)
            if len(self.events_log) > 200:
                self.events_log.pop(0)
        print(f"[{timestamp}] [{level.upper()}] {message}")

    def add_alert(self, title: str, message: str, level: str = "warning", dest_dir: str = ""):
        timestamp = datetime.now().strftime("%H:%M:%S")
        alert = {
            "id": int(time.time() * 1000) + len(self.active_alerts),
            "time": timestamp,
            "title": title,
            "message": message,
            "level": level,
            "dest_dir": dest_dir
        }
        with self.lock:
            self.active_alerts.append(alert)

    def dismiss_alert(self, alert_id: int):
        with self.lock:
            self.active_alerts = [a for a in self.active_alerts if a["id"] != alert_id]

    def clear_alerts(self):
        with self.lock:
            self.active_alerts.clear()

    def process_file(self, pdf_path: str, progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None) -> Dict[str, Any]:
        """Procesa un archivo PDF de forma atómica y segura evitando colisiones concurrentes."""
        with self.processing_lock:
            catalog = load_ocr_catalog()
            fname = os.path.basename(pdf_path)
            self.current_processing_file = fname
            self.update_progress(
                is_processing=True,
                filename=fname,
                phase="Iniciando análisis...",
                current_page=0,
                total_pages=0,
                created_files=[]
            )

            def _on_prog(pdata):
                self.update_progress(
                    is_processing=True,
                    filename=fname,
                    of=pdata.get("of", ""),
                    phase=pdata.get("phase", ""),
                    current_page=pdata.get("current_page", 0),
                    total_pages=pdata.get("total_pages", 0),
                    created_files=pdata.get("created_files", [])
                )
                if progress_callback:
                    try:
                        progress_callback(pdata)
                    except Exception:
                        pass

            try:
                res = process_scanned_pdf(
                    pdf_path=pdf_path,
                    excel_manager=self.excel_manager,
                    catalog=catalog,
                    log_callback=self.log,
                    progress_callback=_on_prog
                )
            finally:
                self.update_progress(is_processing=False)
                self.current_processing_file = None

            if res.get("success"):
                with self.lock:
                    self.last_batch = res
                    batch_entry = {
                        "id": self._next_batch_id,
                        "time": datetime.now().strftime("%H:%M:%S"),
                        "of": res["of"],
                        "order_info": res["order_info"],
                        "dest_dir": res["destination_dir"],
                        "destination_dir": res["destination_dir"],
                        "total_pages": res["total_pages"],
                        "files_count": len(res["created_files"]),
                        "created_files": res["created_files"],
                        "alert": res.get("alert"),
                        "ot_replica_path": res.get("ot_replica_path"),
                        "archive_path": res.get("archive_path")
                    }
                    self._next_batch_id += 1
                    self.batches_history.append(batch_entry)

                if res.get("alert"):
                    self.add_alert(
                        title=f"Aviso en OF {res['of']}",
                        message=res["alert"],
                        level="warning",
                        dest_dir=res["destination_dir"]
                    )
            return res

    def start_listening(self) -> Tuple[bool, str]:
        with self.lock:
            if self.active:
                return True, "El monitor ya está en ejecución."

            cfg = load_config()
            escaneos_folder = cfg.get("escaneos_folder")
            if not os.path.exists(escaneos_folder):
                return False, f"La carpeta de escaneos no es accesible: {escaneos_folder}"

            self.stop_event.clear()
            self.active = True
            self.log(f"Iniciada escucha automática en: '{escaneos_folder}'", "success")

            self.worker_thread = threading.Thread(target=self._monitor_loop, daemon=True)
            self.worker_thread.start()
            return True, "Escucha automática iniciada."

    def stop_listening(self) -> Tuple[bool, str]:
        with self.lock:
            if not self.active:
                return True, "El monitor no estaba activo."
            self.stop_event.set()
            self.active = False
            self.log("Escucha automática pausada por el usuario.", "warning")
            return True, "Escucha automática detenida."

    def get_status(self) -> Dict[str, Any]:
        with self.lock:
            return {
                "active": self.active,
                "current_processing_file": self.current_processing_file,
                "current_progress": dict(self.current_progress),
                "last_batch": self.last_batch,
                "batches_history": list(self.batches_history[-50:]),
                "active_alerts": list(self.active_alerts),
                "events_log": list(self.events_log[-50:]),
                "total_batches_processed": len(self.batches_history),
                "last_event_id": self._next_event_id - 1,
                "last_batch_id": self._next_batch_id - 1
            }

    def _monitor_loop(self):
        while not self.stop_event.is_set():
            try:
                cfg = load_config()
                escaneos_folder = cfg.get("escaneos_folder")

                if not os.path.exists(escaneos_folder):
                    time.sleep(2.0)
                    continue

                candidates = []
                current_files = set()

                with os.scandir(escaneos_folder) as entries:
                    for entry in entries:
                        if entry.is_file() and entry.name.lower().endswith(".pdf"):
                            if entry.name.startswith("_"):
                                continue
                            current_files.add(entry.path)
                            try:
                                stat = entry.stat()
                                sig = (entry.path, stat.st_mtime, stat.st_size)
                                if sig not in self.processed_files_set:
                                    candidates.append((stat.st_mtime, entry.path, entry.name, sig))
                            except Exception:
                                pass

                # Limpieza de archivos que ya no existen en bandeja
                self.processed_files_set = {sig for sig in self.processed_files_set if sig[0] in current_files}

                candidates.sort(key=lambda x: x[0])

                for mtime, fpath, fname, sig in candidates:
                    if self.stop_event.is_set():
                        break

                    if not is_file_ready(fpath):
                        continue

                    self.log(f"📥 Nuevo PDF detectado en bandeja: '{fname}'. Procesando lote...")
                    batch_res = self.process_file(fpath)
                    self.processed_files_set.add(sig)

                    if batch_res.get("success"):
                        try:
                            import winsound
                            winsound.MessageBeep(winsound.MB_ICONASTERISK)
                        except Exception:
                            pass
                    else:
                        err_msg = batch_res.get("error", "Error desconocido")
                        unrec_dir = os.path.join(escaneos_folder, "_NO_IDENTIFICADOS")
                        dest_err = ""
                        try:
                            os.makedirs(to_windows_long_path(unrec_dir), exist_ok=True)
                            dest_err = os.path.join(unrec_dir, fname)
                            if os.path.exists(to_windows_long_path(dest_err)):
                                base_e, ext_e = os.path.splitext(fname)
                                dest_err = os.path.join(unrec_dir, f"{base_e}_{int(time.time())}{ext_e}")
                            shutil.move(to_windows_long_path(fpath), to_windows_long_path(dest_err))
                            self.log(f"📁 Archivo sin OF movido a '_NO_IDENTIFICADOS/{os.path.basename(dest_err)}' para revisión.", "warning")
                        except Exception as e_mv_err:
                            dest_err = ""
                            self.log(f"No se pudo mover el archivo con error: {e_mv_err}", "warning")

                        self.add_alert(
                            title=f"Atención en {fname}",
                            message=err_msg + (" (Movido a carpeta _NO_IDENTIFICADOS)" if dest_err else ""),
                            level="error",
                            dest_dir=unrec_dir if dest_err else ""
                        )
                        self.log(f"Error en escaneo '{fname}': {err_msg}", "error")

            except Exception as e:
                self.log(f"Error en bucle de escaneos: {e}", "error")

            time.sleep(1.0)
