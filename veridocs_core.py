import os
import re
import time
import json
import shutil
import threading
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any, Callable
import openpyxl
import pypdfium2 as pdfium

from config import load_config
from ocr_service import (
    is_ocr_available,
    extract_of_from_header,
    extract_footer_text,
    match_footer_code
)

# Catálogo predeterminado si no se encuentra codigosOCR.xlsx
DEFAULT_FALLBACK_CATALOG = [
    {"codigo": "DC-0018", "documento": "Carros Mat.", "is_incidencia": False},
    {"codigo": "DC-0019", "documento": "Orden de Fabricación", "is_incidencia": False},
    {"codigo": "DC-0020", "documento": "Control Estado (incidencias-Materiales)", "is_incidencia": True},
    {"codigo": "DC-0021", "documento": "Control de pendientes", "is_incidencia": False},
    {"codigo": "DC-0022", "documento": "Control Estado Entrada Verifica", "is_incidencia": False},
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


def sanitize_filename(text: Any) -> str:
    """Limpia caracteres inválidos para nombres de archivos y carpetas en Windows."""
    if text is None:
        return ""
    text = str(text).strip()
    # Reemplazar barras, dos puntos y caracteres ilegales por guión
    text = re.sub(r'[\\/:*?"<>|]', '-', text)
    # Colapsar espacios múltiples y guiones consecutivos
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'-+', '-', text)
    return text.strip(' -.')


def load_ocr_catalog(catalog_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Carga todos los tipos de documentos posibles desde codigosOCR.xlsx.
    Identifica de forma dinámica el código, el nombre y si es tipo incidencia.
    """
    global _cached_catalog, _catalog_mtime

    cfg = load_config()
    target_path = catalog_path or cfg.get("ocr_catalog_path", "codigosOCR.xlsx")

    # Si la ruta es relativa, resolverla respecto al directorio de la aplicación
    if not os.path.isabs(target_path):
        app_dir = os.path.dirname(os.path.abspath(__file__))
        target_path = os.path.join(app_dir, target_path)

    if not os.path.exists(target_path):
        print(f"[load_ocr_catalog] Archivo {target_path} no encontrado. Usando catálogo predeterminado.")
        return DEFAULT_FALLBACK_CATALOG.copy()

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
            
            # Detectar si es tipo Incidencia (código DC-0020 o palabra 'incidencia')
            is_inc = ("incidencia" in doc_str.lower()) or (code_str.upper() == "DC-0020")

            catalog.append({
                "codigo": code_str,
                "documento": doc_str,
                "is_incidencia": is_inc
            })

        wb.close()
        if catalog:
            _cached_catalog = catalog
            _catalog_mtime = current_mtime
            print(f"[load_ocr_catalog] Cargados exitosamente {len(catalog)} códigos de documentos desde Excel.")
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

            # Si el tamaño se mantiene estable y podemos abrir el archivo exclusivamente
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
        self.lock = threading.Lock()

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
                # Abrir en modo solo lectura para no bloquear el archivo
                wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                else:
                    ws = wb.active

                # Columnas (1-based index):
                # A: 1 COMPAÑÍA
                # B: 2 Código cliente
                # C: 3 Cliente
                # D: 4 Pedido CLIENTES
                # E: 5 Pedido ODOO
                # F: 6 Código artículo
                # G: 7 SB… (Subartículo/Subpedido)
                # H: 8 Nº OF
                # I: 9 Descripción
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                    if not row or len(row) < 8:
                        continue

                    of_raw = row[7] # Columna H
                    if not of_raw:
                        continue

                    of_str = str(of_raw).strip()
                    if of_str.upper() == "PTE" or of_str == "-" or of_str == "":
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
        r"""
        Búsqueda robusta de una OF en el listado de pedidos en curso.
        Admite variaciones de separadores (/ o - o \), ceros a la izquierda y número final.
        """
        if not of_query:
            return None

        orders = self.get_orders()
        q_raw = of_query.strip().upper()
        q_norm = q_raw.replace('\\', '/').replace('-', '/')
        q_sanitized = sanitize_filename(q_raw).upper()

        # 1. Coincidencia exacta (original o normalizada)
        for o in orders:
            o_of = o["of"].upper()
            o_norm = o_of.replace('\\', '/').replace('-', '/')
            if o_of == q_raw or o_norm == q_norm or o["of_sanitized"].upper() == q_sanitized:
                return o

        # 2. Coincidencia por número final (ej: '00560' o '560' -> 'AYC/MO/00560')
        # Extraer los dígitos finales de la consulta
        match_digits = re.search(r'(\d{3,6})$', q_raw)
        if match_digits:
            q_num_str = match_digits.group(1)
            q_num_int = int(q_num_str)
            for o in orders:
                parts = o["of"].replace('\\', '/').replace('-', '/').split('/')
                last_part = parts[-1].strip()
                if last_part.isdigit() and int(last_part) == q_num_int:
                    # Si la consulta además contenía prefijo, verificar que coincida
                    prefix_match = re.match(r'^([A-Z0-9]+)', q_raw)
                    if prefix_match:
                        pref = prefix_match.group(1)
                        if pref in o["of"].upper():
                            return o
                    else:
                        return o

        # 3. Coincidencia si termina con la cadena ingresada
        for o in orders:
            if o["of"].upper().endswith(q_raw) or o["of_sanitized"].upper().endswith(q_sanitized):
                return o

        # 4. Subcadena general
        for o in orders:
            if q_raw in o["of"].upper() or q_sanitized in o["of_sanitized"].upper():
                return o

        return None


def process_scanned_pdf(
    pdf_path: str,
    excel_manager: ExcelManager,
    catalog: Optional[List[Dict[str, Any]]] = None,
    log_callback: Optional[Callable[[str, str], None]] = None
) -> Dict[str, Any]:
    """
    Motor central del nuevo flujo de VeriDocs:
    1. Lee la cabecera (30%) de la página 1 con OCR para extraer la OF.
    2. Busca la OF en el Excel de Aktis para obtener Pedido ODOO, Descripción, etc.
    3. Analiza el pie de página de cada hoja del PDF contra codigosOCR.xlsx.
    4. Agrupa páginas con el mismo código en un único PDF.
    5. Para Incidencias (DC-0020), si hay más de 1 página genera documentos separados (1-3, 2-3, 3-3).
    6. Guarda páginas no reconocidas en un PDF de revisión y genera alerta.
    7. Guarda todo en la carpeta destino de la OF y archiva el PDF original escaneado.
    """
    def log(msg: str, level: str = "info"):
        if log_callback:
            log_callback(msg, level)
        else:
            print(f"[{level.upper()}] {msg}")

    cfg = load_config()
    header_pct = int(cfg.get("ocr_header_height_percent", 30))
    footer_pct = int(cfg.get("ocr_footer_height_percent", 15))
    ocr_lang = cfg.get("ocr_lang", "es-ES")
    dest_base = cfg.get("destination_base_folder") or cfg.get("escaneos_folder")

    if catalog is None:
        catalog = load_ocr_catalog()

    fname = os.path.basename(pdf_path)
    log(f"Iniciando procesamiento de archivo escaneado: '{fname}'...")

    # 1. Comprobar que el archivo existe y es legible
    if not os.path.exists(pdf_path):
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

    # 2. OCR de la Cabecera de la Página 1 para detectar la OF
    log(f"🔍 Leyendo cabecera (primeros {header_pct}%) de la página 1...")
    of_res = extract_of_from_header(src_doc, page_index=0, header_height_percent=header_pct, lang=ocr_lang)

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

    # 3. Buscar la OF en el Excel de Aktis
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
        col_e = sanitize_filename(order_info.get("pedido_odoo", ""))
        col_i = sanitize_filename(order_info.get("descripcion", ""))
        col_h = sanitize_filename(of_str)
        log(f"📋 Encontrada en Aktis: Pedido Odoo='{col_e}', Descripción='{col_i}', Cliente='{order_info.get('cliente', '')}'")
    else:
        of_str = best_of
        col_e = ""
        col_i = ""
        col_h = sanitize_filename(best_of)
        alert_warning = f"Aviso: La OF '{best_of}' no se encontró en el listado de pedidos en curso Aktis."
        log(f"⚠️ {alert_warning}", "warning")
        order_info = {
            "of": of_str,
            "of_sanitized": col_h,
            "pedido_odoo": "",
            "descripcion": "",
            "cliente": "NO REGISTRADO EN AKTIS"
        }

    # 4. Carpeta de destino de la OF
    dest_dir = os.path.join(dest_base, col_h)
    try:
        os.makedirs(dest_dir, exist_ok=True)
    except Exception as e_mkdir:
        src_doc.close()
        log(f"Error creando carpeta destino {dest_dir}: {e_mkdir}", "error")
        return {"success": False, "error": f"No se pudo crear la carpeta {dest_dir}: {e_mkdir}"}

    log(f"Carpeta destino preparada: '{dest_dir}'")

    # 5. Análisis del pie de página de cada hoja del PDF
    log(f"🔍 Analizando pies de página de las {total_pages} páginas según 'codigosOCR.xlsx'...")
    page_classifications = []

    for p_idx in range(total_pages):
        res_footer = extract_footer_text(src_doc, page_index=p_idx, footer_height_percent=footer_pct, lang=ocr_lang)
        matched_doc, matched_code = match_footer_code(res_footer.get("text", ""), catalog)

        if matched_doc:
            doc_name = matched_doc.get("documento", matched_code)
            log(f"  Pág {p_idx + 1}/{total_pages}: Código '{matched_code}' -> '{doc_name}'")
        else:
            log(f"  Pág {p_idx + 1}/{total_pages}: ⚠️ Sin código reconocido.", "warning")

        page_classifications.append({
            "page_idx": p_idx,
            "doc_def": matched_doc,
            "code": matched_code,
            "raw_text": res_footer.get("text", "")
        })

    # 6. Agrupación por tipo de documento
    # Estructura: code -> {"doc_def": ..., "pages": [idx1, idx2...]}
    groups = {}
    unrecognized_pages = []

    for item in page_classifications:
        p_idx = item["page_idx"]
        doc_def = item["doc_def"]
        code = item["code"]

        if doc_def and code:
            if code not in groups:
                groups[code] = {
                    "doc_def": doc_def,
                    "pages": []
                }
            groups[code]["pages"].append(p_idx)
        else:
            unrecognized_pages.append(p_idx)

    # 7. Generación de los archivos PDF divididos
    created_files = []

    for code, g in groups.items():
        doc_def = g["doc_def"]
        pages = g["pages"]
        doc_name = sanitize_filename(doc_def.get("documento", code))
        is_inc = doc_def.get("is_incidencia", False)

        # Regla especial Incidencias: si hay más de 1 página, hacer varios PDFs independientes (1-3, 2-3, 3-3)
        if is_inc and len(pages) > 1:
            total_inc = len(pages)
            for k, p in enumerate(pages, start=1):
                # Nomenclatura elegida: 1-3
                parts = [doc_name, f"{k}-{total_inc}"]
                if col_e:
                    parts.append(col_e)
                if col_i:
                    parts.append(col_i)
                parts.append(col_h)
                out_name = " - ".join(parts) + ".pdf"
                out_path = os.path.join(dest_dir, out_name)

                # Si ya existiese, evitar sobreescritura accidental
                if os.path.exists(out_path):
                    base, ext = os.path.splitext(out_name)
                    out_path = os.path.join(dest_dir, f"{base}_{int(time.time())}{ext}")

                sub_doc = pdfium.PdfDocument.new()
                sub_doc.import_pages(src_doc, [p])
                sub_doc.save(out_path)
                sub_doc.close()

                created_files.append({
                    "code": code,
                    "doc_name": f"{doc_name} ({k}-{total_inc})",
                    "filename": os.path.basename(out_path),
                    "path": out_path,
                    "pages_count": 1,
                    "pages_list": [p + 1]
                })
                log(f"  📄 Creado: '{os.path.basename(out_path)}' (Pág {p + 1})")
        else:
            # Documento normal: agrupar todas las páginas en un único PDF
            parts = [doc_name]
            if col_e:
                parts.append(col_e)
            if col_i:
                parts.append(col_i)
            parts.append(col_h)
            out_name = " - ".join(parts) + ".pdf"
            out_path = os.path.join(dest_dir, out_name)

            if os.path.exists(out_path):
                base, ext = os.path.splitext(out_name)
                out_path = os.path.join(dest_dir, f"{base}_{int(time.time())}{ext}")

            sub_doc = pdfium.PdfDocument.new()
            sub_doc.import_pages(src_doc, pages)
            sub_doc.save(out_path)
            sub_doc.close()

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

    # Páginas no reconocidas
    alert_unrecognized = None
    if unrecognized_pages:
        unrec_pages_human = [p + 1 for p in unrecognized_pages]
        unrec_name = f"REVISAR_PAGINAS_NO_IDENTIFICADAS - {col_h}.pdf"
        unrec_path = os.path.join(dest_dir, unrec_name)

        if os.path.exists(unrec_path):
            base, ext = os.path.splitext(unrec_name)
            unrec_path = os.path.join(dest_dir, f"{base}_{int(time.time())}{ext}")

        sub_doc = pdfium.PdfDocument.new()
        sub_doc.import_pages(src_doc, unrecognized_pages)
        sub_doc.save(unrec_path)
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
    if os.path.exists(archive_dest):
        base_orig, ext_orig = os.path.splitext(fname)
        archive_dest = os.path.join(dest_dir, f"_ORIGINAL_{int(time.time())}_{fname}")

    try:
        shutil.move(pdf_path, archive_dest)
        log(f"Archivado escaneo original en: '{os.path.basename(archive_dest)}'")
    except Exception as e_move:
        try:
            shutil.copy2(pdf_path, archive_dest)
            os.remove(pdf_path)
            log(f"Archivado escaneo original en: '{os.path.basename(archive_dest)}'")
        except Exception as e_copy:
            log(f"No se pudo mover el PDF original: {e_copy}", "warning")

    full_alert = alert_unrecognized or alert_warning
    log(f"🎉 Lote de la OF '{of_str}' completado con éxito. Generados {len(created_files)} archivos PDF.", "success")

    return {
        "success": True,
        "of": of_str,
        "order_info": order_info,
        "destination_dir": dest_dir,
        "total_pages": total_pages,
        "created_files": created_files,
        "unrecognized_pages": [p + 1 for p in unrecognized_pages],
        "alert": full_alert,
        "archive_path": archive_dest
    }


class ScanningSession:
    """
    Gestor de escucha automática continua para la carpeta de escaneos.
    Monitoriza la llegada de nuevos PDFs por lotes, los procesa automáticamente
    y mantiene el historial de lotes y las alertas vivas para la interfaz.
    """
    def __init__(self, excel_manager: ExcelManager):
        self.excel_manager = excel_manager
        self.active = False
        self.worker_thread = None
        self.stop_event = threading.Event()
        self.lock = threading.Lock()
        
        self.events_log = []
        self.batches_history = []
        self.active_alerts = []
        self.current_processing_file = None
        self.last_batch = None
        self.processed_files_set = set()

    def log(self, message: str, level: str = "info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        entry = {"time": timestamp, "message": message, "level": level}
        with self.lock:
            self.events_log.append(entry)
            if len(self.events_log) > 100:
                self.events_log.pop(0)
        print(f"[{timestamp}] [{level.upper()}] {message}")

    def add_alert(self, title: str, message: str, level: str = "warning", dest_dir: str = ""):
        timestamp = datetime.now().strftime("%H:%M:%S")
        alert = {
            "id": int(time.time() * 1000),
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
                "last_batch": self.last_batch,
                "batches_history": self.batches_history[-20:], # Últimos 20 lotes
                "active_alerts": list(self.active_alerts),
                "events_log": list(self.events_log[-35:]),
                "total_batches_processed": len(self.batches_history)
            }

    def _monitor_loop(self):
        catalog = load_ocr_catalog()

        while not self.stop_event.is_set():
            try:
                cfg = load_config()
                escaneos_folder = cfg.get("escaneos_folder")

                if not os.path.exists(escaneos_folder):
                    time.sleep(2.0)
                    continue

                candidates = []
                with os.scandir(escaneos_folder) as entries:
                    for entry in entries:
                        if entry.is_file() and entry.name.lower().endswith(".pdf"):
                            # Ignorar archivos que empiecen por '_' (originales o temporales)
                            if entry.name.startswith("_"):
                                continue
                            if entry.path not in self.processed_files_set:
                                try:
                                    stat = entry.stat()
                                    candidates.append((stat.st_mtime, entry.path, entry.name))
                                except Exception:
                                    pass

                # Procesar por orden cronológico (el más antiguo primero)
                candidates.sort(key=lambda x: x[0])

                for mtime, fpath, fname in candidates:
                    if self.stop_event.is_set():
                        break

                    if not is_file_ready(fpath):
                        # Todavía está siendo escrito por el escáner
                        continue

                    self.current_processing_file = fname
                    self.log(f"📥 Nuevo PDF detectado en bandeja: '{fname}'. Procesando lote...")

                    batch_res = process_scanned_pdf(
                        pdf_path=fpath,
                        excel_manager=self.excel_manager,
                        catalog=catalog,
                        log_callback=self.log
                    )

                    self.processed_files_set.add(fpath)
                    self.current_processing_file = None

                    if batch_res["success"]:
                        self.last_batch = batch_res
                        batch_entry = {
                            "time": datetime.now().strftime("%H:%M:%S"),
                            "of": batch_res["of"],
                            "order_info": batch_res["order_info"],
                            "dest_dir": batch_res["destination_dir"],
                            "total_pages": batch_res["total_pages"],
                            "files_count": len(batch_res["created_files"]),
                            "created_files": batch_res["created_files"],
                            "alert": batch_res["alert"]
                        }
                        self.batches_history.append(batch_entry)

                        if batch_res.get("alert"):
                            self.add_alert(
                                title=f"Aviso en OF {batch_res['of']}",
                                message=batch_res["alert"],
                                level="warning",
                                dest_dir=batch_res["destination_dir"]
                            )

                        # Sonido de confirmación en Windows
                        try:
                            import winsound
                            winsound.MessageBeep(winsound.MB_ICONASTERISK)
                        except Exception:
                            pass
                    else:
                        # Error procesando (ej: no se detectó OF)
                        err_msg = batch_res.get("error", "Error desconocido")
                        self.add_alert(
                            title=f"Error procesando {fname}",
                            message=err_msg,
                            level="error"
                        )
                        self.log(f"Error en escaneo '{fname}': {err_msg}", "error")

            except Exception as e:
                self.log(f"Error en bucle de escaneos: {e}", "error")

            time.sleep(1.0)
