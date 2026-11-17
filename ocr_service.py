import os
import re
import asyncio
from typing import Dict, List, Optional, Tuple, Any, Union
from PIL import Image, ImageOps

# Verificación de librerías nativas
try:
    import pypdfium2 as pdfium
    PDFIUM_AVAILABLE = True
except ImportError:
    pdfium = None
    PDFIUM_AVAILABLE = False

try:
    import winocr
    WINOCR_AVAILABLE = True
except ImportError:
    winocr = None
    WINOCR_AVAILABLE = False


def is_ocr_available() -> Tuple[bool, str]:
    """Comprueba si el motor OCR nativo de Windows y el renderizador PDF están listos."""
    if not PDFIUM_AVAILABLE:
        return False, "pypdfium2 no está instalado."
    if not WINOCR_AVAILABLE:
        return False, "winocr (Windows.Media.Ocr) no está disponible."
    return True, "OCR nativo de Windows (Windows.Media.Ocr) activo y listo."


def _prepare_image_for_ocr(img: Image.Image) -> Image.Image:
    """
    Convierte imágenes a RGB con fondo blanco y asegura dimensiones mínimas
    y márgenes para que Windows Media OCR reconozca texto sin problemas.
    """
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        alpha = img.convert("RGBA")
        bg = Image.new("RGBA", alpha.size, (255, 255, 255, 255))
        bg.paste(alpha, mask=alpha.split()[3])
        rgb = bg.convert("RGB")
    else:
        rgb = img.convert("RGB")

    w, h = rgb.size
    # Si la franja es muy estrecha (menos de 60px de alto), duplicar tamaño
    if h < 60:
        scale_factor = max(2, int(60 / max(1, h)) + 1)
        rgb = rgb.resize((w * scale_factor, h * scale_factor), Image.Resampling.LANCZOS)

    # Añadir un margen blanco de 20px alrededor para evitar que las letras pegadas al borde se ignoren
    padded = ImageOps.expand(rgb, border=20, fill="white")
    return padded


def render_page_region(
    pdf_input: Union[str, Any],
    page_index: int = 0,
    region: str = "footer",
    percent: int = 15,
    scale: float = 2.5
) -> Optional[Image.Image]:
    """
    Renderiza la página indicada y recorta la franja superior (cabecera) o inferior (pie de página).
    Acepta tanto la ruta al PDF como una instancia abierta de pdfium.PdfDocument.
    """
    if not PDFIUM_AVAILABLE:
        return None

    close_doc_at_end = False
    doc = None
    if isinstance(pdf_input, str):
        if not os.path.exists(pdf_input):
            return None
        try:
            doc = pdfium.PdfDocument(pdf_input)
            close_doc_at_end = True
        except Exception as e:
            print(f"[ocr_service] Error al abrir PDF {pdf_input}: {e}")
            return None
    else:
        doc = pdf_input

    try:
        if len(doc) <= page_index:
            if close_doc_at_end:
                doc.close()
            return None

        page = doc[page_index]
        pil_img = page.render(scale=scale).to_pil()
        if close_doc_at_end:
            doc.close()

        width, height = pil_img.size
        pct = max(5, min(80, percent))

        if region.lower() == "header":
            crop_box = (0, 0, width, int(height * (pct / 100.0)))
        else: # footer
            crop_top = int(height * (1.0 - (pct / 100.0)))
            crop_box = (0, crop_top, width, height)

        crop_img = pil_img.crop(crop_box)
        return _prepare_image_for_ocr(crop_img)
    except Exception as e:
        print(f"[ocr_service] Error recortando región {region} (pág {page_index}): {e}")
        if close_doc_at_end and doc is not None:
            try:
                doc.close()
            except Exception:
                pass
        return None


def extract_text_from_image(img: Image.Image, lang: str = "es-ES") -> Dict[str, Any]:
    """Ejecuta OCR sobre una imagen PIL y retorna texto completo y líneas."""
    avail, msg = is_ocr_available()
    if not avail:
        return {"success": False, "error": msg, "text": "", "lines": []}

    try:
        prepared = _prepare_image_for_ocr(img)
        result = asyncio.run(winocr.recognize_pil(prepared, lang=lang))
        extracted_text = result.text.strip() if hasattr(result, "text") else ""
        lines = [line.text.strip() for line in getattr(result, "lines", []) if hasattr(line, "text")]
        return {
            "success": True,
            "text": extracted_text,
            "lines": lines,
            "error": None
        }
    except Exception as e:
        print(f"[ocr_service] Error reconociendo texto OCR: {e}")
        return {"success": False, "error": str(e), "text": "", "lines": []}


def extract_header_text(
    pdf_input: Union[str, Any],
    page_index: int = 0,
    header_height_percent: int = 30,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """Extrae el texto de la cabecera (aprox. 30% superior) de la página indicada."""
    img = render_page_region(
        pdf_input,
        page_index=page_index,
        region="header",
        percent=header_height_percent,
        scale=2.5
    )
    if img is None:
        return {"success": False, "error": "No se pudo extraer la cabecera", "text": "", "lines": []}
    return extract_text_from_image(img, lang=lang)


def extract_footer_text(
    pdf_input: Union[str, Any],
    page_index: int = 0,
    footer_height_percent: int = 15,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """Extrae el texto del pie de página (aprox. 15% inferior) de la página indicada."""
    img = render_page_region(
        pdf_input,
        page_index=page_index,
        region="footer",
        percent=footer_height_percent,
        scale=2.5
    )
    if img is None:
        return {"success": False, "error": "No se pudo extraer el pie de página", "text": "", "lines": []}
    return extract_text_from_image(img, lang=lang)


def extract_of_from_header(
    pdf_input: Union[str, Any],
    page_index: int = 0,
    header_height_percent: int = 30,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """
    Analiza la cabecera de la página indicada (normalmente la página 1) y busca la OF.
    Ejemplos esperados: AYC/MO/00560, CESA/MO/00154, LOESS/MO/00087, MES/MO/05310.
    Retorna la lista de candidatas detectadas y la mejor coincidencia.
    """
    ocr_res = extract_header_text(
        pdf_input,
        page_index=page_index,
        header_height_percent=header_height_percent,
        lang=lang
    )

    if not ocr_res["success"]:
        return {
            "success": False,
            "best_of": None,
            "detected_ofs": [],
            "raw_text": "",
            "error": ocr_res["error"]
        }

    raw_text = ocr_res["text"]
    lines = ocr_res.get("lines", [])
    
    # Normalizar espacios y saltos de línea
    clean_text = raw_text.replace('\r', ' ').replace('\n', ' ')

    # Patrón principal: [PREFIJO]/MO/[NUMERO]
    # Permite barras, contrabarras o pipes, y corrige posibles confusiones OCR (O por 0, I/l por 1)
    primary_pat = re.compile(
        r'\b([A-Za-z0-9_-]{2,10})\s*[/\\|]\s*MO\s*[/\\|]\s*([0-9OIl]{3,6})\b',
        re.IGNORECASE
    )

    detected = []
    for match in primary_pat.finditer(clean_text):
        prefix = match.group(1).upper()
        raw_num = match.group(2).upper()
        # Normalizar ceros y unos mal leídos en la sección numérica
        clean_num = raw_num.replace('O', '0').replace('I', '1').replace('L', '1')
        normalized_of = f"{prefix}/MO/{clean_num}"
        if normalized_of not in detected:
            detected.append(normalized_of)

    # Patrón secundario más permisivo si el primero no encontró nada
    if not detected:
        secondary_pat = re.compile(r'([A-Za-z0-9_-]{2,10})/MO/(\d{3,6})', re.IGNORECASE)
        normalized_str = clean_text.replace('\\', '/').replace('|', '/')
        for m in secondary_pat.finditer(normalized_str):
            cand = f"{m.group(1).upper()}/MO/{m.group(2)}"
            if cand not in detected:
                detected.append(cand)

    # Patrón terciario: búsqueda de /MO/ aislado
    if not detected:
        tertiary_pat = re.compile(r'/MO/([0-9OIl]{3,6})', re.IGNORECASE)
        for m in tertiary_pat.finditer(clean_text.replace('\\', '/').replace('|', '/')):
            clean_num = m.group(1).upper().replace('O', '0').replace('I', '1').replace('L', '1')
            detected.append(f"MO/{clean_num}")

    best_of = detected[0] if detected else None

    return {
        "success": True,
        "best_of": best_of,
        "detected_ofs": detected,
        "raw_text": raw_text,
        "lines": lines,
        "error": None if best_of else "No se detectó ningún patrón de OF en la cabecera"
    }


def match_footer_code(ocr_text: str, candidates: List[Dict[str, Any]]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Compara el texto detectado por el OCR en el pie de página con el catálogo de códigos.
    candidates: lista de dicts con al menos 'codigo' (o 'footer_code') y 'documento' (o 'name').
    Retorna (candidato_ganador, código_detectado) o (None, None).
    """
    if not ocr_text or not candidates:
        return None, None

    clean_text = ocr_text.upper()
    compact_text = re.sub(r'[\s\-_]', '', clean_text)

    # Ordenar candidatos dando prioridad a códigos más largos
    def code_len(c):
        code = str(c.get("codigo") or c.get("footer_code") or "")
        return len(code)

    sorted_candidates = sorted(candidates, key=code_len, reverse=True)

    for item in sorted_candidates:
        raw_code = str(item.get("codigo") or item.get("footer_code") or "").strip().upper()
        if not raw_code:
            continue

        escaped = re.escape(raw_code)

        # 1. Búsqueda exacta respetando límites alfanuméricos
        if re.search(rf'(?<![A-Z0-9]){escaped}(?![A-Z0-9])', clean_text):
            return item, raw_code

        # 2. Tolerancia a espacios o guiones alternativos (ej: 'DC 0020' o 'DC_0020' por 'DC-0020')
        spaced_pattern = re.sub(r'[-_]', r'[-_\\s]+', escaped)
        if re.search(rf'(?<![A-Z0-9]){spaced_pattern}(?![A-Z0-9])', clean_text):
            return item, raw_code

        # 3. Coincidencia compacta sin espacios ni guiones (ej: 'DC0020')
        compact_code = re.sub(r'[\s\-_]', '', raw_code)
        if len(compact_code) >= 4 and compact_code in compact_text:
            return item, raw_code

        # 4. En anexos complejos (ej: ANEXO 1_IT0001), permitir coincidencia con 'ANEXO 1' o 'IT0001'
        if "ANEXO" in raw_code:
            parts = raw_code.split('_')
            for part in parts:
                if len(part) >= 4 and re.search(rf'(?<![A-Z0-9]){re.escape(part)}(?![A-Z0-9])', clean_text):
                    return item, raw_code

    return None, None


def identify_document_from_pdf(
    pdf_input: Union[str, Any],
    candidates: List[Dict[str, Any]],
    page_index: int = 0,
    footer_height_percent: int = 15,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """Analiza una página y determina qué documento es según el pie de página."""
    ocr_result = extract_footer_text(
        pdf_input,
        page_index=page_index,
        footer_height_percent=footer_height_percent,
        lang=lang
    )

    if not ocr_result["success"]:
        return {
            "matched": False,
            "document": None,
            "code": None,
            "raw_text": "",
            "error": ocr_result["error"]
        }

    raw_text = ocr_result["text"]
    matched_doc, detected_code = match_footer_code(raw_text, candidates)

    return {
        "matched": matched_doc is not None,
        "document": matched_doc,
        "code": detected_code,
        "raw_text": raw_text,
        "lines": ocr_result.get("lines", []),
        "error": None
    }


def load_ocr_catalog(excel_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Carga el catálogo oficial de tipos de documentos desde codigosOCR.xlsx.
    Retorna una lista de dicts:
    [{'id': idx, 'codigo': 'DC-0020', 'documento': 'Control Estado (incidencias-Materiales)', 'is_incidencia': True}, ...]
    """
    if not excel_path:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(base_dir, "codigosOCR.xlsx")

    if not os.path.exists(excel_path):
        print(f"[ocr_service] No se encontró el catálogo de códigos en: {excel_path}")
        return []

    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        ws = wb.active
        catalog = []
        for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
            if not row or len(row) < 2:
                continue
            code = str(row[0]).strip() if row[0] is not None else ""
            doc_name = str(row[1]).strip() if row[1] is not None else ""
            if code and doc_name:
                is_inc = (code.upper() == "DC-0020" or "INCIDENCIA" in doc_name.upper())
                catalog.append({
                    "id": idx,
                    "codigo": code,
                    "documento": doc_name,
                    "footer_code": code,
                    "name": doc_name,
                    "is_incidencia": is_inc
                })
        wb.close()
        return catalog
    except Exception as e:
        print(f"[ocr_service] Error leyendo {excel_path}: {e}")
        return []


def analyze_batch_pdf(
    pdf_path: str,
    catalog: Optional[List[Dict[str, Any]]] = None,
    header_height_percent: int = 30,
    footer_height_percent: int = 15,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """
    Realiza el análisis completo de un PDF multipágina escaneado:
    1. Abre el PDF con pdfium.
    2. En la página 0 (primera hoja), hace OCR de cabecera y extrae la OF.
    3. En cada una de las páginas (0 a N-1), hace OCR de pie de página y clasifica según codigosOCR.xlsx.
    4. Retorna el resultado estructurado para proceder al corte y agrupación.
    """
    if not PDFIUM_AVAILABLE or not os.path.exists(pdf_path):
        return {"success": False, "error": f"Archivo PDF no accesible: {pdf_path}"}

    if not catalog:
        catalog = load_ocr_catalog()

    try:
        doc = pdfium.PdfDocument(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            doc.close()
            return {"success": False, "error": "El documento PDF está vacío (0 páginas)."}

        # 1. OCR Cabecera página 0 -> Detectar OF
        of_res = extract_of_from_header(
            doc,
            page_index=0,
            header_height_percent=header_height_percent,
            lang=lang
        )
        detected_of = of_res.get("best_of")

        # 2. OCR Pie de página para todas las hojas
        pages_analysis = []
        for p_idx in range(total_pages):
            doc_id_res = identify_document_from_pdf(
                doc,
                candidates=catalog,
                page_index=p_idx,
                footer_height_percent=footer_height_percent,
                lang=lang
            )
            pages_analysis.append({
                "page_index": p_idx,
                "page_number": p_idx + 1,
                "matched": doc_id_res["matched"],
                "document": doc_id_res["document"],
                "code": doc_id_res["code"],
                "raw_footer": doc_id_res.get("raw_text", ""),
                "is_incidencia": bool(doc_id_res["document"] and doc_id_res["document"].get("is_incidencia"))
            })

        doc.close()

        return {
            "success": True,
            "total_pages": total_pages,
            "of_detected": detected_of,
            "of_analysis": of_res,
            "pages": pages_analysis,
            "error": None
        }
    except Exception as e:
        print(f"[ocr_service] Error analizando PDF por lotes: {e}")
        return {"success": False, "error": str(e)}

