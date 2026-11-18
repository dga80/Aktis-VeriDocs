import os
import re
import asyncio
import unicodedata
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


def is_page_blank(
    pdf_input: Union[str, Any],
    page_index: int = 0,
    dark_pixel_threshold: int = 220,
    max_dark_ratio: float = 0.0025
) -> bool:
    """
    Determina si una página escaneada está prácticamente en blanco (típico reverso dúplex).
    Primero comprueba texto embebido nativo; si no hay, renderiza a 72 DPI y evalúa
    la proporción de píxeles oscuros (< 0.25% se considera hoja en blanco).
    """
    if not PDFIUM_AVAILABLE:
        return False

    close_doc_at_end = False
    doc = None
    try:
        if isinstance(pdf_input, str):
            doc = pdfium.PdfDocument(pdf_input)
            close_doc_at_end = True
        else:
            doc = pdf_input

        if page_index < 0 or page_index >= len(doc):
            return False

        page = doc[page_index]
        # 1. Comprobar texto vectorial nativo
        try:
            tp = page.get_textpage()
            txt = tp.get_text_range().strip()
            if len(txt) > 25:
                return False
        except Exception:
            pass

        # 2. Renderizado ultrarrápido a escala 1.0 (72 DPI)
        bitmap = page.render(scale=1.0)
        pil_img = bitmap.to_pil()
        gray = pil_img.convert("L")

        import numpy as np
        arr = np.array(gray)
        dark_pixels = np.count_nonzero(arr < dark_pixel_threshold)
        total_pixels = arr.size
        if total_pixels == 0:
            return True

        ratio = dark_pixels / total_pixels
        return ratio < max_dark_ratio
    except Exception as e:
        print(f"[ocr_service] Error comprobando página en blanco (pág {page_index}): {e}")
        return False
    finally:
        if close_doc_at_end and doc is not None:
            try:
                doc.close()
            except Exception:
                pass


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


def _find_ofs_in_text(raw_text: str) -> List[str]:
    """Busca patrones de OF (ej: AYC/MO/00524) dentro de un texto extraído."""
    if not raw_text:
        return []

    clean_text = raw_text.replace('\r', ' ').replace('\n', ' ')

    # Patrón principal: [PREFIJO]/MO/[NUMERO]
    # Permite barras (/,\), pipes (|) o confusiones OCR verticales (I, 1, l, !)
    primary_pat = re.compile(
        r'\b([A-Za-z0-9_-]{2,8})\s*[/\\|I1l!]\s*MO\s*[/\\|I1l!]\s*([0-9OIl]{3,6})\b',
        re.IGNORECASE
    )

    detected = []
    for match in primary_pat.finditer(clean_text):
        prefix = match.group(1).upper()
        raw_num = match.group(2).upper()
        clean_num = raw_num.replace('O', '0').replace('I', '1').replace('L', '1')
        normalized_of = f"{prefix}/MO/{clean_num}"
        if normalized_of not in detected:
            detected.append(normalized_of)

    # Patrón secundario más permisivo
    if not detected:
        secondary_pat = re.compile(r'([A-Za-z0-9_-]{2,8})[/\\|I1l!]+MO[/\\|I1l!]+(\d{3,6})', re.IGNORECASE)
        for m in secondary_pat.finditer(clean_text):
            cand = f"{m.group(1).upper()}/MO/{m.group(2)}"
            if cand not in detected:
                detected.append(cand)

    # Patrón terciario: búsqueda de MO/[NUMERO]
    if not detected:
        tertiary_pat = re.compile(r'MO\s*[/\\|I1l!]\s*([0-9OIl]{3,6})', re.IGNORECASE)
        for m in tertiary_pat.finditer(clean_text):
            clean_num = m.group(1).upper().replace('O', '0').replace('I', '1').replace('L', '1')
            detected.append(f"MO/{clean_num}")

    return detected


def extract_of_from_header(
    pdf_input: Union[str, Any],
    page_index: int = 0,
    header_height_percent: int = 30,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """
    Analiza la cabecera de la página indicada (normalmente la página 1) y busca la OF.
    Ejemplos esperados: AYC/MO/00560, CESA/MO/00154, LOESS/MO/00087, MES/MO/05310.
    Si la página está apaisada (landscape) o rotada, prueba automáticamente
    las orientaciones de 90°, 270° y 180° para asegurar la detección.
    Retorna la lista de candidatas detectadas y la mejor coincidencia.
    """
    # 1. Intento estándar a 0°
    ocr_res = extract_header_text(
        pdf_input,
        page_index=page_index,
        header_height_percent=header_height_percent,
        lang=lang
    )

    detected = []
    raw_text = ""
    lines = []

    if ocr_res.get("success"):
        raw_text = ocr_res.get("text", "")
        lines = ocr_res.get("lines", [])
        detected = _find_ofs_in_text(raw_text)

    # 2. Si no se encontró OF, probar en orientaciones rotadas (hojas apaisadas)
    if not detected and PDFIUM_AVAILABLE and WINOCR_AVAILABLE:
        close_doc = False
        doc = None
        try:
            if isinstance(pdf_input, str):
                if os.path.exists(pdf_input):
                    doc = pdfium.PdfDocument(pdf_input)
                    close_doc = True
            else:
                doc = pdf_input

            if doc and page_index < len(doc):
                page = doc[page_index]
                pil_img = page.render(scale=2.5).to_pil()
                
                # Probar 90° horario (Image.Transpose.ROTATE_270)
                rot90 = pil_img.transpose(Image.Transpose.ROTATE_270)
                rw, rh = rot90.size
                head90 = rot90.crop((0, 0, rw, int(rh * 0.40)))
                res90 = extract_text_from_image(head90, lang=lang)
                if res90.get("success"):
                    det90 = _find_ofs_in_text(res90.get("text", ""))
                    if det90:
                        detected = det90
                        raw_text = res90.get("text", "")
                        lines = res90.get("lines", [])

                # Si aún no, probar 270° horario (Image.Transpose.ROTATE_90)
                if not detected:
                    rot270 = pil_img.transpose(Image.Transpose.ROTATE_90)
                    rw2, rh2 = rot270.size
                    head270 = rot270.crop((0, 0, rw2, int(rh2 * 0.40)))
                    res270 = extract_text_from_image(head270, lang=lang)
                    if res270.get("success"):
                        det270 = _find_ofs_in_text(res270.get("text", ""))
                        if det270:
                            detected = det270
                            raw_text = res270.get("text", "")
                            lines = res270.get("lines", [])
        except Exception as e_rot_of:
            print(f"[ocr_service] Error buscando OF en orientaciones rotadas: {e_rot_of}")
        finally:
            if close_doc and doc:
                try:
                    doc.close()
                except Exception:
                    pass

    best_of = detected[0] if detected else None

    return {
        "success": bool(best_of),
        "best_of": best_of,
        "detected_ofs": detected,
        "raw_text": raw_text,
        "lines": lines,
        "error": None if best_of else "No se detectó ningún patrón de OF en la cabecera"
    }


DOCUMENT_KEYWORDS: Dict[str, List[str]] = {
    "DC-0018": ["CARROS MAT", "CARROS MATERIAL", "CARRO MATERIAL"],
    "DC-0019": ["ORDEN DE FABRICACION", "ORDEN DE FABRICACIÓN"],
    "DC-0020": ["CONTROL ESTADO (INCIDENCIAS", "INCIDENCIAS-MATERIALES", "INCIDENCIAS MATERIALES"],
    "DC-0021": ["CONTROL DE PENDIENTES", "TRABAJOS PENDIENTES", "CONTROL PENDIENTES"],
    "DC-0022": ["ESTADO ENTRADA VERIFICA", "ENTRADA VERIFICA", "ENTRADA VERIFICACION"],
    "DC-0023": ["PESAJES", "HOJA DE PESAJES"],
    "DC-0024": ["EXPEDICIONES-SUBMER", "EXPEDICIONES SUBMER", "SUBMER"],
    "DC-0025": ["D.CONFORMIDAD UE", "CONFORMIDAD UE", "DECLARACION DE CONFORMIDAD UE"],
    "DC-0026": ["DECL.CONFORMIDAD UL508A", "UL508A", "CONFORMIDAD UL508A", "CONFORMIDAD UL"],
    "DC-0027": ["CERTIFICADO GARANTIA MCD", "CERTIFICADO GARANTÍA MCD", "GARANTIA MCD", "GARANTÍA MCD"],
    "DC-0028": ["VERIFICACION IEC - (ES)", "VERIFICACIÓN IEC - (ES)", "VERIFICACION IEC", "VERIFICACIÓN IEC"],
    "DC-0029": [
        "CONTROLES INTERNOS GENERAL",
        "CONTROLES INTERNOS",
        "INTERNOS GENERAL",
        "DOCUMENTACION INTERNA A RELLENAR",
        "DOCUMENTACION INTERNA",
        "DOCUMENTACIÓN INTERNA"
    ],
    "DC-0030": ["VERIFICACION IEC - (EN)", "ROUTINE VERIFICATION", "VERIFICATION IEC"],
    "DC-0031": ["VERIFICACION UL - (EN)", "VERIFICATION UL"],
    "DC-0032": ["INFORME TECNICO INGENIERO", "INFORME TÉCNICO INGENIERO"],
    "DC-0033": ["INFORME TECNICO OBRA", "INFORME TÉCNICO OBRA"],
    "DC-0034": ["PRUEBAS FAT_(ES)", "PRUEBAS FAT (ES)", "PRUEBAS FAT"],
    "DC-0035": ["PRUEBAS FAT_(EN)", "PRUEBAS FAT (EN)", "FAT TEST"],
    "DC-0036": ["PRUEBAS FAT _CLIENTE_(ES)", "FAT CLIENTE"],
    "DC-0037": ["PRUEBAS FAT _CLIENTE_(EN)", "FAT CLIENT"],
    "DC-0038": ["HOJA ALMACEN", "HOJA ALMACÉN"],
    "DC-0039": ["HOJA REVISION CUBICULOS", "HOJA REVISIÓN CUBÍCULOS", "REVISION CUBICULOS", "REVISIÓN CUBÍCULOS"],
    "ANEXO 1_IT0001": ["ANEXO 1", "TAREAS DE VERIFICACION", "TAREAS DE VERIFICACIÓN"],
    "ANEXO 2_IT0001": ["ANEXO 2", "CHECK MCD", "H.CHECK MCD"]
}


def suppress_table_lines(pil_img: Image.Image) -> Image.Image:
    """
    Detecta y suprime líneas horizontales continuas oscuras (bordes de tablas)
    que suelen tocar o cortar los caracteres de código y texto en pies de página.
    """
    try:
        import numpy as np
        gray = pil_img.convert('L')
        arr = np.array(gray)
        h, w = arr.shape
        if h < 10 or w < 30:
            return pil_img

        dark_mask = arr < 160
        row_ratios = dark_mask.mean(axis=1)
        lines_to_clear = [r for r in range(h) if row_ratios[r] > 0.30]
        if not lines_to_clear:
            return pil_img

        cleaned = arr.copy()
        for r in lines_to_clear:
            r_min = max(0, r - 1)
            r_max = min(h, r + 2)
            cleaned[r_min:r_max, :] = 255

        return Image.fromarray(cleaned).convert('RGB')
    except Exception:
        return pil_img


def score_document_candidate(clean_text: str, compact_text: str, raw_code: str, doc_name: str) -> int:
    """Evalúa la coincidencia de un candidato con el texto OCR asignando puntuación de 0 a 100."""
    escaped = re.escape(raw_code)

    # 1. Búsqueda exacta respetando límites alfanuméricos
    if re.search(rf'(?<![A-Z0-9]){escaped}(?![A-Z0-9])', clean_text):
        return 100

    # 2. Tolerancia a espacios o guiones alternativos (ej: 'DC 0021' o 'DC_0021')
    spaced = re.sub(r'[-_]', r'[-_\\s]+', escaped)
    if re.search(rf'(?<![A-Z0-9]){spaced}(?![A-Z0-9])', clean_text):
        return 90

    # 3. Coincidencia compacta sin espacios ni guiones (ej: 'DC0021')
    compact_code = re.sub(r'[\s\-_]', '', raw_code)
    if len(compact_code) >= 4 and compact_code in compact_text:
        return 85

    # 4. Código numérico con contexto de documento (ej: Documento: 0021, DC-0O21)
    m = re.match(r'DC-(\d{4})', raw_code)
    if m:
        num = m.group(1)
        num_pat = num.replace('0', '[0O]')
        if re.search(rf'(?:DC|C|DOC|DOCUMENTO)[\s\-_.:]*{num_pat}', clean_text):
            return 80

    # 5. Palabras clave y título del documento
    import difflib
    kws = DOCUMENT_KEYWORDS.get(raw_code, [])
    if doc_name and doc_name.upper() not in [k.upper() for k in kws]:
        kws = list(kws) + [doc_name.upper()]

    score = 0
    for kw in kws:
        kw_clean = kw.upper()
        if kw_clean in clean_text:
            score = max(score, 75)
        elif re.sub(r'[\s\-_()]', '', kw_clean) in compact_text:
            score = max(score, 70)
        else:
            kw_words = [w for w in re.findall(r'[A-Z0-9ÁÉÍÓÚ]{4,}', kw_clean)]
            if len(kw_words) >= 2:
                matched_w = 0
                for kw_w in kw_words:
                    if any(kw_w in tw or tw in kw_w or difflib.SequenceMatcher(None, kw_w, tw).ratio() >= 0.75 
                           for tw in re.findall(r'[A-Z0-9ÁÉÍÓÚ]{3,}', clean_text)):
                        matched_w += 1
                if matched_w == len(kw_words):
                    score = max(score, 65)

    return score


def match_footer_code(ocr_text: str, candidates: List[Dict[str, Any]]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Compara el texto detectado por el OCR en el pie de página con el catálogo de códigos.
    Utiliza búsqueda exacta de código, variantes con tolerancia a separadores,
    patrón numérico con prefijo de documento y coincidencia semántica por título/palabras clave.
    Retorna (candidato_ganador, código_detectado) o (None, None).
    """
    if not ocr_text or not candidates:
        return None, None

    clean_text = ocr_text.upper()
    compact_text = re.sub(r'[\s\-_]', '', clean_text)

    best_item = None
    best_code = None
    best_score = 0

    for item in candidates:
        code = str(item.get("codigo") or item.get("footer_code") or "").strip().upper()
        if not code:
            continue
        name = str(item.get("documento") or item.get("name") or "").strip().upper()
        score = score_document_candidate(clean_text, compact_text, code, name)
        if score > best_score:
            best_score = score
            best_item = item
            best_code = code

    if best_score >= 60:
        return best_item, best_code
    return None, None


def extract_footer_and_detect_code(
    pdf_input: Union[str, Any],
    catalog: List[Dict[str, Any]],
    page_index: int = 0,
    footer_height_percent: int = 15,
    lang: str = "es-ES"
) -> Dict[str, Any]:
    """
    Analiza una página para identificar el código de documento oficial en el pie.
    Incluye detección automática para hojas apaisadas (landscape) y volteadas:
    1. Intento estándar: Franja inferior (0°).
    2. Si no coincide, analiza la orientación apaisada 90° horario (pie y cabecera).
    3. Si no coincide, analiza la orientación apaisada 270° horario (pie y cabecera).
    4. Si aún no coincide, analiza la franja superior invertida (180°).
    
    Retorna un diccionario con:
    - matched_doc: Dict o None
    - matched_code: str o None
    - rotation_needed: int (0, 90, 180, 270)
    - text: str
    - strategy: str
    """
    if not PDFIUM_AVAILABLE or not WINOCR_AVAILABLE:
        return {"matched_doc": None, "matched_code": None, "rotation_needed": 0, "text": "", "strategy": "ocr_unavailable"}

    close_doc_at_end = False
    doc = None
    if isinstance(pdf_input, str):
        if not os.path.exists(pdf_input):
            return {"matched_doc": None, "matched_code": None, "rotation_needed": 0, "text": "", "strategy": "file_not_found"}
        try:
            doc = pdfium.PdfDocument(pdf_input)
            close_doc_at_end = True
        except Exception as e:
            return {"matched_doc": None, "matched_code": None, "rotation_needed": 0, "text": "", "strategy": f"open_error: {e}"}
    else:
        doc = pdf_input

    try:
        if len(doc) <= page_index:
            return {"matched_doc": None, "matched_code": None, "rotation_needed": 0, "text": "", "strategy": "index_out_of_range"}

        # 1. Renderizado a escala óptima (2.5) para nitidez OCR
        page = doc[page_index]
        pil_img = page.render(scale=2.5).to_pil()
        width, height = pil_img.size

        # --- ESTRATEGIA 1: Pie de página estándar inferior (0°) ---
        pct = max(5, min(80, footer_height_percent))
        crop_top = int(height * (1.0 - (pct / 100.0)))
        bottom_img = pil_img.crop((0, crop_top, width, height))

        res_bottom = extract_text_from_image(bottom_img, lang=lang)
        matched_doc, matched_code = match_footer_code(res_bottom.get("text", ""), catalog)

        # 1.1 Si no coincide a escala estándar, reescalar 2x (vital para fuentes pequeñas < 6pt)
        if not matched_doc:
            bw, bh = bottom_img.size
            scaled_bottom = bottom_img.resize((bw * 2, bh * 2), Image.Resampling.LANCZOS)
            res_bottom_scaled = extract_text_from_image(scaled_bottom, lang=lang)
            matched_doc, matched_code = match_footer_code(res_bottom_scaled.get("text", ""), catalog)
            if matched_doc:
                res_bottom = res_bottom_scaled

        # 1.2 Si aún no coincide, probar con supresión de bordes de tabla + 2x
        if not matched_doc:
            cleaned_bottom = suppress_table_lines(bottom_img)
            cbw, cbh = cleaned_bottom.size
            cleaned_scaled = cleaned_bottom.resize((cbw * 2, cbh * 2), Image.Resampling.LANCZOS)
            res_bottom_cleaned = extract_text_from_image(cleaned_scaled, lang=lang)
            matched_doc, matched_code = match_footer_code(res_bottom_cleaned.get("text", ""), catalog)
            if matched_doc:
                res_bottom = res_bottom_cleaned

        if matched_doc:
            return {
                "matched_doc": matched_doc,
                "matched_code": matched_code,
                "rotation_needed": 0,
                "text": res_bottom.get("text", ""),
                "strategy": "bottom_standard"
            }

        # Comprobar si la página está en blanco antes de gastar recursos en rotaciones
        if is_page_blank(doc, page_index=page_index):
            return {
                "matched_doc": None,
                "matched_code": "BLANK",
                "rotation_needed": 0,
                "text": "",
                "strategy": "blank_page"
            }

        # --- ESTRATEGIA 2: Hoja apaisada rotada 90° horario (Image.Transpose.ROTATE_270) ---
        margin_pct = 0.22
        right_x = int(width * (1.0 - margin_pct))
        right_strip = pil_img.crop((right_x, 0, width, height))
        rot270_strip = right_strip.transpose(Image.Transpose.ROTATE_270)
        rw270, rh270 = rot270_strip.size
        scaled_r270 = rot270_strip.resize((rw270 * 2, rh270 * 2), Image.Resampling.LANCZOS)
        
        res_r270 = extract_text_from_image(scaled_r270, lang=lang)
        matched_doc, matched_code = match_footer_code(res_r270.get("text", ""), catalog)
        if not matched_doc:
            res_r270_clean = extract_text_from_image(suppress_table_lines(scaled_r270), lang=lang)
            matched_doc, matched_code = match_footer_code(res_r270_clean.get("text", ""), catalog)
            if matched_doc:
                res_r270 = res_r270_clean

        if matched_doc:
            return {
                "matched_doc": matched_doc,
                "matched_code": matched_code,
                "rotation_needed": 90,
                "text": res_r270.get("text", ""),
                "strategy": "right_margin_rot90"
            }

        # Cabecera de la orientación rotada 90° (por si el pie está degradado pero el título es legible)
        rot270_page = pil_img.transpose(Image.Transpose.ROTATE_270)
        rpw270, rph270 = rot270_page.size
        head270 = rot270_page.crop((0, 0, rpw270, int(rph270 * 0.25)))
        res_head270 = extract_text_from_image(head270, lang=lang)
        matched_doc, matched_code = match_footer_code(res_head270.get("text", ""), catalog)
        if matched_doc:
            return {
                "matched_doc": matched_doc,
                "matched_code": matched_code,
                "rotation_needed": 90,
                "text": res_head270.get("text", ""),
                "strategy": "rot90_header"
            }

        # --- ESTRATEGIA 3: Hoja apaisada rotada 270° horario (Image.Transpose.ROTATE_90) ---
        left_w = int(width * margin_pct)
        left_strip = pil_img.crop((0, 0, left_w, height))
        rot90_strip = left_strip.transpose(Image.Transpose.ROTATE_90)
        rw90, rh90 = rot90_strip.size
        scaled_r90 = rot90_strip.resize((rw90 * 2, rh90 * 2), Image.Resampling.LANCZOS)
        
        res_r90 = extract_text_from_image(scaled_r90, lang=lang)
        matched_doc, matched_code = match_footer_code(res_r90.get("text", ""), catalog)
        if not matched_doc:
            res_r90_clean = extract_text_from_image(suppress_table_lines(scaled_r90), lang=lang)
            matched_doc, matched_code = match_footer_code(res_r90_clean.get("text", ""), catalog)
            if matched_doc:
                res_r90 = res_r90_clean

        if matched_doc:
            return {
                "matched_doc": matched_doc,
                "matched_code": matched_code,
                "rotation_needed": 270,
                "text": res_r90.get("text", ""),
                "strategy": "left_margin_rot270"
            }

        # Cabecera de la orientación rotada 270°
        rot90_page = pil_img.transpose(Image.Transpose.ROTATE_90)
        rpw90, rph90 = rot90_page.size
        head90 = rot90_page.crop((0, 0, rpw90, int(rph90 * 0.25)))
        res_head90 = extract_text_from_image(head90, lang=lang)
        matched_doc, matched_code = match_footer_code(res_head90.get("text", ""), catalog)
        if matched_doc:
            return {
                "matched_doc": matched_doc,
                "matched_code": matched_code,
                "rotation_needed": 270,
                "text": res_head90.get("text", ""),
                "strategy": "rot270_header"
            }

        # --- ESTRATEGIA 4: Margen superior invertido 180° ---
        top_h = int(height * (pct / 100.0))
        top_strip = pil_img.crop((0, 0, width, top_h))
        rot_top180 = top_strip.transpose(Image.Transpose.ROTATE_180)
        tw180, th180 = rot_top180.size
        scaled_top180 = rot_top180.resize((tw180 * 2, th180 * 2), Image.Resampling.LANCZOS)
        
        res_top180 = extract_text_from_image(scaled_top180, lang=lang)
        matched_doc, matched_code = match_footer_code(res_top180.get("text", ""), catalog)
        if not matched_doc:
            res_top180_clean = extract_text_from_image(suppress_table_lines(scaled_top180), lang=lang)
            matched_doc, matched_code = match_footer_code(res_top180_clean.get("text", ""), catalog)
            if matched_doc:
                res_top180 = res_top180_clean

        if matched_doc:
            return {
                "matched_doc": matched_doc,
                "matched_code": matched_code,
                "rotation_needed": 180,
                "text": res_top180.get("text", ""),
                "strategy": "top_inverted_180"
            }

        return {
            "matched_doc": None,
            "matched_code": None,
            "rotation_needed": 0,
            "text": res_bottom.get("text", ""),
            "strategy": "not_found"
        }

    except Exception as e:
        print(f"[extract_footer_and_detect_code] Error: {e}")
        return {
            "matched_doc": None,
            "matched_code": None,
            "rotation_needed": 0,
            "text": "",
            "strategy": f"exception: {e}"
        }
    finally:
        if close_doc_at_end and doc is not None:
            try:
                doc.close()
            except Exception:
                pass


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
                is_inc = (code.upper() in ("DC-0020", "DC-0022") or "INCIDENCIA" in doc_name.upper())
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

