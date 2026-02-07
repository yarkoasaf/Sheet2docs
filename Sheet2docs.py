#made by Yarko Bahamonde
#distribution outside of my github page is prohibited
#can be used in non-profit projects
#the complete or partial use of the code for other projects needs to be credited

import os
from docx import Document
from openpyxl import load_workbook
from django.conf import settings
from .models import Trabajador
import openpyxl


#==============================================================================================================#

import glob

def _ls(dirpath):
    """Return a sorted list of entries in dirpath (name + size), or a reason if it fails."""
    try:
        entries = []
        for name in sorted(os.listdir(dirpath)):
            p = os.path.join(dirpath, name)
            try:
                size = os.path.getsize(p) if os.path.isfile(p) else -1
            except Exception:
                size = -1
            entries.append(f"{name}{' ('+str(size)+' B)' if size>=0 else ''}")
        return entries
    except Exception as e:
        return [f"[no se pudo listar: {e}]"]

def _strip_leading_media(relpath: str) -> str:
    """Strip a leading 'media/' or 'media\\' to avoid MEDIA_ROOT/media/... duplication."""
    p = relpath.lstrip("/\\")
    low = p.lower()
    if low.startswith("media/") or low.startswith("media\\"):
        return p.split("/", 1)[1] if "/" in p else p.split("\\", 1)[1]
    return p

def _resolve_template_path(template_hint: str = "plantillas/'template Informe NeuroBit Feria --nombre--.docx'") -> str:
    """
    Resolve a template path robustly:
    - If absolute and exists → return it.
    - If relative, join with MEDIA_ROOT (stripping any leading 'media/').
    - If that exact file doesn't exist, search MEDIA_ROOT/plantillas/*.docx
        and try to pick a sensible candidate.
    """
    # 1) If absolute and exists
    if os.path.isabs(template_hint) and os.path.exists(template_hint):
        print(f"[DEBUG] Using absolute template: {template_hint}")
        return template_hint

    # 2) Normalize relative under MEDIA_ROOT
    hint = _strip_leading_media(template_hint or "")
    candidate = os.path.join(settings.MEDIA_ROOT, hint)
    plantillas_dir = os.path.join(settings.MEDIA_ROOT, "plantillas")

    # If exact candidate exists, done
    if os.path.exists(candidate):
        print(f"[DEBUG] Using exact relative template under MEDIA_ROOT: {candidate}")
        return candidate

    # 3) Fallback: search in plantillas dir
    print(f"[DEBUG] Template not found at expected path: {candidate}")
    print(f"[DEBUG] MEDIA_ROOT: {settings.MEDIA_ROOT} exists={os.path.exists(settings.MEDIA_ROOT)}")
    print(f"[DEBUG] Listing {plantillas_dir}: {_ls(plantillas_dir) if os.path.exists(plantillas_dir) else '[no existe]'}")

    if not os.path.exists(plantillas_dir):
        raise Sheet2DocsError(
            f"[TEMPLATE_DIR_MISSING] No existe la carpeta de plantillas: {plantillas_dir}"
        )

    all_docx = sorted(glob.glob(os.path.join(plantillas_dir, "*.docx")))
    if not all_docx:
        raise Sheet2DocsError(
            f"[NO_DOCX_IN_PLANTILLAS] No se encontraron .docx en {plantillas_dir}. "
            f"Contenido: {_ls(plantillas_dir)}"
        )

    # Try to pick the best match by simple heuristics
    base_lower = os.path.basename(hint).lower()
    preferred = [p for p in all_docx if "template" in os.path.basename(p).lower()]
    preferred = [p for p in preferred if "feria" in os.path.basename(p).lower()] or preferred
    preferred = [p for p in preferred if "--nombre--" in os.path.basename(p).lower()] or preferred
    # If still many, try name similarity
    if len(preferred) > 1 and base_lower:
        preferred = sorted(preferred, key=lambda p: abs(len(os.path.basename(p).lower()) - len(base_lower)))

    if len(preferred) == 1:
        print(f"[DEBUG] Using discovered template: {preferred[0]}")
        return preferred[0]

    # If multiple remain, surface a helpful error
    raise Sheet2DocsError(
        "[AMBIGUOUS_TEMPLATES] Se encontraron múltiples candidatos. "
        f"Por favor especifica el nombre exacto o renombra para que sea único.\n"
        f"Candidatos: {[os.path.basename(p) for p in preferred]}\n"
        f"Todos los .docx: {[os.path.basename(p) for p in all_docx]}"
    )

#==============================================================================================================#

#==============================================================================================================#

import traceback
from openpyxl.utils.exceptions import InvalidFileException

class Sheet2DocsError(RuntimeError):
    """Explicit, debuggable errors for sheet→doc pipeline."""
    pass

def _ensure(cond, code: str, msg: str):
    """Raise a labeled error with a readable message."""
    if not cond:
        raise Sheet2DocsError(f"[{code}] {msg}")

def _abs_media_path(rel_or_abs: str) -> str:
    """
    Return an absolute path under MEDIA_ROOT when given a relative path.
    If already absolute, return it as-is. Also avoids 'media/media/...'.
    """
    _ensure(rel_or_abs, "XLSX_PATH_EMPTY", "ruta_xlsx está vacío o None")
    p = rel_or_abs
    if os.path.isabs(p):
        return p
    # strip leading slashes and a leading 'media/' if present
    p = p.lstrip("/\\")
    if p.lower().startswith("media/") or p.lower().startswith("media\\"):
        p = p.split("/", 1)[1] if "/" in p else p.split("\\", 1)[1]
    return os.path.join(settings.MEDIA_ROOT, p)

def _debug_headers(ws, label="WS HEADERS"):
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    print(f"[DEBUG] {label}: {[h for h in headers if h not in (None, '')]}")
    return headers


#==============================================================================================================#

#==================================== el amigo feria ==========================================================#

def _to_str_with_locale(v, decimal_comma: bool = True):
    """Return value as string; if it's numeric, optionally use comma decimal."""
    if v is None:
        return ""
    s = str(v)
    if decimal_comma:
        try:
            # only swap if it looks like a float with dot
            float(s.replace(",", "."))  # probe
            s = s.replace(".", ",")
        except Exception:
            pass
    return s

def _find_col_by_header(ws, header_name: str):
    """Return 1-based column index of a header in row 1, or None if not found."""
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=c).value == header_name:
            return c
    return None

def append_or_update_datos_feria(ws, datos_feria: dict, decimal_comma: bool = True):
    """
    For each metric in datos_feria (e.g., 'atencion': [valor, interpretacion]),
    upsert two columns:
    - '<metric>'                      -> numeric value
    - '<metric> interpretacion'       -> interpretation text
    If the header already exists, update row 2 in place. Otherwise, append at the end.
    """
    for key, pair in datos_feria.items():
        if not isinstance(pair, (list, tuple)) or len(pair) != 2:
            # skip malformed entries rather than crashing the whole process
            print(f"[WARN] datos_feria['{key}'] debe ser [valor, interpretacion]. Se omite.")
            continue

        valor, interpretacion = pair
        valor_str = _to_str_with_locale(valor, decimal_comma)
        interp_header = f"{key} interpretacion"

        # 1) Upsert main value column
        col_idx = _find_col_by_header(ws, key)
        if col_idx is None:
            col_idx = ws.max_column + 1
            ws.cell(row=1, column=col_idx, value=key)
        ws.cell(row=2, column=col_idx, value=valor_str)

        # 2) Upsert interpretation column
        col_interp = _find_col_by_header(ws, interp_header)
        if col_interp is None:
            col_interp = ws.max_column + 1
            ws.cell(row=1, column=col_interp, value=interp_header)
        ws.cell(row=2, column=col_interp, value=interpretacion or "")


#==============================================================================================================#

#==============================================================================================================#

import re

# --- Extract <<placeholders>> from a .docx (body, tables, headers/footers, text boxes) ---
_PLACEHOLDER_RE = re.compile(r"<<\s*([^<>]+?)\s*>>")

def _extract_from_paragraphs(paragraphs):
    """Yield placeholders found in paragraphs; concatenates runs to survive Word splitting."""
    for p in paragraphs:
        text = "".join(run.text for run in p.runs)
        for m in _PLACEHOLDER_RE.finditer(text):
            yield m.group(1).strip()

def _extract_from_tables(tables):
    """Yield placeholders from (possibly nested) tables."""
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _extract_from_paragraphs(cell.paragraphs)
                yield from _extract_from_tables(cell.tables)

def list_docx_placeholders(template_path):
    """
    Return a list of unique placeholders (without << >>) found in a .docx template.
    Scans body, tables, headers, footers, and text boxes (w:txbxContent).
    """
    doc = Document(template_path)
    found = []

    # Body
    found.extend(_extract_from_paragraphs(doc.paragraphs))
    found.extend(_extract_from_tables(doc.tables))

    # Headers & footers
    for section in doc.sections:
        hdr, ftr = section.header, section.footer
        found.extend(_extract_from_paragraphs(hdr.paragraphs))
        found.extend(_extract_from_tables(hdr.tables))
        found.extend(_extract_from_paragraphs(ftr.paragraphs))
        found.extend(_extract_from_tables(ftr.tables))

    # Text boxes / shapes (w:txbxContent) – compatible with python-docx that lacks namespaces= on .xpath()
    root = doc.part.element
    WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Iterate all <w:txbxContent> elements regardless of where they live
    for txbx in root.iter(f"{{{WNS}}}txbxContent"):
        # Collect all <w:t> descendants inside this text box
        fragments = [t.text for t in txbx.iter(f"{{{WNS}}}t") if t.text]
        if fragments:
            text = "".join(fragments)
            for m in _PLACEHOLDER_RE.finditer(text):
                found.append(m.group(1).strip())

    # De-duplicate preserving order
    seen, unique = set(), []
    for ph in found:
        if ph not in seen:
            seen.add(ph)
            unique.append(ph)
    return unique

# --- Normalize & prune worksheet columns by placeholder set ---
def _norm(text: str) -> str:
    """Normalize for robust matching: strip, collapse spaces, lowercase."""
    return re.sub(r"\s+", " ", (text or "").strip()).lower()

def prune_ws_columns_not_in_placeholders(ws, placeholders, case_insensitive: bool = True):
    """
    Remove every column whose header (row 1) is NOT present in `placeholders`.
    Returns a list of (col_index, header_value) that were deleted.
    """
    if not placeholders:
        return []  # nothing to compare—skip

    allowed = {_norm(p) for p in placeholders} if case_insensitive else {p.strip() for p in placeholders}

    cols_to_delete = []
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        header = ws.cell(row=1, column=col_idx).value
        key = _norm(header) if case_insensitive else (header.strip() if isinstance(header, str) else header)
        if not key or key not in allowed:
            cols_to_delete.append((col_idx, header))

    # Delete right-to-left so indexes don’t shift
    for col_idx, _ in reversed(cols_to_delete):
        ws.delete_cols(col_idx, 1)

    return cols_to_delete

#==============================================================================================================#


# This script reads data from an Excel sheet and generates Word documents based on a template.

def generar_informe(nombre, fecha, mes, anio, cargo, empresa, ruta_xlsx, ruta_eeg, datos_feria):
    try:
        # -----------------------------
        # 0) Resolve inputs and paths
        # -----------------------------
        print("[STEP] iniciar generar_informe()")
        print("[DEBUG] datos_feria:", datos_feria)

        # ⚠️ Evita doble 'media/'. Creamos path absoluto con MEDIA_ROOT si viene relativo.
        ruta_xlsx_abs = _abs_media_path(ruta_xlsx)
        print("[DEBUG] ruta_xlsx_abs:", ruta_xlsx_abs)

        template_path = _resolve_template_path("plantillas/template Informe NeuroBit Feria --nombre--.docx")

        print("[DEBUG] template_path:", template_path)
        _ensure(os.path.exists(template_path), "TEMPLATE_NOT_FOUND",
                f"No existe el template DOCX en: {template_path}")

        _ensure(os.path.exists(ruta_xlsx_abs), "XLSX_NOT_FOUND",
                f"No existe el XLSX fuente en: {ruta_xlsx_abs}")

        # -------------------------------------
        # 1) Seed workbook with base identity
        # -------------------------------------
        wb_seed = openpyxl.Workbook()
        ws_seed = wb_seed.active

        ws_seed.append(["nombre", "fecha", "mes", "año", "cargo", "empresa"])
        ws_seed.append([nombre, fecha, mes, anio, cargo, empresa])

        seed_path = "temp_data.xlsx"
        wb_seed.save(seed_path)
        print("[STEP] seed temp_data.xlsx creado:", seed_path)

        # -------------------------------------
        # 2) Open main workbook and source xlsx
        # -------------------------------------
        try:
            wb_main = load_workbook(seed_path)
        except Exception as e:
            raise Sheet2DocsError(f"[OPEN_SEED_FAIL] No se pudo abrir {seed_path}: {e}")

        ws_main = wb_main.active

        try:
            wb_temp = load_workbook(ruta_xlsx_abs, data_only=True)
        except InvalidFileException as e:
            raise Sheet2DocsError(f"[OPEN_XLSX_INVALID] Archivo inválido para openpyxl: {ruta_xlsx_abs}: {e}")
        except Exception as e:
            raise Sheet2DocsError(f"[OPEN_XLSX_FAIL] No se pudo abrir ruta_xlsx: {ruta_xlsx_abs}: {e}")

        ws_temp = wb_temp.active
        _ensure(ws_temp.max_row >= 2, "XLSX_NO_DATA",
                f"El XLSX fuente no tiene fila de datos (se requiere al menos encabezado + una fila) en {ruta_xlsx_abs}")

        # -------------------------------------
        # 3) Append columns from ruta_xlsx_abs
        # -------------------------------------
        print("[STEP] volcar columnas desde XLSX fuente → ws_main")
        temp_cols = [col for col in ws_temp.iter_cols(min_row=2, max_row=2, values_only=True)]
        start_col = ws_main.max_column + 1
        for idx, col_data in enumerate(temp_cols):
            header = ws_temp.cell(row=1, column=idx + 1).value
            _ensure(header not in (None, ""), "XLSX_EMPTY_HEADER",
                    f"Encabezado vacío en columna {idx+1} de {ruta_xlsx_abs}")
            ws_main.cell(row=1, column=start_col + idx, value=header)
            ws_main.cell(row=2, column=start_col + idx, value=(col_data[0] if col_data else None))

        _debug_headers(ws_main, "HEADERS post-append ruta_xlsx")

        # -------------------------------------
        # 4) Auto-interpret numbers already in ws_main
        # -------------------------------------
        print("[STEP] generar columnas de interpretación auto (si aplica)")
        for col_idx, col in enumerate(ws_main.iter_cols(min_row=2, values_only=True), start=1):
            header = ws_main.cell(row=1, column=col_idx).value
            if header in (None, ""):
                continue
            interp_header = f"{header} interpretacion"

            if interp_header not in [ws_main.cell(row=1, column=i).value for i in range(1, ws_main.max_column + 1)]:
                ws_main.cell(row=1, column=ws_main.max_column + 1).value = interp_header

            # localizar índice de columna de interpretación
            interp_col_idx = None
            for i in range(1, ws_main.max_column + 1):
                if ws_main.cell(row=1, column=i).value == interp_header:
                    interp_col_idx = i
                    break

            for row_idx, cell_value in enumerate(col, start=2):
                # Sólo intenta interpretar si parece número
                if isinstance(cell_value, (int, float)) or (isinstance(cell_value, str) and '.' in cell_value):
                    try:
                        if isinstance(cell_value, str):
                            interp_num = float(cell_value.replace(',', '.'))
                            cell_str = cell_value.replace('.', ',')
                        else:
                            interp_num = float(cell_value)
                            cell_str = str(cell_value).replace('.', ',')
                    except ValueError:
                        interp_num = None
                        cell_str = str(cell_value)

                    if interp_num is not None:
                        if 0 <= interp_num <= 20:   interp_value = "Muy bajo"
                        elif 21 <= interp_num <= 49: interp_value = "Bajo"
                        elif 50 <= interp_num <= 74: interp_value = "Adecuado"
                        elif 75 <= interp_num <= 100: interp_value = "Alto"
                        else: interp_value = ""
                        ws_main.cell(row=row_idx, column=col_idx).value = cell_str
                        ws_main.cell(row=row_idx, column=interp_col_idx).value = interp_value

        # -------------------------------------
        # 5) Upsert datos_feria (values + interpretations)
        # -------------------------------------
        print("[STEP] inyectar datos_feria a ws_main")
        append_or_update_datos_feria(ws_main, datos_feria, decimal_comma=True)
        _debug_headers(ws_main, "HEADERS post-datos_feria")

        # -------------------------------------
        # 6) Read placeholders + prune columns
        # -------------------------------------
        print("[STEP] leer placeholders del template y podar columnas")
        placeholders = list_docx_placeholders(template_path)
        print("[DEBUG] placeholders:", placeholders)
        _ensure(isinstance(placeholders, list), "PLACEHOLDERS_BAD_TYPE",
                f"El extractor devolvió tipo inesperado: {type(placeholders)}")
        _ensure(len(placeholders) > 0, "PLACEHOLDERS_EMPTY",
                "El template no contiene placeholders '<<...>>' o no se pudieron leer.")

        deleted = prune_ws_columns_not_in_placeholders(ws_main, placeholders, case_insensitive=True)
        print("[DEBUG] columnas podadas:", [(i, str(h)) for i, h in deleted])
        kept = _debug_headers(ws_main, "HEADERS finales tras poda")

        # Sanity: al menos 1 header debe quedar, y deben existir los usados por el template
        _ensure(ws_main.max_column >= 1, "ALL_COLUMNS_REMOVED",
                "Tras la poda, no quedó ninguna columna. Revisa nombres de placeholders/headers.")
        # Nota: si quieres forzar que *todos* los placeholders existan como header:
        missing = [p for p in placeholders if _norm(p) not in {_norm(h) for h in kept if h}]
        if missing:
            print("[WARN] placeholders sin columna correspondiente en XLSX:", missing)

        # -------------------------------------
        # 7) Persist the trimmed workbook
        # -------------------------------------
        wb_main.save("temp_data.xlsx")
        print(f"[STEP] guardado temp_data.xlsx OK; columnas={ws_main.max_column}")

        # -------------------------------------
        # 8) Re-open trimmed xlsx for filling
        # -------------------------------------
        wb = load_workbook("temp_data.xlsx")
        ws = wb.active

        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        headers = list(filter(lambda x: x is not None, first_row))
        _ensure(len(headers) > 0, "NO_HEADERS_AFTER_TRIM",
                "No hay encabezados en temp_data.xlsx tras la poda.")

        # -------------------------------------
        # 9) Fill DOCX and save
        # -------------------------------------
        for data_row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
            doc = Document(template_path)

            # Reemplazos en párrafos
            for paragraph in doc.paragraphs:
                for header in headers:
                    ph = f'<<{header}>>'
                    if ph in paragraph.text:
                        idx = headers.index(header)
                        val = data_row[idx]
                        if val in (None, ""):
                            paragraph.text = paragraph.text.replace(ph, ' ')
                        elif "img" in header:
                            paragraph.text = paragraph.text.replace(ph, '')
                            run = paragraph.add_run()
                            run.add_picture(val)
                        else:
                            paragraph.text = paragraph.text.replace(ph, str(val))

            # Reemplazos en tablas
            for table in doc.tables:
                for row_cells in table.rows:
                    for cell in row_cells.cells:
                        for paragraph in cell.paragraphs:
                            for header in headers:
                                ph = f'<<{header}>>'
                                if ph in paragraph.text:
                                    idx = headers.index(header)
                                    val = data_row[idx]
                                    if val in (None, ""):
                                        paragraph.text = paragraph.text.replace(ph, ' ')
                                    elif "img" in header:
                                        paragraph.text = paragraph.text.replace(ph, '')
                                        run = paragraph.add_run()
                                        run.add_picture(val)
                                    else:
                                        paragraph.text = paragraph.text.replace(ph, str(val))

            # Nombre de salida
            output_path = template_path.replace("template ", "").replace("template_", "").replace("template", "").strip().replace(" ", "_")
            for header in headers:
                token = f'--{header}--'
                if token in output_path:
                    output_path = output_path.replace(token, str(data_row[headers.index(header)]))

            doc.save(output_path)
            print(f"[STEP] DOCX guardado: {output_path}")

        # -------------------------------------
        # 10) Cleanup y return
        # -------------------------------------
        try:
            os.remove("temp_data.xlsx")
        except Exception as e:
            print("[WARN] no se pudo borrar temp_data.xlsx:", e)

        return output_path

    except Sheet2DocsError:
        # errores “nuestros” con código claro
        print(traceback.format_exc())
        raise
    except Exception as e:
        # errores inesperados: wrap con etiqueta para el caller
        print(traceback.format_exc())
        raise Sheet2DocsError(f"[UNEXPECTED] {e}")
