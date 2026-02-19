#made by Yarko Bahamonde
#distribution outside of my github page is prohibited
#can be used in non-profit projects
#the complete or partial use of the code for other projects needs to be credited

import os
from openpyxl import load_workbook
from docx import Document


def find_first_xlsx(directory: str) -> str:
    files = sorted(
        f for f in os.listdir(directory)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"No se encontró ningún .xlsx en: {directory}")
    return os.path.join(directory, files[0])


def find_template_docx(directory: str) -> str:
    files = sorted(
        f for f in os.listdir(directory)
        if f.lower().endswith(".docx") and f.lower().startswith("template")
    )
    if not files:
        raise FileNotFoundError(f"No se encontró ningún .docx que empiece por 'template' en: {directory}")
    return os.path.join(directory, files[0])


def header_is_image(header: str) -> bool:
    return "img" in str(header).lower()


def replace_placeholders_in_doc(doc: Document, mapping: dict):
    """
    Reemplaza placeholders <<key>> por mapping[key] (texto),
    y si el header contiene 'img' y el valor es ruta existente, inserta imagen.
    """
    def handle_paragraph(paragraph):
        text = paragraph.text
        if not text:
            return

        # Si hay imágenes, para evitar dejar <<img...>> visible:
        # 1) borramos placeholder en el texto
        # 2) agregamos imagen si la ruta existe
        for k, v in mapping.items():
            ph = f"<<{k}>>"
            if ph not in text:
                continue

            if v is None or v == "":
                text = text.replace(ph, " ")
                continue

            if header_is_image(k):
                img_path = str(v)
                text = text.replace(ph, "")
                paragraph.text = text
                if os.path.exists(img_path) and os.path.isfile(img_path):
                    run = paragraph.add_run()
                    run.add_picture(img_path)
                return  # ya tocamos paragraph.text; salimos

            text = text.replace(ph, str(v))

        if text != paragraph.text:
            paragraph.text = text

    # Body
    for p in doc.paragraphs:
        handle_paragraph(p)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    handle_paragraph(p)


def generate_docs_from_excel(directory: str, output_dir: str | None = None):
    directory = os.path.abspath(directory)
    output_dir = os.path.abspath(output_dir or directory)

    xlsx_path = find_first_xlsx(directory)
    template_path = find_template_docx(directory)

    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    if ws.max_row < 2:
        raise ValueError("El Excel debe tener al menos 2 filas: headers (fila 1) + datos (fila 2).")

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    if any(h in (None, "") for h in headers):
        raise ValueError("Hay headers vacíos en la fila 1 del Excel.")

    headers = [str(h) for h in headers]

    created = []

    # Genera 1 doc por cada fila de datos desde la 2 hacia abajo
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if all(v is None for v in values):
            continue

        mapping = dict(zip(headers, values))

        doc = Document(template_path)
        replace_placeholders_in_doc(doc, mapping)

        # Nombre de salida: intenta usar number_docs y name si existen
        num = mapping.get("number_docs", r - 1)
        name = mapping.get("name", "")
        safe_name = str(name).strip().replace(" ", "_")

        out_name = f"output_{num}_{safe_name}.docx" if safe_name else f"output_{num}.docx"
        out_path = os.path.join(output_dir, out_name)

        doc.save(out_path)
        created.append(out_path)

    print("XLSX usado:", xlsx_path)
    print("Template usado:", template_path)
    print("Docs creados:")
    for p in created:
        print(" -", p)

    return created


if __name__ == "__main__":
    # Ejecuta en el directorio actual por defecto
    generate_docs_from_excel(".")
