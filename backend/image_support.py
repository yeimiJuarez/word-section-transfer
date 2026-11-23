from io import BytesIO

from docx.oxml.ns import qn
from fastapi import HTTPException


def reject(message):
    raise HTTPException(status_code=400, detail=message)


def read_inline_image(drawing, source_document):
    if drawing.find(qn("wp:anchor")) is not None:
        reject(
            "La sección contiene una imagen flotante. "
            "Esta versión solo admite imágenes en línea."
        )

    inline = drawing.find(qn("wp:inline"))

    if inline is None or len(drawing) != 1:
        reject("El objeto gráfico no es una imagen en línea compatible.")

    graphic = inline.find(qn("a:graphic"))
    graphic_data = (
        graphic.find(qn("a:graphicData"))
        if graphic is not None
        else None
    )

    picture_uri = (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/picture"
    )

    if (
        graphic_data is None
        or graphic_data.get("uri") != picture_uri
    ):
        reject(
            "El objeto contiene un gráfico, una figura "
            "u otro elemento todavía no soportado."
        )

    pictures = list(graphic_data.iter(qn("pic:pic")))
    blips = list(graphic_data.iter(qn("a:blip")))

    if len(pictures) != 1 or len(blips) != 1:
        reject("El objeto gráfico contiene una estructura no soportada.")

    if list(graphic_data.iter(qn("a:srcRect"))):
        reject(
            "La imagen contiene un recorte. "
            "Esta versión todavía no conserva imágenes recortadas."
        )

    for transform in graphic_data.iter(qn("a:xfrm")):
        if (
            transform.get("rot", "0") != "0"
            or transform.get("flipH", "0") not in {"0", "false"}
            or transform.get("flipV", "0") not in {"0", "false"}
        ):
            reject(
                "La imagen contiene rotación o inversión "
                "todavía no soportada."
            )

    blip = blips[0]

    if blip.get(qn("r:link")):
        reject("No se admiten imágenes vinculadas externamente.")

    relationship_id = blip.get(qn("r:embed"))

    if not relationship_id:
        reject("No se encontró la imagen incrustada.")

    image_part = source_document.part.related_parts.get(
        relationship_id
    )

    if image_part is None:
        reject("No se encontró el archivo de imagen asociado.")

    if not image_part.content_type.startswith("image/"):
        reject("La relación del objeto no corresponde a una imagen.")

    extent = inline.find(qn("wp:extent"))

    if extent is None:
        reject("La imagen no tiene dimensiones de presentación.")

    width = int(extent.get("cx", "0"))
    height = int(extent.get("cy", "0"))

    if width <= 0 or height <= 0:
        reject("La imagen tiene dimensiones inválidas.")

    return {
        "blob": image_part.blob,
        "width": width,
        "height": height,
    }


def extract_content_runs(paragraph, source_document):
    """
    Conserva el orden del texto y de las imágenes dentro del párrafo.
    La validación del párrafo debe ejecutarse antes.
    """
    result = []

    for run in paragraph.runs:
        text_parts = []

        def flush_text():
            if text_parts:
                result.append(
                    {
                        "text": "".join(text_parts),
                        "bold": run.bold,
                        "italic": run.italic,
                    }
                )
                text_parts.clear()

        for child in run._r:
            if child.tag == qn("w:t"):
                text_parts.append(child.text or "")

            elif child.tag == qn("w:tab"):
                text_parts.append("\t")

            elif child.tag in {qn("w:br"), qn("w:cr")}:
                text_parts.append("\n")

            elif child.tag == qn("w:drawing"):
                flush_text()

                result.append(
                    {
                        "image": read_inline_image(
                            child,
                            source_document,
                        )
                    }
                )

        flush_text()

    return result


def get_target_width(document, title_element):
    """
    Encuentra las propiedades de la sección de Word que contiene
    el título de destino. Las secciones de Word no son lo mismo
    que las secciones identificadas por títulos.
    """
    body_elements = list(document.element.body)
    start = body_elements.index(title_element)
    section_properties = None

    for element in body_elements[start:]:
        if element.tag == qn("w:sectPr"):
            section_properties = element
            break

        if element.tag == qn("w:p"):
            properties = element.find(qn("w:pPr"))

            if properties is not None:
                found = properties.find(qn("w:sectPr"))

                if found is not None:
                    section_properties = found
                    break

    target_section = next(
        (
            section
            for section in document.sections
            if section._sectPr is section_properties
        ),
        None,
    )

    if target_section is None:
        reject("No se pudo determinar el ancho de la sección de B.")

    columns = section_properties.find(qn("w:cols"))

    if columns is not None:
        explicit_columns = columns.findall(qn("w:col"))
        column_count = int(columns.get(qn("w:num"), "1"))

        if column_count > 1 or len(explicit_columns) > 1:
            reject(
                "La inserción de imágenes en secciones "
                "con varias columnas todavía no está soportada."
            )

    if any(
        value is None
        for value in (
            target_section.page_width,
            target_section.left_margin,
            target_section.right_margin,
        )
    ):
        reject("B no tiene dimensiones de página completas.")

    width = (
        target_section.page_width
        - target_section.left_margin
        - target_section.right_margin
        - (target_section.gutter or 0)
    )

    if width <= 0:
        reject("El ancho disponible de B no es válido.")

    return int(width)


def add_image_run(paragraph, image, available_width):
    width = image["width"]
    height = image["height"]

    # Considera también las sangrías del párrafo y sus estilos.
    def inherited_indent(name):
        value = getattr(paragraph.paragraph_format, name)

        if value is not None:
            return int(value)

        style = paragraph.style
        visited = set()

        while style is not None:
            if style.style_id in visited:
                break

            visited.add(style.style_id)
            value = getattr(style.paragraph_format, name)

            if value is not None:
                return int(value)

            style = style.base_style

        return 0

    usable_width = (
        available_width
        - max(0, inherited_indent("left_indent"))
        - max(0, inherited_indent("right_indent"))
    )

    if usable_width <= 0:
        reject("Las sangrías de B no dejan espacio para la imagen.")

    if width > usable_width:
        scale = usable_width / width
        width = usable_width
        height = max(1, round(height * scale))

    try:
        paragraph.add_run().add_picture(
            BytesIO(image["blob"]),
            width=width,
            height=height,
        )
    except Exception:
        reject(
            "No se pudo insertar una imagen. "
            "Su formato puede no ser compatible con python-docx."
        )