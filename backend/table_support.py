from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from fastapi import HTTPException

from image_support import extract_content_runs, add_image_run
from list_support import detect_list, apply_list_format



def reject(message):
    raise HTTPException(
        status_code=400,
        detail=message,
    )


def find_table_style(document):
    for style in document.styles:
        if (
            style.type == WD_STYLE_TYPE.TABLE
            and style.style_id == "TableGrid"
        ):
            return style

    reject(
        "B no contiene el estilo Table Grid. "
        "Aplica ese estilo a una tabla de prueba en B, "
        "guarda el documento y vuelve a cargarlo."
    )


def get_cell_structure(cell):
    properties = cell.find(qn("w:tcPr"))

    span = 1
    vertical_merge = None

    if properties is not None:
        # El formato antiguo hMerge necesita un tratamiento diferente.
        if properties.find(qn("w:hMerge")) is not None:
            reject(
                "La tabla utiliza combinaciones horizontales "
                "en formato antiguo w:hMerge, todavía no soportadas."
            )

        grid_span = properties.find(qn("w:gridSpan"))

        if grid_span is not None:
            span = int(grid_span.get(qn("w:val"), "1"))

        merge = properties.find(qn("w:vMerge"))

        if merge is not None:
            vertical_merge = merge.get(
                qn("w:val"),
                "continue",
            )

            if vertical_merge not in {"restart", "continue"}:
                reject(
                    "La tabla contiene una combinación vertical inválida."
                )

    if span < 1:
        reject("Una celda tiene una extensión horizontal inválida.")

    return span, vertical_merge


def validate_cell_blocks(cell):
    for child in cell:
        if child.tag not in {
            qn("w:tcPr"),
            qn("w:p"),
        }:
            reject(
                "Las celdas solo pueden contener párrafos. "
                "Las tablas anidadas y los controles de contenido "
                "todavía no están soportados."
            )


def continuation_has_content(cell):
    """
    Una celda de continuación debería estar vacía.
    No omitimos silenciosamente texto, imágenes o campos.
    """
    for paragraph in cell.findall(qn("w:p")):
        for child in paragraph:
            if child.tag == qn("w:pPr"):
                continue

            if child.tag in {
                qn("w:bookmarkStart"),
                qn("w:bookmarkEnd"),
                qn("w:proofErr"),
            }:
                continue

            if child.tag != qn("w:r"):
                return True

            for run_child in child:
                if run_child.tag in {
                    qn("w:rPr"),
                    qn("w:lastRenderedPageBreak"),
                }:
                    continue

                if run_child.tag == qn("w:t"):
                    if run_child.text:
                        return True
                    continue

                return True

    return False


def extract_cell_paragraphs(
    cell,
    table,
    source_document,
    validate_paragraph,
    flatten_hyperlinks,
    remove_page_breaks,
    get_heading_level,
    row_index,
    column_index,
):
    wrapped_cell = _Cell(cell, table)

    paragraphs = []
    list_specs = []

    for paragraph in wrapped_cell.paragraphs:
        flatten_hyperlinks(paragraph)
        remove_page_breaks(paragraph)

        if get_heading_level(paragraph) is not None:
            reject(
                f"Celda {row_index + 1},{column_index + 1}: "
                "los títulos dentro de tablas todavía "
                "no están soportados."
            )

        list_spec = detect_list(
            paragraph,
            source_document,
        )

        validate_paragraph(
            paragraph,
            allow_numbering=list_spec is not None,
        )

        content = extract_content_runs(
            paragraph,
            source_document,
        )

      
        paragraphs.append(content)
        list_specs.append(list_spec)

    if not paragraphs:
        reject("Una celda no contiene un párrafo válido.")

    return paragraphs, list_specs

def build_table_plan(
    element,
    source_document,
    target_document,
    validate_paragraph,
    flatten_hyperlinks,
    remove_page_breaks,
    get_heading_level,
):
    table = Table(element, source_document)
    grid = element.find(qn("w:tblGrid"))

    if grid is None:
        reject("La tabla no contiene una cuadrícula válida.")

    column_count = len(grid.findall(qn("w:gridCol")))
    source_rows = element.findall(qn("w:tr"))

    if column_count == 0 or not source_rows:
        reject("La tabla está vacía o no tiene columnas definidas.")

    for child in element:
        if child.tag not in {
            qn("w:tblPr"),
            qn("w:tblGrid"),
            qn("w:tr"),
        }:
            reject(
                "La tabla contiene elementos estructurales "
                "todavía no soportados."
            )

    properties = element.find(qn("w:tblPr"))

    if (
        properties is not None
        and properties.find(qn("w:tblpPr")) is not None
    ):
        reject("Las tablas flotantes todavía no están soportadas.")

    planned_rows = [
        [[] for _ in range(column_count)]
        for _ in source_rows
    ]

    # Cada región describe una celda lógica:
    # inicio y final del rectángulo que ocupa.
    regions = []
    cell_lists = {}

    # Combinaciones verticales activas de la fila anterior.
    active_vertical = {}

    for row_index, row in enumerate(source_rows):
        for child in row:
            if child.tag not in {
                qn("w:trPr"),
                qn("w:tc"),
            }:
                reject(
                    "Una fila contiene elementos no soportados."
                )

        row_properties = row.find(qn("w:trPr"))

        if row_properties is not None:
            for tag in ("w:gridBefore", "w:gridAfter"):
                skipped = row_properties.find(qn(tag))

                if (
                    skipped is not None
                    and int(skipped.get(qn("w:val"), "0")) != 0
                ):
                    reject(
                        "Las filas con celdas omitidas "
                        "todavía no están soportadas."
                    )

        column_index = 0
        next_active = {}

        for cell in row.findall(qn("w:tc")):
            validate_cell_blocks(cell)

            span, vertical_merge = get_cell_structure(cell)
            end_column = column_index + span - 1

            if end_column >= column_count:
                reject(
                    "Una celda combinada excede "
                    "la cuadrícula de la tabla."
                )

            if vertical_merge == "continue":
                region = active_vertical.get(column_index)

                if (
                    region is None
                    or region["end_column"] != end_column
                ):
                    reject(
                        "La combinación vertical no forma "
                        "un rectángulo válido."
                    )

                if continuation_has_content(cell):
                    reject(
                        "Una celda de continuación vertical "
                        "contiene contenido propio. "
                        "No puede omitirse de forma segura."
                    )

                region["end_row"] = row_index
                next_active[column_index] = region

            else:
                paragraphs, list_specs = extract_cell_paragraphs(
                    cell,
                    table,
                    source_document,
                    validate_paragraph,
                    flatten_hyperlinks,
                    remove_page_breaks,
                    get_heading_level,
                    row_index,
                    column_index,
                )

                planned_rows[row_index][column_index] = paragraphs
                cell_lists[(row_index, column_index)] = list_specs

                region = {
                    "start_row": row_index,
                    "start_column": column_index,
                    "end_row": row_index,
                    "end_column": end_column,
                }

                regions.append(region)

                if vertical_merge == "restart":
                    next_active[column_index] = region

            column_index += span

        if column_index != column_count:
            reject(
                "La fila no cubre todas las columnas "
                "de la cuadrícula."
            )

        active_vertical = next_active

    merges = [
        region
        for region in regions
        if (
            region["start_row"] != region["end_row"]
            or region["start_column"] != region["end_column"]
        )
    ]

    return {
        "style": find_table_style(target_document),
        "column_count": column_count,
        "rows": planned_rows,
        "regions": regions,
        "merges": merges,
        "cell_lists": cell_lists,
    }

def get_cell_inner_width(cell, table):
    """
    Obtiene el ancho de la celda menos sus márgenes interiores.
    Consulta propiedades de celda, tabla y estilos heredados.
    """
    if cell.width is None:
        reject("No se pudo determinar el ancho de la celda.")

    margin_sources = []

    cell_properties = cell._tc.find(qn("w:tcPr"))

    if cell_properties is not None:
        margins = cell_properties.find(qn("w:tcMar"))

        if margins is not None:
            margin_sources.append(margins)

    table_properties = table._tbl.tblPr
    table_margins = table_properties.find(qn("w:tblCellMar"))

    if table_margins is not None:
        margin_sources.append(table_margins)

    style = table.style
    visited = set()

    while style is not None:
        if style.style_id in visited:
            break

        visited.add(style.style_id)

        properties = style.element.find(qn("w:tblPr"))

        if properties is not None:
            margins = properties.find(qn("w:tblCellMar"))

            if margins is not None:
                margin_sources.append(margins)

        style = style.base_style

    def read_margin(side):
        aliases = (
            ("start", "left")
            if side == "left"
            else ("end", "right")
        )

        for source in margin_sources:
            for name in aliases:
                margin = source.find(qn(f"w:{name}"))

                if margin is None:
                    continue

                margin_type = margin.get(qn("w:type"), "dxa")

                if margin_type == "nil":
                    return 0

                if margin_type != "dxa":
                    reject(
                        "La celda utiliza márgenes interiores "
                        "con unidades todavía no soportadas."
                    )

                twips = int(margin.get(qn("w:w"), "0"))

                # 1 twip = 635 EMU.
                return max(0, twips) * 635

        # Reserva predeterminada de 0.08 pulgadas por lado.
        # También se aplica explícitamente en el destino.
        return 115 * 635

    left_margin = read_margin("left")
    right_margin = read_margin("right")

    inner_width = (
        int(cell.width)
        - left_margin
        - right_margin
    )

    if inner_width <= 0:
        reject(
            "Los márgenes interiores no dejan espacio "
            "para la imagen."
        )

    return inner_width, left_margin, right_margin

def insert_table(
    document,
    insertion_point,
    plan,
    available_width,
):
    from docx.oxml import OxmlElement

    if available_width <= 0:
        reject(
            "No se pudo determinar el ancho disponible "
            "para la tabla."
        )

    table = document.add_table(
        rows=len(plan["rows"]),
        cols=plan["column_count"],
    )

    table.style = plan["style"]
    table.autofit = False

    column_width = available_width // plan["column_count"]

    for column in table.columns:
        column.width = column_width

    for row in table.rows:
        for cell in row.cells:
            cell.width = column_width

    table_width = table._tbl.tblPr.find(qn("w:tblW"))

    if table_width is not None:
        table_width.set(qn("w:type"), "dxa")
        table_width.set(
            qn("w:w"),
            str(round(available_width / 635)),
        )

    # Combina primero las celdas vacías.
    for region in plan["merges"]:
        first = table.cell(
            region["start_row"],
            region["start_column"],
        )

        last = table.cell(
            region["end_row"],
            region["end_column"],
        )

        first.merge(last)

    normal_style = document.styles["Normal"]

    # Inserta contenido una sola vez por celda lógica.
    for region in plan["regions"]:
        row_index = region["start_row"]
        column_index = region["start_column"]

        cell = table.cell(row_index, column_index)

        span = (
            region["end_column"]
            - region["start_column"]
            + 1
        )

        cell.width = column_width * span

        planned_paragraphs = plan["rows"][
            row_index
        ][column_index]

        has_images = any(
            "image" in content
            for paragraph_content in planned_paragraphs
            for content in paragraph_content
        )

        inner_width = 0

        if has_images:
            (
                inner_width,
                left_margin,
                right_margin,
            ) = get_cell_inner_width(cell, table)

            # Aplica explícitamente los márgenes calculados.
            properties = cell._tc.get_or_add_tcPr()
            margins = properties.find(qn("w:tcMar"))

            if margins is None:
                margins = OxmlElement("w:tcMar")

                properties.insert_element_before(
                    margins,
                    "w:textDirection",
                    "w:tcFitText",
                    "w:vAlign",
                    "w:hideMark",
                    "w:headers",
                    "w:cellIns",
                    "w:cellDel",
                    "w:cellMerge",
                    "w:tcPrChange",
                )

            for side, width in (
                ("left", left_margin),
                ("right", right_margin),
            ):
                # Evita que start/end entren en conflicto
                # con left/right.
                aliases = (
                    ("start", "left")
                    if side == "left"
                    else ("end", "right")
                )

                for alias in aliases:
                    existing = margins.find(qn(f"w:{alias}"))

                    if existing is not None:
                        margins.remove(existing)

                margin = OxmlElement(f"w:{side}")
                margin.set(qn("w:type"), "dxa")
                margin.set(
                    qn("w:w"),
                    str(round(width / 635)),
                )
                margins.append(margin)

        # Elimina párrafos vacíos generados por merge().
        for child in list(cell._tc):
            if child.tag != qn("w:tcPr"):
                cell._tc.remove(child)

        list_specs = plan.get("cell_lists", {}).get(
            (row_index, column_index),
            [None] * len(planned_paragraphs),
        )

        list_state = {}

        for paragraph_index, content in enumerate(
            planned_paragraphs
        ):
            paragraph = cell.add_paragraph()
            paragraph.style = normal_style

            list_spec = list_specs[paragraph_index]

            apply_list_format(
                document,
                paragraph,
                list_spec,
                list_state,
            )

            for item in content:
                if "image" in item:
                    image_width = inner_width

                    if list_spec is not None:
                        level = list_spec.get("level", 0)

                        image_width -= (
                            720 * (level + 1) * 635
                        )

                    if image_width <= 0:
                        reject(
                            "La celda es demasiado estrecha "
                            "para insertar la imagen con "
                            "la sangría de la lista."
                        )

                    add_image_run(
                        paragraph,
                        item["image"],
                        image_width,
                    )
                    continue

                run = paragraph.add_run(item["text"])

                if item["bold"] is not None:
                    run.bold = item["bold"]

                if item["italic"] is not None:
                    run.italic = item["italic"]

    insertion_point.addnext(table._tbl)

    return table._tbl