from io import BytesIO
from uuid import uuid4

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from fastapi import HTTPException, Response
from pydantic import BaseModel
from table_support import build_table_plan, insert_table
from transfer_summary import summarize_plan


from list_support import (
    detect_list,
    apply_list_format,
)

from image_support import (
    extract_content_runs,
    get_target_width,
    add_image_run,
)




MAX_DOCUMENTS = 10


class TransferRequest(BaseModel):
    source_document_id: str
    target_document_id: str
    source_section_id: str
    target_section_id: str


def fail(message: str):
    raise HTTPException(
        status_code=400,
        detail=message,
    )


def get_blocks(document):
    return [
        element
        for element in document.element.body
        if element.tag != qn("w:sectPr")
    ]


def find_section(stored_document, section_id):
    section = next(
        (
            item
            for item in stored_document["sections"]
            if item["id"] == section_id
        ),
        None,
    )

    if section is None:
        fail("No se encontró la sección seleccionada.")

    return section


def validate_section_permissions(stored_document, section):
    if not section.get("transfer_enabled", False):
        fail(
            "La sección seleccionada tiene "
            "la transferencia deshabilitada."
        )

    # También protege las subsecciones incluidas en la selección.
    for candidate in stored_document["sections"]:
        is_inside = (
            section["start_index"]
            <= candidate["start_index"]
            < section["end_index"]
        )

        if is_inside and not candidate.get(
            "transfer_enabled",
            False,
        ):
            fail(
                "La sección seleccionada contiene "
                "una sección protegida: "
                + candidate["title"]
            )


def has_numbering(paragraph):
    properties = paragraph._p.pPr

    if properties is not None:
        if properties.find(qn("w:numPr")) is not None:
            return True

    style = paragraph.style
    visited = set()

    while style is not None:
        if style.style_id in visited:
            break

        visited.add(style.style_id)
        properties = style.element.pPr

        if properties is not None:
            if properties.find(qn("w:numPr")) is not None:
                return True

        style = style.base_style

    return False


def flatten_hyperlinks(paragraph):
    """
    Conserva el texto de los enlaces,
    pero no su comportamiento clicable.
    """
    for hyperlink in list(
        paragraph._p.findall(qn("w:hyperlink"))
    ):
        children = list(hyperlink)

        if any(
            child.tag != qn("w:r")
            for child in children
        ):
            fail(
                "El hipervínculo contiene elementos complejos "
                "todavía no soportados."
            )

        for child in children:
            hyperlink.addprevious(child)

        paragraph._p.remove(hyperlink)


def remove_source_page_breaks(paragraph):
    """
    Elimina los saltos de página explícitos de la copia de A.
    Conserva los saltos de línea.
    """
    for run in paragraph._p.findall(qn("w:r")):
        for page_break in list(
            run.findall(qn("w:br"))
        ):
            if page_break.get(qn("w:type")) == "page":
                run.remove(page_break)


def validate_plain_paragraph(
    paragraph,
    allow_numbering=False,
):
    element = paragraph._p

    visible_text = "".join(
        node.text or ""
        for node in element.iter(qn("w:t"))
    )

    preview = " ".join(visible_text.split())[:220]
    preview = preview or "(Párrafo sin texto visible)"

    style_name = (
        paragraph.style.name
        if paragraph.style is not None
        else "(Sin estilo)"
    )

    def describe_tag(tag):
        if tag.startswith("{"):
            namespace, local_name = tag[1:].split("}", 1)

            if namespace == (
                "http://schemas.openxmlformats.org/"
                "wordprocessingml/2006/main"
            ):
                return f"w:{local_name}"

            return f"{{{namespace}}}{local_name}"

        return tag

    def reject(reason, tag):
        fail(
            f"{reason} "
            f"Elemento XML: {describe_tag(tag)}. "
            f"Estilo: {style_name}. "
            f"Párrafo: «{preview}»."
        )

    if element.find(".//" + qn("w:sectPr")) is not None:
        reject(
            "La sección contiene un salto "
            "de sección no soportado.",
            qn("w:sectPr"),
        )

    if has_numbering(paragraph) and not allow_numbering:
        reject(
            "El párrafo contiene propiedades de numeración "
            "que no se pudieron interpretar.",
            qn("w:numPr"),
        )

    allowed_paragraph_children = {
        qn("w:pPr"),
        qn("w:r"),
        qn("w:bookmarkStart"),
        qn("w:bookmarkEnd"),
        qn("w:proofErr"),
    }

    allowed_run_children = {
        qn("w:rPr"),
        qn("w:t"),
        qn("w:tab"),
        qn("w:br"),
        qn("w:cr"),
        qn("w:lastRenderedPageBreak"),
        qn("w:drawing"),
    }

    for child in element:
        if child.tag not in allowed_paragraph_children:
            reject(
                "El párrafo contiene un elemento no soportado.",
                child.tag,
            )

        if child.tag != qn("w:r"):
            continue

        for run_child in child:
            if run_child.tag not in allowed_run_children:
                reject(
                    "Un fragmento del párrafo contiene "
                    "un elemento no soportado.",
                    run_child.tag,
                )

            if run_child.tag == qn("w:br"):
                break_type = run_child.get(
                    qn("w:type"),
                    "textWrapping",
                )

                if break_type != "textWrapping":
                    reject(
                        f"El párrafo contiene un salto "
                        f"de tipo {break_type} no soportado.",
                        run_child.tag,
                    )


def find_normal_style(document):
    try:
        style = document.styles["Normal"]
    except KeyError:
        fail(
            "B no contiene el estilo Normal "
            "requerido para esta versión."
        )

    if style.type != WD_STYLE_TYPE.PARAGRAPH:
        fail(
            "El estilo Normal de B "
            "no es un estilo de párrafo."
        )

    return style


def find_heading_style(document, level):
    expected_id = f"Heading{level}"

    # Primero busca el estilo integrado correspondiente.
    for style in document.styles:
        if (
            style.type == WD_STYLE_TYPE.PARAGRAPH
            and style.style_id == expected_id
        ):
            return style

    # Respaldo para estilos personalizados que tengan
    # un nivel de esquema explícito.
    for style in document.styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue

        properties = style.element.pPr

        if properties is None:
            continue

        outline = properties.find(qn("w:outlineLvl"))

        if (
            outline is not None
            and outline.get(qn("w:val")) == str(level - 1)
        ):
            return style

    fail(
        f"B no contiene un estilo de título "
        f"compatible con el nivel {level}."
    )


def build_transfer_plan(
    source_document,
    target_document,
    source_section,
    target_section,
    get_heading_level,
):
    source_blocks = get_blocks(source_document)

    content = source_blocks[
        source_section["content_start_index"]:
        source_section["end_index"]
    ]

    if not content:
        fail("La sección de origen no contiene elementos para copiar.")

    normal_style = find_normal_style(target_document)
    plan = []

    for element in content:
        if element.tag in {
            qn("w:bookmarkStart"),
            qn("w:bookmarkEnd"),
        }:
            continue

        if element.tag == qn("w:tbl"):
            table_plan = build_table_plan(
                element,
                source_document,
                target_document,
                validate_plain_paragraph,
                flatten_hyperlinks,
                remove_source_page_breaks,
                get_heading_level,
            )

            plan.append(
                {
                    "table": table_plan,
                    "runs": [],
                    "list": None,
                }
            )
            continue

        if element.tag != qn("w:p"):
            block_type = element.tag.split("}")[-1]

            fail(
                "La sección de origen contiene "
                f"un bloque no soportado: w:{block_type}."
            )

        paragraph = Paragraph(element, source_document)

        flatten_hyperlinks(paragraph)
        remove_source_page_breaks(paragraph)

        source_level = get_heading_level(paragraph)

        list_spec = (
            detect_list(paragraph, source_document)
            if source_level is None
            else None
        )

        validate_plain_paragraph(
            paragraph,
            allow_numbering=(
                source_level is not None
                or list_spec is not None
            ),
        )

        if source_level is None:
            target_style = normal_style
        else:
            target_level = (
                target_section["level"]
                + source_level
                - source_section["level"]
            )

            if not 1 <= target_level <= 9:
                fail("El ajuste de subtítulos excede los niveles 1 a 9.")

            target_style = find_heading_style(
                target_document,
                target_level,
            )

        runs = extract_content_runs(
            paragraph,
            source_document,
        )

        plan.append(
            {
                "style": target_style,
                "runs": runs,
                "list": list_spec,
            }
        )

    if not plan:
        fail("La sección no contiene elementos compatibles para copiar.")

    return plan

def insert_planned_paragraph(
    document,
    insertion_point,
    item,
    list_state,
    available_width,
):
    if "table" in item:
        # Una tabla termina el grupo de lista actual.
        list_state.clear()

        return insert_table(
            document,
            insertion_point,
            item["table"],
            available_width,
        )

    paragraph = document.add_paragraph(
        style=item["style"]
    )

    apply_list_format(
        document,
        paragraph,
        item.get("list"),
        list_state,
    )

    for content in item["runs"]:
        if "image" in content:
            add_image_run(
                paragraph,
                content["image"],
                available_width,
            )
            continue

        run = paragraph.add_run(content["text"])

        if content["bold"] is not None:
            run.bold = content["bold"]

        if content["italic"] is not None:
            run.italic = content["italic"]

    insertion_point.addnext(paragraph._p)

    return paragraph._p

def register_transfer_routes(
    app,
    documents,
    extract_sections,
    get_heading_level,
):
    @app.post("/transfers", status_code=201)
    def transfer_section(request: TransferRequest):
        source_stored = documents.get(
            request.source_document_id
        )

        target_stored = documents.get(
            request.target_document_id
        )

        if source_stored is None or target_stored is None:
            raise HTTPException(
                status_code=404,
                detail=(
                    "No se encontró alguno de los documentos. "
                    "Si reiniciaste Python, vuelve a cargarlos."
                ),
            )

        if len(documents) >= MAX_DOCUMENTS:
            raise HTTPException(
                status_code=409,
                detail=(
                    "Se alcanzó el límite temporal "
                    "de 10 documentos."
                ),
            )

        source_section = find_section(
            source_stored,
            request.source_section_id,
        )

        target_section = find_section(
            target_stored,
            request.target_section_id,
        )

        validate_section_permissions(
            source_stored,
            source_section,
        )

        validate_section_permissions(
            target_stored,
            target_section,
        )

        # Trabaja sobre copias independientes.
        source_document = Document(
            BytesIO(source_stored["content"])
        )

        target_document = Document(
            BytesIO(target_stored["content"])
        )

        target_blocks = get_blocks(target_document)

        target_range = target_blocks[
            target_section["start_index"]:
            target_section["end_index"]
        ]

        for element in target_range:
            if (
                element.find(".//" + qn("w:sectPr"))
                is not None
            ):
                fail(
                    "La sección de destino contiene "
                    "un salto de sección. Esta versión "
                    "no puede reemplazarla de forma segura."
                )

        # Prepara el contenido antes de eliminar el destino.
        plan = build_transfer_plan(
            source_document,
            target_document,
            source_section,
            target_section,
            get_heading_level,
        )

        title_element = target_blocks[
            target_section["start_index"]
        ]

        needs_width = any(
            "table" in item
            or any(
                "image" in content
                for content in item["runs"]
            )
            for item in plan
        )

        available_width = (
            get_target_width(
                target_document,
                title_element,
            )
            if needs_width
            else 0
        )

        old_content = target_blocks[
            target_section["content_start_index"]:
            target_section["end_index"]
        ]

        for element in old_content:
            element.getparent().remove(element)

        insertion_point = title_element
        list_state = {}

        for item in plan:
            insertion_point = insert_planned_paragraph(
                target_document,
                insertion_point,
                item,
                list_state,
                available_width,
            )

        # Guarda únicamente cuando termina la operación.
        output = BytesIO()
        target_document.save(output)

        sections = extract_sections(target_document)
        document_id = str(uuid4())
        filename = "B_actualizado.docx"

        documents[document_id] = {
            "filename": filename,
            "content": output.getvalue(),
            "sections": sections,
        }

        summary = summarize_plan(plan)

        return {
            "document_id": document_id,
            "filename": filename,
            "section_count": len(sections),
            "sections": sections,
            **summary,
            "download_url": (
                f"/documents/{document_id}/download"
            ),
        }

    @app.get("/documents/{document_id}/download")
    def download_document(document_id: str):
        stored = documents.get(document_id)

        if stored is None:
            raise HTTPException(
                status_code=404,
                detail="Documento no encontrado.",
            )

        return Response(
            content=stored["content"],
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            headers={
                "Content-Disposition": (
                    'attachment; filename="documento.docx"'
                ),
                "Cache-Control": "no-store",
            },
        )