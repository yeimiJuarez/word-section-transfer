from io import BytesIO
from uuid import uuid4

from docx import Document
from docx.oxml.ns import qn
from fastapi import HTTPException
from pydantic import BaseModel, Field

from image_support import get_target_width
from transfer_summary import summarize_batch

from transfer_api import (
    MAX_DOCUMENTS,
    build_transfer_plan,
    find_section,
    get_blocks,
    insert_planned_paragraph,
    validate_section_permissions,
)


class SectionReplacement(BaseModel):
    source_section_id: str
    target_section_id: str


class BatchTransferRequest(BaseModel):
    source_document_id: str
    target_document_id: str
    replacements: list[SectionReplacement] = Field(
        min_length=1,
        max_length=50,
    )


def register_batch_transfer_routes(
    app,
    documents,
    extract_sections,
    get_heading_level,
):
    @app.post("/transfers/batch", status_code=201)
    def transfer_batch(request: BatchTransferRequest):
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
                    "Vuelve a cargar A y B si reiniciaste Python."
                ),
            )

        if len(documents) >= MAX_DOCUMENTS:
            raise HTTPException(
                status_code=409,
                detail=(
                    "Se alcanzó el límite temporal "
                    "de documentos."
                ),
            )

        source_document = Document(
            BytesIO(source_stored["content"])
        )

        # Todos los reemplazos se aplican a una única copia de B.
        target_document = Document(
            BytesIO(target_stored["content"])
        )

        target_blocks = get_blocks(target_document)

        prepared = []
        used_targets = set()

        # Primero valida y prepara todas las parejas.
        # Todavía no elimina contenido de B.
        for position, replacement in enumerate(
            request.replacements,
            start=1,
        ):
            try:
                source_section = find_section(
                    source_stored,
                    replacement.source_section_id,
                )

                target_section = find_section(
                    target_stored,
                    replacement.target_section_id,
                )

                if target_section["id"] in used_targets:
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            "Una sección de destino no puede "
                            "aparecer dos veces en el lote."
                        ),
                    )

                used_targets.add(target_section["id"])

                validate_section_permissions(
                    source_stored,
                    source_section,
                )

                validate_section_permissions(
                    target_stored,
                    target_section,
                )

                # Impide reemplazar una sección padre
                # y una subsección dentro de ella.
                for previous in prepared:
                    previous_target = previous["target"]

                    overlaps = (
                        target_section["start_index"]
                        < previous_target["end_index"]
                        and previous_target["start_index"]
                        < target_section["end_index"]
                    )

                    if overlaps:
                        raise HTTPException(
                            status_code=400,
                            detail=(
                                "Los destinos se superponen: "
                                f"«{target_section['title']}» y "
                                f"«{previous_target['title']}». "
                                "Selecciona secciones independientes."
                            ),
                        )

                target_range = target_blocks[
                    target_section["start_index"]:
                    target_section["end_index"]
                ]

                for element in target_range:
                    if (
                        element.find(
                            ".//" + qn("w:sectPr")
                        )
                        is not None
                    ):
                        raise HTTPException(
                            status_code=400,
                            detail=(
                                "La sección de destino contiene "
                                "un salto de sección no soportado."
                            ),
                        )

                plan = build_transfer_plan(
                    source_document,
                    target_document,
                    source_section,
                    target_section,
                    get_heading_level,
                )

                if not plan:
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            "El origen no contiene párrafos "
                            "compatibles para copiar."
                        ),
                    )

                title_element = target_blocks[
                    target_section["start_index"]
                ]

                needs_width = any(
                    "table" in paragraph_data
                    or any(
                        "image" in content
                        for content in paragraph_data["runs"]
                    )
                    for paragraph_data in plan
                )

                available_width = (
                    get_target_width(
                        target_document,
                        title_element,
                    )
                    if needs_width
                    else 0
                )

                prepared.append(
                    {
                        "position": position,
                        "source": source_section,
                        "target": target_section,
                        "plan": plan,
                        "available_width": available_width,
                    }
                )

            except HTTPException as error:
                raise HTTPException(
                    status_code=error.status_code,
                    detail=(
                        f"Transferencia {position}: "
                        f"{error.detail}"
                    ),
                ) from None

        # Aplica los reemplazos desde el final hacia el principio.
        ordered = sorted(
            prepared,
            key=lambda item: item["target"]["start_index"],
            reverse=True,
        )

        for item in ordered:
            try:
                target_section = item["target"]

                title_element = target_blocks[
                    target_section["start_index"]
                ]

                old_content = target_blocks[
                    target_section["content_start_index"]:
                    target_section["end_index"]
                ]

                for element in old_content:
                    element.getparent().remove(element)

                insertion_point = title_element

                # Cada sección reemplazada tiene
                # su propio estado de numeración.
                list_state = {}

                for paragraph_data in item["plan"]:
                    insertion_point = insert_planned_paragraph(
                        target_document,
                        insertion_point,
                        paragraph_data,
                        list_state,
                        item["available_width"],
                    )

            except HTTPException as error:
                # No se guarda la copia de B si falla la inserción.
                raise HTTPException(
                    status_code=error.status_code,
                    detail=(
                        f"Transferencia {item['position']}: "
                        f"{error.detail}"
                    ),
                ) from None

        # Guarda el resultado únicamente después
        # de completar todos los reemplazos.
        output = BytesIO()
        target_document.save(output)

        document_id = str(uuid4())
        filename = "B_actualizado.docx"

        sections = extract_sections(target_document)

        documents[document_id] = {
            "filename": filename,
            "content": output.getvalue(),
            "sections": sections,
        }

        summary = summarize_batch(prepared)

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