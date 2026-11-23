from io import BytesIO
from uuid import uuid4
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware


app = FastAPI(title="Word Section Transfer")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB por archivo
MAX_UNCOMPRESSED_SIZE = 100 * 1024 * 1024
MAX_DOCUMENTS = 10

# Almacenamiento temporal en memoria para el prototipo local.
documents: dict[str, dict] = {}


@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "Backend funcionando",
    }


def validate_docx(content: bytes):
    """Comprueba la estructura básica y limita el tamaño descomprimido."""
    try:
        with ZipFile(BytesIO(content)) as archive:
            entries = archive.infolist()
            names = {entry.filename for entry in entries}

            required = {
                "[Content_Types].xml",
                "word/document.xml",
            }

            if not required.issubset(names):
                raise HTTPException(
                    status_code=400,
                    detail="El archivo no tiene la estructura de un DOCX.",
                )

            if len(entries) > 5000:
                raise HTTPException(
                    status_code=400,
                    detail="El documento contiene demasiados elementos.",
                )

            total_size = sum(entry.file_size for entry in entries)

            if total_size > MAX_UNCOMPRESSED_SIZE:
                raise HTTPException(
                    status_code=413,
                    detail="El contenido descomprimido supera el límite.",
                )

            if any(entry.flag_bits & 1 for entry in entries):
                raise HTTPException(
                    status_code=400,
                    detail="No se admiten archivos ZIP cifrados.",
                )

    except BadZipFile:
        raise HTTPException(
            status_code=400,
            detail="El archivo no es un DOCX válido.",
        ) from None


def get_heading_level(paragraph: Paragraph) -> int | None:
    """Busca el nivel de esquema en el párrafo y sus estilos heredados."""
    paragraph_properties = paragraph._p.pPr

    if paragraph_properties is not None:
        outline = paragraph_properties.find(qn("w:outlineLvl"))

        if outline is not None:
            value = int(outline.get(qn("w:val")))
            return value + 1 if 0 <= value <= 8 else None

    style = paragraph.style
    visited = set()

    while style is not None:
        if style.style_id in visited:
            break

        visited.add(style.style_id)
        properties = style.element.pPr

        if properties is not None:
            outline = properties.find(qn("w:outlineLvl"))

            if outline is not None:
                value = int(outline.get(qn("w:val")))
                return value + 1 if 0 <= value <= 8 else None

        # Respaldo para estilos integrados de título.
        name = (style.name or "").strip().lower()

        for prefix in ("heading ", "título ", "titulo "):
            if name.startswith(prefix):
                suffix = name[len(prefix):]

                if suffix.isdigit() and 1 <= int(suffix) <= 9:
                    return int(suffix)

        style = style.base_style

    return None


def extract_sections(document) -> list[dict]:
    """
    Los índices corresponden a elementos del cuerpo del documento,
    incluyendo párrafos y tablas.
    """
    blocks = [
        element
        for element in document.element.body
        if element.tag != qn("w:sectPr")
    ]

    sections = []
    stack = []

    for index, element in enumerate(blocks):
        if element.tag != qn("w:p"):
            continue

        paragraph = Paragraph(element, document)
        level = get_heading_level(paragraph)

        if level is None:
            continue

        # Cierra las secciones que terminan al llegar a este título.
        while stack and stack[-1]["level"] >= level:
            stack.pop()["end_index"] = index

        section = {
            "id": f"section-{index}",
            "title": paragraph.text.strip() or "(Título sin texto)",
            "transfer_enabled": paragraph.text.strip().casefold() != "contents",
            "level": level,
            "parent_id": stack[-1]["id"] if stack else None,
            "start_index": index,
            "content_start_index": index + 1,
            "end_index": len(blocks),
        }

        sections.append(section)
        stack.append(section)

    return sections


@app.post("/documents", status_code=201)
async def upload_document(file: UploadFile = File(...)):
    filename = file.filename or ""

    if not filename.lower().endswith(".docx"):
        await file.close()
        raise HTTPException(
            status_code=400,
            detail="Solo se admiten archivos .docx.",
        )

    if len(documents) >= MAX_DOCUMENTS:
        await file.close()
        raise HTTPException(
            status_code=409,
            detail="Se alcanzó el límite temporal de 10 documentos.",
        )

    content = bytearray()

    try:
        while chunk := await file.read(1024 * 1024):
            content.extend(chunk)

            if len(content) > MAX_FILE_SIZE:
                raise HTTPException(
                    status_code=413,
                    detail="El archivo supera el límite de 10 MB.",
                )
    finally:
        await file.close()

    if not content:
        raise HTTPException(
            status_code=400,
            detail="El archivo está vacío.",
        )

    content = bytes(content)
    validate_docx(content)

    try:
        document = Document(BytesIO(content))
        sections = extract_sections(document)
    except Exception:
        raise HTTPException(
            status_code=400,
            detail="No se pudo interpretar la estructura del documento.",
        ) from None

    document_id = str(uuid4())

    documents[document_id] = {
        "filename": filename,
        "content": content,
        "sections": sections,
    }

    return {
        "document_id": document_id,
        "filename": filename,
        "section_count": len(sections),
        "sections": sections,
    }


@app.get("/documents/{document_id}/sections")
def list_sections(document_id: str):
    stored_document = documents.get(document_id)

    if stored_document is None:
        raise HTTPException(
            status_code=404,
            detail="Documento no encontrado. Puede haberse reiniciado el servidor.",
        )

    return {
        "document_id": document_id,
        "filename": stored_document["filename"],
        "sections": stored_document["sections"],
    }


from transfer_api import register_transfer_routes

register_transfer_routes(
    app,
    documents,
    extract_sections,
    get_heading_level,
)

from batch_transfer_api import register_batch_transfer_routes

register_batch_transfer_routes(
    app,
    documents,
    extract_sections,
    get_heading_level,
)