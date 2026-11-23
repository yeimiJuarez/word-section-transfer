from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from fastapi import HTTPException


def reject(message):
    raise HTTPException(status_code=400, detail=message)


def find_by_id(parent, tag, attribute, value):
    for element in parent.findall(qn(tag)):
        if element.get(qn(attribute)) == str(value):
            return element

    return None


def get_numbering_properties(paragraph):
    """
    Resuelve numId e ilvl desde el párrafo y sus estilos.
    numId=0 desactiva la numeración.
    """
    num_id = None
    level = None

    sources = [paragraph._p.pPr]

    style = paragraph.style
    visited = set()

    while style is not None:
        if style.style_id in visited:
            break

        visited.add(style.style_id)
        sources.append(style.element.pPr)
        style = style.base_style

    for properties in sources:
        if properties is None:
            continue

        num_properties = properties.find(qn("w:numPr"))

        if num_properties is None:
            continue

        num_element = num_properties.find(qn("w:numId"))
        level_element = num_properties.find(qn("w:ilvl"))

        if num_id is None and num_element is not None:
            num_id = int(num_element.get(qn("w:val")))

        if level is None and level_element is not None:
            level = int(level_element.get(qn("w:val")))

    if num_id is None or num_id == 0:
        return None

    return {
        "num_id": num_id,
        "level": level if level is not None else 0,
    }


def detect_list(paragraph, document):
    properties = get_numbering_properties(paragraph)

    if properties is None:
        return None

    level = properties["level"]
    preview = paragraph.text.strip()[:150] or "(Sin texto visible)"

    if not 0 <= level <= 8:
        reject(
            "La lista tiene un nivel fuera del rango admitido. "
            f"Párrafo: «{preview}»."
        )

    try:
        numbering = document.part.numbering_part.element
    except KeyError:
        reject("El documento no contiene definiciones de numeración.")

    num = find_by_id(
        numbering,
        "w:num",
        "w:numId",
        properties["num_id"],
    )

    if num is None:
        reject(
            "No se encontró la definición de la lista. "
            f"Párrafo: «{preview}»."
        )

    reference = num.find(qn("w:abstractNumId"))

    if reference is None:
        reject("La lista no tiene una definición abstracta válida.")

    abstract = find_by_id(
        numbering,
        "w:abstractNum",
        "w:abstractNumId",
        reference.get(qn("w:val")),
    )

    if abstract is None:
        reject("No se encontró la definición abstracta de la lista.")

    if abstract.find(qn("w:numStyleLink")) is not None:
        reject(
            "La lista utiliza una definición vinculada a un estilo "
            "todavía no soportada."
        )

    level_definition = find_by_id(
        abstract,
        "w:lvl",
        "w:ilvl",
        level,
    )

    override = find_by_id(
        num,
        "w:lvlOverride",
        "w:ilvl",
        level,
    )

    if override is not None:
        overridden_level = override.find(qn("w:lvl"))

        if overridden_level is not None:
            level_definition = overridden_level

    if level_definition is None:
        reject(
            f"No se encontró la definición del nivel {level + 1}."
        )

    if level_definition.find(qn("w:lvlPicBulletId")) is not None:
        reject("Las viñetas basadas en imágenes no están soportadas.")

    format_element = level_definition.find(qn("w:numFmt"))

    if format_element is None:
        reject("La lista no tiene un formato definido.")

    number_format = format_element.get(qn("w:val"))

    if number_format == "bullet":
        kind = "bullet"
    elif number_format in {"decimal", "decimalZero"}:
        kind = "decimal"
    else:
        reject(
            f"El formato «{number_format}» del nivel {level + 1} "
            f"todavía no está soportado. Párrafo: «{preview}»."
        )

    return {
        "kind": kind,
        "source_num_id": properties["num_id"],
        "level": level,
    }


def append_value(parent, tag, value):
    element = OxmlElement(tag)
    element.set(qn("w:val"), str(value))
    parent.append(element)
    return element


def next_id(numbering, tag, attribute, minimum=0):
    ids = [
        int(element.get(qn(attribute)))
        for element in numbering.findall(qn(tag))
    ]

    return max(ids + [minimum - 1]) + 1


def configure_level(level_element, kind, level_index):
    format_element = level_element.find(qn("w:numFmt"))
    text_element = level_element.find(qn("w:lvlText"))

    format_element.set(qn("w:val"), kind)

    bullet_symbols = ("•", "◦", "▪")

    text_element.set(
        qn("w:val"),
        (
            bullet_symbols[level_index % len(bullet_symbols)]
            if kind == "bullet"
            else f"%{level_index + 1}."
        ),
    )

    previous_run_properties = level_element.find(qn("w:rPr"))

    if previous_run_properties is not None:
        level_element.remove(previous_run_properties)

    if kind == "bullet":
        run_properties = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        fonts.set(qn("w:ascii"), "Arial")
        fonts.set(qn("w:hAnsi"), "Arial")
        run_properties.append(fonts)
        level_element.append(run_properties)


def create_default_list(document, kind):
    """
    Crea una definición multinivel independiente.
    No modifica las listas que ya existían en B.
    """
    try:
        numbering = document.part.numbering_part.element
    except KeyError:
        reject(
            "B no contiene una parte de numeración. "
            "Añade una lista de prueba a su plantilla y vuelve a cargarlo."
        )

    abstract_id = next_id(
        numbering,
        "w:abstractNum",
        "w:abstractNumId",
    )

    num_id = next_id(
        numbering,
        "w:num",
        "w:numId",
        minimum=1,
    )

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))

    append_value(abstract, "w:multiLevelType", "multilevel")

    for level_index in range(9):
        level = OxmlElement("w:lvl")
        level.set(qn("w:ilvl"), str(level_index))

        append_value(level, "w:start", 1)
        append_value(level, "w:numFmt", kind)
        append_value(level, "w:lvlText", "")
        append_value(level, "w:suff", "tab")
        append_value(level, "w:lvlJc", "left")

        # El nivel inferior reinicia al avanzar su nivel padre.
        if level_index > 0:
            append_value(level, "w:lvlRestart", level_index)

        left_indent = 720 * (level_index + 1)

        properties = OxmlElement("w:pPr")

        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "num")
        tab.set(qn("w:pos"), str(left_indent))
        tabs.append(tab)
        properties.append(tabs)

        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), str(left_indent))
        indent.set(qn("w:hanging"), "360")
        properties.append(indent)

        level.append(properties)
        configure_level(level, kind, level_index)
        abstract.append(level)

    existing_nums = numbering.findall(qn("w:num"))

    if existing_nums:
        existing_nums[0].addprevious(abstract)
    else:
        cleanup = numbering.find(qn("w:numIdMacAtCleanup"))

        if cleanup is not None:
            cleanup.addprevious(abstract)
        else:
            numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    append_value(num, "w:abstractNumId", abstract_id)

    cleanup = numbering.find(qn("w:numIdMacAtCleanup"))

    if cleanup is not None:
        cleanup.addprevious(num)
    else:
        numbering.append(num)

    return num_id


def set_target_level_format(document, num_id, level_index, kind):
    numbering = document.part.numbering_part.element

    num = find_by_id(
        numbering,
        "w:num",
        "w:numId",
        num_id,
    )

    reference = num.find(qn("w:abstractNumId"))

    abstract = find_by_id(
        numbering,
        "w:abstractNum",
        "w:abstractNumId",
        reference.get(qn("w:val")),
    )

    level = find_by_id(
        abstract,
        "w:lvl",
        "w:ilvl",
        level_index,
    )

    configure_level(level, kind, level_index)


def apply_list_format(document, paragraph, list_spec, state):
    if list_spec is None:
        state.clear()
        return

    # Cambiar de nivel no crea una lista independiente.
    key = list_spec["source_num_id"]
    level = list_spec.get("level", 0)
    kind = list_spec["kind"]

    if state.get("key") != key:
        state.clear()
        state["key"] = key
        state["target_num_id"] = create_default_list(
            document,
            kind,
        )
        state["level_formats"] = {}

    previous_kind = state["level_formats"].get(level)

    if previous_kind is not None and previous_kind != kind:
        reject(
            "Una misma lista cambia de formato dentro del mismo nivel. "
            "Esta estructura todavía no está soportada."
        )

    if previous_kind is None:
        set_target_level_format(
            document,
            state["target_num_id"],
            level,
            kind,
        )
        state["level_formats"][level] = kind

    properties = paragraph._p.get_or_add_pPr()
    previous = properties.find(qn("w:numPr"))

    if previous is not None:
        properties.remove(previous)

    num_properties = OxmlElement("w:numPr")
    append_value(num_properties, "w:ilvl", level)
    append_value(
        num_properties,
        "w:numId",
        state["target_num_id"],
    )

    properties.insert_element_before(
        num_properties,
        "w:suppressLineNumbers",
        "w:pBdr",
        "w:shd",
        "w:tabs",
        "w:spacing",
        "w:ind",
        "w:jc",
        "w:outlineLvl",
        "w:sectPr",
    )