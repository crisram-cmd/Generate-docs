import argparse
import json
import os
import re
import uuid
import traceback
from io import BytesIO
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from google.cloud import storage
from google.oauth2 import service_account
from google.genai import types
from google.adk.agents.callback_context import CallbackContext


FONT_MAIN = "Times New Roman"
MARGIN_CM = 42 * 2.54 / 72 
PT_FOLA_CODE = 10.5
PT_HEADER_LINE = 11
PT_TITLE_SOLICITUD = 12.5
PT_SECTION = 11
PT_BODY = 10.3
PT_BODY_COMPACT = 10.0
PT_SMALL = 9.75
PT_SMALLER = 9.6
PT_TABLE_CELL = 8.45
PT_LEGAL_FOOTER = 8.6
PT_CONTINUATION_TITLE = 10.8
PT_CONSTANCIA = 10.2
PT_FIRMA_BLOCK = 10.5
PT_USER_DECL = 9.3


# =========================
# GCS
# =========================

def build_storage_client():
    key_path = os.environ.get("ADK_CREDENTIALS_PATH")

    if key_path and os.path.exists(key_path):
        creds = service_account.Credentials.from_service_account_file(key_path)
        client = storage.Client(credentials=creds, project=creds.project_id)
        return client, creds

    client = storage.Client()
    return client, None


def upload_to_gcs(bucket_name: str, object_name: str, data: bytes) -> dict:
    client, _ = build_storage_client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(object_name)

    blob.upload_from_string(
        data,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    return {
        "bucket_name": bucket_name,
        "object_name": object_name,
    }


def _build_artifact_filename(case_id: str, safe_dui: str, prefix: str) -> str:
    short_case_id = case_id[:8]
    short_dui = safe_dui[:12]
    return f"{prefix}_{short_dui}_{short_case_id}.docx"


# =========================
# HELPERS DE FORMATO
# =========================

def _apply_font(run, font_name=None, size_pt=None, color_rgb=(0, 0, 0), bold=False, italic=False):
    if font_name is None:
        font_name = FONT_MAIN
    run.font.name = font_name

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

    if size_pt is not None:
        run.font.size = Pt(size_pt)


def _set_default_font(doc: Document, font_name=None, size_pt=None):
    if font_name is None:
        font_name = FONT_MAIN
    if size_pt is None:
        size_pt = PT_BODY
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(size_pt)

    rPr = style.element.get_or_add_rPr()
    rfonts = rPr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rPr.append(rfonts)

    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:cs"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)


def _set_spacing(paragraph, before=0, after=0, line=1.0):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def _normalize_text(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(str(x).strip() for x in value if str(x).strip())
    return str(value).strip()


def _extract_year_20xx_from_text(text) -> str:
    """Primer año 20xx encontrado en texto (p. ej. expediente UDDT-2025-0142)."""
    if not text:
        return ""
    m = re.search(r"(20\d{2})", str(text))
    return m.group(1) if m else ""


def _display_fola_form_year(
    anio_field,
    *,
    expediente_hint: str | None = None,
    trailing_dot: bool = False,
) -> str:
    """
    Año para líneas de expediente / fecha: 2 cifras -> 20xx, 4 cifras tal cual.
    Si anio en JSON no coincide con el año en el número de expediente, se usa el del expediente.
    """
    a = _normalize_text(anio_field)
    hint = _extract_year_20xx_from_text(expediente_hint or "")
    sfx = "." if trailing_dot else ""

    def pack(y: str) -> str:
        if not y:
            return "20____" + sfx
        return y + sfx

    if len(a) == 4 and a.isdigit():
        return pack(a)
    if len(a) == 2 and a.isdigit():
        cand = f"20{a}"
        if hint and cand != hint:
            return pack(hint)
        return pack(cand)
    if hint:
        return pack(hint)
    if not a:
        return pack("")
    return pack("")


def _expediente_header_year_display(anio_field, expediente_hint) -> str:
    """
    Texto del año junto al expediente: '20' y dos posiciones (cifras o guiones bajos).
    Sin año resoluble -> '20__' como en el formulario impreso.
    """
    a = _normalize_text(anio_field)
    hint = _extract_year_20xx_from_text(expediente_hint or "")
    y4 = ""
    if len(a) == 4 and a.isdigit():
        y4 = a
    elif len(a) == 2 and a.isdigit():
        cand = f"20{a}"
        y4 = hint if hint and cand != hint else cand
    elif hint:
        y4 = hint
    if len(y4) == 4 and y4.isdigit():
        return "20" + y4[2:4]
    return "20__"


def _safe(value, default=""):
    value = _normalize_text(value)
    return value if value else default


def _chk(value=False):
    return "☒" if bool(value) else "☐"


def _underline(length=40):
    return "_" * length


def _form_body_width_cm(doc: Document) -> float:
    """Ancho util entre margenes (cm), para tablas de una fila a ancho completo."""
    sec = doc.sections[0]
    return max(round(sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm, 3), 10.0)


def _add_full_width_horizontal_rule(doc, size_pt=None, after=1):
    """Linea en blanco de punta a punta: tabla 1x1 con borde inferior (visible en Word)."""
    if size_pt is None:
        size_pt = PT_BODY
    w_cm = _form_body_width_cm(doc)
    t = doc.add_table(rows=1, cols=1)
    _set_widths(t, [w_cm])
    _form_row_prepare(t)
    c = t.cell(0, 0)
    c.text = ""
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=0, after=0, line=1.0)
    r = p.add_run("\u00a0")
    _apply_font(r, size_pt=size_pt)
    _set_cell_bottom_rule(c)
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    _add_table_spacer(doc, after=after)
    return t


def _usable_body_width(doc: Document):
    """Ancho útil del cuerpo (entre márgenes) para tabs y tablas."""
    sec = doc.sections[0]
    return sec.page_width - sec.left_margin - sec.right_margin


def _add_line_left_right_tab(doc, left_text: str, right_text: str, size_pt=None, after=0):
    """Una línea: bloque izquierdo y bloque alineado a la derecha (tab derecho al margen)."""
    if size_pt is None:
        size_pt = PT_BODY
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=0, after=after, line=1.0)
    _ = p.paragraph_format.tab_stops.add_tab_stop(_usable_body_width(doc), WD_TAB_ALIGNMENT.RIGHT)
    r1 = p.add_run(left_text + "\t")
    _apply_font(r1, size_pt=size_pt)
    r2 = p.add_run(right_text)
    _apply_font(r2, size_pt=size_pt)
    return p


def _set_cell_borders_nil(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        tag = f"w:{edge}"
        el = tcBorders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            tcBorders.append(el)
        el.set(qn("w:val"), "nil")


def _format_borderless_form_table(table):
    """Tabla de formulario sin bordes visibles, alineada al ancho del contenido."""
    _strip_default_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    for row in table.rows:
        for cell in row.cells:
            _set_cell_borders_nil(cell)
            _set_cell_padding(cell, top=12, start=12, bottom=12, end=12)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def _add_form_row_table(doc, cells: list, widths_cm: list, size_pt=None, after=0):
    """Una fila, N columnas, sin bordes (equivalencia a campos repartidos en el PDF)."""
    if size_pt is None:
        size_pt = PT_BODY
    n = len(cells)
    table = doc.add_table(rows=1, cols=n)
    _set_widths(table, widths_cm)
    _format_borderless_form_table(table)
    for i, text in enumerate(cells):
        _cell_text(table.cell(0, i), str(text) if text is not None else "", size_pt=size_pt)
    spacer = doc.add_paragraph()
    _set_spacing(spacer, before=0, after=after, line=1.0)
    return table


def _add_line(doc, text="", size_pt=None, bold=False, italic=False, center=False, justify=False, after=0):
    if size_pt is None:
        size_pt = PT_BODY
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    _set_spacing(p, before=0, after=after, line=1.0)
    r = p.add_run(text)
    _apply_font(r, size_pt=size_pt, bold=bold, italic=italic)
    return p


def _add_section_title(doc, text, size_pt=None, before=0, after=2):
    if size_pt is None:
        size_pt = PT_SECTION
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=before, after=after, line=1.0)
    r = p.add_run(text)
    _apply_font(r, size_pt=size_pt, bold=True)


def _set_cell_border(cell, size="6", color="000000"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("top", "left", "bottom", "right"):
        tag = f"w:{edge}"
        element = tcBorders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)

        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def _set_cell_padding(cell, top=35, start=35, bottom=35, end=35):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)

    for margin, value in {
        "top": top,
        "start": start,
        "bottom": bottom,
        "end": end,
    }.items():
        node = tcMar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tcMar.append(node)

        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _format_table(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for row in table.rows:
        for cell in row.cells:
            _set_cell_border(cell)
            _set_cell_padding(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def _set_widths(table, widths_cm):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])


def _cell_text(cell, text="", bold=False, center=False, size_pt=None, italic=False):
    if size_pt is None:
        size_pt = PT_TABLE_CELL
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=0, after=0, line=1.0)

    r = p.add_run(text)
    _apply_font(r, size_pt=size_pt, bold=bold, italic=italic)


def _set_cell_nowrap(cell):
    """Evita que Word parta la palabra (p. ej. Expedient / e) en columnas estrechas."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    el = tcPr.find(qn("w:noWrap"))
    if el is None:
        el = OxmlElement("w:noWrap")
        tcPr.append(el)


def _set_cell_bottom_rule(cell, sz="10", color="000000"):
    """Borde inferior visible; resto nil (linea de formulario fija)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        tag = f"w:{edge}"
        el = tcBorders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            tcBorders.append(el)
        if edge == "bottom":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), sz)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")


def _display_or_placeholder(value, placeholder="______________"):
    v = _normalize_text(value)
    return v if v else placeholder


def _cell_label_then_centered_over_line(
    cell,
    label,
    value,
    size_pt=None,
    label_cm=2.35,
    field_cm=None,
):
    """
    Misma fila: etiqueta a la izquierda y valor centrado solo sobre la linea de la
    celda derecha (estilo Procuraduria Auxiliar de [___]), no etiqueta arriba y campo abajo.
    """
    if size_pt is None:
        size_pt = PT_BODY
    parent_w = cell.width
    cell.text = ""
    if field_cm is None:
        if parent_w is not None:
            field_cm = max(round(parent_w.cm - label_cm - 0.12, 2), 1.0)
        else:
            field_cm = 6.5
    inner = cell.add_table(rows=1, cols=2)
    _set_widths(inner, [label_cm, field_cm])
    _form_row_prepare_inner(inner)
    _cell_text(inner.cell(0, 0), label, size_pt=size_pt)
    _cell_centered_over_line_only(inner.cell(0, 1), value, size_pt=size_pt)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def _cell_expediente_label_and_short_rule_right(
    parent_cell,
    value,
    *,
    zone_width_cm: float,
    label_width_cm: float = 2.85,
    rule_width_cm: float = 6.55,
    size_pt=None,
):
    """
    'Expediente ' pegado al codigo, con raya corta solo bajo el numero, bloque empujado a la
    derecha de la zona (relleno a la izquierda) antes del año.
    """
    if size_pt is None:
        size_pt = PT_BODY
    pad_w = max(round(zone_width_cm - label_width_cm - rule_width_cm, 2), 0.35)
    parent_cell.text = ""
    inner = parent_cell.add_table(rows=1, cols=3)
    _set_widths(inner, [pad_w, label_width_cm, rule_width_cm])
    _form_row_prepare_inner(inner)
    c_pad = inner.cell(0, 0)
    c_pad.text = ""
    pp = c_pad.paragraphs[0]
    pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(pp, before=0, after=0, line=1.0)
    pp.add_run("\u00a0")
    _apply_font(pp.runs[0], size_pt=size_pt)
    c_lab = inner.cell(0, 1)
    _cell_text(c_lab, "Expediente ", size_pt=size_pt)
    _set_cell_nowrap(c_lab)
    _cell_centered_over_line_only(inner.cell(0, 2), value, size_pt=size_pt)
    for cell in inner.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM


def _cell_centered_over_line_only(cell, value, size_pt=None):
    """Solo valor centrado sobre linea inferior (sin etiqueta en la celda).

    Si el valor esta vacio, no se dibujan guiones bajo el texto: solo el borde
    inferior de la celda (evita doble linea como en campos vacios del PDF).
    """
    if size_pt is None:
        size_pt = PT_BODY
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(p, before=0, after=0, line=1.0)
    v = _normalize_text(value)
    r = p.add_run(v if v else "\u00a0")
    _apply_font(r, size_pt=size_pt)
    _set_cell_bottom_rule(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def _strip_default_table_style(table):
    """Quita w:tblStyle (p. ej. Table Grid) para evitar líneas extra encima de bordes inferiores."""
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        return
    el = tbl_pr.find(qn("w:tblStyle"))
    if el is not None:
        tbl_pr.remove(el)


def _form_row_prepare(table):
    _strip_default_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    for row in table.rows:
        for cell in row.cells:
            _set_cell_borders_nil(cell)
            _set_cell_padding(cell, top=10, start=8, bottom=10, end=8)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def _form_row_prepare_inner(table):
    """Tabla anidada: menos padding para no duplicar aire con la celda padre."""
    _strip_default_table_style(table)
    table.autofit = False
    for row in table.rows:
        for cell in row.cells:
            _set_cell_borders_nil(cell)
            _set_cell_padding(cell, top=2, start=2, bottom=2, end=2)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def _add_table_spacer(doc, after=0):
    p = doc.add_paragraph()
    _set_spacing(p, before=0, after=after, line=1.0)


# =========================
# BLOQUES FOLA03
# =========================

def _build_header(doc, first_page=False):
    """
    Encabezado del formulario.

    Solo la pagina 1 lleva el bloque institucional (Procuraduria / Unidad) y el titulo
    de solicitud; las demas paginas solo repiten la clave FOLA03 a la derecha.
    """
    p_code = doc.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_spacing(p_code, before=0, after=0)
    r_code = p_code.add_run("FOLA03")
    _apply_font(r_code, size_pt=PT_FOLA_CODE, bold=True)

    if not first_page:
        return

    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(p1, before=0, after=2)
    r1 = p1.add_run("PROCURADURIA GENERAL DE LA REPUBLICA")
    _apply_font(r1, size_pt=PT_HEADER_LINE, bold=True)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(p2, before=0, after=2)
    r2 = p2.add_run("UNIDAD DE DEFENSA DE LOS DERECHOS DEL TRABAJADOR")
    _apply_font(r2, size_pt=PT_HEADER_LINE, bold=True)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(p3, before=0, after=0)
    r3 = p3.add_run("SOLICITUD DE ASISTENCIA LEGAL PARA JUICIO DE TRABAJO")
    _apply_font(r3, size_pt=PT_TITLE_SOLICITUD, bold=True)


def _build_case_lines(doc, fola):
    # Expediente: etiqueta y codigo juntos; raya corta solo bajo el numero; año aparte sin raya.
    w_sp, w_mid, w_yr = 0.25, 15.88, 2.5
    t0 = doc.add_table(rows=1, cols=3)
    _set_widths(t0, [w_sp, w_mid, w_yr])
    _form_row_prepare(t0)
    c_pad0 = t0.cell(0, 0)
    c_pad0.text = ""
    pz = c_pad0.paragraphs[0]
    pz.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(pz, before=0, after=0, line=1.0)
    pz.add_run("\u00a0")
    _apply_font(pz.runs[0], size_pt=PT_BODY)
    _cell_expediente_label_and_short_rule_right(
        t0.cell(0, 1),
        fola.get("expediente"),
        zone_width_cm=w_mid,
        label_width_cm=2.85,
        rule_width_cm=6.55,
        size_pt=PT_BODY,
    )
    # Año sin linea inferior; misma linea visual que expediente; formato 20 + dos (__ o cifras).
    c_y = t0.cell(0, 2)
    c_y.text = ""
    py = c_y.paragraphs[0]
    py.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_spacing(py, before=0, after=0, line=1.0)
    yr_show = _expediente_header_year_display(
        fola.get("anio_expediente"),
        fola.get("expediente"),
    )
    ry = py.add_run(yr_show)
    _apply_font(ry, size_pt=PT_BODY)
    _set_cell_nowrap(c_y)
    for cell in t0.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    _add_table_spacer(doc, after=2)

    # Procuraduria auxiliar / horas: campos centrados sobre linea.
    t1 = doc.add_table(rows=1, cols=4)
    _set_widths(t1, [3.65, 7.48, 1.15, 6.35])
    _form_row_prepare(t1)
    _cell_text(t1.cell(0, 0), "Procuraduria Auxiliar de", size_pt=PT_BODY)
    _cell_centered_over_line_only(t1.cell(0, 1), fola.get("procuraduria_auxiliar"))
    _cell_text(t1.cell(0, 2), "a las", size_pt=PT_BODY)
    c_h = t1.cell(0, 3)
    c_h.text = ""
    ph = c_h.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(ph, before=0, after=0, line=1.0)
    h_at = _normalize_text(fola.get("hora_atencion"))
    rh = ph.add_run(f"{h_at} horas" if h_at else "horas")
    _apply_font(rh, size_pt=PT_BODY)
    _set_cell_bottom_rule(c_h)
    _add_table_spacer(doc, after=2)

    # Fecha de atencion: ultima caja "20xx." unificada (mismo criterio que expediente).
    t2 = doc.add_table(rows=1, cols=6)
    _set_widths(t2, [1.85, 3.85, 1.45, 0.48, 3.55, 7.45])
    _form_row_prepare(t2)
    _cell_centered_over_line_only(t2.cell(0, 0), fola.get("minutos_atencion"))
    _cell_text(t2.cell(0, 1), "minutos del dia", size_pt=PT_BODY)
    _cell_centered_over_line_only(t2.cell(0, 2), fola.get("dia_atencion"))
    _cell_text(t2.cell(0, 3), "de", size_pt=PT_BODY)
    _cell_centered_over_line_only(t2.cell(0, 4), fola.get("mes_atencion"))
    cy = t2.cell(0, 5)
    cy.text = ""
    py = cy.paragraphs[0]
    py.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(py, before=0, after=0, line=1.0)
    tail = _display_fola_form_year(
        fola.get("anio_atencion"),
        expediente_hint=None,
        trailing_dot=True,
    )
    ry = py.add_run(tail)
    _apply_font(ry, size_pt=PT_BODY)
    _set_cell_bottom_rule(cy)
    _add_table_spacer(doc, after=2)


def _add_label_and_centered_field_row(doc, label, value, label_col_cm, size_pt=None, after=0):
    """Fila etiqueta (izq.) + campo con valor centrado sobre linea (resto del ancho)."""
    if size_pt is None:
        size_pt = PT_BODY
    w2 = max(18.63 - label_col_cm, 4.0)
    t = doc.add_table(rows=1, cols=2)
    _set_widths(t, [label_col_cm, w2])
    _form_row_prepare(t)
    _cell_text(t.cell(0, 0), label, size_pt=size_pt)
    _cell_centered_over_line_only(t.cell(0, 1), value, size_pt=size_pt)
    _add_table_spacer(doc, after=after)


def _build_user_data(doc, worker, fola):
    _add_section_title(doc, "DATOS DE USUARIO/A", after=4)

    gender = _safe(worker.get("worker_gender")).upper()

    tn = doc.add_table(rows=1, cols=2)
    _set_widths(tn, [9.315, 9.315])
    _form_row_prepare(tn)
    _cell_label_then_centered_over_line(
        tn.cell(0, 0), "Nombre: ", worker.get("worker_name"), size_pt=PT_BODY, label_cm=1.9, field_cm=7.28
    )
    _cell_label_then_centered_over_line(
        tn.cell(0, 1), "Conocido/a por ", fola.get("known_as"), size_pt=PT_BODY, label_cm=2.85, field_cm=6.33
    )
    _add_table_spacer(doc, after=1)

    td = doc.add_table(rows=1, cols=3)
    _set_widths(td, [7.85, 5.45, 5.33])
    _form_row_prepare(td)
    c_de = td.cell(0, 0)
    c_de.text = ""
    pd0 = c_de.paragraphs[0]
    pd0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(pd0, before=0, after=0, line=1.0)
    pd0.add_run(
        f"De {_safe(worker.get('worker_age'), '____')} años de edad. Genero: "
        f"{_chk(gender == 'F')} F  {_chk(gender == 'M')} M"
    )
    _apply_font(pd0.runs[0], size_pt=PT_BODY)
    _cell_label_then_centered_over_line(
        td.cell(0, 1),
        "Estado Familiar: ",
        worker.get("worker_marital_status"),
        size_pt=PT_BODY,
        label_cm=2.85,
        field_cm=2.48,
    )
    _cell_label_then_centered_over_line(
        td.cell(0, 2),
        "Profesion/Oficio: ",
        worker.get("worker_profession_or_trade"),
        size_pt=PT_BODY,
        label_cm=3.05,
        field_cm=2.16,
    )
    _add_table_spacer(doc, after=1)

    dui_mid = (
        f"{_safe(fola.get('dui_issue_day'), '____')} de {_safe(fola.get('dui_issue_month'), '________')} "
        f"de {_safe(fola.get('dui_issue_year'), '______')}"
    )
    td2 = doc.add_table(rows=1, cols=3)
    _set_widths(td2, [5.45, 6.95, 6.23])
    _form_row_prepare(td2)
    _cell_label_then_centered_over_line(
        td2.cell(0, 0), "DUI: ", worker.get("worker_dui"), size_pt=PT_BODY, label_cm=1.35, field_cm=4.0
    )
    _cell_label_then_centered_over_line(
        td2.cell(0, 1),
        "expedido el ",
        dui_mid.strip(),
        size_pt=PT_BODY,
        label_cm=2.0,
        field_cm=4.83,
    )
    _cell_label_then_centered_over_line(
        td2.cell(0, 2),
        "en ",
        fola.get("dui_issue_place"),
        size_pt=PT_BODY,
        label_cm=0.75,
        field_cm=5.36,
    )
    _add_table_spacer(doc, after=1)

    _add_label_and_centered_field_row(
        doc,
        "Otro Documento de identidad (Extranjeros):",
        fola.get("other_document"),
        label_col_cm=8.45,
        size_pt=PT_BODY,
        after=1,
    )

    tnat = doc.add_table(rows=1, cols=3)
    _set_widths(tnat, [4.15, 9.65, 4.83])
    _form_row_prepare(tnat)
    _cell_label_then_centered_over_line(
        tnat.cell(0, 0), "Nacionalidad: ", worker.get("worker_nationality"), size_pt=PT_BODY, label_cm=2.55, field_cm=1.48
    )
    _cell_label_then_centered_over_line(
        tnat.cell(0, 1),
        "Domicilio: ",
        worker.get("worker_address"),
        size_pt=PT_BODY,
        label_cm=2.15,
        field_cm=7.38,
    )
    _cell_label_then_centered_over_line(
        tnat.cell(0, 2),
        "Departamento: ",
        fola.get("worker_department"),
        size_pt=PT_BODY,
        label_cm=2.65,
        field_cm=2.06,
    )
    _add_table_spacer(doc, after=1)

    _add_label_and_centered_field_row(
        doc,
        "Notificaciones:",
        fola.get("notifications_address"),
        label_col_cm=2.8,
        size_pt=PT_BODY,
        after=1,
    )

    ttel = doc.add_table(rows=1, cols=2)
    _set_widths(ttel, [9.315, 9.315])
    _form_row_prepare(ttel)
    _cell_label_then_centered_over_line(
        ttel.cell(0, 0),
        "Telefono Residencia: ",
        fola.get("home_phone"),
        size_pt=PT_BODY,
        label_cm=3.45,
        field_cm=5.68,
    )
    _cell_label_then_centered_over_line(
        ttel.cell(0, 1), "Celular: ", worker.get("worker_phone"), size_pt=PT_BODY, label_cm=1.85, field_cm=7.28
    )
    _add_table_spacer(doc, after=1)

    _add_label_and_centered_field_row(
        doc,
        "Recomendado/a:",
        fola.get("recommended_by"),
        label_col_cm=3.2,
        size_pt=PT_BODY,
        after=1,
    )

    trec = doc.add_table(rows=1, cols=3)
    _set_widths(trec, [6.15, 6.15, 6.33])
    _form_row_prepare(trec)
    _cell_label_then_centered_over_line(
        trec.cell(0, 0), "Telefono(s): ", fola.get("recommended_phone"), size_pt=PT_BODY, label_cm=2.45, field_cm=3.58
    )
    _cell_label_then_centered_over_line(
        trec.cell(0, 1), "Celular(es): ", fola.get("recommended_cellphone"), size_pt=PT_BODY, label_cm=2.2, field_cm=3.83
    )
    _cell_label_then_centered_over_line(
        trec.cell(0, 2), "Residencia: ", "", size_pt=PT_BODY, label_cm=2.05, field_cm=4.16
    )
    _add_table_spacer(doc, after=1)

    _add_label_and_centered_field_row(
        doc,
        "# De Personas que dependen economicamente:",
        fola.get("economic_dependents"),
        label_col_cm=7.2,
        size_pt=PT_BODY,
        after=1,
    )

    tm = doc.add_table(rows=1, cols=3)
    _set_widths(tm, [3.2, 5.0, 10.43])
    _form_row_prepare(tm)
    _cell_label_then_centered_over_line(
        tm.cell(0, 0), "Dia ", fola.get("mintrab_day"), size_pt=PT_BODY, label_cm=0.95, field_cm=2.13
    )
    _cell_label_then_centered_over_line(
        tm.cell(0, 1), "Mes ", fola.get("mintrab_month"), size_pt=PT_BODY, label_cm=0.95, field_cm=3.93
    )
    _cell_label_then_centered_over_line(
        tm.cell(0, 2),
        "Año que comparecio al MINTRAB. ",
        fola.get("mintrab_year"),
        size_pt=PT_BODY,
        label_cm=4.95,
        field_cm=5.36,
    )
    _add_table_spacer(doc, after=1)


def _build_employer_data(doc, emp):
    _add_section_title(doc, "DATOS DE LA O EL EMPLEADOR", after=2)

    employer_type = _safe(emp.get("employer_type")).lower()
    # Orden como PDF: Persona Juridica, luego Persona Natural.
    _add_line(
        doc,
        f"{_chk(employer_type in ['persona juridica', 'persona jurídica'])} Persona Juridica        "
        f"{_chk(employer_type == 'persona natural')} Persona Natural",
        size_pt=PT_BODY_COMPACT,
        after=1,
    )

    place_type = _safe(emp.get("notification_place_type")).lower()

    _add_label_and_centered_field_row(
        doc,
        "Nombre /Razon Social/ denominacion:",
        emp.get("company_defendant"),
        label_col_cm=5.5,
        size_pt=PT_BODY,
        after=1,
    )

    tdom = doc.add_table(rows=1, cols=2)
    _set_widths(tdom, [9.315, 9.315])
    _form_row_prepare(tdom)
    _cell_label_then_centered_over_line(
        tdom.cell(0, 0), "Domicilio ", emp.get("company_address"), size_pt=PT_BODY, label_cm=2.0, field_cm=7.18
    )
    _cell_label_then_centered_over_line(
        tdom.cell(0, 1), "Departamento ", emp.get("company_department"), size_pt=PT_BODY, label_cm=2.55, field_cm=6.63
    )
    _add_table_spacer(doc, after=1)

    lines = [
        f"Representante Legal: {_safe(emp.get('legal_representative_name'), _underline(72))}",
    ]
    for line in lines:
        _add_line(doc, line, size_pt=PT_BODY, after=1)

    tmay = doc.add_table(rows=1, cols=2)
    _set_widths(tmay, [9.315, 9.315])
    _form_row_prepare(tmay)
    _cell_label_then_centered_over_line(
        tmay.cell(0, 0),
        "Mayor de edad y domicilio de ",
        emp.get("legal_representative_address"),
        size_pt=PT_BODY,
        label_cm=4.25,
        field_cm=4.93,
    )
    _cell_label_then_centered_over_line(
        tmay.cell(0, 1),
        "Departamento ",
        emp.get("legal_representative_department"),
        size_pt=PT_BODY,
        label_cm=2.55,
        field_cm=6.63,
    )
    _add_table_spacer(doc, after=1)

    lines = [
        f"Lugar del emplazamiento: {_safe(emp.get('company_notification_address'), _underline(72))}",
        f"{_chk(place_type == 'negocio')} Lugar donde habitualmente atiende sus negocios      "
        f"{_chk(place_type == 'residencia')} Lugar de Residencia      "
        f"{_chk(place_type == 'trabajo')} Lugar de Trabajo",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_BODY, after=1)


def _build_work_relation(doc, emp):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=2, after=2)

    r = p.add_run("RELACION DE TRABAJO")
    _apply_font(r, size_pt=PT_SECTION, bold=True)

    r2 = p.add_run(" " * 10 + f"{_chk(emp.get('substitution_patronal'))} SUSTITUCION PATRONAL")
    _apply_font(r2, size_pt=PT_BODY_COMPACT)

    _add_line(
        doc,
        f"FECHA DE INGRESO: DIA {_safe(emp.get('employment_start_day'), '______')}  "
        f"MES {_safe(emp.get('employment_start_month'), '______________')}  "
        f"AÑO {_safe(emp.get('employment_start_year'), '______')}  "
        f"CARGO: {_safe(emp.get('job_title'), _underline(28))}",
        size_pt=PT_BODY,
        after=1,
    )
    tlab = doc.add_table(rows=1, cols=2)
    _set_widths(tlab, [9.315, 9.315])
    _form_row_prepare(tlab)
    _cell_text(
        tlab.cell(0, 0),
        f"Desarrollo sus labores en: {_chk(emp.get('workplace_is_notification_place'))} Emplazamiento",
        size_pt=PT_BODY,
    )
    _cell_label_then_centered_over_line(
        tlab.cell(0, 1),
        f"{_chk(emp.get('workplace_is_other'))} Otro: ",
        emp.get("workplace"),
        size_pt=PT_BODY,
        label_cm=2.65,
        field_cm=6.53,
    )
    _add_table_spacer(doc, after=1)
    lines = [
        f"Consistian sus labores: {_safe(emp.get('actual_functions'), _underline(68))}",
        f"Jornada Ordinaria de Trabajo: {_chk(emp.get('ordinary_workday', True))} SI  "
        f"{_safe(emp.get('daily_hours'), '____')} Horas diarias        "
        f"{_chk(not emp.get('ordinary_workday', True))} NO  Generalmente laboraba",
        f"Horario: {_safe(emp.get('work_schedule'), _underline(78))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_BODY, after=1)

    for _ in range(4):
        _add_full_width_horizontal_rule(doc, size_pt=PT_BODY, after=1)


def _build_salary_page(doc, emp):
    _add_section_title(doc, "SALARIO:", after=4)

    salary_period = _safe(emp.get("salary_period")).lower()
    payment_period = _safe(emp.get("payment_period")).lower()
    payment_place = _safe(emp.get("payment_place")).lower()
    salary_type = _safe(emp.get("salary_type")).lower()

    lines = [
        f"UNIDAD TIEMPO (base global): $ {_safe(emp.get('salary_amount'), '__________')}  "
        f"{_chk(salary_period == 'mensual')} MENSUAL  {_chk(salary_period == 'quincenal')} QUINCENAL  "
        f"{_chk(salary_period == 'catorcenal')} CATORCENAL  {_chk(salary_period == 'semanal')} SEMANAL  "
        f"{_chk(salary_period == 'diario')} DIARIO",
        "",
        f"FORMA DE PAGO:  "
        f"{_chk(payment_period == 'mensual')} MENSUAL  {_chk(payment_period == 'quincenal')} QUINCENAL  "
        f"{_chk(payment_period == 'catorcenal')} CATORCENAL  {_chk(payment_period == 'semanal')} SEMANAL  "
        f"{_chk(payment_period == 'diario')} DIARIO",
        "",
        f"LUGAR DE PAGO: {_chk(payment_place == 'emplazamiento')} Lugar del emplazamiento  "
        f"{_chk(payment_place == 'trabajo')} Lugar de Trabajo  "
        f"{_chk(payment_place == 'banco')} Deposito en Banco {_safe(emp.get('bank_name'), _underline(28))}",
        "",
        f"SALARIO POR: {_chk(salary_type == 'comision')} 1.Comision  {_chk(salary_type == 'obra')} 2.Obra  "
        f"{_chk(salary_type == 'mixto')} 3.Mixto  {_chk(salary_type == 'destajo')} 4.A Destajo  "
        f"{_chk(salary_type == 'tarea')} 5.Tarea",
        f"            {_chk(salary_type == 'domicilio')} 6.Domicilio  {_chk(salary_type == 'otro')} 7.Otro",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_BODY, after=1)

    _add_section_title(doc, "PARA ESTOS SALARIOS DETALLARLOS:", size_pt=10.6, after=2)

    details = [
        "1. Habiendo devengado en los seis meses anteriores a la fecha de la ultima liquidacion que fue el dia",
        f"Mes {_underline(10)} 20___ la cantidad de: $ {_underline(22)} laborando en dicho periodo dias.",
        "CASO SALARIOS ADEUDADOS POR COMISION: Copia de liquidacion Art 126 d) C. de T.",
        "2. Habiendo devengado en los seis dias anteriores a la fecha de la ultima entrega o recuento respectivo que fue el",
        f"dia___ Mes {_underline(10)} 20___ la cantidad de $ {_underline(22)} laborando dias/horas.",
        "3. Habiendo devengado en los seis dias anteriores a la fecha de la ultima entrega o recuento respectivo que fue el dia",
        f"___ Mes {_underline(10)} 20___ la cantidad de $ {_underline(22)} laborando horas.",
        "4. Habiendo (devengado en la ultima entrega/pactado) que fue el dia ___ Mes 20___",
        f"$ {_underline(14)} (finalizando la obra/devolviendo el producto) el dia ___ mes ___ 20___ laborando horas.",
    ]

    for line in details:
        _add_line(doc, line, size_pt=9.75, after=0)

    _add_line(
        doc,
        f"{' ' * 52}{_chk()} SI     {_chk()} NO",
        size_pt=9.75,
        after=2,
    )


def _build_facts_page(doc, emp, fola):
    _add_section_title(doc, "RELACION DE HECHOS:", after=4)

    resignation = fola.get("voluntary_resignation_claim", {}) or {}

    lines = [
        f"DESPIDO: DIA {_safe(emp.get('dismissal_day'), '__________')}  "
        f"MES {_safe(emp.get('dismissal_month'), '__________')}  "
        f"20{_safe(emp.get('dismissal_year'), '____')}  "
        f"HORA: {_safe(emp.get('dismissal_time_text'), '__________')}  Persona que efectuo el despido:",
        f"{_safe(emp.get('person_who_dismissed_name'), _underline(62))}  Cargo {_safe(emp.get('person_who_dismissed_position'), _underline(22))}",
        "quien tiene facultades para contratar, despedir, dirigir y administrar.",
        "Le manifesto que a partir de ese momento estaba despedido(a) de su trabajo.",
        f"Nombre de la persona que impide el ingreso: {_safe(fola.get('person_who_prevented_entry'), _underline(58))}",
        f"HECHO QUE OCURRIO EN: {_chk(emp.get('dismissal_at_notification_place'))} Lugar del emplazamiento.  "
        f"{_chk(emp.get('dismissal_other_place'))} Otro: {_safe(emp.get('dismissal_place'), _underline(52))}",
        "",
        f"{_chk(resignation.get('enabled'))} Reclamo por incumplimiento a Ley Reguladora de la Prestacion Economica por Renuncia Voluntaria.",
        "",
        f"PRESENTO RENUNCIA EL DIA {_safe(resignation.get('day'), '__________')}  "
        f"MES {_safe(resignation.get('month'), '__________')}  "
        f"20{_safe(resignation.get('year'), '____')}  HORA: {_safe(resignation.get('hour'), '__________')}",
        f"Lugar {_safe(resignation.get('place'), _underline(36))}  Efectiva a partir del Dia: {_safe(resignation.get('effective_day'), '__________')}",
        f"MES {_safe(resignation.get('effective_month'), '__________')}  20{_safe(resignation.get('effective_year'), '____')}  "
        f"y habiendo transcurrido el plazo del Art.8 de la referida Ley, sin pago,",
        "se presume el despido injusto a partir de: DIA ___, MES ___, AÑO ___.",
        f"(Detalle) DIA {_safe(resignation.get('presumed_dismissal_day'), '__________')}  "
        f"MES {_safe(resignation.get('presumed_dismissal_month'), '__________')}  "
        f"AÑO {_safe(resignation.get('presumed_dismissal_year'), '______')}",
        f"Nombre de persona a quien le presento la renuncia {_safe(resignation.get('person'), _underline(48))}  "
        f"cargo {_safe(resignation.get('position'), _underline(20))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_SMALL if "facultades" in line or "manifesto" in line or "despedido" in line else PT_BODY, after=1)


def _build_no_effect_dismissal(doc, fola):
    _add_section_title(doc, "DESPIDO QUE NO SURTE SUS EFECTOS LEGALES POR:", after=2)

    ineffective = fola.get("ineffective_dismissal", {}) or {}
    motives = fola.get("dismissal_motives", {}) or {}
    other = fola.get("other_claim_facts", {}) or {}

    lines = [
        f"{_chk(ineffective.get('pregnancy'))} Encontrarse en estado de embarazo, tal como lo comprueba con medica que adjunta",
        f"a la presente, fecha probable de parto el dia {_safe(ineffective.get('probable_birth_day'), '__________')}  "
        f"mes {_safe(ineffective.get('probable_birth_month'), '________________')}  "
        f"20{_safe(ineffective.get('probable_birth_year'), '____')}.",
        f"{_chk(ineffective.get('union_board_member'))} Ser miembro de la Junta Directiva del Sindicato de "
        f"{_safe(ineffective.get('union_name'), _underline(52))}",
        f"{_chk(other.get('terminacion_contrato'))} TERMINACION DEL CONTRATO Art.53 C.T.  "
        f"{_chk(other.get('otro'))} OTRO: {_safe(other.get('otro_detalle'), _underline(18))}  "
        f"{_chk(other.get('riesgo_profesional'))} RIESGO PROFESIONAL",
        f"Cargo: {_safe(ineffective.get('union_position'), _underline(36))}  "
        f"lo comprueba con la certificacion que adjunta a la presente.",
        f"MOTIVOS/DESPIDO: {_chk(motives.get('embarazo'))} EMBARAZO  {_chk(motives.get('sindicalista'))} SINDICALISTA  "
        f"{_chk(motives.get('vih_sida'))} VIH/SIDA  {_chk(motives.get('acoso_sexual'))} ACOSO SEXUAL  "
        f"{_chk(motives.get('otro'))} OTRO: {_safe(motives.get('otro_detalle'), _underline(28))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_SMALLER, after=1)


def _build_other_facts(doc, fola):
    _add_section_title(doc, "OTROS HECHOS:", after=2)

    other = fola.get("other_claim_facts", {}) or {}

    lines = [
        f"{_chk(other.get('despido_indirecto'))} DESPIDO INDIRECTO. Art.55 Inc.3 o 56 del C.T.  "
        f"{_chk(other.get('terminacion_contrato'))} TERMINACION DEL CTR. Art.53 C.T.",
        f"{_chk(other.get('riesgo_profesional'))} RIESGO PROFESIONAL  "
        f"{_chk(other.get('otro'))} OTRO: {_safe(other.get('otro_detalle'), _underline(36))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_SMALLER, after=1)


def _build_claims_table(doc, fola, law_suggestions):
    _add_section_title(
        doc,
        "PIDE: SE PRESENTE DEMANDA EN CONTRA DE SU EMPLEADOR/A PARA RECLAMARLE:",
        size_pt=10.6,
        after=2,
    )

    table = doc.add_table(rows=11, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_widths(table, [9.595, 9.595])
    _format_table(table)

    left = [
        "Indemnizacion por despido injusto.\n(Art.38 Ord. 11 Cn y 58 C. de T)",
        "Indemnizacion por despido, vacacion y aguinaldo\nproporcional por renuncia voluntaria (Arts.3, 8, 9 y 15).",
        "Salarios no devengados por causa imputable al patrono/a\nhasta que concluya descanso post natal o garantia sindical.",
        "Vacacion y Aguinaldo Proporcional.\n(187-202 C. de T)",
        "Vacacion completa dia___ mes___ 20___.\n(Art.177 C. de T)",
        "Aguinaldo completo: 12 diciembre 2___\nal 11 diciembre 2___",
        "Salarios adeudados por dias\nlaborados y no remunerados.",
        "Horas extraordinarias laboradas\ny no remuneradas.",
        "Dias de descanso semanal laborado\ny no remunerado.",
        "Dias de asueto laborados\ny no remunerados.",
        "",
    ]

    right = [
        "Casos subsidios, servicios medicos, aparatos medicos,\ngastos de traslados por enfermedad/accidente comun.",
        "Indemnizacion por muerte\ndel/la trabajadora.",
        "Indemnizacion por incapacidad\npermanente del/la trabajadora.",
        "Indemnizacion por incapacidades\npermanentes parciales.",
        "Indemnizacion por lesiones\ndesfigurativas.",
        "Indemnizaciones a favor del/la conyuge\no companero/a de vida.",
        "Suspension por actividades\nde representacion gremial.",
        "Suspension del contrato con\nresponsabilidad patronal.",
        "Suspension del contrato sin\nresponsabilidad patronal.",
        "Reduccion de jornada por caso\nfortuito/fuerza mayor.",
        f"Otros reclamos: {_safe(law_suggestions, '______________________')}",
    ]

    for i in range(11):
        _cell_text(table.cell(i, 0), left[i], size_pt=PT_TABLE_CELL)
        _cell_text(table.cell(i, 1), right[i], size_pt=PT_TABLE_CELL)


def _build_documents_complement(doc, fola):
    lines = [
        f"Documentos que presenta: {_safe(fola.get('documents_presented'), _underline(72))}",
        f"Documentos que ofrece: {_safe(fola.get('documents_offered'), _underline(74))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=PT_BODY, after=1)

    _add_section_title(doc, "COMPLEMENTO:", size_pt=10.6, after=2)

    comp = fola.get("complement", {}) or {}

    line_a = (
        f"{_chk(comp.get('sustitucion_patronal'))} 1. SUSTITUCION PATRONAL  "
        f"{_chk(comp.get('horario'))} 2. HORARIO  "
        f"{_chk(comp.get('salario'))} 3. SALARIO  "
        f"{_chk(comp.get('lugar_trabajo'))} 4. LUGAR DE TRABAJO"
    )
    line_b = (
        f"{_chk(comp.get('hechos'))} 5. HECHOS  "
        f"{_chk(comp.get('reclamos'))} 6. RECLAMOS  "
        f"{_chk(comp.get('otro_tipo_hechos'))} 7. OTRO TIPO DE HECHOS"
    )
    _add_line(doc, line_a, size_pt=8.8, after=0)
    _add_line(doc, line_b, size_pt=8.8, after=2)

    det = _safe(comp.get("detalle"), "")
    if det:
        _add_line(doc, det, size_pt=PT_BODY, after=1)
    for _ in range(7):
        _add_full_width_horizontal_rule(doc, size_pt=PT_BODY, after=1)


def _build_final_page(doc, fola):
    _add_section_title(
        doc,
        "CONTINUACION DE COMPLEMENTO / OBSERVACIONES:",
        size_pt=PT_CONTINUATION_TITLE,
        after=2,
    )

    additional_notes = _normalize_text(fola.get("additional_notes"))
    note_lines = additional_notes.split("\n") if additional_notes else []
    max_obs = 21
    for i in range(max_obs):
        if i < len(note_lines) and note_lines[i].strip():
            _add_line(doc, note_lines[i].strip(), size_pt=PT_BODY, after=0)
        else:
            _add_full_width_horizontal_rule(doc, size_pt=PT_BODY, after=0)

    _add_line(doc, "", size_pt=PT_BODY, after=4)

    notice = fola.get("legal_effects_notice", {}) or {}

    _add_line(
        doc,
        "SE HACE CONSTAR QUE SE LE INFORMO Y EXPLICO A LA TRABAJADOR/A LOS EFECTOS LEGALES DE",
        size_pt=PT_CONSTANCIA,
        bold=True,
        after=1,
    )
    _add_line(doc, "PRESENTARSE A LA FECHA:", size_pt=PT_CONSTANCIA, bold=True, after=2)
    _add_line(
        doc,
        f"{_chk(notice.get('accion_prescrita'))} CON ACCION PRESCRITA.",
        size_pt=PT_BODY_COMPACT,
        after=1,
    )
    _add_line(
        doc,
        f"{_chk(notice.get('sin_presuncion_art_414'))} SIN QUE OPEREN PRESUNCIONES DEL ART.414 DEL CODIGO DE TRABAJO.",
        size_pt=PT_BODY_COMPACT,
        after=3,
    )

    _add_line(doc, "PARA CONSTANCIA FIRMA:", size_pt=PT_FIRMA_BLOCK, bold=True, after=6)
    _add_line(doc, "________________________________________", size_pt=PT_BODY, center=True, after=0)
    _add_line(doc, "Firma o huella de la o el trabajador", size_pt=PT_BODY, bold=True, center=True, after=6)

    paragraphs = [
        "COMO USUARIO/A DE ESTA UNIDAD SE ME HA EXPLICADO: LA DURACION APROXIMADA, ETAPAS DEL PROCESO",
        "JUDICIAL; LA PRUEBA QUE DEBO PRESENTAR; LA EXISTENCIA DEL PROCESO DE QUEJAS, RECLAMACIONES Y",
        "SUGERENCIAS, AL QUE PUEDO OPTAR EN EL CASO DE MI INCONFORMIDAD CON EL SERVICIO Y MIS DERECHOS",
        "COMO USUARIO/A DEL SERVICIO, COMPROMETIENDOME A MANTENER ACTUALIZADA LA INFORMACION;",
        "PROPORCIONAR UNA DIRECCION ACCESIBLE PARA LAS NOTIFICACIONES, ASISTIR A LAS CITAS EN LA HORA Y DIA",
        "INDICADOS, PRESENTAR LA PRUEBA REQUERIDA Y TRATAR CON RESPETO Y DIGNIDAD AL PERSONAL DE LA UNIDAD.",
    ]
    for ptxt in paragraphs:
        _add_line(doc, ptxt, size_pt=PT_USER_DECL, after=0)

    _add_line(doc, "", size_pt=PT_USER_DECL, after=2)
    _add_line(
        doc,
        "PARA CONSTANCIA FIRMAMOS: (deja impresa su huella dactilar)",
        size_pt=9.6,
        bold=True,
        after=5,
    )

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_widths(table, [9.595, 9.595])
    _format_table(table)

    _cell_text(
        table.cell(0, 0),
        "______________________________\nFirma o Huella del Usuario/a",
        center=True,
        size_pt=9.8,
        bold=True,
    )
    _cell_text(
        table.cell(0, 1),
        "______________________________\nNombre y Firma de Defensor/a Publico/a Laboral",
        center=True,
        size_pt=9.2,
        bold=True,
    )

    _add_line(doc, "", size_pt=PT_LEGAL_FOOTER, after=4)
    _add_line(
        doc,
        '"El presente formato difiere del generado por el Sistema de Informacion Gerencial, '
        'ya que este ultimo contiene exclusivamente la informacion del caso en concreto"',
        size_pt=PT_LEGAL_FOOTER,
        center=True,
    )


# =========================
# DOCUMENTO
# =========================

def _configure_fola03_page(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21.59)
    section.page_height = Cm(27.94)
    m = Cm(MARGIN_CM)
    section.top_margin = m
    section.bottom_margin = m
    section.left_margin = m
    section.right_margin = m
    _set_default_font(doc)


def _fill_fola03_document(doc: Document, analysis: dict) -> None:
    worker_information = analysis.get("worker_information", {}) or {}
    employment_relationship_data = analysis.get("employment_relationship_data", {}) or {}
    fola03_information = analysis.get("fola03_information", {}) or {}
    law_suggestions = _safe(analysis.get("law_suggestions"))

    _build_header(doc, first_page=True)
    _build_case_lines(doc, fola03_information)
    _build_user_data(doc, worker_information, fola03_information)
    _build_employer_data(doc, employment_relationship_data)
    _build_work_relation(doc, employment_relationship_data)

    doc.add_page_break()

    _build_header(doc, first_page=False)
    _build_salary_page(doc, employment_relationship_data)
    _build_facts_page(doc, employment_relationship_data, fola03_information)

    doc.add_page_break()

    _build_header(doc, first_page=False)
    _build_no_effect_dismissal(doc, fola03_information)
    _build_other_facts(doc, fola03_information)
    _build_claims_table(doc, fola03_information, law_suggestions)
    _build_documents_complement(doc, fola03_information)

    doc.add_page_break()

    _build_header(doc, first_page=False)
    _build_final_page(doc, fola03_information)


def write_fola03_docx(analysis: dict | str, output_path: str | Path) -> Path:
    """
    Genera el FOLA03 en disco (sin ADK). ``analysis`` puede ser un dict o un JSON string.
    """
    if isinstance(analysis, str):
        analysis = json.loads(analysis)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    _configure_fola03_page(doc)
    _fill_fola03_document(doc, analysis)
    doc.save(str(out))
    return out.resolve()


async def fola03_document_maker(
    analysis_json: str,
    tool_context: CallbackContext,
) -> dict:
    try:
        analysis = json.loads(analysis_json) if isinstance(analysis_json, str) else analysis_json

        worker_information = analysis.get("worker_information", {}) or {}

        doc = Document()
        _configure_fola03_page(doc)
        _fill_fola03_document(doc, analysis)

        buffer = BytesIO()
        doc.save(buffer)
        docx_bytes = buffer.getvalue()
        buffer.seek(0)

        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        worker_dui = _safe(worker_information.get("worker_dui"), "sin_dui")
        safe_dui = "".join(ch for ch in str(worker_dui) if ch.isalnum()) or "sin_dui"

        case_id = uuid.uuid4().hex
        output_filename = f"{case_id}_{safe_dui}_fola03_solicitud_asistencia_legal.docx"
        artifact_filename = _build_artifact_filename(case_id, safe_dui, "fola03")

        artifact_part = types.Part(
            inline_data=types.Blob(data=docx_bytes, mime_type=mime_type)
        )

        version = await tool_context.save_artifact(
            filename=artifact_filename,
            artifact=artifact_part
        )

        return {
            "status": "ok",
            "message": f"El documento {output_filename} versión {version} ha sido creado.",
            "artifact_filename": artifact_filename,
        }

    except Exception as e:
        return {
            "status": "error",
            "error_type": type(e).__name__,
            "error_message": str(e),
            "traceback": traceback.format_exc()
        }


def _default_docx_output_path() -> Path:
    return Path(__file__).resolve().parent / "output" / "fola03_generado.docx"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera el FOLA03 en Word (.docx) a partir de un JSON de análisis.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=_default_docx_output_path(),
        help="Ruta del .docx de salida (por defecto: output/fola03_generado.docx).",
    )
    parser.add_argument(
        "--json",
        "-j",
        type=Path,
        default=None,
        help="Archivo JSON con worker_information, employment_relationship_data, fola03_information, etc.",
    )
    args = parser.parse_args()
    if args.json is not None:
        analysis = json.loads(args.json.read_text(encoding="utf-8"))
    else:
        analysis = {}
    path = write_fola03_docx(analysis, args.output)
    print(f"Documento guardado en: {path}")


if __name__ == "__main__":
    main()
