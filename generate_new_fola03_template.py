import json
import os
import uuid
import traceback
from io import BytesIO
from datetime import datetime

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from google.cloud import storage
from google.oauth2 import service_account
from google.genai import types
from google.adk.agents.callback_context import CallbackContext


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

def _apply_font(run, font_name="Arial", size_pt=None, color_rgb=(0, 0, 0), bold=False, italic=False):
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


def _set_default_font(doc: Document, font_name="Arial", size_pt=8.2):
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


def _safe(value, default=""):
    value = _normalize_text(value)
    return value if value else default


def _chk(value=False):
    return "☒" if bool(value) else "☐"


def _underline(length=40):
    return "_" * length


def _add_line(doc, text="", size_pt=8.0, bold=False, italic=False, center=False, justify=False, after=1):
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


def _add_section_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=3, after=1, line=1.0)
    r = p.add_run(text)
    _apply_font(r, size_pt=9.2, bold=True)


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


def _cell_text(cell, text="", bold=False, center=False, size_pt=6.8, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=0, after=0, line=1.0)

    r = p.add_run(text)
    _apply_font(r, size_pt=size_pt, bold=bold, italic=italic)


# =========================
# BLOQUES FOLA03
# =========================

def _build_header(doc):
    p_code = doc.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_spacing(p_code, before=0, after=0)
    r_code = p_code.add_run("FOLA03")
    _apply_font(r_code, size_pt=9, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(p, before=0, after=2)

    for line in [
        "PROCURADURÍA GENERAL DE LA REPÚBLICA",
        "UNIDAD DE DEFENSA DE LOS DERECHOS DEL TRABAJADOR",
        "SOLICITUD DE ASISTENCIA LEGAL PARA JUICIO DE TRABAJO",
    ]:
        r = p.add_run(line + "\n")
        _apply_font(r, size_pt=10, bold=True)


def _build_case_lines(doc, fola):
    _add_line(doc, f"Expediente {_safe(fola.get('expediente'), _underline(28))} 20____", size_pt=8.0)
    _add_line(
        doc,
        f"Procuraduría Auxiliar de {_safe(fola.get('procuraduria_auxiliar'), _underline(28))} "
        f"a las {_safe(fola.get('hora_atencion'), _underline(26))} horas",
        size_pt=8.0,
    )
    _add_line(
        doc,
        f"{_safe(fola.get('minutos_atencion'), _underline(16))} minutos del día "
        f"{_safe(fola.get('dia_atencion'), _underline(18))} de "
        f"{_safe(fola.get('mes_atencion'), _underline(18))} 20{_safe(fola.get('anio_atencion'), '____')}.",
        size_pt=8.0,
    )


def _build_user_data(doc, worker, fola):
    _add_section_title(doc, "DATOS DE USUARIO/A")

    gender = _safe(worker.get("worker_gender")).upper()

    lines = [
        f"Nombre {_safe(worker.get('worker_name'), _underline(92))}",
        f"Conocido/a por {_safe(fola.get('known_as'), _underline(84))}",
        f"Inscrito en ISSS como {_safe(fola.get('isss_registered_as'), _underline(78))}",
        f"De {_safe(worker.get('worker_age'), '____')} años de edad. Género {_chk(gender == 'F')} F  {_chk(gender == 'M')} M  "
        f"Estado Familiar {_safe(worker.get('worker_marital_status'), _underline(24))} "
        f"Profesión/Oficio: {_safe(worker.get('worker_profession_or_trade'), _underline(19))}",
        f"{_underline(20)} DUI {_safe(worker.get('worker_dui'), _underline(26))} "
        f"expedido el {_safe(fola.get('dui_issue_day'), '____')} de {_safe(fola.get('dui_issue_month'), '__________')} "
        f"de {_safe(fola.get('dui_issue_year'), '______')} en {_safe(fola.get('dui_issue_place'), _underline(16))}",
        f"{_underline(25)} Otro Documento de identidad (Extranjeros) {_safe(fola.get('other_document'), _underline(35))}",
        f"ISSS Trabajador/a {_safe(fola.get('worker_isss'), _underline(36))} ISSS Empleador/a {_safe(fola.get('employer_isss'), _underline(36))}",
        f"Nacionalidad {_safe(worker.get('worker_nationality'), _underline(30))} "
        f"Domicilio {_safe(worker.get('worker_address'), _underline(35))} "
        f"Departamento {_safe(fola.get('worker_department'), _underline(25))}",
        f"Notificaciones: {_safe(fola.get('notifications_address'), _underline(88))}",
        f"{_underline(18)} Teléfono Residencia {_safe(fola.get('home_phone'), _underline(28))} "
        f"Celular {_safe(worker.get('worker_phone'), _underline(28))}",
        f"{_underline(20)} Recomendada/a {_safe(fola.get('recommended_by'), _underline(68))}",
        f"{_underline(14)} Teléfono(s) {_safe(fola.get('recommended_phone'), _underline(18))} "
        f"Celular(es) {_safe(fola.get('recommended_cellphone'), _underline(32))} Residencia",
        f"{_underline(92)}",
        f"# De Personas que dependen económicamente de la o el trabajador: {_safe(fola.get('economic_dependents'), _underline(34))}",
        f"Día {_safe(fola.get('mintrab_day'), _underline(12))} Mes {_safe(fola.get('mintrab_month'), _underline(16))} "
        f"Año {_safe(fola.get('mintrab_year'), _underline(12))} que compareció al MINTRAB.",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.8)


def _build_employer_data(doc, emp):
    _add_section_title(doc, "DATOS DE LA O EL EMPLEADOR")

    employer_type = _safe(emp.get("employer_type")).lower()
    _add_line(
        doc,
        f"{_chk(employer_type == 'persona natural')} Persona Natural        "
        f"{_chk(employer_type in ['persona juridica', 'persona jurídica'])} Persona Jurídica",
        size_pt=8.0,
        center=True,
    )

    place_type = _safe(emp.get("notification_place_type")).lower()

    lines = [
        f"Nombre /Razón Social/ denominación {_safe(emp.get('company_defendant'), _underline(72))}",
        f"Domicilio {_safe(emp.get('company_address'), _underline(42))} Departamento {_safe(emp.get('company_department'), _underline(39))}",
        f"Representante Legal {_safe(emp.get('legal_representative_name'), _underline(77))}",
        f"Mayor de edad y domicilio de {_safe(emp.get('legal_representative_address'), _underline(40))} Departamento {_safe(emp.get('legal_representative_department'), _underline(30))}",
        f"Lugar del emplazamiento {_safe(emp.get('company_notification_address'), _underline(78))}",
        f"{_underline(92)}",
        f"{_chk(place_type == 'negocio')} Lugar donde habitualmente atiende sus negocios      "
        f"{_chk(place_type == 'residencia')} Lugar de Residencia      "
        f"{_chk(place_type == 'trabajo')} Lugar de Trabajo",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.8)


def _build_work_relation(doc, emp):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(p, before=4, after=1)

    r = p.add_run("RELACIÓN DE TRABAJO")
    _apply_font(r, size_pt=10, bold=True)

    r2 = p.add_run(" " * 35 + f"{_chk(emp.get('substitution_patronal'))} SUSTITUCIÓN PATRONAL")
    _apply_font(r2, size_pt=7.8)

    lines = [
        f"FECHA DE INGRESO DÍA {_safe(emp.get('employment_start_day'), '______')} "
        f"MES {_safe(emp.get('employment_start_month'), '______________')} "
        f"AÑO {_safe(emp.get('employment_start_year'), '______')} "
        f"CARGO: {_safe(emp.get('job_title'), _underline(32))}",
        f"Desarrollo sus labores en:  {_chk(emp.get('workplace_is_notification_place'))} Emplazamiento   "
        f"{_chk(emp.get('workplace_is_other'))} Otro: {_safe(emp.get('workplace'), _underline(57))}",
        f"{_underline(74)} Consistían sus labores: {_safe(emp.get('actual_functions'), _underline(22))}",
        f"{_underline(92)}",
        f"Jornada Ordinaria de Trabajo:  {_chk(emp.get('ordinary_workday', True))} SÍ "
        f"{_safe(emp.get('daily_hours'), '____')} Horas diarias        "
        f"{_chk(not emp.get('ordinary_workday', True))} NO  Generalmente laboraba",
        f"Horario {_safe(emp.get('work_schedule'), _underline(86))}",
        f"{_underline(92)}",
        f"{_underline(92)}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.8)


def _build_salary_page(doc, emp):
    _add_section_title(doc, "SALARIO:")

    salary_period = _safe(emp.get("salary_period")).lower()
    payment_period = _safe(emp.get("payment_period")).lower()
    payment_place = _safe(emp.get("payment_place")).lower()
    salary_type = _safe(emp.get("salary_type")).lower()

    lines = [
        f"UNIDAD TIEMPO (base global) ${_safe(emp.get('salary_amount'), '__________')}  "
        f"{_chk(salary_period == 'mensual')} MENSUAL  {_chk(salary_period == 'quincenal')} QUINCENAL  "
        f"{_chk(salary_period == 'catorcenal')} CATORCENAL  {_chk(salary_period == 'semanal')} SEMANAL  "
        f"{_chk(salary_period == 'diario')} DIARIO",
        "",
        f"FORMA DE PAGO:                         "
        f"{_chk(payment_period == 'mensual')} MENSUAL  {_chk(payment_period == 'quincenal')} QUINCENAL  "
        f"{_chk(payment_period == 'catorcenal')} CATORCENAL  {_chk(payment_period == 'semanal')} SEMANAL  "
        f"{_chk(payment_period == 'diario')} DIARIO",
        "",
        f"LUGAR DE PAGO: {_chk(payment_place == 'emplazamiento')} Lugar del emplazamiento  "
        f"{_chk(payment_place == 'trabajo')} Lugar de Trabajo  "
        f"{_chk(payment_place == 'banco')} Depósito en Banco {_safe(emp.get('bank_name'), _underline(28))}",
        "",
        f"SALARIO POR:  {_chk(salary_type == 'comision')} 1.Comisión      {_chk(salary_type == 'obra')} 2.Obra      "
        f"{_chk(salary_type == 'mixto')} 3.Mixto      {_chk(salary_type == 'destajo')} 4.A Destajo      "
        f"{_chk(salary_type == 'tarea')} 5.Tarea",
        f"              {_chk(salary_type == 'domicilio')} 6.Domicilio.    {_chk(salary_type == 'otro')} 7.Otro",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.8)

    _add_section_title(doc, "PARA ESTOS SALARIOS DETALLARLOS:")

    details = [
        f"1. Habiendo devengado en los seis meses anteriores a la fecha de la última liquidación que fue el día {_underline(12)}",
        f"Mes {_underline(14)} 20____ la cantidad de: $ {_underline(20)} Laborando en dicho periodo ______ días.",
        f"CASO SALARIOS ADEUDADOS POR COMISIÓN Copia de liquidación Art. 126 d) C. de T.       {_chk()} SÍ     {_chk()} NO",
        "",
        f"2. Habiendo devengado en los seis días anteriores a la última entrega o recuento respectivo que fue el",
        f"día ______ Mes __________ 20____ la cantidad de: $ {_underline(18)} laborando ______ días/horas.",
        "",
        f"3. Habiendo devengado en los seis días anteriores a la última entrega o recuento respectivo que fue el día",
        f"____ Mes __________ 20____ la cantidad de: $ {_underline(18)} laborando ______ horas.",
        "",
        f"4. Habiendo devengado en la última entrega/pactado que fue el día ____ Mes __________ 20____ la cantidad de",
        f"$ {_underline(18)} finalizando la obra / devolviendo el producto el día ______ mes __________",
        f"20____ laborando ______ horas.",
    ]

    for line in details:
        _add_line(doc, line, size_pt=7.4)


def _build_facts_page(doc, emp, fola):
    _add_section_title(doc, "RELACION DE HECHOS")

    resignation = fola.get("voluntary_resignation_claim", {}) or {}

    lines = [
        f"DESPIDO DÍA {_safe(emp.get('dismissal_day'), '__________')} "
        f"MES {_safe(emp.get('dismissal_month'), '__________')} "
        f"20{_safe(emp.get('dismissal_year'), '____')}    "
        f"HORA: {_safe(emp.get('dismissal_time_text'), '__________')} Persona que efectuó el despido:",
        f"{_safe(emp.get('person_who_dismissed_name'), _underline(72))} Cargo {_safe(emp.get('person_who_dismissed_position'), _underline(27))}",
        f"quien tiene facultades para contratar, despedir, dirigir y administrar. Le manifestó que a partir de ese momento estaba",
        f"despedido(a) de su trabajo. Nombre de la persona que impidió el ingreso: {_safe(fola.get('person_who_prevented_entry'), _underline(38))}",
        f"{_underline(44)} HECHO QUE OCURRIÓ EN:    {_chk(emp.get('dismissal_at_notification_place'))} Lugar del emplazamiento.",
        f"{_chk(emp.get('dismissal_other_place'))} Otro: {_safe(emp.get('dismissal_place'), _underline(86))}",
        f"{_underline(92)}",
        "",
        f"{_chk(resignation.get('enabled'))} Reclamo por incumplimiento a Ley Reguladora de la Prestación Económica por Renuncia Voluntaria.",
        "",
        f"PRESENTÓ RENUNCIA EL DÍA {_safe(resignation.get('day'), '__________')} "
        f"MES {_safe(resignation.get('month'), '__________')} "
        f"20{_safe(resignation.get('year'), '____')}    HORA: {_safe(resignation.get('hour'), '__________')}",
        f"Lugar {_safe(resignation.get('place'), _underline(76))} Efectiva a partir del Día: {_safe(resignation.get('effective_day'), '__________')}",
        f"MES {_safe(resignation.get('effective_month'), '__________')} 20{_safe(resignation.get('effective_year'), '____')} y habiendo transcurrido el plazo del Art.8 de la referida Ley, sin que se haya",
        f"realizado el pago y con base al Art.3 inciso 2° del mismo cuerpo legal, se presume el despido injusto a partir de",
        f"DÍA {_safe(resignation.get('presumed_dismissal_day'), '__________')} "
        f"MES {_safe(resignation.get('presumed_dismissal_month'), '__________')} "
        f"AÑO {_safe(resignation.get('presumed_dismissal_year'), '______')} Nombre de persona a quien le presentó la renuncia",
        f"{_safe(resignation.get('person'), _underline(70))} cargo {_safe(resignation.get('position'), _underline(25))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.4)


def _build_no_effect_dismissal(doc, fola):
    _add_section_title(doc, "DESPIDO QUE NO SURTE SUS EFECTOS LEGALES POR:")

    ineffective = fola.get("ineffective_dismissal", {}) or {}
    motives = fola.get("dismissal_motives", {}) or {}

    lines = [
        f"{_chk(ineffective.get('pregnancy'))} Encontrarse en estado de embarazo, tal como lo comprueba con "
        f"{_safe(ineffective.get('medical_constancy'), _underline(28))} médica que adjunta a la",
        f"presente, fecha probable de parto el día {_safe(ineffective.get('probable_birth_day'), '__________')} "
        f"mes {_safe(ineffective.get('probable_birth_month'), '______________________________')} "
        f"20{_safe(ineffective.get('probable_birth_year'), '____')}.",
        f"{_chk(ineffective.get('union_board_member'))} Ser miembro de la Junta Directiva del Sindicato de "
        f"{_safe(ineffective.get('union_name'), _underline(62))}",
        f"Cargo: {_safe(ineffective.get('union_position'), _underline(78))} lo comprueba con la certificación que adjunta a la presente.",
        f"MOTIVOS/DESPIDO:   {_chk(motives.get('embarazo'))} EMBARAZO   {_chk(motives.get('sindicalista'))} SINDICALISTA   "
        f"{_chk(motives.get('vih_sida'))} VIH/SIDA   {_chk(motives.get('acoso_sexual'))} ACOSO SEXUAL",
        f"{_chk(motives.get('otro'))} OTRO: {_safe(motives.get('otro_detalle'), _underline(33))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.4)


def _build_other_facts(doc, fola):
    _add_section_title(doc, "OTROS HECHOS:")

    other = fola.get("other_claim_facts", {}) or {}

    lines = [
        f"{_chk(other.get('despido_indirecto'))} DESPIDO INDIRECTO Art. 55 Inc.3 o 56 del C.T.                         "
        f"{_chk(other.get('terminacion_contrato'))} TERMINACIÓN DEL CTR C/RP. Art.53 C.T.",
        f"{_chk(other.get('riesgo_profesional'))} RIESGO PROFESIONAL                                                   "
        f"{_chk(other.get('otro'))} OTRO {_safe(other.get('otro_detalle'), _underline(38))}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.4)


def _build_claims_table(doc, fola, law_suggestions):
    _add_section_title(doc, "PIDE: SE PRESENTE DEMANDA EN CONTRA DE SU EMPLEADOR/A PARA RECLAMARLE:")

    selected_claims = fola.get("selected_claims", []) or []
    selected_text = " ".join(str(x).lower() for x in selected_claims)

    table = doc.add_table(rows=12, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_widths(table, [9.595, 9.595])
    _format_table(table)

    left = [
        "Indemnización por despido injusto. (Art. 58 Ord. 11 Cn y 58 C. de T.)",
        "Indemnización por despido, vacación y aguinaldo proporcional por incumplimiento a la Ley Reguladora de la Prestación Económica por Renuncia Voluntaria.",
        "Salarios no devengados por causa imputable al patrono desde el día ___ mes ___ 20___ hasta que concluya su descanso post natal.",
        "Prestaciones por maternidad desde seis semanas antes de la fecha probable de parto.",
        "Salarios no devengados por causa imputable al patrono desde el día ___ mes ___ 20___ hasta que concluya su año de garantía sindical.",
        "Vacación y Aguinaldo Proporcional. (187-202 C. de T.)",
        "Vacación completa de ____ mes ____ 20____.",
        "Aguinaldo completo de diciembre de ____ al 11 de diciembre de ____.",
        "Salarios adeudados por días laborados y no remunerados.",
        "Horas extraordinarias laboradas y no remuneradas.",
        "Días de descanso semanal laborados y no remunerados.",
        "Días de asueto laborados y no remunerados.",
    ]

    right = [
        "Casos subsidios, servicios médicos, aparatos médicos, gastos de traslado por enfermedad / accidente común.",
        "Indemnización por muerte del/la trabajador/a.",
        "Indemnización por incapacidad permanente del/la trabajador/a.",
        "Indemnización por incapacidades permanentes parciales.",
        "Indemnización por lesiones desfigurativas.",
        "Indemnizaciones a favor del/la cónyuge o compañero/a de vida.",
        "Suspensión por actividades de representación gremial.",
        "Suspensión del contrato con responsabilidad patronal.",
        "Suspensión del contrato sin responsabilidad patronal.",
        "Reducción de jornada por caso fortuito o fuerza mayor.",
        "Otros reclamos ________________________________",
        _safe(law_suggestions, "______________________________________________"),
    ]

    def mark(text):
        t = text.lower()
        checked = any(word in selected_text for word in t.split()[:3])
        return f"{_chk(checked)} {text}"

    for i in range(12):
        _cell_text(table.cell(i, 0), mark(left[i]), size_pt=6.2)
        _cell_text(table.cell(i, 1), mark(right[i]), size_pt=6.2)


def _build_documents_complement(doc, fola):
    lines = [
        f"Documentos que presenta {_safe(fola.get('documents_presented'), _underline(82))}",
        f"{_underline(92)}",
        f"Documentos que ofrece {_safe(fola.get('documents_offered'), _underline(84))}",
        f"{_underline(92)}",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.6)

    _add_section_title(doc, "COMPLEMENTO:")

    comp = fola.get("complement", {}) or {}

    lines = [
        f"{_chk(comp.get('sustitucion_patronal'))} 1. SUSTITUCIÓN PATRONAL     "
        f"{_chk(comp.get('horario'))} 2. HORARIO     "
        f"{_chk(comp.get('salario'))} 3.SALARIO     "
        f"{_chk(comp.get('lugar_trabajo'))} 4. LUGAR DE TRABAJO",
        f"{_chk(comp.get('hechos'))} 5. HECHOS     "
        f"{_chk(comp.get('reclamos'))} 6.RECLAMOS.     "
        f"{_chk(comp.get('otro_tipo_hechos'))} 7. OTRO TIPO DE HECHOS",
        _safe(comp.get("detalle"), _underline(92)),
        _underline(92),
        _underline(92),
        _underline(92),
        _underline(92),
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.6)


def _build_final_page(doc, fola):
    additional_notes = _normalize_text(fola.get("additional_notes"))

    if additional_notes:
        for line in additional_notes.split("\n"):
            _add_line(doc, line, size_pt=7.6)
    else:
        for _ in range(17):
            _add_line(doc, _underline(92), size_pt=7.6)

    _add_line(doc, "", size_pt=7.6)

    notice = fola.get("legal_effects_notice", {}) or {}

    lines = [
        "SE HACE CONSTAR QUE SE LE INFORMÓ Y EXPLICÓ AL/LA TRABAJADOR/A LOS EFECTOS LEGALES DE",
        "PRESENTARSE A LA FECHA:",
        f"{_chk(notice.get('accion_prescrita'))} CON ACCIÓN PRESCRITA",
        f"{_chk(notice.get('sin_presuncion_art_414'))} SIN QUE OPEREN PRESUNCIONES DEL ART.414 DEL CÓDIGO DE TRABAJO.",
        "PARA CONSTANCIA FIRMA:",
        "",
        "________________________________________",
        "Firma o huella de la o el trabajador",
    ]

    for line in lines:
        _add_line(doc, line, size_pt=7.4, center=("____" in line or line.startswith("Firma")))

    _add_line(doc, "", size_pt=7.4)

    _add_line(
        doc,
        "COMO USUARIO/A DE ESTA UNIDAD SE ME HA EXPLICADO LA DURACIÓN APROXIMADA O ETAPAS DEL PROCESO "
        "JUDICIAL, LA PRUEBA QUE DEBO PRESENTAR, LA EXISTENCIA DEL PROCESO DE QUEJAS, RECLAMACIONES Y "
        "SUGERENCIAS, AL QUE PUEDO OPTAR EN EL CASO DE MI INCONFORMIDAD CON EL SERVICIO Y MIS DERECHOS "
        "COMO USUARIO/A DEL SERVICIO. COMPROMETIÉNDOME A MANTENER ACTUALIZADA LA INFORMACIÓN; "
        "PROPORCIONAR UNA DIRECCIÓN ACCESIBLE PARA LAS NOTIFICACIONES, ASISTIR A LAS CITAS EN LA HORA Y DÍA "
        "INDICADOS, PRESENTAR LA PRUEBA REQUERIDA Y TRATAR CON RESPETO Y DIGNIDAD AL PERSONAL DE LA UNIDAD. "
        "PARA CONSTANCIA FIRMAMOS: (letra impresa si huella dactilar)",
        size_pt=7.2,
        bold=True,
        italic=True,
        justify=True,
    )

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_widths(table, [9.595, 9.595])
    _format_table(table)

    _cell_text(table.cell(0, 0), "______________________________\nFirma o Huella del Usuario/a", center=True, size_pt=7.4)
    _cell_text(table.cell(0, 1), "______________________________\nNombre y Firma de Defensor/a Público/a Laboral", center=True, size_pt=7.4)

    _add_line(doc, "", size_pt=7.4)
    _add_line(
        doc,
        '"El presente formato difiere del generado por el Sistema de Información Gerencial, ya que este último contiene',
        size_pt=7.2,
        center=True,
    )
    _add_line(
        doc,
        'exclusivamente la información del caso en concreto"',
        size_pt=7.2,
        center=True,
    )


# =========================
# DOCUMENTO
# =========================

async def fola03_document_maker(
    analysis_json: str,
    tool_context: CallbackContext,
) -> dict:
    try:
        analysis = json.loads(analysis_json) if isinstance(analysis_json, str) else analysis_json

        worker_information = analysis.get("worker_information", {}) or {}
        employment_relationship_data = analysis.get("employment_relationship_data", {}) or {}
        fola03_information = analysis.get("fola03_information", {}) or {}
        law_suggestions = _safe(analysis.get("law_suggestions"))

        doc = Document()

        section = doc.sections[0]
        section.page_width = Cm(21.59)
        section.page_height = Cm(27.94)
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.2)
        section.left_margin = Cm(1.2)
        section.right_margin = Cm(1.2)

        _set_default_font(doc, font_name="Arial", size_pt=8.2)

        _build_header(doc)
        _build_case_lines(doc, fola03_information)
        _build_user_data(doc, worker_information, fola03_information)
        _build_employer_data(doc, employment_relationship_data)
        _build_work_relation(doc, employment_relationship_data)

        doc.add_page_break()

        _build_header(doc)
        _build_salary_page(doc, employment_relationship_data)
        _build_facts_page(doc, employment_relationship_data, fola03_information)

        doc.add_page_break()

        _build_header(doc)
        _build_no_effect_dismissal(doc, fola03_information)
        _build_other_facts(doc, fola03_information)
        _build_claims_table(doc, fola03_information, law_suggestions)
        _build_documents_complement(doc, fola03_information)

        doc.add_page_break()

        _build_header(doc)
        _build_final_page(doc, fola03_information)

        footer_para = doc.sections[0].footer.paragraphs[0]
        footer_para.text = f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if footer_para.runs:
            _apply_font(
                footer_para.runs[0],
                size_pt=8,
                color_rgb=(110, 110, 110)
            )

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
