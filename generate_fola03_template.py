"""Generador independiente de plantilla FOLA-03 en PDF.

Genera una plantilla base multi-pagina inspirada en el formato oficial
FOLA-03, con campos en blanco, lineas de llenado y checkboxes dibujados.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def _draw_centered(c: canvas.Canvas, text: str, y: float, font: str = "Times-Bold", size: int = 11) -> None:
    width, _ = LETTER
    c.setFont(font, size)
    c.drawCentredString(width / 2, y, text)


def _draw_right(c: canvas.Canvas, text: str, y: float, right_margin: float = 54, font: str = "Times-Bold", size: int = 10) -> None:
    width, _ = LETTER
    c.setFont(font, size)
    text_width = c.stringWidth(text, font, size)
    c.drawString(width - right_margin - text_width, y, text)


def _draw_line(c: canvas.Canvas, x1: float, y: float, x2: float, thickness: float = 0.55) -> None:
    c.setLineWidth(thickness)
    c.line(x1, y, x2, y)


def _draw_label(c: canvas.Canvas, x: float, y: float, text: str, size: int = 10.3, bold: bool = False) -> None:
    c.setFont("Times-Bold" if bold else "Times-Roman", size)
    c.drawString(x, y, text)


def _draw_checkbox(c: canvas.Canvas, x: float, y: float, label: str, size: float = 8, font_size: int = 10) -> float:
    c.setLineWidth(0.7)
    # Alineacion vertical: centra visualmente el texto respecto al checkbox.
    box_y = y - (size * 0.72)
    c.rect(x, box_y, size, size, stroke=1, fill=0)
    box_center_y = box_y + (size / 2)
    text_y = box_center_y - (font_size * 0.31)
    _draw_label(c, x + size + 4, text_y, label, size=font_size)
    return x + size + 4 + c.stringWidth(label, "Times-Roman", font_size)


def _draw_fill_after_label(
    c: canvas.Canvas,
    x: float,
    y: float,
    label: str,
    x_end: float,
    size: int = 10.3,
    line_thickness: float = 0.55,
) -> None:
    _draw_label(c, x, y, label, size=size)
    start = x + c.stringWidth(label, "Times-Roman", size) + 5
    _draw_line(c, start, y - 1.5, x_end, thickness=line_thickness)


def _draw_header(c: canvas.Canvas, y_start: float = 752) -> float:
    y = y_start
    _draw_right(c, "FOLA03", y, font="Times-Bold", size=10.5)
    y -= 22
    _draw_centered(c, "PROCURADURIA GENERAL DE LA REPUBLICA", y, size=11)
    y -= 14
    _draw_centered(c, "UNIDAD DE DEFENSA DE LOS DERECHOS DEL TRABAJADOR", y, size=11)
    y -= 20
    return y


def _draw_page_number(c: canvas.Canvas, page_number: int, total_pages: int) -> None:
    width, _ = LETTER
    c.setFont("Times-Roman", 9)
    c.drawCentredString(width / 2, 18, f"-- {page_number} of {total_pages} --")


def _draw_page1(c: canvas.Canvas) -> None:
    left = 42
    right = 570
    y = _draw_header(c)

    _draw_centered(c, "SOLICITUD DE ASISTENCIA LEGAL PARA JUICIO DE TRABAJO", y, size=12.5)
    y -= 22

    # Bloque superior derecho (descomprimido, con mas aire vertical/horizontal).
    block_x = 280
    _draw_fill_after_label(c, block_x, y, "Expediente", 450, line_thickness=0.72)
    _draw_label(c, 456, y, "20___", size=10.3)
    y -= 22

    # Estas dos filas deben ir alineadas a la izquierda como el resto del formulario.
    _draw_fill_after_label(c, left, y, "Procuraduria Auxiliar de", 365, line_thickness=0.72)
    _draw_fill_after_label(c, 372, y, "a las", 500, line_thickness=0.72)
    _draw_label(c, 506, y, "horas", size=10.3)
    y -= 22

    # Tercera linea ajustada como referencia:
    
    _draw_line(c, left, y - 1.5, left + 88, thickness=0.72)
    _draw_fill_after_label(c, left + 94, y, "minutos del dia", 250, line_thickness=0.72)
    _draw_fill_after_label(c, 257, y, "de", 340, line_thickness=0.72)
    _draw_fill_after_label(c, 347, y, "20", 430, line_thickness=0.72)
    y -= 26

    _draw_label(c, left, y, "DATOS DE USUARIO/A", size=11, bold=True)
    y -= 14

    _draw_fill_after_label(c, left, y, "Nombre:", 290, line_thickness=0.72)
    _draw_fill_after_label(c, 296, y, "Conocido/a por", right, line_thickness=0.72)
    y -= 17

    _draw_fill_after_label(c, left, y, "De", 90)
    _draw_label(c, 96, y, "años de edad. Genero:", size=10.3)
    x = 210
    x = _draw_checkbox(c, x, y + 1, "F", size=8.2, font_size=10) + 10
    x = _draw_checkbox(c, x, y + 1, "M", size=8.2, font_size=10) + 10
    _draw_fill_after_label(c, x, y, "Estado Familiar:", 430, line_thickness=0.72)
    _draw_fill_after_label(c, 435, y, "Profesion/Oficio:", right, line_thickness=0.72)
    y -= 15

    _draw_fill_after_label(c, left, y, "DUI:", 145, line_thickness=0.72)
    _draw_fill_after_label(c, 150, y, "expedido el", 220, line_thickness=0.72)
    _draw_fill_after_label(c, 226, y, "de", 300, line_thickness=0.72)
    _draw_fill_after_label(c, 306, y, "de", 362, line_thickness=0.72)
    _draw_fill_after_label(c, 368, y, "en", right, line_thickness=0.72)
    y -= 17

    _draw_fill_after_label(c, left, y, "Otro Documento de identidad (Extranjeros):", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "Nacionalidad:", 180, line_thickness=0.72)
    _draw_fill_after_label(c, 185, y, "Domicilio:", 360, line_thickness=0.72)
    _draw_fill_after_label(c, 365, y, "Departamento:", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "Notificaciones:", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "Telefono Residencia:", 250, line_thickness=0.72)
    _draw_fill_after_label(c, 255, y, "Celular:", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "Recomendado/a:", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "Telefono(s):", 215, line_thickness=0.72)
    _draw_fill_after_label(c, 220, y, "Celular(es):", 390, line_thickness=0.72)
    _draw_fill_after_label(c, 395, y, "Residencia:", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "# De Personas que dependen economicamente:", right, line_thickness=0.72)
    y -= 17
    _draw_fill_after_label(c, left, y, "Dia", 90, line_thickness=0.72)
    _draw_fill_after_label(c, 95, y, "Mes", 150, line_thickness=0.72)
    _draw_fill_after_label(c, 155, y, "Año que comparecio al MINTRAB.", right, line_thickness=0.72)
    y -= 19

    _draw_label(c, left, y, "DATOS DE LA O EL EMPLEADOR", size=11, bold=True)
    x = left + 208
    x = _draw_checkbox(c, x, y + 1, "Persona Juridica", size=8.2, font_size=10) + 12
    _draw_checkbox(c, x, y + 1, "Persona Natural", size=8.2, font_size=10)
    y -= 15

    _draw_fill_after_label(c, left, y, "Nombre /Razon Social/ denominacion:", right, line_thickness=0.72)
    y -= 15
    _draw_fill_after_label(c, left, y, "Domicilio", 330, line_thickness=0.72)
    _draw_fill_after_label(c, 335, y, "Departamento", right, line_thickness=0.72)
    y -= 15
    _draw_fill_after_label(c, left, y, "Representante Legal:", right, line_thickness=0.72)
    y -= 15
    _draw_fill_after_label(c, left, y, "Mayor de edad y domicilio de", 352, line_thickness=0.72)
    _draw_fill_after_label(c, 357, y, "Departamento", right, line_thickness=0.72)
    y -= 15
    _draw_fill_after_label(c, left, y, "Lugar del emplazamiento:", right, line_thickness=0.72)
    y -= 15

    x = left
    x = _draw_checkbox(c, x, y + 1, "Lugar donde habitualmente atiende sus negocios", size=8.2, font_size=9.6) + 10
    x = _draw_checkbox(c, x, y + 1, "Lugar de Residencia", size=8.2, font_size=9.6) + 10
    _draw_checkbox(c, x, y + 1, "Lugar de Trabajo", size=8.2, font_size=9.6)
    y -= 18

    _draw_label(c, left, y, "RELACION DE TRABAJO", size=11, bold=True)
    _draw_checkbox(c, left + 162, y + 1, "SUSTITUCION PATRONAL", size=8.2, font_size=10)
    y -= 16
    _draw_fill_after_label(c, left, y, "FECHA DE INGRESO: DIA", 165, line_thickness=0.72)
    _draw_fill_after_label(c, 170, y, "MES", 225, line_thickness=0.72)
    _draw_fill_after_label(c, 230, y, "AÑO", 285, line_thickness=0.72)
    _draw_fill_after_label(c, 290, y, "CARGO:", right, line_thickness=0.72)
    y -= 15
    _draw_label(c, left, y, "Desarrollo sus labores en:", size=10.3)
    x = left + 145
    x = _draw_checkbox(c, x, y + 1, "Emplazamiento", size=8.2, font_size=10) + 12
    _draw_checkbox(c, x, y + 1, "Otro:", size=8.2, font_size=10)
    _draw_line(c, x + 36, y - 1.5, right)
    y -= 15
    _draw_fill_after_label(c, left, y, "Consistian sus labores:", right, line_thickness=0.72)
    y -= 15
    _draw_fill_after_label(c, left, y, "Jornada Ordinaria de Trabajo:", 242, line_thickness=0.72)
    x = _draw_checkbox(c, 248, y + 1, "SI", size=8.2, font_size=10) + 4
    _draw_line(c, x, y - 1.5, x + 62)
    _draw_label(c, x + 66, y, "Horas diarias", size=10.3)
    _draw_checkbox(c, 470, y + 1, "NO", size=8.2, font_size=10)
    y -= 15
    _draw_fill_after_label(c, left, y, "Horario:", right, line_thickness=0.72)

    # En el formato de referencia queda un bloque de lineas vacias al final de la pagina 1.
    y -= 10
    for _ in range(4):
        _draw_line(c, left, y, right, thickness=0.72)
        y -= 14


def _draw_page2(c: canvas.Canvas) -> None:
    left = 42
    right = 570
    y = _draw_header(c)

    _draw_label(c, left, y, "SALARIO:", size=11, bold=True)
    y -= 16
    _draw_fill_after_label(c, left, y, "UNIDAD TIEMPO (base global): $", 230)
    x = _draw_checkbox(c, 236, y - 0.8, "MENSUAL", size=8.2, font_size=9.8) + 10
    x = _draw_checkbox(c, x, y - 0.8, "QUINCENAL", size=8.2, font_size=9.8) + 10
    x = _draw_checkbox(c, x, y - 0.8, "CATORCENAL", size=8.2, font_size=9.8) + 10
    x = _draw_checkbox(c, x, y - 0.8, "SEMANAL", size=8.2, font_size=9.8) + 10
    _draw_checkbox(c, x, y - 0.8, "DIARIO", size=8.2, font_size=9.8)
    y -= 35

    _draw_label(c, left, y, "FORMA DE PAGO:", size=10.3)
    x = _draw_checkbox(c, 148, y + 0.6, "MENSUAL", size=8.2, font_size=9.8) + 10
    x = _draw_checkbox(c, x, y + 0.6, "QUINCENAL", size=8.2, font_size=9.8) + 10
    x = _draw_checkbox(c, x, y + 0.6, "CATORCENAL", size=8.2, font_size=9.8) + 10
    x = _draw_checkbox(c, x, y + 0.6, "SEMANAL", size=8.2, font_size=9.8) + 10
    _draw_checkbox(c, x, y + 0.6, "DIARIO", size=8.2, font_size=9.8)
    y -= 15

    _draw_label(c, left, y, "LUGAR DE PAGO:", size=10.3)
    x = _draw_checkbox(c, 124, y + 0.6, "Lugar del emplazamiento", size=8.2, font_size=9.6) + 12
    x = _draw_checkbox(c, x, y + 0.6, "Lugar de Trabajo", size=8.2, font_size=9.6) + 12
    x2 = _draw_checkbox(c, x, y + 0.6, "Deposito en Banco", size=8.2, font_size=9.6) + 8
    _draw_line(c, x2, y - 1.5, 475)
    y -= 15

    _draw_label(c, left, y, "SALARIO POR:", size=10.3)
    x = _draw_checkbox(c, 124, y + 0.6, "1.Comision", size=8.2, font_size=9.8) + 12
    x = _draw_checkbox(c, x, y + 0.6, "2.Obra", size=8.2, font_size=9.8) + 12
    x = _draw_checkbox(c, x, y + 0.6, "3.Mixto", size=8.2, font_size=9.8) + 12
    x = _draw_checkbox(c, x, y + 0.6, "4.A Destajo", size=8.2, font_size=9.8) + 12
    _draw_checkbox(c, x, y + 0.6, "5.Tarea", size=8.2, font_size=9.8)
    y -= 14
    x = _draw_checkbox(c, 124, y + 0.6, "6.Domicilio", size=8.2, font_size=9.8) + 12
    _draw_checkbox(c, x, y + 0.6, "7.Otro", size=8.2, font_size=9.8)
    y -= 17

    _draw_label(c, left, y, "PARA ESTOS SALARIOS DETALLARLOS:", size=10.6, bold=True)
    y -= 14
    for text in [
        "1. Habiendo devengado en los seis meses anteriores a la fecha de la ultima liquidacion que fue el dia",
        "Mes 20___ la cantidad de: $_______________________ laborando en dicho periodo dias.",
        "CASO SALARIOS ADEUDADOS POR COMISION: Copia de liquidacion Art 126 d) C. de T.",
        "2. Habiendo devengado en los seis dias anteriores a la fecha de la ultima entrega o recuento respectivo que fue el",
        "dia___ Mes 20___ la cantidad de $_______________________ laborando dias/horas.",
        "3. Habiendo devengado en los seis dias anteriores a la fecha de la ultima entrega o recuento respectivo que fue el dia",
        "___ Mes 20___ la cantidad de $_______________________ laborando horas.",
        "4. Habiendo (devengado en la ultima entrega/pactado) que fue el dia ___ Mes 20___",
        "$___________ (finalizando la obra/devolviendo el producto) el dia ___ mes ___ 20___ laborando horas.",
    ]:
        _draw_label(c, left, y, text, size=9.75)
        y -= 12
    _draw_checkbox(c, left + 420, y + 11.5, "SI", size=8.2, font_size=9.8)
    _draw_checkbox(c, left + 468, y + 11.5, "NO", size=8.2, font_size=9.8)

    y -= 4
    _draw_label(c, left, y, "RELACION DE HECHOS:", size=11, bold=True)
    y -= 14
    _draw_fill_after_label(c, left, y, "DESPIDO: DIA", 145)
    _draw_fill_after_label(c, 150, y, "MES", 205)
    _draw_fill_after_label(c, 210, y, "20", 255)
    _draw_fill_after_label(c, 260, y, "HORA:", 318)
    _draw_fill_after_label(c, 323, y, "Persona que efectuo el despido:", right)
    y -= 15
    _draw_fill_after_label(c, left, y, "Cargo", 120)
    _draw_label(c, 125, y, "quien tiene facultades para contratar, despedir, dirigir y administrar.", size=9.75)
    y -= 12
    _draw_label(c, left, y, "Le manifesto que a partir de ese momento estaba despedido(a) de su trabajo.", size=9.75)
    y -= 12
    _draw_fill_after_label(c, left, y, "Nombre de la persona que impide el ingreso:", right)
    y -= 15
    _draw_label(c, left, y, "HECHO QUE OCURRIO EN:", size=10.3)
    x = _draw_checkbox(c, 170, y - 0.8, "Lugar del emplazamiento", size=8.2, font_size=9.7) + 10
    _draw_checkbox(c, x, y - 0.8, "Otro:", size=8.2, font_size=9.7)
    _draw_line(c, x + 36, y - 1.5, right)
    y -= 15
    _draw_checkbox(
        c,
        left,
        y - 0.8,
        "Reclamo por incumplimiento a Ley Reguladora de la Prestacion Economica por Renuncia Voluntaria.",
        size=8.2,
        font_size=9.7,
    )
    y -= 15
    _draw_fill_after_label(c, left, y, "PRESENTO RENUNCIA EL DIA", 210)
    _draw_fill_after_label(c, 215, y, "Lugar", 310)
    _draw_fill_after_label(c, 315, y, "Efectiva a partir del Dia:", right)
    y -= 15
    _draw_fill_after_label(c, left, y, "MES", 90)
    _draw_fill_after_label(c, 95, y, "20", 140)
    _draw_label(c, 145, y, "y habiendo transcurrido el plazo del Art.8 de la referida Ley, sin pago,", size=9.4)
    y -= 12
    _draw_label(c, left, y, "se presume el despido injusto a partir de: DIA ___, MES ___, AÑO ___.", size=9.4)
    y -= 12
    _draw_fill_after_label(c, left, y, "Nombre de persona a quien le presento la renuncia", 410)
    _draw_fill_after_label(c, 415, y, "cargo", right)


def _draw_page3(c: canvas.Canvas) -> None:
    left = 42
    right = 570
    y = _draw_header(c)

    _draw_label(c, left, y, "DESPIDO QUE NO SURTE SUS EFECTOS LEGALES POR:", size=11, bold=True)
    y -= 15
    _draw_checkbox(c, left, y + 1, "Encontrarse en estado de embarazo, tal como lo comprueba con medica que adjunta", size=8.2, font_size=9.6)
    y -= 13
    _draw_fill_after_label(c, left + 18, y, "a la presente, fecha probable de parto el dia", 285)
    _draw_fill_after_label(c, 290, y, "mes", right)
    y -= 14
    _draw_checkbox(c, left, y + 1, "Ser miembro de la Junta Directiva del Sindicato de", size=8.2, font_size=9.6)
    _draw_line(c, 300, y - 1.5, right)
    y -= 14
    _draw_checkbox(c, left, y + 1, "TERMINACION DEL CONTRATO Art.53 C.T.", size=8.2, font_size=9.6)
    _draw_checkbox(c, left + 240, y + 1, "OTRO:", size=8.2, font_size=9.6)
    _draw_line(c, left + 288, y - 1.5, 430)
    _draw_checkbox(c, 438, y + 1, "RIESGO PROFESIONAL", size=8.2, font_size=9.6)
    y -= 14
    _draw_fill_after_label(c, left, y, "Cargo:", 230)
    _draw_label(c, 235, y, "lo comprueba con la certificacion que adjunta a la presente.", size=9.6)
    y -= 14
    _draw_label(c, left, y, "MOTIVOS/DESPIDO:", size=10.3)
    x = _draw_checkbox(c, 130, y + 1, "EMBARAZO", size=8.2, font_size=9.6) + 8
    x = _draw_checkbox(c, x, y + 1, "SINDICALISTA", size=8.2, font_size=9.6) + 8
    x = _draw_checkbox(c, x, y + 1, "VIH/SIDA", size=8.2, font_size=9.6) + 8
    x = _draw_checkbox(c, x, y + 1, "ACOSO SEXUAL", size=8.2, font_size=9.6) + 8
    _draw_checkbox(c, x, y + 1, "OTRO:", size=8.2, font_size=9.6)
    _draw_line(c, x + 52, y - 1.5, right)
    y -= 16

    _draw_label(c, left, y, "OTROS HECHOS:", size=11, bold=True)
    y -= 14
    # Formato 2x2 como solicitado.
    _draw_checkbox(c, left, y + 0.6, "DESPIDO INDIRECTO. Art.55 Inc.3 o 56 del C.T.", size=8.2, font_size=9.6)
    _draw_checkbox(c, left + 250, y + 0.6, "TERMINACION DEL CTR. Art.53 C.T.", size=8.2, font_size=9.6)
    y -= 14
    x_otros = _draw_checkbox(c, left, y + 0.6, "RIESGO PROFESIONAL", size=8.2, font_size=9.6) + 12
    _draw_checkbox(c, left + 250, y + 0.6, "OTRO:", size=8.2, font_size=9.6)
    _draw_line(c, left + 300, y - 1.5, right)
    y -= 16

    _draw_label(c, left, y, "PIDE: SE PRESENTE DEMANDA EN CONTRA DE SU EMPLEADOR/A PARA RECLAMARLE:", size=10.6, bold=True)
    y -= 12
    # Tabla de pretensiones (sin checkbox por fila), con celdas multilinea.
    table_left = left
    table_right = right
    table_top = y
    col_split = table_left + 265
    row_h = 26
    rows = 11
    table_bottom = table_top - (rows * row_h)

    _draw_line(c, table_left, table_top, table_right, thickness=0.7)
    _draw_line(c, table_left, table_bottom, table_right, thickness=0.7)
    _draw_line(c, table_left, table_top, table_left, thickness=0.7)
    _draw_line(c, table_right, table_top, table_right, thickness=0.7)
    _draw_line(c, col_split, table_top, col_split, thickness=0.7)

    for i in range(1, rows):
        row_y = table_top - (i * row_h)
        _draw_line(c, table_left, row_y, table_right, thickness=0.45)

    left_claims = [
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
    right_claims = [
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
        "Otros reclamos: ________________________",
    ]

    for i in range(rows):
        row_top = table_top - (i * row_h)
        left_lines = [line for line in left_claims[i].split("\n") if line] if left_claims[i] else []
        right_lines = [line for line in right_claims[i].split("\n") if line] if right_claims[i] else []

        # Render multilinea compacta dentro de celda.
        for j, line in enumerate(left_lines):
            _draw_label(c, table_left + 4, row_top - 10 - (j * 9.4), line, size=8.45)
        for j, line in enumerate(right_lines):
            _draw_label(c, col_split + 4, row_top - 10 - (j * 9.4), line, size=8.45)

    y = table_bottom - 10
    _draw_fill_after_label(c, left, y, "Documentos que presenta:", right)
    y -= 13
    _draw_fill_after_label(c, left, y, "Documentos que ofrece:", right)
    y -= 13
    _draw_label(c, left, y, "COMPLEMENTO:", size=10.6, bold=True)
    y -= 13
    x = left
    for item in [
        "1. SUSTITUCION PATRONAL",
        "2. HORARIO",
        "3. SALARIO",
        "4. LUGAR DE TRABAJO",
        "5. HECHOS",
        "6. RECLAMOS",
        "7. OTRO TIPO DE HECHOS",
    ]:
        x = _draw_checkbox(c, x, y + 1, item, size=8, font_size=8.8) + 8
        if x > 510:
            y -= 11
            x = left

    # Espacio de escritura adicional solicitado bajo COMPLEMENTO (7 lineas).
    y -= 10
    for _ in range(7):
        _draw_line(c, left, y, right, thickness=0.6)
        y -= 13


def _draw_page4(c: canvas.Canvas) -> None:
    left = 42
    right = 570
    y = _draw_header(c)

    # Continuacion de complemento en la misma pagina final (21 lineas),
    # para mantener el documento total en 4 paginas.
    _draw_label(c, left, y, "CONTINUACION DE COMPLEMENTO / OBSERVACIONES:", size=10.8, bold=True)
    y -= 12
    for _ in range(21):
        _draw_line(c, left, y, right, thickness=0.6)
        y -= 10.5

    y -= 10
    _draw_label(c, left, y, "SE HACE CONSTAR QUE SE LE INFORMO Y EXPLICO A LA TRABAJADOR/A LOS EFECTOS LEGALES DE", size=10.2, bold=True)
    y -= 12
    _draw_label(c, left, y, "PRESENTARSE A LA FECHA:", size=10.2, bold=True)
    y -= 16
    _draw_checkbox(c, left, y + 1, "CON ACCION PRESCRITA.", size=8.2, font_size=10)
    y -= 14
    _draw_checkbox(c, left, y + 1, "SIN QUE OPEREN PRESUNCIONES DEL ART.414 DEL CODIGO DE TRABAJO.", size=8.2, font_size=10)
    y -= 20

    _draw_label(c, left, y, "PARA CONSTANCIA FIRMA:", size=10.5, bold=True)
    y -= 44
    _draw_line(c, 170, y, 410, thickness=0.7)
    _draw_centered(c, "Firma o huella de la o el trabajador", y - 16, font="Times-Bold", size=10)
    y -= 44

    paragraphs = [
        "COMO USUARIO/A DE ESTA UNIDAD SE ME HA EXPLICADO: LA DURACION APROXIMADA, ETAPAS DEL PROCESO",
        "JUDICIAL; LA PRUEBA QUE DEBO PRESENTAR; LA EXISTENCIA DEL PROCESO DE QUEJAS, RECLAMACIONES Y",
        "SUGERENCIAS, AL QUE PUEDO OPTAR EN EL CASO DE MI INCONFORMIDAD CON EL SERVICIO Y MIS DERECHOS",
        "COMO USUARIO/A DEL SERVICIO, COMPROMETIENDOME A MANTENER ACTUALIZADA LA INFORMACION;",
        "PROPORCIONAR UNA DIRECCION ACCESIBLE PARA LAS NOTIFICACIONES, ASISTIR A LAS CITAS EN LA HORA Y DIA",
        "INDICADOS, PRESENTAR LA PRUEBA REQUERIDA Y TRATAR CON RESPETO Y DIGNIDAD AL PERSONAL DE LA UNIDAD.",
    ]
    for p in paragraphs:
        _draw_label(c, left, y, p, size=9.3)
        y -= 11.2

    y -= 8
    _draw_label(c, left, y, "PARA CONSTANCIA FIRMAMOS: (deja impresa su huella dactilar)", size=9.6, bold=True)
    y -= 36
    _draw_line(c, 85, y, 265, thickness=0.7)
    _draw_line(c, 345, y, 525, thickness=0.7)
    _draw_label(c, 100, y - 15, "Firma o Huella del Usuario/a", size=9.8, bold=True)
    _draw_label(c, 355, y - 15, "Nombre y Firma de Defensor/a Publico/a Laboral", size=9.2, bold=True)

    y -= 56
    legal_footer = (
        '"El presente formato difiere del generado por el Sistema de Informacion Gerencial, '
        'ya que este ultimo contiene exclusivamente la informacion del caso en concreto"'
    )
    _draw_centered(c, legal_footer, y, font="Times-Roman", size=8.6)


def build_blank_fola03_pdf(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)

    total_pages = 4

    _draw_page1(c)
    _draw_page_number(c, 1, total_pages)
    c.showPage()

    _draw_page2(c)
    _draw_page_number(c, 2, total_pages)
    c.showPage()

    _draw_page3(c)
    _draw_page_number(c, 3, total_pages)
    c.showPage()

    _draw_page4(c)
    _draw_page_number(c, 4, total_pages)
    c.showPage()

    c.save()
    return output_path


def _default_output_path() -> Path:
    return DEFAULT_OUTPUT_DIR / "fola03_template.pdf"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera la plantilla en blanco del FOLA-03 en PDF.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=_default_output_path(),
        help=(
            "Ruta de salida del PDF. Por defecto: "
            "output/fola03_template.pdf dentro de esta carpeta."
        ),
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    final_path = build_blank_fola03_pdf(args.output)
    print(f"Plantilla FOLA-03 PDF generada en: {final_path}")


if __name__ == "__main__":
    main()
