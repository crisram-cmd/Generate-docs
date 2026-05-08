"""Generador independiente de la plantilla en blanco del FOLA-01 en PDF.

Este script produce un .pdf con la estructura visual del FOLA-01,
con campos vacios (lineas y placeholders), sin dependencias de GCS ni ADK.

Uso:
    python generate_fola01_template.py
    python generate_fola01_template.py --output C:/ruta/personalizada/plantilla.pdf
"""

from __future__ import annotations

import argparse
from pathlib import Path

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

PLACEHOLDER_LINE = "___________________"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def _draw_centered(
    c: canvas.Canvas, text: str, y: float, font: str = "Helvetica", size: int = 11
) -> None:
    c.setFont(font, size)
    width, _ = LETTER
    text_width = c.stringWidth(text, font, size)
    c.drawString((width - text_width) / 2, y, text)


def _draw_right(
    c: canvas.Canvas,
    text: str,
    y: float,
    right_margin: float,
    font: str = "Helvetica-Bold",
    size: int = 9,
) -> None:
    c.setFont(font, size)
    text_width = c.stringWidth(text, font, size)
    width, _ = LETTER
    c.drawString(width - right_margin - text_width, y, text)


def _draw_label(
    c: canvas.Canvas, text: str, x: float, y: float, font: str = "Times-Roman", size: int = 11
) -> float:
    c.setFont(font, size)
    c.drawString(x, y, text)
    return y - 16


def _draw_underline_line(c: canvas.Canvas, x: float, y: float, width: float) -> float:
    c.setLineWidth(0.55)
    c.line(x, y, x + width, y)
    return y - 16


def _draw_multiline_block(
    c: canvas.Canvas,
    label: str,
    x: float,
    y: float,
    width: float,
    lines: int,
    inline_first_line: bool = True,
    font: str = "Times-Roman",
    size: int = 11,
) -> float:
    c.setFont(font, size)
    c.drawString(x, y, label)
    if inline_first_line:
        # Primera linea en la misma fila del label (como formato de referencia)
        label_width = c.stringWidth(label, font, size)
        first_line_start = x + label_width + 6
        c.setLineWidth(0.55)
        c.line(first_line_start, y - 1, x + width, y - 1)
        y -= 16
        remaining_lines = max(lines - 1, 0)
    else:
        # Dejar salto de linea completo antes de iniciar lineas (caso "Requisitos...")
        y -= 16
        remaining_lines = max(lines, 0)

    for _ in range(remaining_lines):
        y = _draw_underline_line(c, x, y, width)
    return y - 4


def _draw_signature_block(c: canvas.Canvas, text: str, x_center: float, y: float) -> float:
    line_half = 118
    c.setLineWidth(0.7)
    c.line(x_center - line_half, y, x_center + line_half, y)

    c.setFont("Times-Bold", 10)
    text_width = c.stringWidth(text, "Times-Bold", 10)
    c.drawString(x_center - text_width / 2, y - 16, text)
    return y - 34


def _draw_inline_field(
    c: canvas.Canvas,
    x: float,
    y: float,
    label: str,
    field_width: float,
    trailing_text: str = "",
    font: str = "Times-Roman",
    size: int = 11,
) -> None:
    c.setFont(font, size)
    c.drawString(x, y, label)
    label_width = c.stringWidth(label, font, size)
    line_start = x + label_width + 2
    c.setLineWidth(0.55)
    c.line(line_start, y - 1, line_start + field_width, y - 1)
    if trailing_text:
        c.drawString(line_start + field_width + 4, y, trailing_text)


def build_blank_fola01_pdf(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    width, height = LETTER

    left_margin = 56
    right_margin = 56
    content_width = width - left_margin - right_margin

    y = height - 42

    _draw_right(c, "FOLA01", y, right_margin)
    y -= 22

    _draw_centered(c, "PROCURADURIA GENERAL DE LA REPUBLICA", y, font="Times-Bold", size=14)
    y -= 14
    _draw_centered(
        c,
        "UNIDAD DE DEFENSA DE LOS DERECHOS DEL TRABAJADOR",
        y,
        font="Times-Bold",
        size=12,
    )
    y -= 30

    _draw_centered(c, "REGISTRO DE ASESORIA INDIVIDUAL", y, font="Times-Bold", size=13)
    y -= 30

    _draw_inline_field(
        c,
        left_margin,
        y,
        "Procuraduria Auxiliar de",
        field_width=230,
        trailing_text="a las",
        size=10.8,
    )
    _draw_inline_field(
        c,
        left_margin + 365,
        y,
        "",
        field_width=94,
        trailing_text="horas",
        size=10.8,
    )
    y -= 16

    _draw_inline_field(
        c,
        left_margin + 88,
        y,
        "",
        field_width=78,
        trailing_text="minutos del dia",
        size=10.8,
    )
    _draw_inline_field(
        c,
        left_margin + 248,
        y,
        "",
        field_width=84,
        trailing_text="de",
        size=10.8,
    )
    _draw_inline_field(
        c,
        left_margin + 348,
        y,
        "",
        field_width=36,
        trailing_text=",20",
        size=10.8,
    )
    _draw_inline_field(
        c,
        left_margin + 408,
        y,
        "",
        field_width=74,
        trailing_text=".",
        size=10.8,
    )
    y -= 20

    y = _draw_label(c, "Nombre       del/la    usuario/a:", left_margin, y, size=10.8)
    y = _draw_underline_line(c, left_margin, y + 4, content_width)

    y = _draw_label(
        c,
        "Genero ________de______ anos de edad, con Documento Unico de Identidad Numero__________",
        left_margin,
        y,
        size=10.8,
    )

    y = _draw_label(
        c,
        "Extendido en _______________________ al dia __________________ mes _______________, 20____.",
        left_margin,
        y,
        size=10.8,
    )

    y = _draw_multiline_block(
        c, "Problema planteado:", left_margin, y, content_width, lines=3
    )
    y = _draw_multiline_block(
        c, "Asesoria para Juicio /Proceso:", left_margin, y, content_width, lines=2
    )
    y = _draw_multiline_block(c, "Documentos requeridos:", left_margin, y, content_width, lines=3)
    y = _draw_multiline_block(
        c,
        "Requisitos que faltan para conceder la asistencia legal:",
        left_margin,
        y,
        content_width,
        lines=5,
        inline_first_line=False,
    )

    y -= 10
    y = _draw_signature_block(c, "Firma o huella del usuario/a", width / 2, y)
    y -= 10
    _draw_signature_block(c, "Nombre y firma Defensor/a Publico Laboral", width / 2, y)

    legal_text = (
        '"*El presente formato difiere del generado por el Sistema de Informacion '
        "Gerencial, ya que este ultimo contiene exclusivamente la informacion "
        'del caso en concreto"'
    )
    _draw_centered(c, legal_text, 56, font="Times-Roman", size=8.5)

    c.showPage()
    c.save()
    return output_path


def _default_output_path() -> Path:
    return DEFAULT_OUTPUT_DIR / "fola01_template.pdf"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera la plantilla en blanco del FOLA-01 en PDF, sin datos.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=_default_output_path(),
        help=(
            "Ruta de salida del PDF. Por defecto: "
            "output/fola01_template.pdf dentro de esta carpeta."
        ),
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    final_path = build_blank_fola01_pdf(args.output)
    print(f"Plantilla FOLA-01 PDF generada en: {final_path}")


if __name__ == "__main__":
    main()
