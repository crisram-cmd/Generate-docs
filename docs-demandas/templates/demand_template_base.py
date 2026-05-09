"""Base compartida para plantillas PDF de demandas laborales (ReportLab canvas).

Salida por defecto: carpeta ``output`` en la raíz del proyecto Generate-docs.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph as RLParagraph

# Generate-docs/ (docs-demandas/templates -> .. -> .. -> ..)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
OUTPUT_DIR = _PROJECT_ROOT / "output"

PH = "____________________"


def bold_field_caps(s: str) -> str:
    """Marcador de campo en negrita (placeholder para datos del caso)."""
    return f"<b>{s.upper()}</b>"


def parse_args(default_filename: str, description: str) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=OUTPUT_DIR / default_filename,
        help=f"Ruta del PDF generado (por defecto: output/{default_filename})",
    )
    return parser.parse_args()


class DemandLayout:
    """Coordenadas en puntos PDF; página LETTER con márgenes ~3 cm lateral."""

    def __init__(self, c: canvas.Canvas) -> None:
        self.c = c
        self.width, self.height = LETTER
        self.left = 72
        self.right_edge = self.width - 72
        self.max_width = self.right_edge - self.left
        self.bottom_margin = 56
        self.y = self.height - 72

    def _ensure(self, points_needed: float) -> None:
        if self.y - points_needed < self.bottom_margin:
            self.c.showPage()
            self.y = self.height - 72

    def skip(self, pts: float = 12) -> None:
        self._ensure(pts)
        self.y -= pts

    def centered_bold(self, text: str, size: float = 12) -> None:
        font = "Helvetica-Bold"
        lead = size + 8
        self._ensure(lead)
        self.c.setFont(font, size)
        tw = self.c.stringWidth(text, font, size)
        self.c.drawString((self.width - tw) / 2, self.y, text)
        self.y -= lead

    def heading(self, text: str, size: float = 12) -> None:
        self.heading_left(text, size=size)

    def heading_left(self, text: str, size: float = 12) -> None:
        self.paragraph(text, font_bold=True, size=size)
        self.skip(6)

    def paragraph(self, text: str, font_bold: bool = False, size: float = 12) -> None:
        font = "Helvetica-Bold" if font_bold else "Helvetica"
        leading = size + 4
        lines = simpleSplit(text, font, size, self.max_width)
        for line in lines:
            self._ensure(leading)
            self.c.setFont(font, size)
            self.c.drawString(self.left, self.y, line)
            self.y -= leading

    def rich_paragraph(
        self,
        text: str,
        space_after: float = 12,
        alignment=TA_LEFT,
    ) -> None:
        style = ParagraphStyle(
            name="RichPara",
            fontName="Helvetica",
            fontSize=12,
            leading=16,
            alignment=alignment,
        )
        p = RLParagraph(text, style)
        w, h = p.wrap(self.max_width, 99999)
        self._ensure(h + space_after)
        p.drawOn(self.c, self.left, self.y - h)
        self.y -= h + space_after

    def bullet_lines(self, count: int = 6, size: float = 12) -> None:
        font = "Helvetica"
        leading = size + 6
        bullet = "\u2022 "
        self.c.setFont(font, size)
        bw = self.c.stringWidth(bullet, font, size)
        uw = self.c.stringWidth("_", font, size) or 5
        avail = self.max_width - bw - 4
        n_underscores = max(12, int(avail / uw))
        underscore_fill = "_" * n_underscores
        for _ in range(count):
            line = bullet + underscore_fill
            self._ensure(leading)
            self.c.setFont(font, size)
            self.c.drawString(self.left, self.y, line)
            self.y -= leading

    def blank_bullets(self, count: int = 6, size: float = 12) -> None:
        self.bullet_lines(count=count, size=size)


def draw_common_sections(layout: DemandLayout) -> None:
    layout.centered_bold("SEÑOR JUEZ DE LO LABORAL", size=12)
    layout.skip(10)
    layout.centered_bold("COMPARECENCIA DEL ABOGADO", size=12)
    layout.skip(10)

    intro = (
        f"{PH}, de {PH} años de edad, Abogado y del domicilio de {PH} con Documento "
        f"Único de Identidad número {PH}, Tarjeta de Identificación de la Abogacía {PH}, "
        "a usted atentamente EXPONGO:"
    )
    layout.paragraph(intro)
    layout.skip(16)

    layout.heading_left("PARTE EXPOSITIVA")
    parte_expositiva = (
        f"En mi calidad de Defensor Público Laboral, vengo a promover {PH}, en nombre y "
        f"representación del trabajador {PH}, de {PH} años de edad, {PH}, {PH}, de "
        f"nacionalidad {PH} y del domicilio de {PH}, con Documento Único de Identidad "
        f"número {PH} contra {PH}, y del domicilio de {PH}, representada legalmente por "
        f"{PH}, mayor de edad y del domicilio de {PH}, pudiendo ser citada, notificada y "
        "emplazada dicha persona jurídica por medio de su representante legal en "
        f"{PH}, para mejor ubicación agrego croquis, lugar donde habitualmente atiende "
        "sus negocios, para reclamarle prestaciones laborales."
    )
    layout.paragraph(parte_expositiva)
    layout.skip(12)

    layout.heading_left("RELACIÓN DE TRABAJO")
    relacion_trabajo = (
        f"Mi representado ingresó a laborar para y a las órdenes de {PH}, el {PH}, con el "
        f"cargo de {PH}; desarrollando sus labores en {PH} y las cuales consistían en "
        f"{PH}; estando sujeto a una jornada ordinaria de trabajo de {PH}, y un horario "
        f"de trabajo {PH}; devengando por sus servicios un salario de {PH}, los cuales "
        f"eran cancelados {PH}."
    )
    layout.paragraph(relacion_trabajo)
    layout.skip(12)


def draw_common_petitoria_and_footer(layout: DemandLayout) -> None:
    layout.heading_left("PARTE PETITORIA")
    layout.paragraph("Por lo antes expuesto, a usted respetuosamente PIDO:")
    layout.skip(12)

    petitoria_intro = (
        f"Me admita la presente demanda, me tenga por parte en la calidad en que comparezco, "
        f"cite a conciliación a {PH} por medio de su representante legal antes mencionado, y "
        "si no llegásemos a ningún acuerdo en dicha audiencia, previo los trámites legales y "
        "las pruebas que oportunamente aportaré, sea condenada en la sentencia definitiva a "
        "pagarle a mi representado:"
    )
    layout.paragraph(petitoria_intro)
    layout.skip(8)
    layout.bullet_lines(count=6)

    layout.skip(14)

    layout.paragraph(
        "Legítimo mi personería con copia certificada por notario de la Credencial Única y "
        "sus respectivas copias de ley."
    )
    layout.skip(12)
    layout.paragraph("Señalo para oír notificaciones, uddt.sansalvador@pgr.gob.sv")
    layout.skip(16)

    layout.paragraph(f"{PH} de {PH} de {PH}.")
    layout.skip(10)

