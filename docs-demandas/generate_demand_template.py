"""Generador independiente de plantilla PDF para demanda laboral (en blanco).

Molde unificado: encabezado y comparecencia, PARTE EXPOSITIVA, RELACIÓN DE TRABAJO,
RELACIÓN DE HECHOS, PARTE PETITORIA y cierre. Textos alineados a `complaint_maker`.
Sin GCS ni ADK.

Uso:
    python generate_demand_template.py
    python generate_demand_template.py --output C:/ruta/demanda_template.pdf
"""

from __future__ import annotations

import argparse
from pathlib import Path

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT_DIR = _PROJECT_ROOT / "output"

PH = "____________________"


class _DemandLayout:
    """Coordenadas en puntos PDF; página LETTER con márgenes similares al DOCX (~3 cm lateral)."""

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


def build_blank_demanda_pdf(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    L = _DemandLayout(c)

    # Encabezado y comparecencia del abogado
    L.centered_bold("SEÑOR JUEZ DE LO LABORAL", size=12)
    L.skip(10)
    L.centered_bold("COMPARECENCIA DEL ABOGADO", size=12)
    L.skip(10)

    # Identificación del abogado + EXPONGO
    intro = (
        f"{PH}, de {PH} años de edad, Abogado y del domicilio de {PH} con Documento "
        f"Único de Identidad número {PH}, Tarjeta de Identificación de la Abogacía {PH}, "
        "a usted atentamente EXPONGO:"
    )
    L.paragraph(intro)
    L.skip(16)

    # PARTE EXPOSITIVA
    L.heading_left("PARTE EXPOSITIVA")
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
    L.paragraph(parte_expositiva)
    L.skip(12)

    # RELACIÓN DE TRABAJO
    L.heading_left("RELACIÓN DE TRABAJO")
    relacion_trabajo = (
        f"Mi representado ingresó a laborar para y a las órdenes de {PH}, el {PH}, con el "
        f"cargo de {PH}; desarrollando sus labores en {PH} y las cuales consistían en "
        f"{PH}; estando sujeto a una jornada ordinaria de trabajo de {PH}, y un horario "
        f"de trabajo {PH}; devengando por sus servicios un salario de {PH}, los cuales "
        f"eran cancelados {PH}."
    )
    L.paragraph(relacion_trabajo)
    L.skip(12)

    # RELACIÓN DE HECHOS
    L.heading_left("RELACIÓN DE HECHOS")
    relacion_hechos = (
        f"En las condiciones de trabajo antes mencionadas laboró mi patrocinado para y a "
        f"las órdenes de {PH}, desde la fecha de su ingreso hasta el {PH}, fecha en la "
        f"cual como a eso de las {PH} {PH}, {PH}, quien tiene facultades para contratar, "
        "despedir, dirigir y administrar trabajadores, le manifestó que a partir de ese "
        "momento estaba despedido de su trabajo, hecho que ocurrió en "
        f"{PH}."
    )
    L.paragraph(relacion_hechos)
    L.skip(16)

    # PARTE PETITORIA
    L.heading_left("PARTE PETITORIA")
    L.paragraph("Por lo antes expuesto, a usted respetuosamente PIDO:")
    L.skip(12)

    petitoria_intro = (
        f"Me admita la presente demanda, me tenga por parte en la calidad en que comparezco, "
        f"cite a conciliación a {PH} por medio de su representante legal antes mencionado, y "
        "si no llegásemos a ningún acuerdo en dicha audiencia, previo los trámites legales y "
        "las pruebas que oportunamente aportaré, sea condenada en la sentencia definitiva a "
        "pagarle a mi representado:"
    )
    L.paragraph(petitoria_intro)
    L.skip(8)
    L.bullet_lines(count=6)

    L.skip(14)

    # Textos fijos (como en demand.py)
    L.paragraph(
        "Legítimo mi personería con copia certificada por notario de la Credencial Única y "
        "sus respectivas copias de ley."
    )
    L.skip(12)
    L.paragraph("Señalo para oír notificaciones, uddt.sansalvador@pgr.gob.sv")
    L.skip(16)

    # Fecha (plantilla)
    L.paragraph(f"{PH} de {PH} de {PH}.")
    L.skip(10)

    c.save()
    return output_path


def _default_output_path() -> Path:
    return DEFAULT_OUTPUT_DIR / "demanda_template.pdf"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera la plantilla PDF en blanco de demanda laboral (sin datos).",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=_default_output_path(),
        help="Ruta de salida del PDF. Por defecto: output/demanda_template.pdf",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    final_path = build_blank_demanda_pdf(args.output)
    print(f"Plantilla demanda PDF generada en: {final_path}")


if __name__ == "__main__":
    main()
