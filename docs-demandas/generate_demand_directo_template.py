"""Generador de plantilla PDF para demanda de despido directo."""

from __future__ import annotations

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

from templates.demand_template_base import (
    DemandLayout,
    PH,
    draw_common_petitoria_and_footer,
    draw_common_sections,
    parse_args,
)


def build_blank_demand_directo_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    layout.paragraph(
        f"En las condiciones de trabajo antes mencionadas laboró mi patrocinado para y a "
        f"las órdenes de {PH}, desde la fecha de su ingreso hasta el {PH}, fecha en la "
        f"cual como a eso de las {PH}, {PH}, {PH}, quien tiene facultades para contratar, "
        "despedir, dirigir y administrar trabajadores, le manifestó que a partir de ese "
        f"momento estaba despedido de su trabajo, hecho que ocurrió en {PH}."
    )
    layout.skip(16)

    draw_common_petitoria_and_footer(layout)
    c.save()
    return output_path


def main() -> None:
    args = parse_args(
        "demanda_despido_directo_template.pdf",
        "Genera la plantilla PDF en blanco de demanda por despido directo.",
    )
    final_path = build_blank_demand_directo_pdf(args.output)
    print(f"Plantilla demanda despido directo generada en: {final_path}")


if __name__ == "__main__":
    main()
