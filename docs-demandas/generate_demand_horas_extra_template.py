"""Generador de plantilla PDF para demanda de despido y horas extra.

RELACIÓN DE HECHOS: relato tipo modelo (negritas/mayúsculas en datos del caso) + reclamo de horas extra y salarios.
"""

from __future__ import annotations

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

from templates.demand_template_base import (
    DemandLayout,
    PH,
    bold_field_caps,
    draw_common_petitoria_and_footer,
    draw_common_sections,
    parse_args,
)


def build_blank_demand_horas_extra_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    e = bold_field_caps(PH)
    layout.rich_paragraph(
        "En las condiciones "
        "de trabajo antes mencionadas laboró mi patrocinado para y a las órdenes de "
        f"{e}, desde la fecha de su ingreso hasta el {e}, fecha en la cual como a eso de "
        f"las {e}, {e}, {e}, quien tiene facultades para contratar, despedir, dirigir y "
        "administrar trabajadores, le manifestó que a partir de ese momento estaba despedido "
        f"de su trabajo, hecho que ocurrió específicamente en {e},",
        space_after=6,
    )

    draw_common_petitoria_and_footer(layout)
    c.save()
    return output_path


def main() -> None:
    args = parse_args(
        "demanda_despido_horas_extra_template.pdf",
        "Genera la plantilla PDF en blanco de demanda por despido y horas extra.",
    )
    final_path = build_blank_demand_horas_extra_pdf(args.output)
    print(f"Plantilla demanda despido horas extra generada en: {final_path}")


if __name__ == "__main__":
    main()
