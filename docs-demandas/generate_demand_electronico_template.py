"""Generador de plantilla PDF para demanda de despido por medio electrónico.

Sin bloque aparte de prueba digital: el relato en RELACIÓN DE HECHOS basta para este modelo.
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


def build_blank_demand_electronico_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    e = bold_field_caps(PH)
    layout.rich_paragraph(
        "En las condiciones de trabajo antes mencionadas laboró mi patrocinado para y a "
        f"las órdenes de {e}, desde la fecha de su ingreso hasta el {e}, fecha en la "
        f"cual como a eso de las {e}, {e}, {e}, quien tiene facultades para contratar, "
        "despedir, dirigir y administrar trabajadores, le manifestó que a partir de ese "
        f"momento estaba despedido de su trabajo, hecho que ocurrió específicamente en "
        f"{e},",
        space_after=6,
    )
    layout.rich_paragraph(
        "estando presente mi representado por habérsele solicitado su permanencia, cuando "
        f"recibió un mensaje de WhatsApp del número {e} de {e} a su número {e}, en el "
        "cual le manifestó que estaba despedido.",
        space_after=16,
    )

    draw_common_petitoria_and_footer(layout)
    c.save()
    return output_path


def main() -> None:
    args = parse_args(
        "demanda_despido_electronico_template.pdf",
        "Genera la plantilla PDF en blanco de demanda por despido electrónico.",
    )
    final_path = build_blank_demand_electronico_pdf(args.output)
    print(f"Plantilla demanda despido electrónico generada en: {final_path}")


if __name__ == "__main__":
    main()
