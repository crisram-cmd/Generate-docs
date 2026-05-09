"""Generador de plantilla PDF para demanda de despido con salario por comisión.

Relación de hechos en dos párrafos: modalidad por comisión y relato del despido.
Sin sección aparte de cuantificación salarial.
"""

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


def build_blank_demand_comision_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    layout.skip(10)
    layout.paragraph(
        f"En las condiciones de trabajo antes mencionadas laboró mi patrocinado para y a "
        f"las órdenes de {PH}, desde la fecha de su ingreso hasta el {PH}, fecha en la "
        f"cual como a eso de las {PH}, {PH}, {PH}, quien tiene facultades para contratar, "
        "despedir, dirigir y administrar trabajadores, le manifestó que a partir de ese "
        "momento estaba despedido de su trabajo, hecho que ocurrió en el lugar señalado "
        f"para el emplazamiento específicamente en {PH}, por haber sido requerido de su "
        "presencia."
    )
    layout.skip(16)

    draw_common_petitoria_and_footer(layout)
    c.save()
    return output_path


def main() -> None:
    args = parse_args(
        "demanda_despido_comision_template.pdf",
        "Genera la plantilla PDF en blanco de demanda por despido con salario por comisión.",
    )
    final_path = build_blank_demand_comision_pdf(args.output)
    print(f"Plantilla demanda despido comisión generada en: {final_path}")


if __name__ == "__main__":
    main()
