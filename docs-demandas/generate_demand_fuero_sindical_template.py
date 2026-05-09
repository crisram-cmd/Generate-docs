"""Generador de plantilla PDF para demanda de despido con fuero sindical."""

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


def build_blank_demand_fuero_sindical_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    e = bold_field_caps(PH)
    layout.rich_paragraph(
        "En las condiciones de trabajo antes mencionadas laboró mi patrocinado para y a las "
        f"órdenes de {e}, desde la fecha de su ingreso hasta el {e}, fecha en la cual como "
        f"a eso de las {e}, {e}, {e}, quien tiene facultades para contratar, despedir, dirigir "
        "y administrar trabajadores, le manifestó que a partir de ese momento estaba despedido "
        "de su trabajo, hecho que ocurrió en el lugar señalado para el emplazamiento; "
        f"específicamente en {e}, despido que no surte sus efectos legales por ser mi "
        f"representado miembro de la Junta Directiva del {e}, con el cargo de {e}.",
        space_after=16,
    )

    layout.heading("FUNDAMENTO JURÍDICO")
    layout.paragraph(
        "Existe violación al derecho de sindicación ya que siendo el trabajador demandante, a "
        "la fecha del despido miembro de la Junta Directiva de un sindicato legalmente "
        "constituido, de conformidad con los artículos:"
    )
    layout.skip(8)
    layout.blank_bullets(count=12)
    layout.skip(16)

    draw_common_petitoria_and_footer(layout)
    c.save()
    return output_path


def main() -> None:
    args = parse_args(
        "demanda_fuero_sindical_template.pdf",
        "Genera la plantilla PDF en blanco de demanda por despido de directivo sindical.",
    )
    final_path = build_blank_demand_fuero_sindical_pdf(args.output)
    print(f"Plantilla demanda fuero sindical generada en: {final_path}")


if __name__ == "__main__":
    main()
