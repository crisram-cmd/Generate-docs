"""Generador de plantilla PDF para demanda de despido en garantía extendida postnatal.

RELACIÓN DE HECHOS: relato del despido (patrocinada, datos en negritas) + párrafo centrado
sobre ilegalidad por garantía extendida y partida de nacimiento. Después, mismo esquema que
embarazo (orientación + viñetas) para argumentos y medida cautelar.
"""

from __future__ import annotations

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

from reportlab.lib.enums import TA_CENTER

from templates.demand_template_base import (
    DemandLayout,
    PH,
    bold_field_caps,
    draw_common_petitoria_and_footer,
    draw_common_sections,
    parse_args,
)


def build_blank_demand_garantia_extendida_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    e = bold_field_caps(PH)
    layout.rich_paragraph(
        "En las condiciones de trabajo antes mencionadas laboró mi patrocinada para y a las "
        f"órdenes de {e}, desde la fecha de su ingreso hasta el {e}, fecha en la cual como "
        f"a eso de las {e}, {e}, {e}, quien tiene facultades para contratar, despedir, dirigir "
        "y administrar trabajadores, le manifestó que a partir de ese momento estaba despedida "
        "de su trabajo, hecho que ocurrió en el lugar señalado para el emplazamiento "
        f"específicamente en {e}.",
        space_after=10,
    )
    layout.rich_paragraph(
        "El despido del que fue objeto la trabajadora demandante no produce los efectos "
        "legales de terminación del contrato, por encontrarse dentro de su garantía extendida "
        "de estabilidad laboral, tal como lo compruebo con la certificación de partida de "
        "nacimiento que adjunto a la presente demanda.",
        space_after=16,
        alignment=TA_CENTER,
    )

    layout.heading(
        "ARGUMENTOS DE DERECHO Y NORMAS JURÍDICAS QUE FUNDAMENTAN LA PRETENSIÓN"
    )
    layout.paragraph(
        "Fundamentos constitucionales y legales aplicables al caso concreto (estabilidad "
        "por gravidez, normativa laboral y tratados pertinentes). Redacte según los hechos "
        "y la prueba de este expediente:"
    )
    layout.skip(8)
    layout.blank_bullets(count=10)

    layout.heading("MEDIDA CAUTELAR")
    layout.paragraph(
        "Petición cautelar y fundamento de procedencia (apariencia de buen derecho, peligro "
        "en la demora, medidas solicitadas y criterios jurisprudenciales que correspondan). "
        "Redacte según el caso:"
    )
    layout.skip(8)
    layout.blank_bullets(count=13)

    layout.skip(16)
    draw_common_petitoria_and_footer(layout)
    c.save()
    return output_path


def main() -> None:
    args = parse_args(
        "demanda_garantia_extendida_template.pdf",
        "Genera la plantilla PDF en blanco de demanda en garantía extendida postnatal.",
    )
    final_path = build_blank_demand_garantia_extendida_pdf(args.output)
    print(f"Plantilla demanda garantía extendida generada en: {final_path}")


if __name__ == "__main__":
    main()
