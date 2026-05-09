"""Generador de plantilla PDF para demanda de despido mujer embarazada.

Los argumentos de derecho y la medida cautelar dependen de cada expediente: esta plantilla
deja espacio en blanco para redactarlos (o puedes sustituirlos al generar el PDF definitivo
desde datos / borrador). Un texto modelo solo sirve de referencia documental, no como única
forma posible del caso.
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


def build_blank_demand_embarazo_pdf(output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=LETTER)
    layout = DemandLayout(c)

    draw_common_sections(layout)
    layout.heading("RELACIÓN DE HECHOS")
    layout.paragraph(
        f"La trabajadora se encontraba en estado de gravidez al momento del despido, "
        f"con constancia médica de fecha {PH} y fecha probable de parto {PH}. El despido "
        "se produjo sin autorización judicial, por lo que no produce efectos legales."
    )
    layout.skip(16)

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
        "demanda_despido_embarazo_template.pdf",
        "Genera la plantilla PDF en blanco de demanda por despido de mujer embarazada.",
    )
    final_path = build_blank_demand_embarazo_pdf(args.output)
    print(f"Plantilla demanda despido embarazo generada en: {final_path}")


if __name__ == "__main__":
    main()
