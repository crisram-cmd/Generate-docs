# docs-generator

Generadores de plantillas: PDF con ReportLab y FOLA03 en Word con python-docx a partir de JSON. La carpeta de salida por defecto es **`output/`** (`Generate-docs/output/`).

## Requisitos

```bash
cd Generate-docs
python -m pip install -r requirements.txt
```

Dependencia principal: **reportlab**.

## Plantillas FOLA

```bash
python generate_fola01_template.py
python generate_fola03_template.py
```

## FOLA03 en Word desde JSON (`generate_new_fola03_template.py`)

Genera el formulario **FOLA03** como **`.docx`** (Times, 4 páginas, layout alineado a la plantilla PDF de referencia).

### CLI

Desde `Generate-docs`:

```bash
# Plantilla con campos vacíos (JSON implícito {})
python generate_new_fola03_template.py

# Con datos: ruta a un JSON de análisis y archivo de salida
python generate_new_fola03_template.py --json sample_fola03_analysis.json -o output/mi_fola03.docx
```

| Opción            | Descripción                                                                           |
| ----------------- | ------------------------------------------------------------------------------------- |
| `--json` / `-j`   | Archivo JSON con los datos del caso. Si se omite, se usa `{}` (formulario en blanco). |
| `--output` / `-o` | Ruta del `.docx` a crear. Por defecto: `output/fola03_generado.docx`.                 |

### Formato del JSON

El documento espera un objeto raíz con (al menos) estas claves:

| Clave                          | Contenido                                                                                    |
| ------------------------------ | -------------------------------------------------------------------------------------------- |
| `worker_information`           | Datos de la persona usuaria (nombre, DUI, domicilio, teléfono, etc.).                        |
| `employment_relationship_data` | Empleador, relación de trabajo, salario, despido, etc.                                       |
| `fola03_information`           | Expediente, fechas de atención, complemento, hechos, documentos, observaciones finales.      |
| `law_suggestions`              | _(Opcional.)_ Texto que se inserta en la celda “Otros reclamos” de la tabla de pretensiones. |

Hay un ejemplo listo para probar: **`sample_fola03_analysis.json`** (misma carpeta que el script).

### Uso desde código

```python
from generate_new_fola03_template import write_fola03_docx

write_fola03_docx({"worker_information": {...}, ...}, "salida.docx")
```

`write_fola03_docx` acepta un **dict** o un **string** con JSON. La función async **`fola03_document_maker`** (mismo módulo) está pensada para integración con **Google ADK** y guarda el archivo como artefacto del agente, no como ruta local.

## Plantillas de demandas laborales

```bash
python docs-demandas/generate_demand_template.py #Template base que todas las referencias comparten
python docs-demandas/generate_demand_directo_template.py #3. Demanda despido directo
python docs-demandas/generate_demand_comision_template.py #5. Demanda despido, salario comisión
python docs-demandas/generate_demand_electronico_template.py #2. Demanda despido por medio electrónico
python docs-demandas/generate_demand_embarazo_template.py #8. Demanda despido mujer embarazada
python docs-demandas/generate_demand_horas_extra_template.py #6. Demanda por despido y horas extra
python docs-demandas/generate_demand_garantia_extendida_template.py #7. Demanda despido en garantía extendida
python docs-demandas/generate_demand_fuero_sindical_template.py #9. Demanda despido directivo sindical
```

Los PDF se guardan por defecto en **`output/`** con nombres como `demanda_template.pdf`, `demanda_despido_directo_template.pdf`, etc.

## Estructura relevante

| Ruta                              | Descripción                                     |
| --------------------------------- | ----------------------------------------------- |
| `requirements.txt`                | Dependencias pip (ReportLab, etc.)              |
| `output/`                         | PDF y DOCX generados (por defecto)              |
| `generate_new_fola03_template.py` | FOLA03 en Word; CLI y `write_fola03_docx` / ADK |
| `sample_fola03_analysis.json`     | JSON de ejemplo para probar el generador Word   |
| `docs-demandas/`                  | Scripts de demandas                             |
| `docs-demandas/templates/`        | Código compartido (`demand_template_base.py`)   |
