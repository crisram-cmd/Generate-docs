# docs-generator

Generadores de plantillas PDF con Python y ReportLab. La carpeta de salida por defecto es **`output/`** en la raíz del proyecto (`Generate-docs/output/`).

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

| Ruta                       | Descripción                                   |
| -------------------------- | --------------------------------------------- |
| `requirements.txt`         | Dependencias pip                              |
| `output/`                  | PDF generados (por defecto)                   |
| `docs-demandas/`           | Scripts de demandas                           |
| `docs-demandas/templates/` | Código compartido (`demand_template_base.py`) |
