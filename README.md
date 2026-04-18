# Agente PBS - Generador de Presentaciones desde Excel

Agente de IA que lee datos de un archivo Excel y genera automáticamente una presentación PowerPoint profesional respetando una plantilla proporcionada por el usuario.

## Arquitectura

```
main.py                    # Punto de entrada CLI
src/
  agent.py                 # Agente Claude con tool use (agentic loop)
  excel_reader.py          # Lectura y estructuración de datos Excel
  template_reader.py       # Análisis de plantillas PowerPoint
  presentation_builder.py  # Construcción del archivo .pptx final
examples/
  crear_ejemplo_excel.py   # Genera un Excel de ejemplo
  crear_plantilla.py       # Genera una plantilla PowerPoint de ejemplo
tests/
  test_excel_reader.py
  test_presentation_builder.py
```

## Flujo del agente

```
Usuario → main.py
            ↓
        agent.py (Claude claude-sonnet-4-6 + tools)
            ↓
    ┌───────┴────────┐
    ↓                ↓
read_excel_file   read_template_file
    ↓                ↓
    └───────┬────────┘
            ↓
     (Claude analiza datos + plantilla)
            ↓
     create_presentation
            ↓
        output/*.pptx
```

## Instalación

```bash
pip install -r requirements.txt
cp .env.example .env
# Edita .env y agrega tu ANTHROPIC_API_KEY
```

## Uso rápido

```bash
# Generar archivos de ejemplo
python examples/crear_ejemplo_excel.py
python examples/crear_plantilla.py

# Ejecutar el agente
python main.py \
  --excel examples/datos_ejemplo.xlsx \
  --template examples/plantilla_ejemplo.pptx \
  --output output/mi_presentacion.pptx
```

## Opciones CLI

| Parámetro | Descripción | Requerido |
|-----------|-------------|-----------|
| `--excel` | Ruta al archivo Excel con datos | Sí |
| `--template` | Ruta a la plantilla PowerPoint (.pptx) | Sí |
| `--output` | Ruta de salida (default: `output/presentacion_generada.pptx`) | No |
| `--instructions` | Instrucciones adicionales de personalización | No |

## Ejemplo con instrucciones personalizadas

```bash
python main.py \
  --excel datos/ventas_2024.xlsx \
  --template templates/corporativo.pptx \
  --output presentaciones/ventas_q4.pptx \
  --instructions "Enfócate en los KPIs de ventas regionales, usa lenguaje ejecutivo y limita a 10 diapositivas"
```

## Herramientas del agente (Tool Use)

El agente Claude dispone de tres herramientas:

- **`read_excel_file`** — lee todas las hojas del Excel y devuelve un resumen estructurado
- **`read_template_file`** — analiza la plantilla PowerPoint: layouts, placeholders y diapositivas existentes
- **`create_presentation`** — construye el archivo `.pptx` final con el plan de diapositivas generado

## Tests

```bash
pytest tests/ -v
```

## Variables de entorno

| Variable | Descripción |
|----------|-------------|
| `ANTHROPIC_API_KEY` | API Key de Anthropic (requerida) |
