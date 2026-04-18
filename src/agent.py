"""
PBS Presentation Agent
Uses Claude claude-sonnet-4-6 with tool use to orchestrate Excel → PowerPoint generation.
"""

import json
import os
from pathlib import Path
from typing import Any

import anthropic
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn

from excel_reader import read_excel, summarize_data
from template_reader import read_template, summarize_template
from presentation_builder import build_presentation

load_dotenv()
console = Console()

MODEL = "claude-sonnet-4-6"

# ─── Tool definitions ────────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "read_excel_file",
        "description": "Lee un archivo Excel y devuelve su contenido estructurado con todas las hojas, columnas y filas de datos.",
        "input_schema": {
            "type": "object",
            "properties": {
                "excel_path": {
                    "type": "string",
                    "description": "Ruta absoluta o relativa al archivo Excel (.xlsx)"
                }
            },
            "required": ["excel_path"]
        }
    },
    {
        "name": "read_template_file",
        "description": "Lee una plantilla de PowerPoint (.pptx) y devuelve información sobre sus layouts, placeholders y diapositivas existentes.",
        "input_schema": {
            "type": "object",
            "properties": {
                "template_path": {
                    "type": "string",
                    "description": "Ruta absoluta o relativa al archivo de plantilla PowerPoint (.pptx)"
                }
            },
            "required": ["template_path"]
        }
    },
    {
        "name": "create_presentation",
        "description": "Crea un archivo PowerPoint a partir de la plantilla y un plan de diapositivas generado con los datos del Excel.",
        "input_schema": {
            "type": "object",
            "properties": {
                "template_path": {
                    "type": "string",
                    "description": "Ruta a la plantilla PowerPoint base"
                },
                "slide_plan": {
                    "type": "array",
                    "description": "Lista de diapositivas a crear",
                    "items": {
                        "type": "object",
                        "properties": {
                            "layout": {"type": "string", "description": "Nombre del layout de la plantilla"},
                            "title": {"type": "string", "description": "Título de la diapositiva"},
                            "subtitle": {"type": "string", "description": "Subtítulo (opcional)"},
                            "content": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Lista de puntos o texto de contenido"
                            },
                            "notes": {"type": "string", "description": "Notas del presentador (opcional)"}
                        },
                        "required": ["layout", "title"]
                    }
                },
                "output_path": {
                    "type": "string",
                    "description": "Ruta donde se guardará la presentación generada (ej: output/presentacion.pptx)"
                }
            },
            "required": ["template_path", "slide_plan", "output_path"]
        }
    }
]

# ─── Tool executor ────────────────────────────────────────────────────────────

def execute_tool(tool_name: str, tool_input: dict) -> Any:
    if tool_name == "read_excel_file":
        data = read_excel(tool_input["excel_path"])
        summary = summarize_data(data)
        return {"summary": summary, "raw": data}

    elif tool_name == "read_template_file":
        info = read_template(tool_input["template_path"])
        summary = summarize_template(info)
        return {"summary": summary, "raw": info}

    elif tool_name == "create_presentation":
        output_path = build_presentation(
            template_path=tool_input["template_path"],
            slide_plan=tool_input["slide_plan"],
            output_path=tool_input["output_path"],
        )
        return {"output_path": output_path, "slides_created": len(tool_input["slide_plan"])}

    else:
        return {"error": f"Unknown tool: {tool_name}"}

# ─── Agent loop ───────────────────────────────────────────────────────────────

def run_agent(excel_path: str, template_path: str, output_path: str, instructions: str = "") -> str:
    """
    Runs the PBS Presentation Agent.
    Returns the path to the generated .pptx file.
    """
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    user_message = f"""Eres un agente experto en creación de presentaciones profesionales.

Tu tarea es:
1. Leer el archivo Excel en: {excel_path}
2. Leer la plantilla PowerPoint en: {template_path}
3. Analizar los datos del Excel y la estructura de la plantilla
4. Diseñar un plan de diapositivas que presente los datos de forma clara y profesional
5. Crear la presentación final en: {output_path}

{f"Instrucciones adicionales: {instructions}" if instructions else ""}

Usa las herramientas disponibles en orden. Crea diapositivas con:
- Una portada (título + subtítulo)
- Una diapositiva de agenda/índice
- Diapositivas de contenido para cada sección de datos relevante
- Una diapositiva de conclusiones o cierre

Adapta el contenido y la cantidad de diapositivas a los datos del Excel.
Usa los layouts de la plantilla de forma apropiada."""

    messages = [{"role": "user", "content": user_message}]

    console.print(Panel("[bold blue]PBS Presentation Agent iniciado[/bold blue]"))

    with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), transient=True) as progress:
        task = progress.add_task("Procesando...", total=None)

        while True:
            progress.update(task, description="Consultando al agente Claude...")

            response = client.messages.create(
                model=MODEL,
                max_tokens=8096,
                tools=TOOLS,
                messages=messages,
            )

            # Add assistant response to history
            messages.append({"role": "assistant", "content": response.content})

            if response.stop_reason == "end_turn":
                # Extract final text
                for block in response.content:
                    if hasattr(block, "text"):
                        console.print(Panel(block.text, title="[green]Agente[/green]"))
                break

            if response.stop_reason == "tool_use":
                tool_results = []
                for block in response.content:
                    if block.type == "tool_use":
                        progress.update(task, description=f"Ejecutando: {block.name}...")
                        console.print(f"[cyan]→ Herramienta:[/cyan] {block.name}")

                        result = execute_tool(block.name, block.input)

                        if block.name == "create_presentation":
                            console.print(f"[green]✓ Presentación creada:[/green] {result['output_path']} ({result['slides_created']} diapositivas)")
                        else:
                            console.print(f"[dim]  Completado.[/dim]")

                        tool_results.append({
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": json.dumps(result, ensure_ascii=False, default=str),
                        })

                messages.append({"role": "user", "content": tool_results})

    # Find output path from tool results
    for msg in reversed(messages):
        if isinstance(msg["content"], list):
            for item in msg["content"]:
                if isinstance(item, dict) and item.get("type") == "tool_result":
                    try:
                        data = json.loads(item["content"])
                        if "output_path" in data:
                            return data["output_path"]
                    except Exception:
                        pass

    return output_path
