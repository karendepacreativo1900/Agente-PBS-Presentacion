#!/usr/bin/env python3
"""
PBS Presentation Agent - Entry point
Usage: python main.py --excel datos.xlsx --template plantilla.pptx [--output presentacion.pptx] [--instructions "..."]
"""

import argparse
import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from agent import run_agent
from rich.console import Console

console = Console()


def main():
    parser = argparse.ArgumentParser(
        description="Agente PBS: genera presentaciones PowerPoint desde datos de Excel"
    )
    parser.add_argument("--excel", required=True, help="Ruta al archivo Excel con los datos")
    parser.add_argument("--template", required=True, help="Ruta a la plantilla PowerPoint (.pptx)")
    parser.add_argument(
        "--output",
        default="output/presentacion_generada.pptx",
        help="Ruta de salida para la presentación (default: output/presentacion_generada.pptx)"
    )
    parser.add_argument(
        "--instructions",
        default="",
        help="Instrucciones adicionales para personalizar la presentación"
    )

    args = parser.parse_args()

    # Validate inputs
    if not Path(args.excel).exists():
        console.print(f"[red]Error:[/red] No se encontró el archivo Excel: {args.excel}")
        sys.exit(1)

    if not Path(args.template).exists():
        console.print(f"[red]Error:[/red] No se encontró la plantilla: {args.template}")
        sys.exit(1)

    if not os.environ.get("ANTHROPIC_API_KEY"):
        console.print("[red]Error:[/red] Falta la variable de entorno ANTHROPIC_API_KEY")
        console.print("Copia .env.example a .env y agrega tu API key.")
        sys.exit(1)

    console.print(f"[bold]Excel:[/bold] {args.excel}")
    console.print(f"[bold]Plantilla:[/bold] {args.template}")
    console.print(f"[bold]Salida:[/bold] {args.output}")

    try:
        output_path = run_agent(
            excel_path=args.excel,
            template_path=args.template,
            output_path=args.output,
            instructions=args.instructions,
        )
        console.print(f"\n[bold green]✓ Presentación generada exitosamente:[/bold green] {output_path}")
    except Exception as e:
        console.print(f"[red]Error durante la generación:[/red] {e}")
        raise


if __name__ == "__main__":
    main()
