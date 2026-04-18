"""Reads and structures data from Excel files for presentation generation."""

import pandas as pd
from pathlib import Path


def read_excel(excel_path: str) -> dict:
    """
    Reads all sheets from an Excel file and returns structured data.
    Returns a dict with sheet names as keys and list-of-dicts as values.
    """
    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    xl = pd.ExcelFile(excel_path)
    data = {}

    for sheet in xl.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet)
        df = df.dropna(how="all").reset_index(drop=True)
        data[sheet] = {
            "columns": list(df.columns),
            "rows": df.to_dict(orient="records"),
            "shape": {"rows": len(df), "cols": len(df.columns)},
        }

    return data


def summarize_data(data: dict) -> str:
    """Converts structured Excel data into a readable text summary for the LLM."""
    lines = []
    for sheet_name, sheet in data.items():
        lines.append(f"## Hoja: {sheet_name}")
        lines.append(f"Columnas: {', '.join(str(c) for c in sheet['columns'])}")
        lines.append(f"Filas de datos: {sheet['shape']['rows']}")
        lines.append("")
        for i, row in enumerate(sheet["rows"][:50]):  # max 50 rows in summary
            lines.append(f"  Fila {i+1}: " + " | ".join(f"{k}: {v}" for k, v in row.items()))
        if sheet["shape"]["rows"] > 50:
            lines.append(f"  ... y {sheet['shape']['rows'] - 50} filas más")
        lines.append("")
    return "\n".join(lines)
