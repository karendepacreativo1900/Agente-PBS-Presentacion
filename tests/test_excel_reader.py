"""Tests for excel_reader module."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import pytest
import openpyxl
import tempfile
import os

from excel_reader import read_excel, summarize_data


@pytest.fixture
def sample_excel(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ventas"
    ws.append(["Mes", "Ingresos", "Clientes"])
    ws.append(["Enero", 1000, 50])
    ws.append(["Febrero", 1200, 60])
    path = tmp_path / "test.xlsx"
    wb.save(path)
    return str(path)


def test_read_excel_returns_sheets(sample_excel):
    data = read_excel(sample_excel)
    assert "Ventas" in data


def test_read_excel_columns(sample_excel):
    data = read_excel(sample_excel)
    assert "Mes" in data["Ventas"]["columns"]
    assert "Ingresos" in data["Ventas"]["columns"]


def test_read_excel_row_count(sample_excel):
    data = read_excel(sample_excel)
    assert data["Ventas"]["shape"]["rows"] == 2


def test_read_excel_missing_file():
    with pytest.raises(FileNotFoundError):
        read_excel("/no/existe/archivo.xlsx")


def test_summarize_data(sample_excel):
    data = read_excel(sample_excel)
    summary = summarize_data(data)
    assert "Ventas" in summary
    assert "Enero" in summary
