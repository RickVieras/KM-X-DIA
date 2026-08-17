"""Processa a aba PROGRAMADO e gera o relatório KM X DIA sem Microsoft Excel."""

from __future__ import annotations

import io
import re
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COL_GARAGEM = 3
COL_FROTA = {"U": 11, "S": 12, "D": 13}
COL_VIAGENS = {"U": 14, "S": 15, "D": 16}
COL_OPERACIONAL = 17
COL_MORTA = 18
CALENDARIO_INICIO = 31
CALENDARIO_FIM = 61

HEADER = "123B5B"
SATURDAY = "CFE2F3"
SUNDAY = "F4CCCC"
MIXED = "D9EAD3"
THIN_GRAY = Side(style="thin", color="D9D9D9")


def number(value: object) -> float:
    try:
        return 0.0 if value is None or isinstance(value, bool) else float(value)
    except (TypeError, ValueError):
        return 0.0


def to_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\\/:*?\[\]]", " ", name).strip()[:31] or "SEM EMPRESA"
    candidate, index = base, 2
    while candidate.casefold() in used:
        suffix = f" ({index})"
        candidate = base[: 31 - len(suffix)] + suffix
        index += 1
    used.add(candidate.casefold())
    return candidate


def is_transporta(ws, row: int) -> bool:
    return any("transporta" in str(ws.cell(row, col).value or "").casefold() for col in range(1, COL_MORTA + 1))


def source_export_url(sheet_id: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"


def download_source(sheet_id: str) -> bytes:
    response = requests.get(source_export_url(sheet_id), timeout=60)
    response.raise_for_status()
    if not response.content.startswith(b"PK"):
        raise ValueError("Não foi possível baixar a planilha. Verifique se ela está compartilhada para visualização.")
    return response.content


def calculate(ws, day_start: int, day_end: int):
    dates = []
    for col in range(CALENDARIO_INICIO, CALENDARIO_FIM + 1):
        current = to_date(ws.cell(2, col).value)
        if current and day_start <= current.day <= day_end:
            dates.append((col, current))
    if not dates:
        raise ValueError("Nenhuma data foi encontrada entre AE e BI para o período informado.")

    values = defaultdict(lambda: defaultdict(lambda: [0.0] * 5))
    types = defaultdict(lambda: defaultdict(set))
    for row in range(3, ws.max_row + 1):
        company = str(ws.cell(row, COL_GARAGEM).value or "").strip()
        transporta = is_transporta(ws, row)
        if not company and not transporta:
            continue
        company = company or "TRANSPORTA"
        operational = number(ws.cell(row, COL_OPERACIONAL).value)
        dead = number(ws.cell(row, COL_MORTA).value)
        for col, current in dates:
            kind = str(ws.cell(row, col).value or "").strip().upper()
            if kind not in COL_FROTA:
                continue
            fleet = number(ws.cell(row, COL_FROTA[kind]).value)
            trips = number(ws.cell(row, COL_VIAGENS[kind]).value)
            km_operational = operational if transporta else trips * operational
            km_dead = fleet * dead
            total = values[company][current]
            total[0] += fleet
            total[1] += trips
            total[2] += km_operational
            total[3] += km_dead
            total[4] += km_operational + km_dead
            types[company][current].add(kind)
    if not values:
        raise ValueError("Nenhum dado de empresa foi encontrado na aba PROGRAMADO.")
    return values, types


def decorate(ws, title: str, records, types):
    ws.merge_cells("A1:F1")
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    headers = ["Data", "Frota", "Viagens", "KM Operacional", "KM x Morta", "KM Total"]
    for column, header in enumerate(headers, 1):
        cell = ws.cell(3, column, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 30
    for row, (current, metrics) in enumerate(records, 4):
        row_types = types.get(current, set())
        fill = MIXED if len(row_types) > 1 else SATURDAY if row_types == {"S"} else SUNDAY if row_types == {"D"} else None
        row_values = [current, *metrics]
        for col, value in enumerate(row_values, 1):
            cell = ws.cell(row, col, value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
            cell.number_format = "dd/mmm" if col == 1 else "#,##0" if col in (2, 3) else "#,##0.00"
    for column, width in {"A": 14, "B": 12, "C": 12, "D": 19, "E": 16, "F": 16}.items():
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A4"


def build_report(sheet_id: str, destination: Path, day_start: int = 1, day_end: int = 31) -> dict:
    if not 1 <= day_start <= day_end <= 31:
        raise ValueError("Informe dias válidos entre 1 e 31.")
    workbook = load_workbook(io.BytesIO(download_source(sheet_id)), data_only=True, read_only=False)
    if "PROGRAMADO" not in workbook.sheetnames:
        raise ValueError("A planilha de origem não possui a aba PROGRAMADO.")
    values, types = calculate(workbook["PROGRAMADO"], day_start, day_end)
    output = Workbook()
    output.remove(output.active)
    used = {"total por empresa"}
    total_by_day = defaultdict(lambda: [0.0] * 5)
    total_types = defaultdict(set)
    companies = []
    for company in sorted(values, key=str.casefold):
        records = sorted(values[company].items())
        ws = output.create_sheet(safe_sheet_name(company, used))
        decorate(ws, f"KM X DIA — {company} | Dias {day_start:02d} a {day_end:02d}", records, types[company])
        aggregate = [0.0] * 5
        for current, metrics in records:
            for index, value in enumerate(metrics):
                aggregate[index] += value
                total_by_day[current][index] += value
            total_types[current].update(types[company][current])
        companies.append({"empresa": company, "frota": aggregate[0] / len(records), "viagens": aggregate[1], "km_operacional": aggregate[2], "km_morta": aggregate[3], "km_total": aggregate[4]})
    total = output.create_sheet("TOTAL POR EMPRESA")
    decorate(total, f"KM X DIA — TOTAL POR DIA | Dias {day_start:02d} a {day_end:02d}", sorted(total_by_day.items()), total_types)
    destination.parent.mkdir(parents=True, exist_ok=True)
    output.save(destination)
    daily = [{"data": current.isoformat(), "frota": metrics[0], "viagens": metrics[1], "km_operacional": metrics[2], "km_morta": metrics[3], "km_total": metrics[4]} for current, metrics in sorted(total_by_day.items())]
    return {"companies": companies, "daily": daily}
