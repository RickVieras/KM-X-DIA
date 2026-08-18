"""Processamento leve da aba PROGRAMADO para o painel KM x Dia."""
from __future__ import annotations
import io
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
import requests
from openpyxl import Workbook, load_workbook

COL_EMPRESA=3
COL_FROTA=(11,12,13)
COL_VIAGENS=(14,15,16)
COL_KM_OPERACIONAL=17
COL_KM_MORTA=18
COL_DIA_INICIAL=31
COL_DIA_FINAL=61

def number(value):
    try: return float(value or 0)
    except (TypeError,ValueError): return 0.0

def to_date(value):
    if isinstance(value,datetime): return value.date()
    if isinstance(value,date): return value
    if isinstance(value,str):
        for fmt in ("%Y-%m-%d","%d/%m/%Y","%d/%m/%y"):
            try: return datetime.strptime(value.strip(),fmt).date()
            except ValueError: pass
    return None

def transporta(row):
    return any("transporta" in str(value or "").casefold() for value in row[:COL_KM_MORTA])

def dates_in_period(ws,start_day,end_day):
    dates={}
    for row in ws.iter_rows(min_row=1,max_row=10,min_col=COL_DIA_INICIAL,max_col=COL_DIA_FINAL,values_only=True):
        for offset,value in enumerate(row,COL_DIA_INICIAL):
            current=to_date(value)
            if current and start_day<=current.day<=end_day:
                dates[offset]=current
    return sorted(dates.items())

def calculate(ws,start_day,end_day):
    dates=dates_in_period(ws,start_day,end_day)
    totals=defaultdict(lambda:{"km_operacional":0.0,"km_morta":0.0,"km_transporta":0.0,"km_total":0.0,"viagens":0.0})
    daily=defaultdict(lambda:{"frota":0.0,"viagens":0.0,"km_operacional":0.0,"km_morta":0.0,"km_transporta":0.0,"km_total":0.0})
    by_company=defaultdict(lambda:defaultdict(lambda:{"frota":0.0,"viagens":0.0,"km_operacional":0.0,"km_morta":0.0,"km_transporta":0.0,"km_total":0.0}))
    for row in ws.iter_rows(min_row=3,values_only=True):
        raw=row[COL_EMPRESA-1] if len(row)>=COL_EMPRESA else None
        company=str(raw).strip() if raw is not None else ""
        if not company or company.casefold() in {"none","nan","null","-"}: continue
        op=number(row[COL_KM_OPERACIONAL-1] if len(row)>=COL_KM_OPERACIONAL else 0)
        dead=number(row[COL_KM_MORTA-1] if len(row)>=COL_KM_MORTA else 0)
        is_transporta=transporta(row)
        totals[company]["km_operacional"]+=op
        totals[company]["km_morta"]+=dead
        totals[company]["km_transporta"]+=op if is_transporta else 0
        totals[company]["km_total"]+=op+dead
        totals[company]["viagens"]+=sum(number(row[c-1] if len(row)>=c else 0) for c in COL_VIAGENS)
        fleet=sum(number(row[c-1] if len(row)>=c else 0) for c in COL_FROTA)
        trips=sum(number(row[c-1] if len(row)>=c else 0) for c in COL_VIAGENS)
        for column,current in dates:
            km=number(row[column-1] if len(row)>=column else 0)
            if not km: continue
            key=current.isoformat()
            for target in (daily[key],by_company[company][key]):
                target["frota"]+=fleet
                target["viagens"]+=trips
                target["km_operacional"]+=km
                target["km_transporta"]+=km if is_transporta else 0
                target["km_total"]+=km
    companies=[{"empresa":name,**values} for name,values in sorted(totals.items())]
    daily_rows=[{"data":key,**values} for key,values in sorted(daily.items())]
    company_daily={name:[{"data":key,**values} for key,values in sorted(rows.items())] for name,rows in by_company.items()}
    return companies,daily_rows,company_daily

def download_source(sheet_id):
    response=requests.get(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",timeout=90)
    response.raise_for_status()
    return response.content

def build_report_from_bytes(source,destination:Path,start_day=1,end_day=31):
    workbook=load_workbook(io.BytesIO(source),data_only=True,read_only=True)
    if "PROGRAMADO" not in workbook.sheetnames: raise ValueError("A planilha precisa ter a aba PROGRAMADO.")
    companies,daily,company_daily=calculate(workbook["PROGRAMADO"],start_day,end_day)
    output=Workbook();output.remove(output.active);summary=output.create_sheet("TOTAL POR EMPRESA")
    summary.append(["Empresa","KM Operacional","KM Morta","KM Transporta","KM Total","Viagens"])
    for company in companies: summary.append([company["empresa"],company["km_operacional"],company["km_morta"],company["km_transporta"],company["km_total"],company["viagens"]])
    destination.parent.mkdir(parents=True,exist_ok=True);output.save(destination)
    return {"companies":companies,"daily":daily,"company_daily":company_daily}

def build_report(sheet_id,destination:Path,start_day=1,end_day=31):
    return build_report_from_bytes(download_source(sheet_id),destination,start_day,end_day)
