"""Persistência do histórico mensal no Supabase via APIs REST oficiais."""
from __future__ import annotations

import json
from uuid import uuid4
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

import requests

TABLE = "relatorios_mensais"
BUCKET = "relatorios-km"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class SupabaseHistory:
    def __init__(self, url: str, secret_key: str):
        self.url = url.rstrip("/")
        self.secret_key = secret_key

    @property
    def enabled(self) -> bool:
        return bool(self.url and self.secret_key)

    def headers(self, **extra) -> dict[str, str]:
        return {"apikey": self.secret_key, "Authorization": f"Bearer {self.secret_key}", **extra}

    @staticmethod
    def check(response: requests.Response) -> requests.Response:
        if response.ok:
            return response
        try: detail = response.json()
        except ValueError: detail = response.text
        raise RuntimeError(f"Supabase respondeu {response.status_code}: {detail}")

    def list_periods(self) -> list[dict]:
        response = requests.get(f"{self.url}/rest/v1/{TABLE}", headers=self.headers(), params={"select":"ano,mes,nome_arquivo,atualizado_em","order":"ano.desc,mes.desc"}, timeout=30)
        return self.check(response).json()

    def get_month(self, year: int, month: int) -> dict | None:
        response = requests.get(f"{self.url}/rest/v1/{TABLE}", headers=self.headers(), params={"ano":f"eq.{year}","mes":f"eq.{month}","select":"ano,mes,nome_arquivo,caminho_arquivo,dados_dashboard,atualizado_em","limit":"1"}, timeout=30)
        rows = self.check(response).json()
        return rows[0] if rows else None

    def latest(self) -> dict | None:
        response = requests.get(f"{self.url}/rest/v1/{TABLE}", headers=self.headers(), params={"select":"ano,mes,nome_arquivo,caminho_arquivo,dados_dashboard,atualizado_em","order":"ano.desc,mes.desc","limit":"1"}, timeout=30)
        rows = self.check(response).json()
        return rows[0] if rows else None

    def save_month(self, year: int, month: int, filename: str, report: Path, payload: dict) -> dict:
        storage_path = f"{year}/{month:02d}/relatorios/{uuid4().hex}/KM X DIA.xlsx"
        with report.open("rb") as handle:
            upload = requests.post(f"{self.url}/storage/v1/object/{BUCKET}/{quote(storage_path, safe='/')}", headers=self.headers(**{"Content-Type":XLSX_MIME,"x-upsert":"true"}), data=handle, timeout=90)
        self.check(upload)
        record = {"ano":year,"mes":month,"nome_arquivo":filename,"caminho_arquivo":storage_path,"dados_dashboard":payload,"atualizado_em":datetime.now(timezone.utc).isoformat()}
        response = requests.post(f"{self.url}/rest/v1/{TABLE}", headers=self.headers(**{"Content-Type":"application/json","Prefer":"resolution=merge-duplicates,return=representation"}), params={"on_conflict":"ano,mes"}, data=json.dumps(record,ensure_ascii=False), timeout=30)
        rows = self.check(response).json()
        return rows[0] if rows else record

    def download(self, storage_path: str) -> bytes:
        response = requests.get(f"{self.url}/storage/v1/object/authenticated/{BUCKET}/{quote(storage_path, safe='/')}", headers=self.headers(), timeout=90)
        return self.check(response).content


    def save_original(self, year, month, source_type, filename, local_file, start, end):
        if source_type not in {"custo_variavel", "quilometragem", "nao_realizada"}:
            raise ValueError("Fonte inválida.")
        extension = Path(filename).suffix.lower()
        if extension not in {".xlsx", ".xlsm"}:
            raise ValueError("Formato de arquivo inválido.")
        storage_path = f"originais/{year}/{month:02d}/{source_type}/{uuid4().hex}{extension}"
        mime = XLSX_MIME if extension == ".xlsx" else "application/vnd.ms-excel.sheet.macroEnabled.12"
        with Path(local_file).open("rb") as handle:
            response = requests.post(
                f"{self.url}/storage/v1/object/{BUCKET}/{quote(storage_path, safe='/')}",
                headers=self.headers(**{"Content-Type": mime, "x-upsert": "false"}),
                data=handle, timeout=90)
        self.check(response)
        return {"path": storage_path, "filename": Path(filename).name,
                "uploaded_at": datetime.now(timezone.utc).isoformat(),
                "size_bytes": Path(local_file).stat().st_size,
                "data_inicial": start.isoformat(), "data_final": end.isoformat()}
