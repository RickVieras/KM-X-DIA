"""Aplicação web: atualização manual, painel e download do KM X DIA."""
from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile

from flask import Flask, jsonify, render_template, request, send_file

from km_processor import build_report, build_report_from_bytes

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

OUTPUT_DIR = Path("data")
REPORT = OUTPUT_DIR / "KM X DIA.xlsx"
DATA = OUTPUT_DIR / "dashboard.json"
SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "")
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


@app.get("/")
def index():
    return render_template("index.html", google_sheets_enabled=bool(SHEET_ID))


@app.get("/api/dashboard")
def dashboard():
    if not DATA.exists():
        return jsonify({"ready": False, "message": "Ainda não há uma atualização processada."})
    return jsonify({"ready": True, **json.loads(DATA.read_text(encoding="utf-8"))})


def authorized() -> bool:
    return not ADMIN_TOKEN or request.headers.get("X-Admin-Token") == ADMIN_TOKEN


def save_update(payload: dict, temporary_report: Path, source_name: str, source_type: str) -> None:
    """Publica relatório e JSON somente depois do processamento completo."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    payload["update_info"] = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "source_name": source_name,
        "source_type": source_type,
    }

    with NamedTemporaryFile("w", encoding="utf-8", dir=OUTPUT_DIR, delete=False) as handle:
        json.dump(payload, handle, ensure_ascii=False)
        temporary_json = Path(handle.name)

    temporary_report.replace(REPORT)
    temporary_json.replace(DATA)


@app.post("/api/atualizar")
def update():
    if not authorized():
        return jsonify({"error": "Acesso não autorizado."}), 401

    temporary_report = OUTPUT_DIR / "KM X DIA.processando.xlsx"
    try:
        start = int(request.form.get("dia_inicial", 1))
        end = int(request.form.get("dia_final", 31))
        upload = request.files.get("planilha")

        if upload and upload.filename:
            extension = Path(upload.filename).suffix.casefold()
            if extension not in ALLOWED_EXTENSIONS:
                return jsonify({"error": "Envie uma planilha no formato .xlsx ou .xlsm."}), 400
            source = upload.read()
            if not source:
                return jsonify({"error": "O arquivo enviado está vazio."}), 400
            payload = build_report_from_bytes(source, temporary_report, start, end)
            source_name = Path(upload.filename).name
            source_type = "upload"
        else:
            body = request.get_json(silent=True) or {}
            start = int(body.get("dia_inicial", start))
            end = int(body.get("dia_final", end))
            if not SHEET_ID:
                return jsonify({"error": "Envie um arquivo ou configure GOOGLE_SHEET_ID."}), 400
            payload = build_report(SHEET_ID, temporary_report, start, end)
            source_name = "Google Sheets"
            source_type = "google_sheets"

        save_update(payload, temporary_report, source_name, source_type)
        return jsonify({"ok": True, "message": "Base atualizada com sucesso.", "update_info": payload["update_info"]})
    except Exception as error:
        app.logger.exception("Falha ao processar a planilha KM X DIA")
        return jsonify({"error": f"Falha ao processar a planilha: {error}"}), 500
    finally:
        temporary_report.unlink(missing_ok=True)


@app.errorhandler(413)
def file_too_large(_error):
    return jsonify({"error": "A planilha excede o limite de 25 MB."}), 413


@app.get("/baixar-planilha")
def download():
    if not REPORT.exists():
        return jsonify({"error": "Atualize a base antes de baixar a planilha."}), 404
    return send_file(REPORT, as_attachment=True, download_name="KM X DIA.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
