"""AplicaÃ§Ã£o web: histÃ³rico mensal, painel e download do KM X DIA."""
from __future__ import annotations

import io
import json
import os
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

from flask import Flask, jsonify, render_template, request, send_file

from km_processor import apply_non_operated, build_report, build_report_from_file, line_efficiency, line_values, write_report_from_payload
from supabase_history import SupabaseHistory

app = Flask(__name__, template_folder=".")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
OUTPUT_DIR = Path("data")
REPORT = OUTPUT_DIR / "KM X DIA.xlsx"
DATA = OUTPUT_DIR / "dashboard.json"
SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "")
HISTORY = SupabaseHistory(os.environ.get("SUPABASE_URL", ""), os.environ.get("SUPABASE_SECRET_KEY", ""))
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


@app.get("/")
def index():
    return render_template("index.html", google_sheets_enabled=bool(SHEET_ID))


def month_from_payload(payload: dict) -> tuple[int, int]:
    periods = set()
    for item in payload.get("daily", []):
        value = datetime.fromisoformat(item["data"])
        periods.add((value.year, value.month))
    if not periods:
        raise ValueError("NÃ£o foi possÃ­vel identificar o mÃªs pelas datas da planilha.")
    if len(periods) != 1:
        raise ValueError("A planilha deve conter datas de apenas um mÃªs.")
    return periods.pop()


def normalized_payload(row: dict) -> dict:
    payload = row.get("dados_dashboard") or {}
    payload["periodo"] = {"ano": row["ano"], "mes": row["mes"]}
    payload["update_info"] = {"updated_at":row["atualizado_em"],"source_name":row["nome_arquivo"],"source_type":"supabase"}
    return payload


@app.get("/api/periodos")
def periods():
    if not HISTORY.enabled:
        return jsonify({"error": "Supabase nÃ£o configurado no Render."}), 500
    try:
        return jsonify({"periodos": HISTORY.list_periods()})
    except Exception as error:
        app.logger.exception("Falha ao listar perÃ­odos")
        return jsonify({"error": str(error)}), 500


@app.get("/api/dashboard")
def dashboard():
    try:
        if HISTORY.enabled:
            year = request.args.get("ano", type=int)
            month = request.args.get("mes", type=int)
            row = HISTORY.get_month(year, month) if year and month else HISTORY.latest()
            if not row:
                return jsonify({"ready": False, "message": "Ainda nÃ£o hÃ¡ um mÃªs salvo no histÃ³rico."})
            return jsonify({"ready": True, **normalized_payload(row)})
        if not DATA.exists():
            return jsonify({"ready": False, "message": "Ainda nÃ£o hÃ¡ uma atualizaÃ§Ã£o processada."})
        return jsonify({"ready": True, **json.loads(DATA.read_text(encoding="utf-8"))})
    except Exception as error:
        app.logger.exception("Falha ao carregar o painel")
        return jsonify({"error": f"Falha ao carregar o histÃ³rico: {error}"}), 500


def authorized() -> bool:
    return not ADMIN_TOKEN or request.headers.get("X-Admin-Token") == ADMIN_TOKEN


def save_local(payload: dict, temporary_report: Path) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile("w", encoding="utf-8", dir=OUTPUT_DIR, delete=False) as handle:
        json.dump(payload, handle, ensure_ascii=False)
        temporary_json = Path(handle.name)
    temporary_report.replace(REPORT)
    temporary_json.replace(DATA)


@app.post("/api/atualizar")
def update():
    if not authorized():
        return jsonify({"error": "Acesso nÃ£o autorizado."}), 401
    if not HISTORY.enabled:
        return jsonify({"error": "Configure SUPABASE_URL e SUPABASE_SECRET_KEY no Render."}), 500
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    temporary_report = OUTPUT_DIR / "KM X DIA.processando.xlsx"
    temporary_upload = None
    try:
        start_raw = request.form.get("data_inicial")
        end_raw = request.form.get("data_final")
        mode = request.form.get("modo", "programado")
        upload = request.files.get("planilha")
        non_operated_upload = request.files.get("km_nao_realizada")
        if start_raw and end_raw:
            start = datetime.fromisoformat(start_raw).date()
            end = datetime.fromisoformat(end_raw).date()
        else:
            body = request.get_json(silent=True) or {}
            start_raw = body.get("data_inicial")
            end_raw = body.get("data_final")
            if not start_raw or not end_raw:
                return jsonify({"error": "Escolha a data inicial e a data final."}), 400
            start = datetime.fromisoformat(start_raw).date()
            end = datetime.fromisoformat(end_raw).date()
        if start.year != end.year or start.month != end.month:
            return jsonify({"error": "Para salvar no histórico, escolha datas do mesmo mês."}), 400
        year, month = start.year, start.month

        if mode == "nao_realizada":
            if not non_operated_upload or not non_operated_upload.filename:
                return jsonify({"error": "Selecione a planilha de KM Não Realizada."}), 400
            extension = Path(non_operated_upload.filename).suffix.casefold()
            if extension not in ALLOWED_EXTENSIONS:
                return jsonify({"error": "Envie a planilha de KM Não Realizada no formato .xlsx ou .xlsm."}), 400
            row = HISTORY.get_month(year, month)
            if not row:
                return jsonify({"error": "Atualize primeiro a planilha PROGRAMADO deste mês."}), 400
            payload = row.get("dados_dashboard") or {}
            previous = payload.get("parameters") or {}
            if previous.get("bm1") != start.isoformat() or previous.get("bn1") != end.isoformat():
                return jsonify({"error": "Use o mesmo período já salvo na leitura PROGRAMADO."}), 400
            with NamedTemporaryFile("wb", suffix=extension, dir=OUTPUT_DIR, delete=False) as handle:
                temporary_upload = Path(handle.name)
                non_operated_upload.save(handle)
            if not temporary_upload.stat().st_size:
                return jsonify({"error": "O arquivo de KM Não Realizada está vazio."}), 400
            apply_non_operated(payload, temporary_upload, start, end)
            write_report_from_payload(payload, temporary_report)
            source_name = Path(non_operated_upload.filename).name
        elif upload and upload.filename:
            extension = Path(upload.filename).suffix.casefold()
            if extension not in ALLOWED_EXTENSIONS:
                return jsonify({"error": "Envie a planilha PROGRAMADO no formato .xlsx ou .xlsm."}), 400
            with NamedTemporaryFile("wb", suffix=extension, dir=OUTPUT_DIR, delete=False) as handle:
                temporary_upload = Path(handle.name)
                upload.save(handle)
            if not temporary_upload.stat().st_size:
                return jsonify({"error": "O arquivo enviado está vazio."}), 400
            payload = build_report_from_file(temporary_upload, temporary_report, start, end)
            old = HISTORY.get_month(year, month)
            old_payload = old.get("dados_dashboard") if old else None
            old_params = old_payload.get("parameters") if old_payload else None
            if old_params and old_params.get("bm1") == start.isoformat() and old_params.get("bn1") == end.isoformat() and old_payload.get("line_non_operated"):
                payload["line_non_operated"] = old_payload["line_non_operated"]
                payload["lines"] = line_efficiency(line_values(payload["line_programmed"]), line_values(payload["line_non_operated"]))
                payload["has_non_operated_data"] = True
                write_report_from_payload(payload, temporary_report)
            source_name = Path(upload.filename).name
        else:
            if not SHEET_ID:
                return jsonify({"error": "Envie uma planilha PROGRAMADO."}), 400
            payload = build_report(SHEET_ID, temporary_report, start, end)
            source_name = "Google Sheets"
        calculated_year, calculated_month = month_from_payload(payload)
        if (calculated_year, calculated_month) != (year, month):
            return jsonify({"error": "As datas do período não correspondem às datas presentes na PROGRAMADO."}), 400
        saved = HISTORY.save_month(year, month, source_name, temporary_report, payload)
        payload["periodo"] = {"ano": year, "mes": month}
        payload["update_info"] = {"updated_at":saved["atualizado_em"],"source_name":source_name,"source_type":"supabase"}
        save_local(payload, temporary_report)
        return jsonify({"ok":True,"message":"MÃªs salvo no histÃ³rico com sucesso.","periodo":payload["periodo"]})
    except Exception as error:
        app.logger.exception("Falha ao processar a planilha KM X DIA")
        return jsonify({"error": f"Falha ao processar a planilha: {error}"}), 500
    finally:
        temporary_report.unlink(missing_ok=True)
        if temporary_upload:
            temporary_upload.unlink(missing_ok=True)


@app.errorhandler(413)
def file_too_large(_error):
    return jsonify({"error": "A planilha excede o limite de 25 MB."}), 413


@app.get("/baixar-planilha")
def download():
    try:
        if HISTORY.enabled:
            year = request.args.get("ano", type=int)
            month = request.args.get("mes", type=int)
            row = HISTORY.get_month(year, month) if year and month else HISTORY.latest()
            if not row:
                return jsonify({"error": "O perÃ­odo selecionado nÃ£o foi encontrado."}), 404
            content = HISTORY.download(row["caminho_arquivo"])
            return send_file(io.BytesIO(content), as_attachment=True, download_name=f"KM X DIA {row['mes']:02d}-{row['ano']}.xlsx")
        if not REPORT.exists():
            return jsonify({"error": "Atualize a base antes de baixar a planilha."}), 404
        return send_file(REPORT, as_attachment=True, download_name="KM X DIA.xlsx")
    except Exception as error:
        app.logger.exception("Falha ao baixar relatÃ³rio")
        return jsonify({"error": f"Falha ao baixar o relatÃ³rio: {error}"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
