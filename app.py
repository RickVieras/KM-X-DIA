"""Aplicação web: histórico mensal, painel e download do KM X DIA."""
from __future__ import annotations

import io
from copy import deepcopy
import json
import os
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

from flask import Flask, jsonify, render_template, request, send_file

from km_processor import merge_source, apply_non_operated, build_report, build_report_from_file, line_efficiency, line_values, write_report_from_payload
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
        raise ValueError("Não foi possível identificar o mês pelas datas da planilha.")
    if len(periods) != 1:
        raise ValueError("A planilha deve conter datas de apenas um mês.")
    return periods.pop()


def normalized_payload(row: dict) -> dict:
    payload = row.get("dados_dashboard") or {}
    payload["periodo"] = {"ano": row["ano"], "mes": row["mes"]}
    payload["update_info"] = {"updated_at":row["atualizado_em"],"source_name":row["nome_arquivo"],"source_type":"supabase"}
    return payload


@app.get("/api/periodos")
def periods():
    if not HISTORY.enabled:
        return jsonify({"error": "Supabase não configurado no Render."}), 500
    try:
        return jsonify({"periodos": HISTORY.list_periods()})
    except Exception as error:
        app.logger.exception("Falha ao listar períodos")
        return jsonify({"error": str(error)}), 500


@app.get("/api/dashboard")
def dashboard():
    try:
        if HISTORY.enabled:
            year = request.args.get("ano", type=int)
            month = request.args.get("mes", type=int)
            row = HISTORY.get_month(year, month) if year and month else HISTORY.latest()
            if not row:
                return jsonify({"ready": False, "message": "Ainda não há um mês salvo no histórico."})
            return jsonify({"ready": True, **normalized_payload(row)})
        if not DATA.exists():
            return jsonify({"ready": False, "message": "Ainda não há uma atualização processada."})
        return jsonify({"ready": True, **json.loads(DATA.read_text(encoding="utf-8"))})
    except Exception as error:
        app.logger.exception("Falha ao carregar o painel")
        return jsonify({"error": f"Falha ao carregar o histórico: {error}"}), 500


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
        return jsonify({"error": "Acesso não autorizado."}), 401
    if not HISTORY.enabled:
        return jsonify({"error": "Configure SUPABASE_URL e SUPABASE_SECRET_KEY no Render."}), 500
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(suffix=".xlsx", dir=OUTPUT_DIR, delete=False) as handle:
        temporary_report = Path(handle.name)
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
        if start > end:
            raise ValueError("A data inicial não pode ser maior que a final.")
        if mode not in {"programado", "quilometragem", "nao_realizada"}:
            raise ValueError("Tipo de atualização inválido.")
        source_type = "quilometragem" if mode == "quilometragem" else "custo_variavel"
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
            container = deepcopy(row.get("dados_dashboard") or {})
            payload = container.get("sources", {}).get("custo_variavel") if container.get("sources") else container
            if not payload:
                raise ValueError("Atualize primeiro o Custo Variável deste mês.")
            previous = payload.get("parameters") or {}
            if previous.get("bm1") != start.isoformat() or previous.get("bn1") != end.isoformat():
                return jsonify({"error": "Use o mesmo período já salvo na leitura PROGRAMADO."}), 400
            with NamedTemporaryFile("wb", suffix=extension, dir=OUTPUT_DIR, delete=False) as handle:
                temporary_upload = Path(handle.name)
                non_operated_upload.save(handle)
            if not temporary_upload.stat().st_size:
                return jsonify({"error": "O arquivo de KM Não Realizada está vazio."}), 400
            apply_non_operated(payload, temporary_upload, start, end)
            if container.get("sources"):
                container["sources"]["custo_variavel"] = payload
                payload = {**payload, "sources": container["sources"], "primary_source": "custo_variavel", "originals": container.get("originals", {})}
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
            fresh = build_report_from_file(temporary_upload, temporary_report, start, end, source_type=source_type)
            old = HISTORY.get_month(year, month)
            source_name = Path(upload.filename).name
            payload = merge_source(old.get("dados_dashboard") if old else {}, fresh, source_type, source_name)
            write_report_from_payload(payload, temporary_report)
        else:
            if not SHEET_ID:
                return jsonify({"error": "Envie uma planilha PROGRAMADO."}), 400
            payload = build_report(SHEET_ID, temporary_report, start, end)
            source_name = "Google Sheets"
            old = HISTORY.get_month(year, month)
            payload = merge_source(old.get("dados_dashboard") if old else {}, payload, "custo_variavel", source_name)
            write_report_from_payload(payload, temporary_report)
        if temporary_upload:
            original_type = "nao_realizada" if mode == "nao_realizada" else source_type
            original = HISTORY.save_original(year, month, original_type, source_name, temporary_upload, start, end)
            payload.setdefault("originals", {})[original_type] = original
        else:
            # A new Google Sheets reading has no uploaded Excel original attached.
            payload.setdefault("originals", {}).pop("custo_variavel", None)
        calculated_year, calculated_month = month_from_payload(payload)
        if (calculated_year, calculated_month) != (year, month):
            return jsonify({"error": "As datas do período não correspondem às datas presentes na PROGRAMADO."}), 400
        saved = HISTORY.save_month(year, month, source_name, temporary_report, payload)
        payload["periodo"] = {"ano": year, "mes": month}
        payload["update_info"] = {"updated_at":saved["atualizado_em"],"source_name":source_name,"source_type":"supabase"}
        save_local(payload, temporary_report)
        return jsonify({"ok":True,"message":"Mês salvo no histórico com sucesso.","periodo":payload["periodo"]})
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception as error:
        app.logger.exception("Falha ao processar a planilha KM X DIA")
        return jsonify({"error": f"Falha ao processar a planilha: {error}"}), 500
    finally:
        temporary_report.unlink(missing_ok=True)
        if temporary_upload:
            temporary_upload.unlink(missing_ok=True)


def original_context(values):
    source = values.get("fonte", "")
    if source not in {"custo_variavel", "quilometragem", "nao_realizada"}:
        raise ValueError("Selecione uma fonte válida.")
    year, month = int(values.get("ano", 0)), int(values.get("mes", 0))
    if not 1 <= month <= 12 or not 1900 <= year <= 2200:
        raise ValueError("Selecione um mês válido.")
    row = HISTORY.get_month(year, month)
    payload = deepcopy((row or {}).get("dados_dashboard") or {})
    original = payload.get("originals", {}).get(source)
    if not original:
        raise ValueError("Este mês ainda não possui o original desta fonte. Envie a planilha uma vez.")
    return year, month, source, payload, original


@app.get("/api/original")
def download_original():
    if not authorized():
        return jsonify({"error": "Acesso não autorizado."}), 401
    try:
        _, _, _, _, original = original_context(request.args)
        content = HISTORY.download(original["path"])
        response = send_file(io.BytesIO(content), as_attachment=True, download_name=original["filename"])
        response.headers["Cache-Control"] = "no-store"
        return response
    except (ValueError, TypeError) as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Falha ao baixar original")
        return jsonify({"error": "Não foi possível baixar o original. Confira os logs do servidor."}), 500


@app.post("/api/reprocessar")
def reprocess_original():
    if not authorized():
        return jsonify({"error": "Acesso não autorizado."}), 401
    temporary_report = None
    try:
        year, month, source, previous, original = original_context(request.get_json(silent=True) or {})
        start = datetime.fromisoformat(original["data_inicial"]).date()
        end = datetime.fromisoformat(original["data_final"]).date()
        if start > end or (start.year, start.month) != (year, month) or (end.year, end.month) != (year, month):
            raise ValueError("O período do original não corresponde ao mês selecionado.")
        content = io.BytesIO(HISTORY.download(original["path"]))
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(suffix=".xlsx", dir=OUTPUT_DIR, delete=False) as handle:
            temporary_report = Path(handle.name)
        if source == "nao_realizada":
            sources = previous.get("sources", {})
            primary = sources.get("custo_variavel") if sources else previous
            if not primary or primary.get("parameters") != {"bm1": start.isoformat(), "bn1": end.isoformat()}:
                raise ValueError("O período do Custo Variável deve ser igual ao do original de KM Não Realizada.")
            apply_non_operated(primary, content, start, end)
            payload = {**primary, "sources": sources, "originals": previous.get("originals", {})} if sources else primary
        else:
            fresh = build_report_from_file(content, temporary_report, start, end, source_type=source)
            payload = merge_source(previous, fresh, source, original["filename"])
        payload["originals"][source]["reprocessed_at"] = datetime.now().astimezone().isoformat()
        write_report_from_payload(payload, temporary_report)
        saved = HISTORY.save_month(year, month, original["filename"], temporary_report, payload)
        payload["periodo"] = {"ano": year, "mes": month}
        payload["update_info"] = {"updated_at": saved["atualizado_em"], "source_name": original["filename"], "source_type": "supabase"}
        save_local(payload, temporary_report)
        return jsonify({"ok": True, "periodo": payload["periodo"], "message": "Arquivo salvo reprocessado com sucesso."})
    except (ValueError, TypeError) as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Falha ao reprocessar original")
        return jsonify({"error": "Não foi possível reprocessar o original. Os dados anteriores foram preservados se a gravação não foi concluída. Confira os logs."}), 500
    finally:
        if temporary_report:
            temporary_report.unlink(missing_ok=True)


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
                return jsonify({"error": "O período selecionado não foi encontrado."}), 404
            content = HISTORY.download(row["caminho_arquivo"])
            return send_file(io.BytesIO(content), as_attachment=True, download_name=f"KM X DIA {row['mes']:02d}-{row['ano']}.xlsx")
        if not REPORT.exists():
            return jsonify({"error": "Atualize a base antes de baixar a planilha."}), 404
        return send_file(REPORT, as_attachment=True, download_name="KM X DIA.xlsx")
    except Exception as error:
        app.logger.exception("Falha ao baixar relatório")
        return jsonify({"error": f"Falha ao baixar o relatório: {error}"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
