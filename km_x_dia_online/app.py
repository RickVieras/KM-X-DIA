"""Aplicação web: atualização manual, painel e download do KM X DIA."""
from __future__ import annotations
import json
import os
from pathlib import Path
from flask import Flask, jsonify, render_template, request, send_file
from km_processor import build_report

app = Flask(__name__)
OUTPUT_DIR = Path("data")
REPORT = OUTPUT_DIR / "KM X DIA.xlsx"
DATA = OUTPUT_DIR / "dashboard.json"
SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "")

@app.get("/")
def index(): return render_template("index.html")

@app.get("/api/dashboard")
def dashboard():
    if not DATA.exists(): return jsonify({"ready": False, "message": "Ainda não há uma atualização processada."})
    return jsonify({"ready": True, **json.loads(DATA.read_text(encoding="utf-8"))})

@app.post("/api/atualizar")
def update():
    if not SHEET_ID: return jsonify({"error": "GOOGLE_SHEET_ID não configurado."}), 500
    if ADMIN_TOKEN and request.headers.get("X-Admin-Token") != ADMIN_TOKEN: return jsonify({"error": "Acesso não autorizado."}), 401
    start, end = int(request.json.get("dia_inicial", 1)), int(request.json.get("dia_final", 31))
    payload = build_report(SHEET_ID, REPORT, start, end)
    DATA.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return jsonify({"ok": True, "message": "Base atualizada com sucesso."})

@app.get("/baixar-planilha")
def download():
    if not REPORT.exists(): return jsonify({"error": "Atualize a base antes de baixar a planilha."}), 404
    return send_file(REPORT, as_attachment=True, download_name="KM X DIA.xlsx")

if __name__ == "__main__": app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
