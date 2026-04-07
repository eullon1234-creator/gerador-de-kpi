import os
import uuid
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from kpi_generator import gerar_kpi

app = Flask(__name__)
app.secret_key = os.urandom(24)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "outputs")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def _extensao_valida(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
    if "arquivo" not in request.files:
        flash("Nenhum arquivo enviado.", "erro")
        return redirect(url_for("index"))

    arquivo = request.files["arquivo"]

    if not arquivo.filename:
        flash("Nenhum arquivo selecionado.", "erro")
        return redirect(url_for("index"))

    if not _extensao_valida(arquivo.filename):
        flash("Formato inválido. Envie um arquivo .xlsx ou .xls", "erro")
        return redirect(url_for("index"))

    uid = uuid.uuid4().hex
    caminho_entrada = os.path.join(UPLOAD_FOLDER, f"{uid}_estoque.xlsx")
    caminho_saida   = os.path.join(OUTPUT_FOLDER,  f"{uid}_kpi.xlsx")

    arquivo.save(caminho_entrada)

    # Ler datas do formulário (formato "YYYY-MM" do input type="month")
    data_inicio_str = request.form.get("data_inicio", "").strip()
    data_fim_str    = request.form.get("data_fim", "").strip()

    try:
        data_inicio = pd.Timestamp(f"{data_inicio_str}-01") if data_inicio_str else None
        data_fim    = (pd.Timestamp(f"{data_fim_str}-01") + pd.offsets.MonthEnd(0)) if data_fim_str else None
    except Exception:
        data_inicio = None
        data_fim    = None

    try:
        gerar_kpi(caminho_entrada, caminho_saida, data_inicio=data_inicio, data_fim=data_fim)
    except ValueError as e:
        flash(str(e), "erro")
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"Erro inesperado ao processar a planilha: {e}", "erro")
        return redirect(url_for("index"))
    finally:
        if os.path.exists(caminho_entrada):
            os.remove(caminho_entrada)

    return send_file(
        caminho_saida,
        as_attachment=True,
        download_name="KPI_Estoque.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
