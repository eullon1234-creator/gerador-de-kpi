import os
import time
import uuid
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from kpi_generator import gerar_kpi
from kpi_rm_generator import gerar_kpi_rm

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Na Vercel, o único diretório gravável é o /tmp
UPLOAD_FOLDER = "/tmp"
OUTPUT_FOLDER = "/tmp"
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def _extensao_valida(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def _limpar_arquivos_antigos():
    """Remove arquivos com mais de 1 hora na pasta outputs para economizar espaço."""
    agora = time.time()
    if os.path.exists(OUTPUT_FOLDER):
        for f in os.listdir(OUTPUT_FOLDER):
            caminho = os.path.join(OUTPUT_FOLDER, f)
            if os.path.isfile(caminho):
                if (agora - os.path.getmtime(caminho)) > 3600:
                    try:
                        os.remove(caminho)
                    except:
                        pass


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

    # Limpeza preventiva de arquivos antigos
    _limpar_arquivos_antigos()

    # Parâmetros customizáveis
    try:
        meses_morto = int(request.form.get("meses_morto", 3))
        abc_a = float(request.form.get("abc_a", 80)) / 100.0
        abc_b = float(request.form.get("abc_b", 95)) / 100.0
    except (ValueError, TypeError):
        meses_morto = 3
        abc_a = 0.80
        abc_b = 0.95

    try:
        gerar_kpi(
            caminho_entrada, caminho_saida, 
            data_inicio=data_inicio, data_fim=data_fim,
            meses_morto=meses_morto, limite_abc_a=abc_a, limite_abc_b=abc_b
        )
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


@app.route("/gerar-rm", methods=["POST"])
def gerar_rm():
    """Gera o KPI do RM a partir da planilha de estoque (LOCESTOQUE, GRUPO, etc)."""
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
    caminho_entrada = os.path.join(UPLOAD_FOLDER, f"{uid}_rm.xlsx")
    caminho_saida   = os.path.join(OUTPUT_FOLDER,  f"{uid}_kpi_rm.xlsx")

    arquivo.save(caminho_entrada)

    _limpar_arquivos_antigos()

    try:
        top_grupos = int(request.form.get("top_grupos", 10))
        top_itens  = int(request.form.get("top_itens", 50))
        abc_a = float(request.form.get("abc_a_rm", 80)) / 100.0
        abc_b = float(request.form.get("abc_b_rm", 95)) / 100.0
    except (ValueError, TypeError):
        top_grupos = 10
        top_itens  = 50
        abc_a = 0.80
        abc_b = 0.95

    try:
        gerar_kpi_rm(
            caminho_entrada, caminho_saida,
            top_grupos=top_grupos, top_itens=top_itens,
            limite_abc_a=abc_a, limite_abc_b=abc_b,
        )
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
        download_name="KPI_do_RM.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
