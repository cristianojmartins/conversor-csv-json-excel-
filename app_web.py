from flask import Flask, request, render_template_string, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
from io import BytesIO
import subprocess
import tempfile
import sys
import json
import csv

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

BASE_DIR = Path(__file__).resolve().parent
CONVERSOR = BASE_DIR / "conversor.py"

HTML = """
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Conversor CSV, JSON e Excel</title>
<style>
* { box-sizing: border-box; }
body {
    font-family: Arial, sans-serif;
    background: #f1f4f8;
    margin: 0;
    padding: 20px;
    color: #202938;
}
main {
    max-width: 560px;
    margin: 25px auto;
    background: white;
    padding: 25px;
    border-radius: 12px;
    box-shadow: 0 3px 12px #00000015;
}
h1 { font-size: 25px; margin-top: 0; }
p { line-height: 1.5; }
label { display: block; margin: 18px 0 8px; font-weight: bold; }
input[type=file], select {
    width: 100%;
    padding: 12px;
    border: 1px solid #aeb8c5;
    border-radius: 6px;
    background: white;
    font-size: 16px;
}
button {
    width: 100%;
    margin-top: 22px;
    padding: 14px;
    border: 0;
    border-radius: 6px;
    background: #1769aa;
    color: white;
    font-size: 17px;
    font-weight: bold;
}
.aviso {
    margin-top: 18px;
    padding: 12px;
    border-radius: 6px;
    background: #fff0f0;
    color: #9b1c1c;
    overflow-wrap: anywhere;
}
.info { color: #526174; font-size: 14px; }
</style>
</head>
<body>
<main>
<h1>Conversor CSV, JSON e Excel</h1>
<p>Selecione um arquivo CSV ou JSON e escolha a conversão desejada.</p>

<form method="post" enctype="multipart/form-data">
    <label for="arquivo">Arquivo de entrada</label>
    <input type="file" id="arquivo" name="arquivo"
           accept=".csv,.json,text/csv,application/json" required>

    <label for="formato">Conversão</label>
    <select id="formato" name="formato">
        <option value="csv2json">CSV → JSON</option>
        <option value="csv2xlsx">CSV → Excel (.xlsx)</option>
        <option value="json2csv">JSON → CSV</option>
        <option value="json2xlsx">JSON → Excel (.xlsx)</option>
    </select>

    <button type="submit">Converter e baixar arquivo</button>
</form>

{% if mensagem %}
<div class="aviso"><strong>Não foi possível concluir:</strong><br>
{{ mensagem }}</div>
{% endif %}

<p class="info">
CSV deve ter cabeçalho. JSON deve ser lista de objetos. Arquivos de até 5 MB.
</p>
</main>
</body>
</html>
"""


def pagina(mensagem=None, status=200):
    return render_template_string(HTML, mensagem=mensagem), status


def criar_excel(registros, cabecalhos):
    pasta = Workbook()
    planilha = pasta.active
    planilha.title = "Dados"
    planilha.freeze_panes = "A2"
    planilha.sheet_view.showGridLines = False

    cor_cabecalho = "174A70"
    cor_linha = "EAF2F8"
    cor_borda = "D5DEE7"

    preenchimento_cabecalho = PatternFill(
        fill_type="solid", fgColor=cor_cabecalho
    )
    preenchimento_alternado = PatternFill(
        fill_type="solid", fgColor=cor_linha
    )
    borda_fina = Border(
        bottom=Side(style="thin", color=cor_borda)
    )

    for coluna, nome in enumerate(cabecalhos, start=1):
        celula = planilha.cell(row=1, column=coluna, value=nome)
        celula.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
        celula.fill = preenchimento_cabecalho
        celula.alignment = Alignment(
            horizontal="left", vertical="center"
        )
        celula.border = borda_fina

    planilha.row_dimensions[1].height = 30

    for numero_linha, registro in enumerate(registros, start=2):
        for coluna, nome in enumerate(cabecalhos, start=1):
            valor = registro.get(nome, "")
            celula = planilha.cell(
                row=numero_linha, column=coluna, value=valor
            )
            celula.font = Font(name="Arial", size=13)
            celula.alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            celula.border = borda_fina

            if numero_linha % 2 == 0:
                celula.fill = preenchimento_alternado

        planilha.row_dimensions[numero_linha].height = 24

    for coluna, nome in enumerate(cabecalhos, start=1):
        maior = len(str(nome))

        for registro in registros:
            valor = registro.get(nome, "")
            maior = max(maior, len(str(valor)))

        planilha.column_dimensions[
            get_column_letter(coluna)
        ].width = min(max(maior + 4, 16), 35)

    if cabecalhos:
        planilha.auto_filter.ref = (
            f"A1:{get_column_letter(len(cabecalhos))}"
            f"{max(len(registros) + 1, 1)}"
        )

    planilha.sheet_properties.pageSetUpPr.fitToPage = True
    planilha.page_setup.fitToWidth = 1
    planilha.page_setup.fitToHeight = 0
    planilha.page_setup.orientation = "landscape"

    arquivo = BytesIO()
    pasta.save(arquivo)
    arquivo.seek(0)
    return arquivo


def achatar_valor(valor):
    """Converte valores complexos (dict, list) em string JSON para preservar dados."""
    if isinstance(valor, (dict, list)):
        return json.dumps(valor, ensure_ascii=False)
    return valor


def normalizar_registros_json(registros):
    """Normaliza registros JSON para Excel/CSV: achata objetos/listas e garante chaves uniformes."""
    if not registros:
        return [], []

    # Coletar todos os cabeçalhos
    cabecalhos = []
    vistos = set()
    for registro in registros:
        for chave in registro.keys():
            if chave not in vistos:
                vistos.add(chave)
                cabecalhos.append(chave)

    # Normalizar e achatar
    normalizados = []
    for registro in registros:
        novo = {}
        for cabecalho in cabecalhos:
            valor = registro.get(cabecalho, "")
            novo[cabecalho] = achatar_valor(valor)
        normalizados.append(novo)

    return normalizados, cabecalhos


@app.route("/", methods=["GET", "POST"])
def inicio():
    if request.method == "GET":
        return pagina()[0]

    arquivo = request.files.get("arquivo")
    formato = request.form.get("formato", "csv2json")

    formatos_validos = ("csv2json", "csv2xlsx", "json2csv", "json2xlsx")
    if formato not in formatos_validos:
        return pagina("Formato de conversão inválido.", 400)

    if not arquivo or not arquivo.filename:
        return pagina("Selecione um arquivo.", 400)

    nome_seguro = secure_filename(arquivo.filename)
    ext_entrada = Path(nome_seguro).suffix.lower()

    if ext_entrada not in (".csv", ".json"):
        return pagina("Selecione um arquivo .csv ou .json.", 400)

    # Validar extensão vs formato
    if formato.startswith("csv2") and ext_entrada != ".csv":
        return pagina("Para conversão de CSV, o arquivo deve ser .csv.", 400)
    if formato.startswith("json2") and ext_entrada != ".json":
        return pagina("Para conversão de JSON, o arquivo deve ser .json.", 400)

    if not CONVERSOR.is_file():
        return pagina("O arquivo conversor.py não foi encontrado.", 500)

    try:
        with tempfile.TemporaryDirectory() as temporario:
            pasta_temp = Path(temporario)
            entrada = pasta_temp / nome_seguro

            arquivo.save(entrada)

            if formato == "csv2json":
                saida_json = pasta_temp / "resultado.json"
                resultado = subprocess.run(
                    [
                        "/home/CristianoMartins/.local/bin/python3",
                        str(CONVERSOR),
                        "csv2json",
                        str(entrada),
                        str(saida_json),
                    ],
                    cwd=str(BASE_DIR),
                    capture_output=True,
                    text=True,
                    timeout=30,
                )

                if resultado.returncode != 0 or not saida_json.is_file():
                    mensagem = (
                        resultado.stderr.strip()
                        or resultado.stdout.strip()
                        or "O CSV não passou na validação."
                    )
                    return pagina(mensagem, 400)

                conteudo_json = saida_json.read_bytes()
                arquivo_memoria = BytesIO(conteudo_json)
                return send_file(
                    arquivo_memoria,
                    mimetype="application/json",
                    as_attachment=True,
                    download_name=f"{Path(nome_seguro).stem}.json",
                )

            elif formato == "csv2xlsx":
                saida_json = pasta_temp / "resultado.json"
                resultado = subprocess.run(
                    [
                        "/home/CristianoMartins/.local/bin/python3",
                        str(CONVERSOR),
                        "csv2json",
                        str(entrada),
                        str(saida_json),
                    ],
                    cwd=str(BASE_DIR),
                    capture_output=True,
                    text=True,
                    timeout=30,
                )

                if resultado.returncode != 0 or not saida_json.is_file():
                    mensagem = (
                        resultado.stderr.strip()
                        or resultado.stdout.strip()
                        or "O CSV não passou na validação."
                    )
                    return pagina(mensagem, 400)

                registros = json.loads(
                    saida_json.read_text(encoding="utf-8")
                )

                # Recupera os cabeçalhos do CSV original
                with entrada.open(
                    "r", encoding="utf-8-sig", newline=""
                ) as arquivo_csv:
                    # Detectar delimitador
                    amostra = arquivo_csv.read(4096)
                    arquivo_csv.seek(0)
                    delimitador = max(
                        [",", ";", "\t", "|"],
                        key=lambda d: amostra.count(d)
                    )
                    leitor = csv.reader(arquivo_csv, delimiter=delimitador)
                    cabecalhos = next(leitor, [])

                planilha = criar_excel(registros, cabecalhos)

                return send_file(
                    planilha,
                    mimetype=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    as_attachment=True,
                    download_name=f"{Path(nome_seguro).stem}.xlsx",
                )

            elif formato == "json2csv":
                saida_csv = pasta_temp / "resultado.csv"
                resultado = subprocess.run(
                    [
                        "/home/CristianoMartins/.local/bin/python3",
                        str(CONVERSOR),
                        "json2csv",
                        str(entrada),
                        str(saida_csv),
                    ],
                    cwd=str(BASE_DIR),
                    capture_output=True,
                    text=True,
                    timeout=30,
                )

                if resultado.returncode != 0 or not saida_csv.is_file():
                    mensagem = (
                        resultado.stderr.strip()
                        or resultado.stdout.strip()
                        or "O JSON não passou na validação."
                    )
                    return pagina(mensagem, 400)

                conteudo_csv = saida_csv.read_bytes()
                arquivo_memoria = BytesIO(conteudo_csv)
                return send_file(
                    arquivo_memoria,
                    mimetype="text/csv",
                    as_attachment=True,
                    download_name=f"{Path(nome_seguro).stem}.csv",
                )

            elif formato == "json2xlsx":
                saida_json = pasta_temp / "validado.json"
                resultado = subprocess.run(
                    [
                        "/home/CristianoMartins/.local/bin/python3",
                        str(CONVERSOR),
                        "json2csv",
                        str(entrada),
                        str(saida_json),  # usa json2csv só para validar
                    ],
                    cwd=str(BASE_DIR),
                    capture_output=True,
                    text=True,
                    timeout=30,
                )

                # O json2csv valida e escreve CSV, mas precisamos dos dados para Excel
                # Vamos ler o JSON original e validar/normalizar aqui
                try:
                    with entrada.open("r", encoding="utf-8") as f:
                        registros = json.load(f)
                except (json.JSONDecodeError, UnicodeError) as erro:
                    return pagina(f"Erro ao ler JSON: {erro}", 400)

                if not isinstance(registros, list):
                    return pagina("Erro: JSON deve ser uma lista de objetos.", 400)

                for i, item in enumerate(registros):
                    if not isinstance(item, dict):
                        return pagina(f"Erro: item {i} não é um objeto.", 400)

                registros_normalizados, cabecalhos = normalizar_registros_json(registros)

                planilha = criar_excel(registros_normalizados, cabecalhos)

                return send_file(
                    planilha,
                    mimetype=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    as_attachment=True,
                    download_name=f"{Path(nome_seguro).stem}.xlsx",
                )

    except subprocess.TimeoutExpired:
        return pagina("A conversão demorou demais e foi interrompida.", 500)
    except (OSError, UnicodeError, ValueError, csv.Error) as erro:
        return pagina(f"Erro ao processar o arquivo: {erro}", 400)
    except Exception:
        return pagina(
            "Ocorreu um erro interno ao processar o arquivo.",
            500,
        )


@app.errorhandler(413)
def arquivo_grande(erro):
    return pagina("O arquivo ultrapassa o limite de 5 MB.", 413)


if __name__ == "__main__":
    print("Conversor web iniciado.")
    print("Abra no navegador: http://127.0.0.1:8000")
    app.run(host="127.0.0.1", port=8000, debug=False)