import csv
import json
import sys
from pathlib import Path
from io import StringIO


def detectar_delimitador(amostra, delimitadores=(",", ";", "\t", "|")):
    """Detecta o delimitador mais provável em uma amostra de CSV."""
    contagens = {d: amostra.count(d) for d in delimitadores}
    return max(contagens, key=contagens.get) if any(contagens.values()) else ","


def achatar_valor(valor):
    """Converte valores complexos (dict, list) em string JSON para preservar dados."""
    if isinstance(valor, (dict, list)):
        return json.dumps(valor, ensure_ascii=False)
    return valor


def achatar_registro(registro):
    """Achata um registro, convertendo objetos/listas aninhados em strings JSON."""
    return {chave: achatar_valor(valor) for chave, valor in registro.items()}


def validar_cabecalhos(fieldnames):
    """Valida cabeçalhos do CSV."""
    if not fieldnames:
        return "Erro: o CSV está vazio ou sem cabeçalho."

    if any(not coluna.strip() for coluna in fieldnames):
        return "Erro: existe uma coluna sem nome."

    if len(set(fieldnames)) != len(fieldnames):
        return "Erro: existem nomes de colunas repetidos."

    return None


def ler_csv_com_validacao(caminho_csv):
    """Lê CSV com detecção de delimitador e validações completas."""
    with caminho_csv.open("r", encoding="utf-8-sig", newline="") as entrada:
        amostra = entrada.read(4096)
        entrada.seek(0)
        delimitador = detectar_delimitador(amostra)
        leitor = csv.DictReader(entrada, delimiter=delimitador)

        erro = validar_cabecalhos(leitor.fieldnames)
        if erro:
            return None, None, erro

        cabecalhos = list(leitor.fieldnames)
        registros = []

        for numero_linha, linha in enumerate(leitor, start=2):
            if None in linha:
                return None, None, (
                    f"Erro: linha {numero_linha} tem valores extras."
                )

            if any(valor is None for valor in linha.values()):
                return None, None, (
                    f"Erro: linha {numero_linha} tem campos ausentes."
                )

            registros.append(linha)

    return cabecalhos, registros, None


def converter_csv_para_json(arquivo_csv, arquivo_json):
    caminho_csv = Path(arquivo_csv)
    caminho_json = Path(arquivo_json)

    if not caminho_csv.is_file():
        print(f"Erro: arquivo não encontrado: {caminho_csv}")
        return False

    try:
        cabecalhos, registros, erro = ler_csv_com_validacao(caminho_csv)
        if erro:
            print(erro)
            return False

        # Não converter campos automaticamente - preservar strings originais
        with caminho_json.open("w", encoding="utf-8") as saida:
            json.dump(
                registros,
                saida,
                ensure_ascii=False,
                indent=2
            )

        print(f"Conversão concluída: {caminho_json}")
        print(f"Registros convertidos: {len(registros)}")
        return True

    except (OSError, csv.Error, UnicodeError, ValueError) as erro:
        print(f"Erro durante a conversão: {erro}")
        return False


def converter_csv_para_excel(arquivo_csv, arquivo_excel):
    """Converte CSV para Excel usando openpyxl."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("Erro: openpyxl não está instalado.")
        return False

    caminho_csv = Path(arquivo_csv)
    caminho_excel = Path(arquivo_excel)

    if not caminho_csv.is_file():
        print(f"Erro: arquivo não encontrado: {caminho_csv}")
        return False

    try:
        cabecalhos, registros, erro = ler_csv_com_validacao(caminho_csv)
        if erro:
            print(erro)
            return False

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
                    vertical="center", wrap_text=True
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

        pasta.save(caminho_excel)
        print(f"Conversão concluída: {caminho_excel}")
        print(f"Registros convertidos: {len(registros)}")
        return True

    except (OSError, csv.Error, UnicodeError, ValueError) as erro:
        print(f"Erro durante a conversão: {erro}")
        return False


def ler_json_tabular(caminho_json):
    """Lê e valida JSON tabular (lista de objetos)."""
    try:
        with caminho_json.open("r", encoding="utf-8") as entrada:
            dados = json.load(entrada)
    except (OSError, json.JSONDecodeError, UnicodeError) as erro:
        return None, f"Erro ao ler JSON: {erro}"

    if not isinstance(dados, list):
        return None, "Erro: JSON deve ser uma lista de objetos."

    if not dados:
        return [], None

    for i, item in enumerate(dados):
        if not isinstance(item, dict):
            return None, f"Erro: item {i} não é um objeto."

    return dados, None


def coletar_todos_cabecalhos(registros):
    """Coleta todos os cabeçalhos únicos mantendo ordem de aparição."""
    vistos = set()
    cabecalhos = []
    for registro in registros:
        for chave in registro.keys():
            if chave not in vistos:
                vistos.add(chave)
                cabecalhos.append(chave)
    return cabecalhos


def normalizar_registros(registros, cabecalhos):
    """Garante que todos os registros tenham todas as chaves, preenchendo com string vazia."""
    normalizados = []
    for registro in registros:
        novo = {}
        for cabecalho in cabecalhos:
            valor = registro.get(cabecalho, "")
            novo[cabecalho] = achatar_valor(valor)
        normalizados.append(novo)
    return normalizados


def converter_json_para_csv(arquivo_json, arquivo_csv):
    caminho_json = Path(arquivo_json)
    caminho_csv = Path(arquivo_csv)

    if not caminho_json.is_file():
        print(f"Erro: arquivo não encontrado: {caminho_json}")
        return False

    try:
        registros, erro = ler_json_tabular(caminho_json)
        if erro:
            print(erro)
            return False

        if not registros:
            print("Erro: JSON não contém dados.")
            return False

        cabecalhos = coletar_todos_cabecalhos(registros)
        registros_normalizados = normalizar_registros(registros, cabecalhos)

        with caminho_csv.open("w", encoding="utf-8", newline="") as saida:
            escritor = csv.DictWriter(saida, fieldnames=cabecalhos)
            escritor.writeheader()
            escritor.writerows(registros_normalizados)

        print(f"Conversão concluída: {caminho_csv}")
        print(f"Registros convertidos: {len(registros)}")
        return True

    except (OSError, csv.Error, UnicodeError, ValueError) as erro:
        print(f"Erro durante a conversão: {erro}")
        return False


def converter_json_para_excel(arquivo_json, arquivo_excel):
    """Converte JSON para Excel usando openpyxl."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("Erro: openpyxl não está instalado.")
        return False

    caminho_json = Path(arquivo_json)
    caminho_excel = Path(arquivo_excel)

    if not caminho_json.is_file():
        print(f"Erro: arquivo não encontrado: {caminho_json}")
        return False

    try:
        registros, erro = ler_json_tabular(caminho_json)
        if erro:
            print(erro)
            return False

        if not registros:
            print("Erro: JSON não contém dados.")
            return False

        cabecalhos = coletar_todos_cabecalhos(registros)
        registros_normalizados = normalizar_registros(registros, cabecalhos)

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

        for numero_linha, registro in enumerate(registros_normalizados, start=2):
            for coluna, nome in enumerate(cabecalhos, start=1):
                valor = registro.get(nome, "")
                celula = planilha.cell(
                    row=numero_linha, column=coluna, value=valor
                )
                celula.font = Font(name="Arial", size=13)
                celula.alignment = Alignment(
                    vertical="center", wrap_text=True
                )
                celula.border = borda_fina

                if numero_linha % 2 == 0:
                    celula.fill = preenchimento_alternado

            planilha.row_dimensions[numero_linha].height = 24

        for coluna, nome in enumerate(cabecalhos, start=1):
            maior = len(str(nome))
            for registro in registros_normalizados:
                valor = registro.get(nome, "")
                maior = max(maior, len(str(valor)))
            planilha.column_dimensions[
                get_column_letter(coluna)
            ].width = min(max(maior + 4, 16), 35)

        if cabecalhos:
            planilha.auto_filter.ref = (
                f"A1:{get_column_letter(len(cabecalhos))}"
                f"{max(len(registros_normalizados) + 1, 1)}"
            )

        planilha.sheet_properties.pageSetUpPr.fitToPage = True
        planilha.page_setup.fitToWidth = 1
        planilha.page_setup.fitToHeight = 0
        planilha.page_setup.orientation = "landscape"

        pasta.save(caminho_excel)
        print(f"Conversão concluída: {caminho_excel}")
        print(f"Registros convertidos: {len(registros)}")
        return True

    except (OSError, csv.Error, UnicodeError, ValueError) as erro:
        print(f"Erro durante a conversão: {erro}")
        return False


def main():
    if len(sys.argv) < 4:
        print(
            "Uso: python conversor.py <modo> <entrada> <saida>\n"
            "Modos:\n"
            "  csv2json   - CSV para JSON\n"
            "  csv2xlsx   - CSV para Excel\n"
            "  json2csv   - JSON para CSV\n"
            "  json2xlsx  - JSON para Excel"
        )
        sys.exit(1)

    modo = sys.argv[1]
    entrada = sys.argv[2]
    saida = sys.argv[3]

    conversores = {
        "csv2json": converter_csv_para_json,
        "csv2xlsx": converter_csv_para_excel,
        "json2csv": converter_json_para_csv,
        "json2xlsx": converter_json_para_excel,
    }

    if modo not in conversores:
        print(f"Erro: modo inválido '{modo}'.")
        print("Modos válidos: csv2json, csv2xlsx, json2csv, json2xlsx")
        sys.exit(1)

    # Validar extensões
    ext_entrada = Path(entrada).suffix.lower()
    ext_saida = Path(saida).suffix.lower()

    if modo.startswith("csv2") and ext_entrada != ".csv":
        print("Erro: arquivo de entrada deve ser .csv")
        sys.exit(1)
    if modo.startswith("json2") and ext_entrada != ".json":
        print("Erro: arquivo de entrada deve ser .json")
        sys.exit(1)
    if modo.endswith("json") and ext_saida != ".json":
        print("Erro: arquivo de saída deve ser .json")
        sys.exit(1)
    if modo.endswith("xlsx") and ext_saida != ".xlsx":
        print("Erro: arquivo de saída deve ser .xlsx")
        sys.exit(1)

    sucesso = conversores[modo](entrada, saida)
    sys.exit(0 if sucesso else 1)


if __name__ == "__main__":
    main()