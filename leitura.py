"""
Funções comuns de leitura e limpeza, usadas pelos scripts do projeto.

Tudo aqui roda localmente: nenhuma função envia dados para fora do computador.
"""

import csv
import re
import unicodedata
from datetime import date
from pathlib import Path

# Codificações que as exportações costumam usar:
# - Vanguard: UTF-8 (às vezes com BOM, um caractere invisível no início)
# - Argus: Windows-1252 (o padrão antigo do Excel no Brasil)
CODIFICACOES = ["utf-8-sig", "cp1252"]


def normalizar_nome(nome):
    """
    Deixa um nome pronto para comparação entre sistemas.

    Remove acentos, passa para maiúsculas e tira espaços extras.
    Exemplo: "  Anna Júlia  Barreto " -> "ANNA JULIA BARRETO"

    Isso NÃO é correspondência por semelhança: "EMILY" e "EMILLY"
    continuam diferentes e precisam ser ligados na aba Apelidos.
    """
    if nome is None:
        return ""
    texto = str(nome)
    # Separa as letras dos acentos e descarta os acentos
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ascii", "ignore").decode("ascii")
    # split() sem argumento corta todos os espaços repetidos
    return " ".join(texto.upper().split())


def descobrir_codificacao(caminho):
    """Testa as codificações conhecidas e devolve a primeira que funciona."""
    dados = Path(caminho).read_bytes()
    for codificacao in CODIFICACOES:
        try:
            dados.decode(codificacao)
            return codificacao
        except UnicodeDecodeError:
            continue
    # cp1252 quase nunca falha, mas por garantia:
    raise ValueError(f"Não consegui descobrir a codificação de {caminho}")


def ler_csv(caminho):
    """
    Lê um CSV exportado (separador ;) e devolve (cabecalho, linhas, codificacao).

    Todos os valores são lidos como texto, exatamente como estão no arquivo,
    para não perder nada (ex.: "1,3" continua "1,3").
    """
    codificacao = descobrir_codificacao(caminho)
    with open(caminho, encoding=codificacao, newline="") as arquivo:
        leitor = csv.reader(arquivo, delimiter=";")
        todas = list(leitor)
    if not todas:
        return [], [], codificacao
    cabecalho = todas[0]
    # Ignora linhas totalmente vazias (o Argus coloca uma no final)
    linhas = [linha for linha in todas[1:] if any(c.strip() for c in linha)]
    return cabecalho, linhas, codificacao


def limpar_cabecalho(coluna):
    """
    Padroniza o nome de uma coluna para encontrá-la mesmo com pequenas variações.

    O Argus às vezes coloca símbolos antes do nome (ex.: "↓ Nome" vira "? Nome").
    Exemplo: "? Nome" -> "NOME"; "Valor R$" -> "VALOR R$"
    """
    texto = normalizar_nome(coluna)
    # Remove símbolos soltos no começo (?, ↓, *, etc.)
    while texto and not texto[0].isalnum() and texto[0] not in "<[":
        texto = texto[1:].strip()
    return texto


def ler_tabela(caminho):
    """
    Lê uma exportação em CSV, XLSX ou XLS e devolve (cabecalho, linhas).

    Os valores vêm sempre como texto, para cada script converter do seu jeito.
    No Excel, a primeira linha preenchida é considerada o cabeçalho.
    """
    caminho = Path(caminho)
    if caminho.suffix.lower() == ".csv":
        cabecalho, linhas, _ = ler_csv(caminho)
        return cabecalho, linhas

    import pandas as pd  # só é carregado quando há planilha Excel

    tabela = pd.read_excel(caminho, header=None, dtype=str).fillna("")
    linhas = [list(linha) for linha in tabela.itertuples(index=False)]
    linhas = [linha for linha in linhas if any(str(c).strip() for c in linha)]
    if not linhas:
        return [], []
    return [str(c) for c in linhas[0]], [[str(c) for c in linha] for linha in linhas[1:]]


def converter_valor_br(texto):
    """
    Converte valores no formato brasileiro em número.

    Exemplos: "R$ 1.234,56" -> 1234.56 | "10,50" -> 10.5 | "1,3" -> 1.3
              "1.500" -> 1500.0 (ponto de milhar) | "" -> 0.0
    """
    if texto is None:
        return 0.0
    if isinstance(texto, (int, float)):
        return float(texto)
    limpo = str(texto).replace("R$", "").replace(" ", "").replace("\xa0", "").strip()
    if not limpo or limpo == "-":
        return 0.0
    negativo = limpo.startswith("-")
    limpo = limpo.lstrip("-")
    if "," in limpo:
        # Vírgula é o separador decimal; pontos são de milhar
        limpo = limpo.replace(".", "").replace(",", ".")
    elif limpo.count(".") > 1 or re.fullmatch(r"\d{1,3}\.\d{3}", limpo):
        # Só pontos: "1.234.567" ou "1.500" são milhares
        limpo = limpo.replace(".", "")
    valor = float(limpo)
    return -valor if negativo else valor


def tempo_para_segundos(texto):
    """
    Converte "hh:mm:ss" (ou "mm:ss") em segundos. Vazio vira 0.

    Exemplos: "00:01:04" -> 64 | "01:02:05" -> 3725 | "" -> 0
    """
    texto = str(texto or "").strip()
    if not texto:
        return 0
    if " " in texto:  # "0 days 00:01:04" (vindo do Excel)
        texto = texto.split()[-1]
    partes = [float(p) for p in texto.split(":")]
    segundos = 0
    for parte in partes:
        segundos = segundos * 60 + parte
    return int(round(segundos))


def converter_data(texto):
    """
    Lê a data de um texto e devolve um `date` (ou None se não houver data).

    Aceita "19/10/2026", "19/10/2026 14:35:00", "2026-10-19" e "2026-10-19 14:35:00".
    """
    texto = str(texto or "").strip()
    achou = re.match(r"(\d{2})/(\d{2})/(\d{4})", texto)
    if achou:
        dia, mes, ano = (int(x) for x in achou.groups())
        return date(ano, mes, dia)
    achou = re.match(r"(\d{4})-(\d{2})-(\d{2})", texto)
    if achou:
        ano, mes, dia = (int(x) for x in achou.groups())
        return date(ano, mes, dia)
    return None
