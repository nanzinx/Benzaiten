"""
Funções comuns de leitura e limpeza, usadas pelos scripts do projeto.

Tudo aqui roda localmente: nenhuma função envia dados para fora do computador.
"""

import csv
import unicodedata
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
