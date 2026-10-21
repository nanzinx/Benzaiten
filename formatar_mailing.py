"""
Formatador de planilha para mailing (CSV de CPF, Benefício e Nome).

Versão em Python do "Formatador de Planilha para CSV" do projeto BENHERMES
(https://github.com/ProfetaSaru/benhermes, src/utils/spreadsheetFormatter.ts),
com as mesmas regras e o mesmo arquivo de saída:

  CPF;BENEFICIO;NOME;TELEFONE1;TELEFONE2;TELEFONE3;TELEFONE4;TELEFONE5

  - aceita .xlsx, .xls, .csv, .tsv e .txt (o separador do texto é descoberto sozinho);
  - acha a linha do cabeçalho nas 10 primeiras linhas (pode haver título antes);
  - identifica as colunas de CPF, Benefício e Nome pelo nome do cabeçalho;
  - as 5 colunas de telefone saem vazias, para preenchimento;
  - UTF-8 com BOM e quebra de linha do Windows (abre direto no Excel).

Melhorias em relação ao original:
  - nome exato do cabeçalho tem preferência ("Beneficio" ganha de "Desbloqueio NB")
    e a mesma coluna não é usada para dois campos;
  - CPF que perdeu o zero à esquerda no Excel é completado com zeros até 11 dígitos
    (use --manter-cpf para não mexer).

Tudo roda no seu computador; nada é enviado para fora.

Como usar:
  python formatar_mailing.py planilha.xlsx                -> saida/mailing/planilha_formatado.csv
  python formatar_mailing.py planilha.xlsx --aba "Base 2"  -> escolhe a aba
  python formatar_mailing.py planilha.csv --cpf "Documento" --beneficio "NB" --nome "Cliente"
                                                          -> escolhe as colunas à mão
  python formatar_mailing.py                              -> formata tudo o que estiver em entrada/mailing/
"""

import argparse
import csv
import io
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from leitura import descobrir_codificacao

PASTA_PROJETO = Path(__file__).resolve().parent
PASTA_ENTRADA = PASTA_PROJETO / "entrada" / "mailing"
PASTA_SAIDA = PASTA_PROJETO / "saida" / "mailing"

EXTENSOES = (".xlsx", ".xls", ".csv", ".tsv", ".txt")
CABECALHO_SAIDA = "CPF;BENEFICIO;NOME;TELEFONE1;TELEFONE2;TELEFONE3;TELEFONE4;TELEFONE5"
LINHAS_PARA_ACHAR_CABECALHO = 10

# Padrões de cabeçalho (sem acento, minúsculo, só letras e números), iguais aos do BENHERMES
PADROES = {
    "cpf": ["cpf", "documento", "doc", "cpfcliente", "numcpf", "nrcpf", "nrdocumento",
            "cpftitular", "numerocpf"],
    "beneficio": ["beneficio", "benef", "nb", "nubeneficio", "nrbeneficio", "numbeneficio",
                  "numerobeneficio", "numdocbeneficio", "matricula", "beneficiario"],
    "nome": ["nome", "cliente", "nomecompleto", "titular", "nometitular", "nomecliente",
             "segurado", "razaosocial"],
}
CAMPOS = ("cpf", "beneficio", "nome")


# ---------------------------------------------------------------------------
# Leitura da planilha
# ---------------------------------------------------------------------------

def normalizar_cabecalho(texto):
    """'Nº do Benefício' -> 'nodobeneficio' (sem acento, minúsculo, só letras e números)."""
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", texto.lower())


def como_texto(valor):
    """Valor da célula como texto. Números inteiros vindos do Excel perdem o '.0'."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    texto = str(valor).strip()
    return texto[:-2] if re.fullmatch(r"\d+\.0", texto) else texto


def ler_texto(caminho):
    """Lê .csv/.tsv/.txt descobrindo a codificação e o separador (; , tab ou |)."""
    conteudo = Path(caminho).read_bytes().decode(descobrir_codificacao(caminho))
    amostra = "\n".join(conteudo.splitlines()[:20])
    try:
        separador = csv.Sniffer().sniff(amostra, delimiters=";,\t|").delimiter
    except csv.Error:
        separador = "\t" if Path(caminho).suffix.lower() == ".tsv" else ";"
    return [list(linha) for linha in csv.reader(io.StringIO(conteudo), delimiter=separador)]


def ler_planilha(caminho, aba=None):
    """
    Devolve (abas, aba_usada, linhas). As linhas vêm como listas de texto, sem tratamento.
    Arquivos de texto têm uma "aba" só, com o nome do arquivo.
    """
    caminho = Path(caminho)
    sufixo = caminho.suffix.lower()
    if sufixo not in EXTENSOES:
        raise ValueError(f"Formato não suportado: {caminho.name}. Use .xlsx, .xls, .csv, .tsv ou .txt.")
    if sufixo in (".csv", ".tsv", ".txt"):
        return [caminho.name], caminho.name, [[como_texto(c) for c in l] for l in ler_texto(caminho)]

    import pandas as pd  # só para planilhas Excel

    abas = list(pd.ExcelFile(caminho).sheet_names)
    if aba and aba not in abas:
        raise ValueError(f"A aba '{aba}' não existe em {caminho.name}. Abas: {', '.join(abas)}")
    aba_usada = aba or abas[0]
    tabela = pd.read_excel(caminho, sheet_name=aba_usada, header=None, dtype=object)
    linhas = [[como_texto(None if pd.isna(c) else c) for c in linha] for linha in tabela.itertuples(index=False)]
    return abas, aba_usada, linhas


def linha_tem_dados(linha):
    return any(str(c).strip() for c in linha)


def achar_linha_do_cabecalho(linhas):
    """
    Procura o cabeçalho nas 10 primeiras linhas: cada célula que parece CPF, Benefício
    ou Nome vale 3 pontos e a linha com mais pontos vence (como no BENHERMES).
    """
    melhor, maior = 0, -1
    for i, linha in enumerate(linhas[:LINHAS_PARA_ACHAR_CABECALHO]):
        if not linha_tem_dados(linha):
            continue
        pontos = 0
        for celula in linha:
            texto = normalizar_cabecalho(celula)
            if not texto:
                continue
            pontos += sum(3 for campo in CAMPOS if any(p in texto for p in PADROES[campo]))
        if pontos > maior:
            melhor, maior = i, pontos
    return melhor if maior > 0 else 0


def detectar_colunas(cabecalhos):
    """
    Escolhe as colunas de CPF, Benefício e Nome. Primeiro procura o nome EXATO
    ('cpf', 'beneficio', 'nome'...), depois um cabeçalho que CONTENHA o padrão.
    Uma coluna já escolhida não é usada de novo.
    """
    normalizados = [normalizar_cabecalho(c) for c in cabecalhos]
    escolhidas = {}
    for campo in CAMPOS:
        usadas = set(escolhidas.values())
        candidatos = [i for i in range(len(cabecalhos)) if cabecalhos[i] not in usadas]
        exato = next((i for i in candidatos if normalizados[i] in PADROES[campo]), None)
        parcial = next((i for i in candidatos if any(p in normalizados[i] for p in PADROES[campo])), None)
        indice = exato if exato is not None else parcial
        escolhidas[campo] = cabecalhos[indice] if indice is not None else ""
    return escolhidas


# ---------------------------------------------------------------------------
# Formatação
# ---------------------------------------------------------------------------

@dataclass
class Resultado:
    arquivo: str
    aba: str
    abas: list
    linha_cabecalho: int          # 1 = primeira linha da planilha
    colunas: dict                 # {"cpf": cabeçalho, "beneficio": ..., "nome": ...}
    linhas: list = field(default_factory=list)   # [(cpf, beneficio, nome)]
    cpfs_completados: int = 0
    sem_cpf: int = 0


def escapar(valor):
    """Padrão RFC 4180: aspas em volta se tiver ;, quebra de linha ou aspas (como no BENHERMES)."""
    texto = "" if valor is None else str(valor)
    if any(c in texto for c in ';\n\r"'):
        return '"' + texto.replace('"', '""') + '"'
    return texto


def completar_cpf(cpf):
    """'1234567890' -> '01234567890'. Só mexe em CPF com 9 ou 10 dígitos e nada mais."""
    if re.fullmatch(r"\d{9,10}", cpf):
        return cpf.zfill(11), True
    return cpf, False


def formatar(caminho, aba=None, colunas=None, manter_cpf=False):
    """Lê a planilha e devolve um Resultado pronto para gravar."""
    abas, aba_usada, brutas = ler_planilha(caminho, aba)
    if not any(linha_tem_dados(l) for l in brutas):
        raise ValueError(f"{Path(caminho).name}: a planilha parece estar vazia.")

    i_cab = achar_linha_do_cabecalho(brutas)
    cabecalhos = [str(c).strip() or f"Coluna {n + 1}" for n, c in enumerate(brutas[i_cab])]
    escolhidas = detectar_colunas(cabecalhos)
    for campo, valor in (colunas or {}).items():
        if valor:
            if valor not in cabecalhos:
                raise ValueError(f"A coluna '{valor}' não existe. Colunas: {', '.join(cabecalhos)}")
            escolhidas[campo] = valor
    faltando = [c.upper() for c in CAMPOS if not escolhidas[c]]
    if faltando:
        raise ValueError(
            f"{Path(caminho).name}: não identifiquei a(s) coluna(s) {', '.join(faltando)}.\n"
            f"Colunas da planilha: {', '.join(cabecalhos)}\n"
            'Informe à mão, por exemplo: --cpf "Documento" --beneficio "NB" --nome "Cliente"'
        )

    resultado = Resultado(Path(caminho).name, aba_usada, abas, i_cab + 1, escolhidas)
    indices = {c: cabecalhos.index(escolhidas[c]) for c in CAMPOS}
    for linha in (l for l in brutas[i_cab + 1:] if linha_tem_dados(l)):
        valor = lambda campo: linha[indices[campo]] if indices[campo] < len(linha) else ""
        cpf = valor("cpf")
        if not manter_cpf:
            cpf, completou = completar_cpf(cpf)
            resultado.cpfs_completados += completou
        resultado.sem_cpf += not cpf
        resultado.linhas.append((cpf, valor("beneficio"), valor("nome")))
    return resultado


def texto_csv(resultado):
    """Conteúdo do CSV final (sem o BOM), com as 5 colunas de telefone vazias."""
    linhas = [CABECALHO_SAIDA] + [f"{escapar(c)};{escapar(b)};{escapar(n)};;;;;" for c, b, n in resultado.linhas]
    return "\r\n".join(linhas)


def nome_de_saida(arquivo_original):
    """'Base Outubro (1).xlsx' -> 'Base_Outubro__1__formatado.csv' (como no BENHERMES)."""
    base = re.sub(r"[^a-zA-Z0-9_\-]", "_", Path(arquivo_original).stem)
    return f"{base}_formatado.csv"


def gravar(resultado, pasta_saida):
    destino = Path(pasta_saida) / nome_de_saida(resultado.arquivo)
    destino.parent.mkdir(parents=True, exist_ok=True)
    # UTF-8 com BOM: o Excel no Windows abre com os acentos certos
    destino.write_bytes(("﻿" + texto_csv(resultado)).encode("utf-8"))
    return destino


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def main(argumentos=None):
    parser = argparse.ArgumentParser(description="Formata planilhas para o CSV de mailing (CPF, Benefício, Nome).")
    parser.add_argument("arquivos", nargs="*", help="planilhas a formatar (padrão: tudo em entrada/mailing/)")
    parser.add_argument("--aba", help="aba da planilha Excel (padrão: a primeira)")
    parser.add_argument("--cpf", help="nome da coluna de CPF, se não for identificada sozinha")
    parser.add_argument("--beneficio", help="nome da coluna de Benefício")
    parser.add_argument("--nome", help="nome da coluna de Nome")
    parser.add_argument("--manter-cpf", action="store_true", help="não completar CPF com zeros à esquerda")
    parser.add_argument("--saida", default=str(PASTA_SAIDA), help=argparse.SUPPRESS)
    args = parser.parse_args(argumentos)

    arquivos = [Path(a) for a in args.arquivos]
    if not arquivos:
        arquivos = sorted(p for p in PASTA_ENTRADA.glob("*") if p.suffix.lower() in EXTENSOES and not p.name.startswith("~$"))
        if not arquivos:
            print(f"Nenhuma planilha em {PASTA_ENTRADA}. Coloque o arquivo lá ou informe o caminho:\n"
                  "  python formatar_mailing.py minha_planilha.xlsx")
            return 1

    erros = 0
    for caminho in arquivos:
        if not caminho.exists():
            print(f"\nERRO: não encontrei {caminho}")
            erros += 1
            continue
        try:
            resultado = formatar(caminho, args.aba,
                                 {"cpf": args.cpf, "beneficio": args.beneficio, "nome": args.nome},
                                 args.manter_cpf)
        except ValueError as erro:
            print(f"\nERRO: {erro}")
            erros += 1
            continue
        destino = gravar(resultado, args.saida)

        print(f"\nMAILING — {resultado.arquivo}")
        if len(resultado.abas) > 1:
            print(f"  Aba: {resultado.aba} (outras: {', '.join(a for a in resultado.abas if a != resultado.aba)}; "
                  "use --aba para escolher)")
        if resultado.linha_cabecalho > 1:
            print(f"  Cabeçalho encontrado na linha {resultado.linha_cabecalho}")
        for campo in CAMPOS:
            print(f"  {campo.upper():<10} <- coluna '{resultado.colunas[campo]}'")
        print(f"  Registros: {len(resultado.linhas)}")
        if resultado.cpfs_completados:
            print(f"  CPFs completados com zero à esquerda: {resultado.cpfs_completados}")
        if resultado.sem_cpf:
            print(f"  AVISO: {resultado.sem_cpf} registro(s) sem CPF")
        print(f"  Arquivo: {destino}")
    return 1 if erros else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PermissionError as erro:
        print(f"\nERRO: não consegui gravar {erro.filename}.\n"
              "Se esse arquivo estiver aberto no Excel, feche-o e rode de novo.")
        sys.exit(1)
