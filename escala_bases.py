"""
Escala diária de bases do Ativo 2.

Sorteia as bases (lotes/notas do INSS) entre as pessoas cadastradas
(o grupo operacional de cada supervisor), do mesmo jeito que o botão
"Sortear bases de hoje" da planilha Escala_Bases_Ativo2.xlsm.

Regras do sorteio:
  - Cada base recebe um responsável.
  - Se houver menos pessoas que bases, algumas pessoas ficam com mais de uma base.
  - Se houver mais pessoas que bases, algumas ficam sem base (aparece um aviso).
  - Ninguém repete a mesma base dos últimos JANELA_SORTEIOS sorteios.
    Se não existir nenhuma combinação possível, a janela diminui e um aviso aparece.
  - Pessoas ausentes no dia ficam fora do sorteio.

Como usar (no terminal, dentro da pasta do projeto):
  python escala_bases.py                          -> sorteia hoje
  python escala_bases.py --ausentes "FULANO"       -> sorteia hoje sem o FULANO
  python escala_bases.py --data 21/10/2026        -> sorteia outra data
  python escala_bases.py --importar-xlsm config/Escala_Bases_Ativo2.xlsm
                                                  -> cria config/bases.xlsx e copia o histórico da planilha antiga
"""

import argparse
import csv
import random
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from calendario import eh_feriado, ler_feriados
from leitura import normalizar_nome

# ---------------------------------------------------------------------------
# Configurações (mude aqui se precisar)
# ---------------------------------------------------------------------------

PASTA_PROJETO = Path(__file__).resolve().parent
ARQUIVO_CADASTRO = PASTA_PROJETO / "config" / "bases.xlsx"
ARQUIVO_HISTORICO = PASTA_PROJETO / "historico" / "historico_escalas.csv"
PASTA_SAIDA = PASTA_PROJETO / "saida"

# Quantos sorteios anteriores são olhados para evitar repetir a mesma base.
# Com 4 bases e 4 pessoas, 2 ainda deixa o sorteio aleatório;
# 3 deixaria só uma combinação possível (vira rodízio fixo).
JANELA_SORTEIOS = 2

# Quantas vezes embaralhar antes de desistir de uma janela (igual à macro)
TENTATIVAS = 1000

# Cores da empresa (as mesmas da planilha .xlsm)
AZUL = "0A2F6B"
AMARELO = "FFC20E"
AMARELO_CLARO = "FFF4CC"
BRANCO = "FFFFFF"

DIAS_DA_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
                  "sexta-feira", "sábado", "domingo"]

COLUNAS_HISTORICO = ["Data", "Base", "Responsavel"]


# ---------------------------------------------------------------------------
# Leitura do cadastro e do histórico
# ---------------------------------------------------------------------------

def limpar_lista(valores):
    """Tira espaços, linhas vazias e nomes repetidos, mantendo a ordem."""
    resultado = []
    vistos = set()
    for valor in valores:
        if pd.isna(valor):
            continue
        texto = " ".join(str(valor).split())
        if texto and normalizar_nome(texto) not in vistos:
            vistos.add(normalizar_nome(texto))
            resultado.append(texto)
    return resultado


def ler_cadastro(caminho=ARQUIVO_CADASTRO):
    """
    Lê config/bases.xlsx e devolve (bases, pessoas).

    O arquivo tem duas abas:
      - "Bases":   coluna "Base"   (uma base por linha)
      - "Pessoas": coluna "Pessoa" (uma pessoa por linha)
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(
            f"Não encontrei {caminho}.\n"
            "Crie o cadastro a partir da planilha antiga com:\n"
            "  python escala_bases.py --importar-xlsm config/Escala_Bases_Ativo2.xlsm"
        )
    bases = limpar_lista(pd.read_excel(caminho, sheet_name="Bases")["Base"])
    pessoas = limpar_lista(pd.read_excel(caminho, sheet_name="Pessoas")["Pessoa"])
    return bases, pessoas


def ler_historico(caminho=ARQUIVO_HISTORICO):
    """
    Lê o histórico acumulado e devolve uma lista de registros:
      [{"data": date, "base": "...", "responsavel": "..."}, ...]
    """
    caminho = Path(caminho)
    if not caminho.exists():
        return []
    registros = []
    with open(caminho, encoding="utf-8-sig", newline="") as arquivo:
        for linha in csv.DictReader(arquivo, delimiter=";"):
            registros.append({
                "data": datetime.strptime(linha["Data"], "%d/%m/%Y").date(),
                "base": linha["Base"],
                "responsavel": linha["Responsavel"],
            })
    return registros


def salvar_historico(registros, caminho=ARQUIVO_HISTORICO):
    """
    Grava o histórico inteiro, ordenado por data.

    Usa ; e UTF-8 com BOM para o Excel abrir com acentos certos.
    """
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    # sorted é estável: dentro do mesmo dia, mantém a ordem das bases
    ordenados = sorted(registros, key=lambda r: r["data"])
    with open(caminho, "w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        escritor.writerow(COLUNAS_HISTORICO)
        for r in ordenados:
            escritor.writerow([r["data"].strftime("%d/%m/%Y"), r["base"], r["responsavel"]])


# ---------------------------------------------------------------------------
# Sorteio
# ---------------------------------------------------------------------------

def sorteios_anteriores(historico, data_sorteio, quantidade):
    """
    Devolve os últimos `quantidade` sorteios ANTES da data informada.

    Cada sorteio é um dicionário {base_normalizada: pessoa_normalizada}.
    Conta sorteios registrados (não dias do calendário), então o fim de
    semana não "apaga" a sexta-feira.
    """
    datas = sorted({r["data"] for r in historico if r["data"] < data_sorteio}, reverse=True)
    datas = datas[:quantidade]
    anteriores = []
    for d in datas:
        anteriores.append({
            normalizar_nome(r["base"]): normalizar_nome(r["responsavel"])
            for r in historico if r["data"] == d
        })
    return anteriores


def repete_base(escolha, anteriores):
    """True se alguém pegou a mesma base em algum dos sorteios anteriores."""
    for base, pessoa in escolha.items():
        for sorteio in anteriores:
            if sorteio.get(normalizar_nome(base)) == normalizar_nome(pessoa):
                return True
    return False


def sortear(bases, pessoas, historico, data_sorteio, janela=JANELA_SORTEIOS, rng=None):
    """
    Faz o sorteio e devolve (escolha, janela_usada).

    escolha = {base: pessoa}, na ordem das bases do cadastro.
    janela_usada = quantos sorteios anteriores foram respeitados
                   (menor que `janela` quando não havia combinação possível).

    `rng` serve só para os testes repetirem o mesmo resultado;
    no uso normal fica None e o sorteio é aleatório de verdade.
    """
    if not bases:
        raise ValueError("Nenhuma base cadastrada.")
    if not pessoas:
        raise ValueError("Nenhuma pessoa disponível para o sorteio.")
    rng = rng or random.Random()

    # Tenta com a janela completa; se não der, diminui até 0 (sem restrição)
    for janela_atual in range(janela, -1, -1):
        anteriores = sorteios_anteriores(historico, data_sorteio, janela_atual)
        for _ in range(TENTATIVAS):
            embaralhadas = list(pessoas)
            rng.shuffle(embaralhadas)
            # Monta uma fila com uma vaga por base. Se faltar gente, volta para o
            # começo da lista, então ninguém fica com mais de 1 base a mais que outro.
            fila = [embaralhadas[i % len(embaralhadas)] for i in range(len(bases))]
            # Embaralha a fila de novo para quem tem 2 bases não pegar sempre as mesmas posições
            rng.shuffle(fila)
            escolha = dict(zip(bases, fila))
            if not repete_base(escolha, anteriores):
                return escolha, janela_atual
    # Só chega aqui se nem sem restrição funcionar, o que não deve acontecer
    raise RuntimeError("Não foi possível fazer o sorteio.")


def aplicar_ausencias(pessoas, ausentes):
    """
    Tira as pessoas ausentes da lista.

    Dá erro se algum nome de ausente não existir no cadastro,
    para um erro de digitação não passar despercebido.
    """
    cadastradas = {normalizar_nome(p): p for p in pessoas}
    desconhecidos = [a for a in ausentes if normalizar_nome(a) not in cadastradas]
    if desconhecidos:
        raise ValueError(
            f"Ausente(s) não encontrado(s) no cadastro: {', '.join(desconhecidos)}.\n"
            f"Pessoas cadastradas: {', '.join(pessoas)}"
        )
    fora = {normalizar_nome(a) for a in ausentes}
    return [p for p in pessoas if normalizar_nome(p) not in fora]


# ---------------------------------------------------------------------------
# Planilha de saída
# ---------------------------------------------------------------------------

def gerar_planilha(escolha, data_sorteio, ausentes, janela_usada, historico, caminho):
    """Cria saida/escala_AAAA-MM-DD.xlsx com a escala do dia e os últimos sorteios."""
    azul = PatternFill("solid", fgColor=AZUL)
    amarelo = PatternFill("solid", fgColor=AMARELO)
    amarelo_claro = PatternFill("solid", fgColor=AMARELO_CLARO)
    fonte_branca = Font(bold=True, color=BRANCO, size=14)
    fonte_azul = Font(bold=True, color=AZUL)
    borda = Border(*(Side(style="thin", color=AZUL),) * 4)

    wb = openpyxl.Workbook()

    # --- Aba 1: escala do dia ---
    ws = wb.active
    ws.title = "Escala"
    ws.merge_cells("A1:B1")
    ws["A1"] = "ESCALA DE BASES — ATIVO 2"
    ws["A1"].fill = azul
    ws["A1"].font = fonte_branca
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:B2")
    dia_semana = DIAS_DA_SEMANA[data_sorteio.weekday()]
    ws["A2"] = f"{data_sorteio:%d/%m/%Y} ({dia_semana})"
    ws["A2"].font = fonte_azul
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.append([])
    ws.append(["Base", "Responsável"])
    for celula in ws[4]:
        celula.fill = amarelo
        celula.font = fonte_azul
        celula.border = borda
    for base, pessoa in escolha.items():
        ws.append([base, pessoa])
        for celula in ws[ws.max_row]:
            celula.border = borda

    ws.append([])
    ws.append([f"Ausentes: {', '.join(ausentes) if ausentes else 'nenhum'}"])
    ws.append([f"Regra: ninguém repetiu a base dos últimos {janela_usada} sorteio(s)."])
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 22

    # --- Aba 2: últimos 10 sorteios, uma coluna por base (como na .xlsm) ---
    wh = wb.create_sheet("Últimos sorteios")
    bases = list(escolha)
    wh.append(["Data"] + bases)
    for celula in wh[1]:
        celula.fill = azul
        celula.font = Font(bold=True, color=BRANCO)
    datas = sorted({r["data"] for r in historico}, reverse=True)[:10]
    for d in datas:
        do_dia = {normalizar_nome(r["base"]): r["responsavel"] for r in historico if r["data"] == d}
        wh.append([d.strftime("%d/%m/%Y")] + [do_dia.get(normalizar_nome(b), "") for b in bases])
        if d == data_sorteio:
            for celula in wh[wh.max_row]:
                celula.fill = amarelo_claro
    wh.column_dimensions["A"].width = 12
    for i in range(len(bases)):
        wh.column_dimensions[openpyxl.utils.get_column_letter(i + 2)].width = 32

    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho)


# ---------------------------------------------------------------------------
# Importação da planilha antiga (.xlsm)
# ---------------------------------------------------------------------------

def importar_xlsm(caminho_xlsm, arquivo_cadastro=ARQUIVO_CADASTRO, arquivo_historico=ARQUIVO_HISTORICO):
    """
    Lê a planilha Escala_Bases_Ativo2.xlsm (sem alterá-la) e:
      1. cria config/bases.xlsx com as bases e pessoas da aba Cadastro
         (só se ele ainda não existir);
      2. copia para o histórico os dias da aba Escala que ainda não estão lá.
    """
    caminho_xlsm = Path(caminho_xlsm)
    if not caminho_xlsm.exists():
        raise FileNotFoundError(
            f"Não encontrei a planilha {caminho_xlsm}.\n"
            "A pasta config/ não vem junto com o projeto (ela guarda dados e fica fora do GitHub).\n"
            "Crie a pasta config dentro da pasta do projeto, copie a planilha para lá e rode de novo."
        )
    # keep_vba=False: só lemos os valores; as macros não são executadas
    wb = openpyxl.load_workbook(caminho_xlsm, read_only=True, data_only=True)

    # 1. Cadastro: bases na coluna A e pessoas na coluna C, a partir da linha 5
    cadastro = wb["Cadastro"]
    bases, pessoas = [], []
    for linha in cadastro.iter_rows(min_row=5, max_col=3, values_only=True):
        bases.append(linha[0])
        pessoas.append(linha[2])
    bases, pessoas = limpar_lista(bases), limpar_lista(pessoas)

    arquivo_cadastro = Path(arquivo_cadastro)
    if arquivo_cadastro.exists():
        print(f"- {arquivo_cadastro.name} já existe; não foi alterado.")
    else:
        arquivo_cadastro.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(arquivo_cadastro) as escritor:
            pd.DataFrame({"Base": bases}).to_excel(escritor, sheet_name="Bases", index=False)
            pd.DataFrame({"Pessoa": pessoas}).to_excel(escritor, sheet_name="Pessoas", index=False)
        print(f"- Criado {arquivo_cadastro} com {len(bases)} bases e {len(pessoas)} pessoas.")

    # 2. Histórico: procura a linha de cabeçalho "Data | base 1 | base 2 ..."
    escala = wb["Escala"]
    linhas = list(escala.iter_rows(values_only=True))
    inicio = next(i for i, l in enumerate(linhas) if l and l[0] == "Data")
    cabecalho = linhas[inicio]

    historico = ler_historico(arquivo_historico)
    datas_existentes = {r["data"] for r in historico}
    novos = 0
    for linha in linhas[inicio + 1:]:
        if not isinstance(linha[0], datetime):
            continue
        d = linha[0].date()
        if d in datas_existentes:
            continue
        for base, pessoa in zip(cabecalho[1:], linha[1:]):
            if base and pessoa:
                historico.append({"data": d, "base": str(base).strip(), "responsavel": str(pessoa).strip()})
        datas_existentes.add(d)
        novos += 1
    salvar_historico(historico, arquivo_historico)
    print(f"- {novos} dia(s) importado(s) para {arquivo_historico}.")


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def confirmar(pergunta):
    """Pergunta sim/não no terminal. Sem resposta, considera 'não'."""
    try:
        return input(f"{pergunta} (s/N): ").strip().lower() in ("s", "sim")
    except EOFError:
        return False


def main(argumentos=None):
    parser = argparse.ArgumentParser(description="Sorteio diário das bases do Ativo 2.")
    parser.add_argument("--data", help="data do sorteio no formato dd/mm/aaaa (padrão: hoje)")
    parser.add_argument("--ausentes", default="", help='pessoas ausentes, separadas por vírgula. Ex.: "FULANO, BELTRANO"')
    parser.add_argument("--substituir", action="store_true", help="substitui sem perguntar se a data já foi sorteada")
    parser.add_argument("--importar-xlsm", metavar="ARQUIVO", help="importa cadastro e histórico da planilha antiga")
    args = parser.parse_args(argumentos)

    if args.importar_xlsm:
        importar_xlsm(args.importar_xlsm, ARQUIVO_CADASTRO, ARQUIVO_HISTORICO)
        return 0

    try:
        data_sorteio = datetime.strptime(args.data, "%d/%m/%Y").date() if args.data else date.today()
    except ValueError:
        raise ValueError(f"Data inválida: {args.data}. Use o formato dd/mm/aaaa, ex.: 21/10/2026")
    ausentes = [a.strip() for a in args.ausentes.split(",") if a.strip()]

    bases, pessoas = ler_cadastro(ARQUIVO_CADASTRO)
    disponiveis = aplicar_ausencias(pessoas, ausentes)
    # Mostra os ausentes com a grafia do cadastro (ex.: "fulano" -> "FULANO")
    ausentes = [p for p in pessoas if p not in disponiveis]
    historico = ler_historico(ARQUIVO_HISTORICO)

    print(f"\nSORTEIO DE BASES — ATIVO 2 — {data_sorteio:%d/%m/%Y} ({DIAS_DA_SEMANA[data_sorteio.weekday()]})")
    print(f"Bases: {len(bases)} | Pessoas disponíveis: {len(disponiveis)} | Ausentes: {', '.join(ausentes) or 'nenhum'}")

    avisos = []
    if data_sorteio.weekday() >= 5:
        avisos.append("a data escolhida cai no fim de semana.")
    feriado = eh_feriado(data_sorteio, ler_feriados())
    if feriado:
        avisos.append(f"a data escolhida é feriado ({feriado}).")
    if any(r["data"] == data_sorteio for r in historico):
        if not args.substituir and not confirmar(f"Já existe sorteio para {data_sorteio:%d/%m/%Y}. Substituir?"):
            print("Nada foi alterado.")
            return 1
        historico = [r for r in historico if r["data"] != data_sorteio]
    if len(disponiveis) < len(bases):
        avisos.append(f"{len(disponiveis)} pessoa(s) para {len(bases)} base(s): alguém fica com mais de uma base.")

    escolha, janela_usada = sortear(bases, disponiveis, historico, data_sorteio)

    if janela_usada == 0:
        avisos.append("não havia como evitar que alguém repetisse a base do sorteio anterior.")
    elif janela_usada < JANELA_SORTEIOS:
        avisos.append(f"não havia combinação sem repetir base nos últimos {JANELA_SORTEIOS} sorteios; "
                      f"só o último foi respeitado.")
    sem_base = [p for p in disponiveis if p not in escolha.values()]
    if sem_base:
        avisos.append(f"mais pessoas que bases; sem base hoje: {', '.join(sem_base)}. "
                      "Cadastre novas bases em config/bases.xlsx se precisar.")

    # Grava histórico e planilha (arquivos novos; o cadastro não é alterado)
    for base, pessoa in escolha.items():
        historico.append({"data": data_sorteio, "base": base, "responsavel": pessoa})
    salvar_historico(historico, ARQUIVO_HISTORICO)
    saida = PASTA_SAIDA / f"escala_{data_sorteio:%Y-%m-%d}.xlsx"
    gerar_planilha(escolha, data_sorteio, ausentes, janela_usada, historico, saida)

    print()
    largura = max(len(b) for b in escolha)
    for base, pessoa in escolha.items():
        print(f"  {base:<{largura}}  =  {pessoa}")
    print()
    for aviso in avisos:
        print(f"AVISO: {aviso}")
    print(f"Planilha: {saida}")
    print(f"Histórico: {ARQUIVO_HISTORICO}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PermissionError as erro:
        print(f"\nERRO: não consegui gravar {erro.filename}.\n"
              "Se esse arquivo estiver aberto no Excel, feche-o e rode de novo.")
        sys.exit(1)
    except (FileNotFoundError, ValueError) as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
