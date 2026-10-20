"""
Baixa da API do Argus os dados do relatório da URA e salva em entrada/.

O que é baixado (referência: ONTEM, ou a data informada):
  - Desempenho dos operadores (TMA e TME) ......... entrada/argus_desempenho_AAAA-MM-DD.csv
  - Leads atendidos do grupo RECEPTIVO - URA ....... entrada/argus_leads_ura_AAAA-MM-DD.csv
  - Cadastro de operadores (nome, grupo, supervisor) entrada/argus_operadores_AAAA-MM-DD.csv

Os arquivos têm as mesmas colunas das exportações manuais do Argus, então o
relatório funciona igual com os dois.

O que NÃO dá para baixar: a TARIFAÇÃO (custos VENDITORE e 4NET) não existe na
API. Ela continua sendo exportada à mão (Argus > Tarifação > Painel de Visualização).

O token da API fica SÓ no seu computador, em config/argus_token.txt
(ou na variável de ambiente ARGUS_TOKEN). Nunca envie esse arquivo para ninguém.

Como usar:
  python argus_api.py                    -> baixa os dados de ontem
  python argus_api.py --data 17/10/2026  -> baixa os dados de outro dia
  python argus_api.py --diagnostico      -> mostra os grupos e os NOMES dos campos que a API devolve
"""

import argparse
import csv
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

from leitura import normalizar_nome

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------

PASTA_PROJETO = Path(__file__).resolve().parent
ARQUIVO_TOKEN = PASTA_PROJETO / "config" / "argus_token.txt"
PASTA_ENTRADA = PASTA_PROJETO / "entrada"

URL_BASE = "https://argus.app.br/apiargus"
FUSO = "-03:00"  # horário de Brasília

# Nome do grupo operacional da URA (comparado sem acentos e sem maiúsculas)
GRUPO_URA = "RECEPTIVO - URA"

# Como reconhecer cada informação no relatório de desempenho.
# A documentação do Argus não lista os nomes dos campos, então primeiro
# procuramos um destes nomes exatos; se nenhum existir, procuramos um campo
# que CONTENHA todos os pedaços da lista "contem" e nenhum da lista "evita".
CAMPOS_DESEMPENHO = {
    "operador": {"exatos": ["nomeUsuario", "nomeOperador", "usuario", "operador", "nome"],
                 "contem": ["nome"], "evita": ["supervisor"]},
    "supervisor": {"exatos": ["nomeUsuarioSupervisor", "supervisor", "nomeSupervisor"],
                   "contem": ["supervisor"]},
    "tma": {"exatos": ["tma", "tempoMedioAtendimento"], "contem": ["tma"]},
    "tme": {"exatos": ["tme", "tempoMedioEspera", "tmo"], "contem": ["tme"]},
    "leads": {"exatos": ["leadsAtendidos", "qtdeLeadsAtendidos", "leadsAtend"],
              "contem": ["lead", "atend"]},
}


class ErroArgus(Exception):
    """Erro devolvido pela API (token inválido, parâmetro errado etc.)."""


# ---------------------------------------------------------------------------
# Comunicação com a API
# ---------------------------------------------------------------------------

def ler_token(arquivo=ARQUIVO_TOKEN):
    """Lê o token da variável ARGUS_TOKEN ou do arquivo config/argus_token.txt."""
    token = os.environ.get("ARGUS_TOKEN", "").strip()
    if not token and Path(arquivo).exists():
        token = Path(arquivo).read_text(encoding="utf-8").strip()
    if not token:
        raise ErroArgus(
            "Token da API não encontrado.\n"
            "1. No Argus, abra Configurações > Desenvolvedor > Acesso API e copie a chave.\n"
            f"2. Crie o arquivo {arquivo} e cole a chave dentro dele (só a chave, nada mais)."
        )
    return token


def chamar_api(caminho, corpo, token):
    """
    Faz uma chamada POST à API do Argus e devolve a resposta (dicionário).

    `caminho` é a parte final do endereço, ex.: "report/desempenhoresumido".
    """
    requisicao = urllib.request.Request(
        f"{URL_BASE}/{caminho}",
        data=json.dumps(corpo).encode("utf-8"),
        headers={"Token-Signature": token, "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(requisicao, timeout=120) as resposta:
            dados = json.loads(resposta.read().decode("utf-8"))
    except urllib.error.HTTPError as erro:
        if erro.code == 403:
            raise ErroArgus("O Argus recusou o acesso (403). Confira se o token está correto.")
        raise ErroArgus(f"O Argus respondeu com erro {erro.code} em {caminho}.")
    except urllib.error.URLError as erro:
        raise ErroArgus(f"Não consegui conectar ao Argus ({erro.reason}). Confira a internet.")

    # codStatus 1 = sucesso; menor que 1 = falha
    if dados.get("codStatus", 1) < 1:
        raise ErroArgus(f"O Argus devolveu uma falha em {caminho}: {dados.get('descStatus')}")
    return dados


def periodo_do_dia(dia):
    """Início e fim do dia no formato que a API pede (ISO 8601 com fuso)."""
    return f"{dia:%Y-%m-%d}T00:00:00{FUSO}", f"{dia:%Y-%m-%d}T23:59:59{FUSO}"


def buscar_desempenho(dia, token, id_grupo=None):
    """Relatório de desempenho resumido dos operadores em um dia."""
    inicio, fim = periodo_do_dia(dia)
    corpo = {"periodoInicial": inicio, "periodoFinal": fim}
    if id_grupo is not None:
        corpo["idGrupoUsuario"] = id_grupo
    dados = chamar_api("report/desempenhoresumido", corpo, token)
    if dados.get("endOfTable") is False:
        print("AVISO: o Argus avisou que o desempenho veio incompleto (endOfTable = false).")
    return dados.get("desempenhosResumidos") or []


def buscar_grupos(token):
    """Lista de grupos operacionais (idTipoGrupo 1)."""
    return chamar_api("cmd/listargrupos", {"idTipoGrupo": 1}, token).get("grupos") or []


def buscar_usuarios(token):
    """Lista de operadores (idTipoUsuario 2), ativos e inativos."""
    return chamar_api("cmd/listarusuarios", {"idTipoUsuario": 2}, token).get("usuarios") or []


# ---------------------------------------------------------------------------
# Tratamento dos dados
# ---------------------------------------------------------------------------

def achar_campo(chaves, informacao):
    """
    Descobre qual campo do Argus traz a informação pedida (ex.: "tma").

    Devolve o nome do campo ou None se não encontrar.
    """
    regra = CAMPOS_DESEMPENHO[informacao]
    por_minusculo = {c.lower(): c for c in chaves}
    for nome in regra["exatos"]:
        if nome.lower() in por_minusculo:
            return por_minusculo[nome.lower()]
    for chave in chaves:
        tem_tudo = all(pedaco in chave.lower() for pedaco in regra["contem"])
        tem_proibido = any(pedaco in chave.lower() for pedaco in regra.get("evita", []))
        if tem_tudo and not tem_proibido:
            return chave
    return None


def mapear_campos(registros, informacoes):
    """
    Monta {informação: campo do Argus} para a lista de informações pedida.

    Dá erro explicando quais campos existem, se alguma informação não for encontrada.
    """
    chaves = list(registros[0].keys()) if registros else []
    mapa = {info: achar_campo(chaves, info) for info in informacoes}
    faltando = [info for info, campo in mapa.items() if campo is None]
    if faltando:
        raise ErroArgus(
            f"Não encontrei no desempenho do Argus: {', '.join(faltando)}.\n"
            f"Campos que vieram: {', '.join(chaves) or '(nenhum registro)'}\n"
            "Ajuste CAMPOS_DESEMPENHO no início do argus_api.py."
        )
    return mapa


def formatar_tempo(valor):
    """
    Deixa o tempo no formato hh:mm:ss, igual à exportação manual.

    A API pode mandar segundos (64) ou o texto pronto ("00:01:04").
    """
    if valor is None or valor == "":
        return ""
    if isinstance(valor, str) and ":" in valor:
        return valor
    segundos = int(round(float(valor)))
    return f"{segundos // 3600:02d}:{segundos % 3600 // 60:02d}:{segundos % 60:02d}"


def supervisores_por_operador(usuarios):
    """{NOME NORMALIZADO: nome do supervisor}, a partir do cadastro de usuários."""
    mapa = {}
    for usuario in usuarios:
        supervisores = [c.get("nomeUsuarioSupervisor") for c in usuario.get("campanhas") or []]
        supervisores = [s for s in supervisores if s]
        if supervisores:
            mapa[normalizar_nome(usuario.get("nome"))] = supervisores[0]
    return mapa


def ids_grupo_ura(grupos):
    """Códigos dos grupos operacionais cujo nome contém RECEPTIVO - URA."""
    alvo = normalizar_nome(GRUPO_URA)
    return [g["idGrupoUsuario"] for g in grupos if alvo in normalizar_nome(g.get("grupoUsuarioDesc"))]


def gravar_csv(caminho, cabecalho, linhas):
    """Grava um CSV com ; e UTF-8 com BOM (abre direto no Excel)."""
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho, "w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";", quoting=csv.QUOTE_ALL)
        escritor.writerow(cabecalho)
        escritor.writerows(linhas)


def linhas_desempenho(registros, dia, supervisores):
    """Linhas no formato Data; Operador; Supervisor; TME; TMA."""
    if not registros:
        return []
    campos = mapear_campos(registros, ["operador", "tma", "tme"])
    campo_sup = achar_campo(list(registros[0].keys()), "supervisor")
    linhas = []
    for r in registros:
        operador = r.get(campos["operador"]) or ""
        supervisor = (r.get(campo_sup) if campo_sup else None) or supervisores.get(normalizar_nome(operador), "")
        linhas.append([f"{dia:%d/%m/%Y}", operador, supervisor,
                       formatar_tempo(r.get(campos["tme"])), formatar_tempo(r.get(campos["tma"]))])
    return linhas


def linhas_leads(registros_por_grupo, dia, supervisores):
    """
    Linhas no formato Data; Operador; Supervisor; Leads. Atend.

    Se houver mais de um grupo URA, os leads do mesmo operador são somados.
    """
    todos = [r for registros in registros_por_grupo for r in registros]
    if not todos:
        return []
    campos = mapear_campos(todos, ["operador", "leads"])
    campo_sup = achar_campo(list(todos[0].keys()), "supervisor")
    somas, nomes, sups = {}, {}, {}
    for r in todos:
        operador = r.get(campos["operador"]) or ""
        chave = normalizar_nome(operador)
        nomes.setdefault(chave, operador)
        sups.setdefault(chave, (r.get(campo_sup) if campo_sup else None) or supervisores.get(chave, ""))
        somas[chave] = somas.get(chave, 0) + int(r.get(campos["leads"]) or 0)
    return [[f"{dia:%d/%m/%Y}", nomes[c], sups[c], somas[c]] for c in nomes]


def linhas_operadores(usuarios):
    """
    Cadastro de operadores com as colunas da exportação manual.

    CPF, e-mail, login, RE e ramal que a API devolve são DESCARTADOS aqui.
    """
    linhas = []
    for u in usuarios:
        campanhas = u.get("campanhas") or []
        grupos = ", ".join(c.get("grupoUsuarioDesc") or "" for c in campanhas if c.get("grupoUsuarioDesc"))
        supervisores = ", ".join(c.get("nomeUsuarioSupervisor") or "" for c in campanhas if c.get("nomeUsuarioSupervisor"))
        status = "ATIVO" if u.get("ativo") else "INATIVO"
        linhas.append([u.get("idUsuario", ""), u.get("nome", ""), grupos, supervisores, status])
    return linhas


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def diagnostico(dia, token):
    """Mostra os grupos e os NOMES e TIPOS dos campos, sem mostrar os dados."""
    grupos = buscar_grupos(token)
    print("\nGRUPOS OPERACIONAIS:")
    for g in grupos:
        print(f"  {g.get('idGrupoUsuario')}: {g.get('grupoUsuarioDesc')} ({g.get('campanhaDesc')})")
    print(f"Grupo(s) da URA encontrado(s): {ids_grupo_ura(grupos) or 'NENHUM'}")

    registros = buscar_desempenho(dia, token)
    print(f"\nDESEMPENHO de {dia:%d/%m/%Y}: {len(registros)} registro(s)")
    if registros:
        print("Campos (nome: tipo):")
        for chave, valor in registros[0].items():
            print(f"  {chave}: {type(valor).__name__}")
        for info in CAMPOS_DESEMPENHO:
            print(f"  -> {info} será lido do campo: {achar_campo(list(registros[0].keys()), info)}")


def baixar(dia, token, pasta=PASTA_ENTRADA):
    """Baixa tudo e grava os 3 CSVs. Devolve a lista de arquivos gravados."""
    usuarios = buscar_usuarios(token)
    supervisores = supervisores_por_operador(usuarios)
    grupos_ura = ids_grupo_ura(buscar_grupos(token))
    if not grupos_ura:
        raise ErroArgus(f"Não encontrei o grupo operacional '{GRUPO_URA}' no Argus.")

    desempenho = buscar_desempenho(dia, token)
    leads = [buscar_desempenho(dia, token, id_grupo) for id_grupo in grupos_ura]

    pasta = Path(pasta)
    arquivos = {
        "desempenho": pasta / f"argus_desempenho_{dia:%Y-%m-%d}.csv",
        "leads": pasta / f"argus_leads_ura_{dia:%Y-%m-%d}.csv",
        "operadores": pasta / f"argus_operadores_{dia:%Y-%m-%d}.csv",
    }
    linhas_d = linhas_desempenho(desempenho, dia, supervisores)
    linhas_l = linhas_leads(leads, dia, supervisores)
    linhas_o = linhas_operadores(usuarios)
    gravar_csv(arquivos["desempenho"], ["Data", "Operador", "Supervisor", "TME", "TMA"], linhas_d)
    gravar_csv(arquivos["leads"], ["Data", "Operador", "Supervisor", "Leads. Atend."], linhas_l)
    gravar_csv(arquivos["operadores"], ["Cód. Usuário", "Nome", "[Campanha] Grupo Associado",
                                        "Supervisor(es)", "Status"], linhas_o)

    print(f"\nARGUS — dados de {dia:%d/%m/%Y}")
    print(f"- Desempenho (TMA/TME): {len(linhas_d)} operador(es) -> {arquivos['desempenho']}")
    print(f"- Leads URA: {len(linhas_l)} operador(es), {sum(l[3] for l in linhas_l)} leads -> {arquivos['leads']}")
    print(f"- Operadores: {len(linhas_o)} cadastrado(s) -> {arquivos['operadores']}")
    if desempenho:
        mapa = mapear_campos(desempenho, ["operador", "tma", "tme"])
        print(f"  (campos usados: operador={mapa['operador']}, TMA={mapa['tma']}, TME={mapa['tme']})")
    print("\nLembrete: a TARIFAÇÃO (VENDITORE e 4NET) não está na API; exporte à mão para entrada/.")
    return list(arquivos.values())


def main(argumentos=None):
    parser = argparse.ArgumentParser(description="Baixa os dados do Argus para o relatório da URA.")
    parser.add_argument("--data", help="dia de referência dd/mm/aaaa (padrão: ontem)")
    parser.add_argument("--diagnostico", action="store_true", help="mostra grupos e nomes dos campos, sem gravar nada")
    args = parser.parse_args(argumentos)

    try:
        dia = datetime.strptime(args.data, "%d/%m/%Y").date() if args.data else date.today() - timedelta(days=1)
    except ValueError:
        raise ErroArgus(f"Data inválida: {args.data}. Use o formato dd/mm/aaaa, ex.: 17/10/2026")

    token = ler_token(ARQUIVO_TOKEN)
    if args.diagnostico:
        diagnostico(dia, token)
    else:
        baixar(dia, token, PASTA_ENTRADA)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ErroArgus as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
