"""
Baixa da API do Argus os dados do relatório da URA e salva em entrada/.

O que é baixado (referência: o último dia útil, ou a data informada):
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
  python argus_api.py                    -> baixa os dados do último dia útil
                                            (ontem; na segunda-feira, a sexta)
  python argus_api.py --data 17/10/2026  -> baixa os dados de outro dia
  python argus_api.py --diagnostico      -> mostra os grupos e os NOMES dos campos que a API devolve
  python argus_api.py --bruto            -> grava as linhas do desempenho como vieram (conferência)
"""

import argparse
import csv
import json
import os
import ssl
import sys
import urllib.error
import urllib.request
import warnings
from datetime import date, datetime, timedelta
from pathlib import Path

from leitura import normalizar_nome

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------

PASTA_PROJETO = Path(__file__).resolve().parent
ARQUIVO_TOKEN = PASTA_PROJETO / "config" / "argus_token.txt"
# Opcional: se a Benconsig usar outro endereço da API, escreva-o neste arquivo
# (ex.: https://argus4com.com.br:23001/apiargus)
ARQUIVO_URL = PASTA_PROJETO / "config" / "argus_url.txt"
PASTA_ENTRADA = PASTA_PROJETO / "entrada"

URL_PADRAO = "https://argus.app.br/apiargus"
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
    # No grupo RECEPTIVO - URA, cada atendimento é um lead atendido
    "leads": {"exatos": ["leadsAtendidos", "qtdeLeadsAtendidos", "leadsAtend", "qtdeAtendimentoTotal"],
              "contem": ["lead", "atend"]},
    # Quantidade de atendimentos: usada como peso ao juntar as linhas do mesmo operador
    "atendimentos": {"exatos": ["qtdeAtendimentoTotal"], "contem": ["qtde", "atendimento", "total"]},
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
        bruto = Path(arquivo).read_bytes()
        # O PowerShell pode gravar em UTF-16 e o Bloco de Notas antigo coloca um
        # caractere invisível (BOM) no início; os dois casos são tratados aqui.
        if bruto.startswith((b"\xff\xfe", b"\xfe\xff")):
            token = bruto.decode("utf-16").strip()
        else:
            token = bruto.decode("utf-8-sig").strip()
    if not token:
        raise ErroArgus(
            "Token da API não encontrado.\n"
            "1. No Argus, abra Configurações > Desenvolvedor > Acesso API e copie a chave.\n"
            f"2. Crie o arquivo {arquivo} e cole a chave dentro dele (só a chave, nada mais)."
        )
    return token


def url_base(arquivo=ARQUIVO_URL):
    """Endereço da API: o de config/argus_url.txt, se existir; senão, o padrão."""
    if Path(arquivo).exists():
        url = Path(arquivo).read_bytes().decode("utf-8-sig").strip()
        if url:
            return url.rstrip("/")
    return URL_PADRAO


def contexto_compativel():
    """
    Conexão segura aceitando servidores com configuração antiga.

    O Python recente recusa algumas cifras e versões antigas do TLS que certos
    servidores ainda usam. Aqui elas são liberadas, mas o certificado do
    servidor CONTINUA sendo verificado (a conexão segue segura contra impostores).
    """
    contexto = ssl.create_default_context()
    contexto.set_ciphers("DEFAULT:@SECLEVEL=0")
    with warnings.catch_warnings():  # o TLS 1.0 é obsoleto, mas alguns servidores ainda usam
        warnings.simplefilter("ignore", DeprecationWarning)
        contexto.minimum_version = ssl.TLSVersion.TLSv1
    # Aceita servidores sem "renegociação segura" (OP_LEGACY_SERVER_CONNECT)
    contexto.options |= getattr(ssl, "OP_LEGACY_SERVER_CONNECT", 0x4)
    return contexto


def abrir(requisicao, contexto):
    """Envia a requisição e devolve o JSON da resposta."""
    with urllib.request.urlopen(requisicao, timeout=120, context=contexto) as resposta:
        return json.loads(resposta.read().decode("utf-8"))


def chamar_api(caminho, corpo, token):
    """
    Faz uma chamada POST à API do Argus e devolve a resposta (dicionário).

    `caminho` é a parte final do endereço, ex.: "report/desempenhoresumido".
    """
    endereco = f"{url_base()}/{caminho}"
    requisicao = urllib.request.Request(
        endereco,
        data=json.dumps(corpo).encode("utf-8"),
        headers={"Token-Signature": token, "Content-Type": "application/json"},
        method="POST",
    )
    try:
        try:
            dados = abrir(requisicao, ssl.create_default_context())
        except urllib.error.URLError as erro:
            # Falha na conexão segura: tenta de novo no modo compatível
            if not isinstance(erro.reason, ssl.SSLError) or isinstance(erro, urllib.error.HTTPError):
                raise
            dados = abrir(requisicao, contexto_compativel())
    except urllib.error.HTTPError as erro:
        if erro.code == 403:
            raise ErroArgus("O Argus recusou o acesso (403). Confira se o token está correto.")
        raise ErroArgus(f"O Argus respondeu com erro {erro.code} em {endereco}.")
    except urllib.error.URLError as erro:
        raise ErroArgus(
            f"Não consegui conectar ao Argus em {endereco}\n({erro.reason}).\n"
            "Se o endereço da API da Benconsig for outro, escreva-o em config\\argus_url.txt."
        )

    # codStatus 1 = sucesso; menor que 1 = falha
    if dados.get("codStatus", 1) < 1:
        raise ErroArgus(f"O Argus devolveu uma falha em {caminho}: {dados.get('descStatus')}")
    return dados


def dia_util_anterior(hoje):
    """
    Último dia útil antes de `hoje` (segunda a sexta).

    Segunda-feira -> sexta; sábado e domingo -> sexta; outros dias -> ontem.
    Feriados não são considerados: nesses casos use --data.
    """
    dia = hoje - timedelta(days=1)
    while dia.weekday() >= 5:  # 5 = sábado, 6 = domingo
        dia -= timedelta(days=1)
    return dia


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


def numero(valor):
    """Converte o valor da API em número (vazio vira 0; "hh:mm:ss" vira segundos)."""
    if isinstance(valor, str) and valor.count(":") == 2:
        horas, minutos, segundos = (int(p) for p in valor.split(":"))
        return float(horas * 3600 + minutos * 60 + segundos)
    try:
        return float(valor or 0)
    except (TypeError, ValueError):
        return 0.0


def media_ponderada(pares):
    """
    Média de (valor, peso). Ex.: [(60, 10), (30, 30)] -> 37.5

    Se nenhum peso for maior que zero, faz a média simples dos valores.
    """
    total_peso = sum(peso for _, peso in pares)
    if total_peso > 0:
        return sum(valor * peso for valor, peso in pares) / total_peso
    return sum(valor for valor, _ in pares) / len(pares) if pares else 0


def agrupar_por_operador(registros, campo_operador):
    """
    O Argus devolve várias linhas por operador (uma por sessão/grupo).
    Junta em {NOME NORMALIZADO: (nome original, [linhas])}, na ordem em que aparecem.
    """
    grupos = {}
    for r in registros:
        nome = r.get(campo_operador) or ""
        chave = normalizar_nome(nome)
        grupos.setdefault(chave, (nome, []))[1].append(r)
    return grupos


def linhas_desempenho(registros, dia, supervisores):
    """
    Linhas no formato Data; Operador; Supervisor; TME; TMA (uma por operador).

    TMA e TME de cada operador = média das linhas dele, ponderada pela
    quantidade de atendimentos de cada linha.
    """
    if not registros:
        return []
    campos = mapear_campos(registros, ["operador", "tma", "tme"])
    chaves = list(registros[0].keys())
    campo_sup = achar_campo(chaves, "supervisor")
    campo_peso = achar_campo(chaves, "atendimentos")
    linhas = []
    for chave, (nome, linhas_op) in agrupar_por_operador(registros, campos["operador"]).items():
        peso = [numero(r.get(campo_peso)) if campo_peso else 1 for r in linhas_op]
        tma = media_ponderada([(numero(r.get(campos["tma"])), p) for r, p in zip(linhas_op, peso)])
        tme = media_ponderada([(numero(r.get(campos["tme"])), p) for r, p in zip(linhas_op, peso)])
        supervisor = (linhas_op[0].get(campo_sup) if campo_sup else None) or supervisores.get(chave, "")
        linhas.append([f"{dia:%d/%m/%Y}", nome, supervisor, formatar_tempo(tme), formatar_tempo(tma)])
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


def gravar_bruto(dia, token, pasta=PASTA_ENTRADA):
    """
    Grava as linhas do desempenho exatamente como o Argus devolve (para conferência).

    Não tem dados de clientes: só operadores, tempos e quantidades.
    """
    grupos_ura = ids_grupo_ura(buscar_grupos(token))
    arquivos = []
    for nome, registros in [("geral", buscar_desempenho(dia, token))] + \
            [(f"grupo{g}", buscar_desempenho(dia, token, g)) for g in grupos_ura]:
        if not registros:
            continue
        caminho = Path(pasta) / f"argus_bruto_{nome}_{dia:%Y-%m-%d}.csv"
        cabecalho = list(registros[0].keys())
        gravar_csv(caminho, cabecalho, [[r.get(c, "") for c in cabecalho] for r in registros])
        print(f"- {len(registros)} linha(s) -> {caminho}")
        arquivos.append(caminho)
    return arquivos


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
    parser.add_argument("--data", help="dia de referência dd/mm/aaaa (padrão: último dia útil)")
    parser.add_argument("--diagnostico", action="store_true", help="mostra grupos e nomes dos campos, sem gravar nada")
    parser.add_argument("--bruto", action="store_true", help="grava as linhas do desempenho como vieram do Argus (conferência)")
    args = parser.parse_args(argumentos)

    try:
        dia = datetime.strptime(args.data, "%d/%m/%Y").date() if args.data else dia_util_anterior(date.today())
    except ValueError:
        raise ErroArgus(f"Data inválida: {args.data}. Use o formato dd/mm/aaaa, ex.: 17/10/2026")

    token = ler_token(ARQUIVO_TOKEN)
    if args.diagnostico:
        diagnostico(dia, token)
    elif args.bruto:
        gravar_bruto(dia, token, PASTA_ENTRADA)
    else:
        baixar(dia, token, PASTA_ENTRADA)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ErroArgus as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
