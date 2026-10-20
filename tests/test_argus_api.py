"""
Testes do argus_api.py.

A API é SIMULADA: nenhuma chamada sai para a internet. Nomes, CPFs e e-mails
abaixo são inventados.
"""

import csv
from datetime import date

import pytest

import argus_api as ap

DIA = date(2026, 10, 19)

USUARIOS = [
    {"idUsuario": 1, "nome": "ANA TESTE", "login": "ANA.T", "ramal": "100", "re": "9",
     "cpf": "12345678901", "email": "ana@exemplo.com", "ativo": True, "idTipoUsuario": 2,
     "campanhas": [{"idCampanha": 7, "campanhaDesc": "BENC", "idGrupoUsuario": 50,
                    "grupoUsuarioDesc": "[BENC] RECEPTIVO - URA", "idUsuarioSupervisor": 3,
                    "nomeUsuarioSupervisor": "SUPERVISORA UM"}]},
    {"idUsuario": 2, "nome": "BRUNO TESTE", "cpf": "10987654321", "email": "b@exemplo.com",
     "ativo": False, "campanhas": [{"grupoUsuarioDesc": "[BENC] DIGITAL",
                                    "nomeUsuarioSupervisor": "SUPERVISOR DOIS"}]},
]
GRUPOS = [
    {"idGrupoUsuario": 50, "grupoUsuarioDesc": "[BENC] RECEPTIVO - URA", "campanhaDesc": "BENC"},
    {"idGrupoUsuario": 51, "grupoUsuarioDesc": "[BENC] DIGITAL", "campanhaDesc": "BENC"},
]
# Campos com nomes "inventados" para testar o reconhecimento automático
DESEMPENHO = [
    {"nomeUsuario": "ANA TESTE", "tma": 56, "tme": 47, "qtdeLeadsAtendidos": 0},
    {"nomeUsuario": "BRUNO TESTE", "tma": "00:01:30", "tme": 3725, "qtdeLeadsAtendidos": 0},
]
DESEMPENHO_URA = [{"nomeUsuario": "ANA TESTE", "tma": 50, "tme": 40, "qtdeLeadsAtendidos": 17}]


@pytest.fixture
def api_falsa(monkeypatch):
    """Substitui chamar_api por respostas fixas e guarda as chamadas feitas."""
    chamadas = []

    def falsa(caminho, corpo, token):
        chamadas.append((caminho, corpo))
        if caminho == "cmd/listarusuarios":
            return {"codStatus": 1, "usuarios": USUARIOS}
        if caminho == "cmd/listargrupos":
            return {"codStatus": 1, "grupos": GRUPOS}
        if caminho == "report/desempenhoresumido":
            lista = DESEMPENHO_URA if corpo.get("idGrupoUsuario") == 50 else DESEMPENHO
            return {"codStatus": 1, "endOfTable": True, "desempenhosResumidos": lista}
        raise AssertionError(caminho)

    monkeypatch.setattr(ap, "chamar_api", falsa)
    return chamadas


def ler(caminho):
    with open(caminho, encoding="utf-8-sig", newline="") as arquivo:
        return list(csv.reader(arquivo, delimiter=";"))


# --- Funções pequenas -------------------------------------------------------

def test_periodo_de_ontem_no_formato_iso_com_fuso():
    assert ap.periodo_do_dia(DIA) == ("2026-10-19T00:00:00-03:00", "2026-10-19T23:59:59-03:00")


@pytest.mark.parametrize("valor, esperado", [
    (64, "00:01:04"), (3725, "01:02:05"), (59.6, "00:01:00"), ("00:00:47", "00:00:47"),
    (None, ""), ("", ""),
])
def test_formatar_tempo(valor, esperado):
    assert ap.formatar_tempo(valor) == esperado


def test_reconhece_campos_por_nome_exato_ou_parcial():
    chaves = ["nomeUsuario", "nomeUsuarioSupervisor", "tmaSegundos", "tmeSegundos", "qtdeLeadsAtendidos"]
    assert ap.achar_campo(chaves, "operador") == "nomeUsuario"
    assert ap.achar_campo(chaves, "supervisor") == "nomeUsuarioSupervisor"
    assert ap.achar_campo(chaves, "tma") == "tmaSegundos"
    assert ap.achar_campo(chaves, "leads") == "qtdeLeadsAtendidos"


def test_nome_do_operador_nao_confunde_com_supervisor():
    assert ap.achar_campo(["nomeUsuarioSupervisor", "nomeAgente"], "operador") == "nomeAgente"


def test_campo_inexistente_explica_quais_vieram():
    with pytest.raises(ap.ErroArgus, match="Campos que vieram: nome, outro"):
        ap.mapear_campos([{"nome": "X", "outro": 1}], ["tma"])


def test_acha_grupo_da_ura_sem_diferenciar_acentos():
    assert ap.ids_grupo_ura(GRUPOS + [{"idGrupoUsuario": 9, "grupoUsuarioDesc": "Receptivo - URA 2"}]) == [50, 9]


def test_token_ausente_explica_onde_pegar(tmp_path, monkeypatch):
    monkeypatch.delenv("ARGUS_TOKEN", raising=False)
    with pytest.raises(ap.ErroArgus, match="Acesso API"):
        ap.ler_token(tmp_path / "nao_existe.txt")


def test_token_lido_do_arquivo_sem_espacos(tmp_path, monkeypatch):
    monkeypatch.delenv("ARGUS_TOKEN", raising=False)
    arquivo = tmp_path / "argus_token.txt"
    arquivo.write_text("  abc123 \n", encoding="utf-8")
    assert ap.ler_token(arquivo) == "abc123"


@pytest.mark.parametrize("codificacao", ["utf-8", "utf-8-sig", "utf-16"])
def test_token_gravado_pelo_bloco_de_notas_ou_powershell(tmp_path, monkeypatch, codificacao):
    monkeypatch.delenv("ARGUS_TOKEN", raising=False)
    arquivo = tmp_path / "argus_token.txt"
    arquivo.write_text("abc123\r\n", encoding=codificacao)
    assert ap.ler_token(arquivo) == "abc123"


def test_leads_de_varios_grupos_ura_sao_somados():
    linhas = ap.linhas_leads(
        [[{"nomeUsuario": "ANA", "leadsAtendidos": 10}], [{"nomeUsuario": "Ána", "leadsAtendidos": 5}]],
        DIA, {"ANA": "SUP"})
    assert linhas == [["19/10/2026", "ANA", "SUP", 15]]


# --- Download completo (API simulada) ----------------------------------------

def test_baixar_gera_csvs_no_formato_da_exportacao_manual(api_falsa, tmp_path):
    arquivos = ap.baixar(DIA, "token-falso", tmp_path)

    desempenho, leads, operadores = (ler(a) for a in arquivos)
    assert desempenho[0] == ["Data", "Operador", "Supervisor", "TME", "TMA"]
    assert desempenho[1] == ["19/10/2026", "ANA TESTE", "SUPERVISORA UM", "00:00:47", "00:00:56"]
    assert desempenho[2] == ["19/10/2026", "BRUNO TESTE", "SUPERVISOR DOIS", "01:02:05", "00:01:30"]

    assert leads == [["Data", "Operador", "Supervisor", "Leads. Atend."],
                     ["19/10/2026", "ANA TESTE", "SUPERVISORA UM", "17"]]

    assert operadores[0] == ["Cód. Usuário", "Nome", "[Campanha] Grupo Associado", "Supervisor(es)", "Status"]
    assert operadores[1] == ["1", "ANA TESTE", "[BENC] RECEPTIVO - URA", "SUPERVISORA UM", "ATIVO"]
    assert operadores[2][-1] == "INATIVO"


def test_cpf_e_email_dos_operadores_nao_sao_gravados(api_falsa, tmp_path):
    for arquivo in ap.baixar(DIA, "token-falso", tmp_path):
        texto = arquivo.read_text(encoding="utf-8-sig")
        assert "12345678901" not in texto
        assert "@" not in texto


def test_leads_sao_pedidos_filtrando_pelo_grupo_da_ura(api_falsa, tmp_path):
    ap.baixar(DIA, "token-falso", tmp_path)
    corpos = [corpo for caminho, corpo in api_falsa if caminho == "report/desempenhoresumido"]
    assert {"periodoInicial": "2026-10-19T00:00:00-03:00", "periodoFinal": "2026-10-19T23:59:59-03:00",
            "idGrupoUsuario": 50} in corpos


def test_sem_grupo_ura_da_erro(api_falsa, tmp_path, monkeypatch):
    monkeypatch.setattr(ap, "GRUPO_URA", "GRUPO QUE NAO EXISTE")
    with pytest.raises(ap.ErroArgus, match="GRUPO QUE NAO EXISTE"):
        ap.baixar(DIA, "token-falso", tmp_path)


def test_diagnostico_mostra_nomes_dos_campos_e_nao_os_dados(api_falsa, capsys):
    ap.diagnostico(DIA, "token-falso")
    saida = capsys.readouterr().out
    assert "tma: int" in saida
    assert "-> leads será lido do campo: qtdeLeadsAtendidos" in saida
    assert "ANA TESTE" not in saida


@pytest.mark.parametrize("hoje, esperado", [
    (date(2026, 10, 19), date(2026, 10, 16)),  # segunda -> sexta
    (date(2026, 10, 20), date(2026, 10, 19)),  # terça -> segunda
    (date(2026, 10, 24), date(2026, 10, 23)),  # sábado -> sexta
    (date(2026, 10, 25), date(2026, 10, 23)),  # domingo -> sexta
])
def test_dia_util_anterior(hoje, esperado):
    assert ap.dia_util_anterior(hoje) == esperado


def test_ultimo_dia_util_e_o_padrao(api_falsa, tmp_path, monkeypatch):
    monkeypatch.setattr(ap, "ARQUIVO_TOKEN", tmp_path / "t.txt")
    monkeypatch.setattr(ap, "PASTA_ENTRADA", tmp_path / "entrada")
    monkeypatch.setenv("ARGUS_TOKEN", "token-falso")
    ap.main([])
    esperado = ap.dia_util_anterior(date.today()).strftime("%Y-%m-%d")
    assert any(esperado in p.name for p in (tmp_path / "entrada").iterdir())


# --- Conexão (sem internet: a abertura da conexão é simulada) ----------------

def test_endereco_padrao_e_endereco_do_arquivo(tmp_path):
    arquivo = tmp_path / "argus_url.txt"
    assert ap.url_base(arquivo) == "https://argus.app.br/apiargus"
    arquivo.write_text("https://argus4com.com.br:23001/apiargus/\n", encoding="utf-8")
    assert ap.url_base(arquivo) == "https://argus4com.com.br:23001/apiargus"


def test_falha_ssl_tenta_de_novo_no_modo_compativel(monkeypatch):
    contextos = []

    def abrir_falso(requisicao, contexto):
        contextos.append(contexto)
        if len(contextos) == 1:
            raise ap.urllib.error.URLError(ap.ssl.SSLError("SSLV3_ALERT_HANDSHAKE_FAILURE"))
        return {"codStatus": 1, "grupos": []}

    monkeypatch.setattr(ap, "abrir", abrir_falso)
    assert ap.chamar_api("cmd/listargrupos", {}, "token-falso") == {"codStatus": 1, "grupos": []}
    assert len(contextos) == 2
    assert contextos[1].verify_mode == ap.ssl.CERT_REQUIRED  # certificado continua verificado


def test_sem_internet_nao_tenta_de_novo_e_explica(monkeypatch):
    chamadas = []

    def abrir_falso(requisicao, contexto):
        chamadas.append(1)
        raise ap.urllib.error.URLError("sem rede")

    monkeypatch.setattr(ap, "abrir", abrir_falso)
    with pytest.raises(ap.ErroArgus, match="argus_url.txt"):
        ap.chamar_api("cmd/listargrupos", {}, "token-falso")
    assert len(chamadas) == 1


def test_token_recusado_vira_mensagem_clara(monkeypatch):
    def abrir_falso(requisicao, contexto):
        raise ap.urllib.error.HTTPError(requisicao.full_url, 403, "Forbidden", {}, None)

    monkeypatch.setattr(ap, "abrir", abrir_falso)
    with pytest.raises(ap.ErroArgus, match="token"):
        ap.chamar_api("cmd/listargrupos", {}, "token-falso")
