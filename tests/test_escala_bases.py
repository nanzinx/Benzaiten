"""
Testes do sorteio de bases (escala_bases.py).

Todos os dados aqui são FICTÍCIOS e criados dentro do próprio teste,
numa pasta temporária. Nada é lido de config/ ou historico/.
"""

import random
from collections import Counter
from datetime import date

import openpyxl
import pandas as pd
import pytest

import escala_bases as eb

BASES = ["BASE A", "BASE B", "BASE C", "BASE D"]
PESSOAS = ["ANA", "BRUNO", "CARLA", "DIEGO"]


def registro(d, base, pessoa):
    return {"data": d, "base": base, "responsavel": pessoa}


def historico_de(*dias):
    """Monta um histórico a partir de pares (data, {base: pessoa})."""
    regs = []
    for d, escolha in dias:
        regs += [registro(d, b, p) for b, p in escolha.items()]
    return regs


# Dois sorteios anteriores (sexta 16/10 e segunda 19/10)
HISTORICO = historico_de(
    (date(2026, 10, 16), {"BASE A": "ANA", "BASE B": "BRUNO", "BASE C": "CARLA", "BASE D": "DIEGO"}),
    (date(2026, 10, 19), {"BASE A": "BRUNO", "BASE B": "ANA", "BASE C": "DIEGO", "BASE D": "CARLA"}),
)


# --- Cadastro ---------------------------------------------------------------

def test_ler_cadastro_ignora_vazios_espacos_e_repetidos(tmp_path):
    arquivo = tmp_path / "bases.xlsx"
    with pd.ExcelWriter(arquivo) as escritor:
        pd.DataFrame({"Base": ["BASE A", None, "  BASE B ", "base a"]}).to_excel(escritor, sheet_name="Bases", index=False)
        pd.DataFrame({"Pessoa": ["ANA", "", "BRUNO"]}).to_excel(escritor, sheet_name="Pessoas", index=False)

    bases, pessoas = eb.ler_cadastro(arquivo)

    assert bases == ["BASE A", "BASE B"]
    assert pessoas == ["ANA", "BRUNO"]


def test_ler_cadastro_inexistente_explica_o_que_fazer(tmp_path):
    with pytest.raises(FileNotFoundError, match="--importar-xlsm"):
        eb.ler_cadastro(tmp_path / "nao_existe.xlsx")


# --- Regras do sorteio ------------------------------------------------------

def test_cada_base_tem_um_responsavel_e_ninguem_repete_no_dia():
    escolha, _ = eb.sortear(BASES, PESSOAS, [], date(2026, 10, 20), rng=random.Random(1))

    assert list(escolha) == BASES                      # todas as bases, na ordem do cadastro
    assert sorted(escolha.values()) == sorted(PESSOAS)  # cada pessoa com exatamente uma base


@pytest.mark.parametrize("semente", range(200))
def test_ninguem_repete_base_dos_ultimos_2_sorteios(semente):
    escolha, janela = eb.sortear(BASES, PESSOAS, HISTORICO, date(2026, 10, 20), rng=random.Random(semente))

    assert janela == 2
    for base, pessoa in escolha.items():
        anteriores = [r["responsavel"] for r in HISTORICO if r["base"] == base]
        assert pessoa not in anteriores


def test_sorteio_e_aleatorio_e_nao_um_rodizio_fixo():
    resultados = {
        tuple(eb.sortear(BASES, PESSOAS, HISTORICO, date(2026, 10, 20), rng=random.Random(s))[0].values())
        for s in range(100)
    }
    assert len(resultados) > 1


def test_janela_conta_sorteios_e_nao_dias_do_calendario():
    # De sexta (16/10) para segunda (19/10) passam 3 dias, mas é o sorteio anterior
    sexta = historico_de((date(2026, 10, 16), {"BASE A": "ANA"}))
    anteriores = eb.sorteios_anteriores(sexta, date(2026, 10, 19), 1)
    assert anteriores == [{"BASE A": "ANA"}]


def test_sorteio_do_mesmo_dia_ou_futuro_nao_conta_como_anterior():
    hist = historico_de(
        (date(2026, 10, 20), {"BASE A": "ANA"}),
        (date(2026, 10, 21), {"BASE A": "BRUNO"}),
    )
    assert eb.sorteios_anteriores(hist, date(2026, 10, 20), 2) == []


def test_comparacao_com_historico_ignora_acentos_e_maiusculas():
    hist = historico_de((date(2026, 10, 19), {"base a": "Ána"}))
    assert eb.repete_base({"BASE A": "ANA"}, eb.sorteios_anteriores(hist, date(2026, 10, 20), 1))


def test_menos_pessoas_que_bases_distribui_sem_diferenca_maior_que_1():
    escolha, _ = eb.sortear(BASES, ["ANA", "BRUNO", "CARLA"], [], date(2026, 10, 20), rng=random.Random(3))

    quantidade = Counter(escolha.values())
    assert len(escolha) == 4                      # nenhuma base sem responsável
    assert set(quantidade) == {"ANA", "BRUNO", "CARLA"}
    assert max(quantidade.values()) - min(quantidade.values()) <= 1


def test_com_2_pessoas_para_4_bases_ainda_respeita_o_dia_anterior():
    # Ontem: ANA com A e C, BRUNO com B e D. Hoje a única saída é inverter.
    hist = historico_de((date(2026, 10, 19), {"BASE A": "ANA", "BASE B": "BRUNO", "BASE C": "ANA", "BASE D": "BRUNO"}))
    for semente in range(50):
        escolha, janela = eb.sortear(BASES, ["ANA", "BRUNO"], hist, date(2026, 10, 20), rng=random.Random(semente))
        assert janela >= 1
        assert escolha == {"BASE A": "BRUNO", "BASE B": "ANA", "BASE C": "BRUNO", "BASE D": "ANA"}


def test_mais_pessoas_que_bases_alguem_fica_sem_base():
    pessoas = PESSOAS + ["ELISA"]
    escolha, _ = eb.sortear(BASES, pessoas, [], date(2026, 10, 20), rng=random.Random(4))

    assert len(escolha) == 4
    assert len(set(escolha.values())) == 4        # ninguém com duas bases
    assert len(set(pessoas) - set(escolha.values())) == 1


def test_sem_combinacao_possivel_reduz_a_janela():
    # Com 2 bases e 2 pessoas, quem já teve as duas bases não tem opção na janela de 2
    hist = historico_de(
        (date(2026, 10, 16), {"BASE A": "ANA", "BASE B": "BRUNO"}),
        (date(2026, 10, 19), {"BASE A": "BRUNO", "BASE B": "ANA"}),
    )
    escolha, janela = eb.sortear(["BASE A", "BASE B"], ["ANA", "BRUNO"], hist, date(2026, 10, 20), rng=random.Random(0))

    assert janela == 1
    assert escolha == {"BASE A": "ANA", "BASE B": "BRUNO"}  # única opção sem repetir o dia anterior


def test_sem_bases_ou_sem_pessoas_da_erro():
    with pytest.raises(ValueError):
        eb.sortear([], PESSOAS, [], date(2026, 10, 20))
    with pytest.raises(ValueError):
        eb.sortear(BASES, [], [], date(2026, 10, 20))


# --- Ausências --------------------------------------------------------------

def test_ausentes_saem_do_sorteio_sem_diferenciar_maiusculas():
    assert eb.aplicar_ausencias(PESSOAS, ["carla", " Diego "]) == ["ANA", "BRUNO"]


def test_ausente_com_nome_errado_da_erro():
    with pytest.raises(ValueError, match="CARLLA"):
        eb.aplicar_ausencias(PESSOAS, ["CARLLA"])


# --- Histórico e execução completa -------------------------------------------

def test_historico_grava_e_le_igual(tmp_path):
    arquivo = tmp_path / "historico.csv"
    eb.salvar_historico(HISTORICO, arquivo)
    assert eb.ler_historico(arquivo) == HISTORICO


@pytest.fixture
def projeto_ficticio(tmp_path, monkeypatch):
    """Aponta o script para uma pasta temporária com um cadastro fictício."""
    cadastro = tmp_path / "config" / "bases.xlsx"
    cadastro.parent.mkdir()
    with pd.ExcelWriter(cadastro) as escritor:
        pd.DataFrame({"Base": BASES}).to_excel(escritor, sheet_name="Bases", index=False)
        pd.DataFrame({"Pessoa": PESSOAS}).to_excel(escritor, sheet_name="Pessoas", index=False)
    monkeypatch.setattr(eb, "ARQUIVO_CADASTRO", cadastro)
    monkeypatch.setattr(eb, "ARQUIVO_HISTORICO", tmp_path / "historico" / "historico_escalas.csv")
    monkeypatch.setattr(eb, "PASTA_SAIDA", tmp_path / "saida")
    return tmp_path


def test_execucao_completa_gera_planilha_e_historico(projeto_ficticio):
    assert eb.main(["--data", "21/10/2026", "--ausentes", "ana"]) == 0

    planilha = projeto_ficticio / "saida" / "escala_2026-10-21.xlsx"
    escala = openpyxl.load_workbook(planilha)["Escala"]
    assert escala["A1"].value == "ESCALA DE BASES — ATIVO 2"
    assert escala["A2"].value == "21/10/2026 (quarta-feira)"
    responsaveis = [escala.cell(row=r, column=2).value for r in range(5, 9)]
    assert "ANA" not in responsaveis
    assert any("Ausentes: ANA" == str(c.value) for c in escala["A"])

    historico = eb.ler_historico(eb.ARQUIVO_HISTORICO)
    assert len(historico) == 4
    assert eb.ler_cadastro(eb.ARQUIVO_CADASTRO) == (BASES, PESSOAS)  # cadastro não foi alterado


def test_data_ja_sorteada_nao_e_substituida_sem_confirmar(projeto_ficticio, monkeypatch):
    eb.main(["--data", "21/10/2026"])
    antes = eb.ler_historico(eb.ARQUIVO_HISTORICO)

    monkeypatch.setattr("builtins.input", lambda _: "n")
    assert eb.main(["--data", "21/10/2026"]) == 1
    assert eb.ler_historico(eb.ARQUIVO_HISTORICO) == antes


def test_substituir_troca_o_dia_sem_duplicar(projeto_ficticio):
    eb.main(["--data", "21/10/2026"])
    eb.main(["--data", "21/10/2026", "--substituir"])
    assert len(eb.ler_historico(eb.ARQUIVO_HISTORICO)) == 4


def test_data_invalida_da_erro_claro(projeto_ficticio):
    with pytest.raises(ValueError, match="dd/mm/aaaa"):
        eb.main(["--data", "2026-10-21"])


# --- Importação da planilha antiga ------------------------------------------

def test_importar_xlsm_cria_cadastro_e_historico(tmp_path):
    # Planilha fictícia com a mesma estrutura da Escala_Bases_Ativo2.xlsm
    wb = openpyxl.Workbook()
    escala = wb.active
    escala.title = "Escala"
    escala["A16"], escala["B16"], escala["C16"] = "Data", "BASE A", "BASE B"
    escala["A17"], escala["B17"], escala["C17"] = pd.Timestamp("2026-10-16").to_pydatetime(), "ANA", "BRUNO"
    escala["A18"], escala["B18"], escala["C18"] = pd.Timestamp("2026-10-19").to_pydatetime(), "BRUNO", "ANA"
    cadastro = wb.create_sheet("Cadastro")
    cadastro["A4"], cadastro["C4"] = "Bases", "Pessoas"
    cadastro["A5"], cadastro["C5"] = "BASE A", "ANA"
    cadastro["A6"], cadastro["C6"] = "BASE B", "BRUNO"
    antiga = tmp_path / "antiga.xlsx"
    wb.save(antiga)

    arq_cadastro = tmp_path / "config" / "bases.xlsx"
    arq_historico = tmp_path / "historico" / "historico.csv"
    eb.importar_xlsm(antiga, arq_cadastro, arq_historico)
    eb.importar_xlsm(antiga, arq_cadastro, arq_historico)  # rodar 2 vezes não duplica

    assert eb.ler_cadastro(arq_cadastro) == (["BASE A", "BASE B"], ["ANA", "BRUNO"])
    historico = eb.ler_historico(arq_historico)
    assert len(historico) == 4
    assert registro(date(2026, 10, 19), "BASE A", "BRUNO") in historico
