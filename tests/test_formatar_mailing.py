"""
Testes do formatar_mailing.py (versão em Python do formatador do BENHERMES).

Os CPFs, benefícios e nomes aqui são INVENTADOS, criados dentro do teste.
"""

import openpyxl
import pytest

import formatar_mailing as fm

CABECALHO = "CPF;BENEFICIO;NOME;TELEFONE1;TELEFONE2;TELEFONE3;TELEFONE4;TELEFONE5"


def ler_saida(caminho):
    dados = caminho.read_bytes()
    assert dados.startswith("﻿".encode("utf-8"))      # BOM para o Excel
    return dados.decode("utf-8-sig").split("\r\n")


def criar_xlsx(caminho, abas):
    livro = openpyxl.Workbook()
    livro.remove(livro.active)
    for nome, linhas in abas.items():
        aba = livro.create_sheet(nome)
        for linha in linhas:
            aba.append(linha)
    livro.save(caminho)


# --- detecção ---------------------------------------------------------------

def test_normalizar_cabecalho():
    assert fm.normalizar_cabecalho("Nº do Benefício") == "nodobeneficio"
    assert fm.normalizar_cabecalho(" CPF Cliente ") == "cpfcliente"
    assert fm.normalizar_cabecalho(None) == ""


def test_detecta_colunas_pelos_nomes_comuns():
    assert fm.detectar_colunas(["Documento", "Nr Benefício", "Nome do Cliente", "Valor"]) == {
        "cpf": "Documento", "beneficio": "Nr Benefício", "nome": "Nome do Cliente"}


def test_nome_exato_ganha_de_nome_parecido():
    # No BENHERMES a primeira coluna que CONTÉM o padrão vencia ("Desbloqueio NB" virava Benefício)
    colunas = fm.detectar_colunas(["Desbloqueio NB", "CPF", "Beneficio", "Nome"])
    assert colunas["beneficio"] == "Beneficio"


def test_mesma_coluna_nao_e_usada_duas_vezes():
    # "Nome Cliente" serviria para Nome e para CPF? Não; e "Cliente" sozinho não rouba o Nome
    colunas = fm.detectar_colunas(["CPF Cliente", "NB", "Cliente"])
    assert colunas == {"cpf": "CPF Cliente", "beneficio": "NB", "nome": "Cliente"}


def test_acha_cabecalho_abaixo_do_titulo():
    linhas = [["RELATÓRIO DE CLIENTES"], ["Gerado em 20/10/2026"], [],
              ["Nome", "CPF", "Benefício"], ["FULANO", "12345678901", "1112223334"]]
    assert fm.achar_linha_do_cabecalho(linhas) == 3


# --- formatação -------------------------------------------------------------

def test_xlsx_com_titulo_gera_o_csv_do_benhermes(tmp_path):
    origem = tmp_path / "Base Outubro (1).xlsx"
    criar_xlsx(origem, {"Plan1": [
        ["Mailing de teste"],
        [],
        ["Nome do Cliente", "Valor", "CPF", "Benefício"],
        ["FULANO DE TAL", 1500.5, 12345678901, 1112223334],
        ["CICLANA SOUSA", 300, "023.456.789-01", "5556667778"],
    ]})
    original = origem.read_bytes()

    assert fm.main([str(origem), "--saida", str(tmp_path / "saida")]) == 0

    destino = tmp_path / "saida" / "Base_Outubro__1__formatado.csv"
    assert ler_saida(destino) == [
        CABECALHO,
        "12345678901;1112223334;FULANO DE TAL;;;;;",
        "023.456.789-01;5556667778;CICLANA SOUSA;;;;;",
    ]
    assert fm.formatar(origem).linha_cabecalho == 3      # conta o título e a linha em branco
    assert origem.read_bytes() == original                # original intacto


def test_cpf_sem_zero_a_esquerda_e_completado(tmp_path):
    origem = tmp_path / "base.xlsx"
    criar_xlsx(origem, {"Plan1": [["CPF", "NB", "Nome"], [1234567890, 111, "FULANO"]]})

    resultado = fm.formatar(origem)
    assert resultado.linhas == [("01234567890", "111", "FULANO")]
    assert resultado.cpfs_completados == 1

    assert fm.formatar(origem, manter_cpf=True).linhas == [("1234567890", "111", "FULANO")]


@pytest.mark.parametrize("separador", [";", ",", "\t", "|"])
def test_csv_com_qualquer_separador(tmp_path, separador):
    origem = tmp_path / "base.csv"
    linhas = [["Documento", "Matricula", "Titular"], ["12345678901", "999", "JOANA SOUSA"]]
    origem.write_text("\n".join(separador.join(l) for l in linhas), encoding="cp1252")

    assert fm.formatar(origem).linhas == [("12345678901", "999", "JOANA SOUSA")]


def test_texto_com_ponto_e_virgula_vai_entre_aspas(tmp_path):
    origem = tmp_path / "base.csv"
    origem.write_text('CPF,Beneficio,Nome\n12345678901,1,"SOUSA; JOANA ""JO"""\n', encoding="utf-8")

    assert fm.texto_csv(fm.formatar(origem)).split("\r\n")[1] == '12345678901;1;"SOUSA; JOANA ""JO""";;;;;'


def test_escolhe_aba_e_colunas_a_mao(tmp_path):
    origem = tmp_path / "base.xlsx"
    criar_xlsx(origem, {
        "Resumo": [["Total", 1]],
        "Clientes": [["Doc Cliente", "Num Ben", "Pessoa"], ["12345678901", "7", "FULANO"]],
    })

    resultado = fm.formatar(origem, aba="Clientes",
                            colunas={"cpf": "Doc Cliente", "beneficio": "Num Ben", "nome": "Pessoa"})
    assert resultado.abas == ["Resumo", "Clientes"]
    assert resultado.linhas == [("12345678901", "7", "FULANO")]


def test_avisa_quando_falta_coluna(tmp_path):
    origem = tmp_path / "base.csv"
    origem.write_text("CPF;Valor\n12345678901;10\n", encoding="utf-8")

    with pytest.raises(ValueError, match="BENEFICIO, NOME"):
        fm.formatar(origem)


def test_aba_ou_coluna_inexistente(tmp_path):
    origem = tmp_path / "base.xlsx"
    criar_xlsx(origem, {"Plan1": [["CPF", "NB", "Nome"], ["1", "2", "X"]]})

    with pytest.raises(ValueError, match="aba 'Outra'"):
        fm.formatar(origem, aba="Outra")
    with pytest.raises(ValueError, match="coluna 'Doc'"):
        fm.formatar(origem, colunas={"cpf": "Doc"})


def test_formato_nao_suportado(tmp_path):
    with pytest.raises(ValueError, match="não suportado"):
        fm.formatar(tmp_path / "base.pdf")


def test_sem_argumentos_formata_a_pasta_de_entrada(tmp_path, monkeypatch, capsys):
    entrada = tmp_path / "entrada" / "mailing"
    entrada.mkdir(parents=True)
    (entrada / "a.csv").write_text("CPF;NB;Nome\n12345678901;1;FULANO\n", encoding="utf-8")
    (entrada / "b.csv").write_text("CPF;Valor\n1;2\n", encoding="utf-8")     # sem colunas: erro
    monkeypatch.setattr(fm, "PASTA_ENTRADA", entrada)

    assert fm.main(["--saida", str(tmp_path / "saida")]) == 1                 # houve um erro...
    assert (tmp_path / "saida" / "a_formatado.csv").exists()                   # ...mas o outro saiu
    assert "não identifiquei" in capsys.readouterr().out


def test_pasta_de_entrada_vazia(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(fm, "PASTA_ENTRADA", tmp_path / "nada")
    assert fm.main([]) == 1
    assert "Nenhuma planilha" in capsys.readouterr().out
