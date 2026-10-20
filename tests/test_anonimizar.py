"""
Testes do anonimizar.py e das funções de leitura.

Os CPFs, nomes e e-mails aqui são INVENTADOS, criados dentro do teste.
"""

import anonimizar as an
from leitura import ler_csv, limpar_cabecalho, normalizar_nome

CABECALHO_ESTEIRA = ["Proposta", "Data/Hora", "Etapa", "CPF", "Nome", "Beneficio",
                     "Agente", "Equipe", "Valor Contrato", "CPF Agente", "Acompanhamento"]


def escrever_csv(caminho, linhas, codificacao="utf-8-sig"):
    texto = "\n".join(";".join(f'"{c}"' for c in linha) for linha in linhas) + "\n"
    caminho.write_text(texto, encoding=codificacao)


# --- leitura.py -------------------------------------------------------------

def test_normalizar_nome_tira_acentos_maiusculas_e_espacos():
    assert normalizar_nome("  José   da Conceição ") == "JOSE DA CONCEICAO"
    assert normalizar_nome("CAUÃ") == "CAUA"
    assert normalizar_nome(None) == ""


def test_normalizar_nome_nao_junta_grafias_diferentes():
    # Isso é trabalho da aba Apelidos, não de correspondência automática
    assert normalizar_nome("JOANA SOUSA") != normalizar_nome("JOANNA SOUSA")


def test_limpar_cabecalho_remove_simbolo_do_argus():
    assert limpar_cabecalho("? Nome") == "NOME"
    assert limpar_cabecalho("↓ Nome") == "NOME"
    assert limpar_cabecalho("Cód. Usuário") == "COD. USUARIO"
    assert limpar_cabecalho("[Campanha] Grupo Associado") == "[CAMPANHA] GRUPO ASSOCIADO"


def test_ler_csv_entende_as_duas_codificacoes(tmp_path):
    utf8 = tmp_path / "utf8.csv"
    windows = tmp_path / "windows.csv"
    escrever_csv(utf8, [["Descrição"], ["CAUÃ"]], "utf-8-sig")
    escrever_csv(windows, [["Descrição"], ["CAUÃ"]], "cp1252")

    assert ler_csv(utf8)[:2] == (["Descrição"], [["CAUÃ"]])
    assert ler_csv(windows)[:2] == (["Descrição"], [["CAUÃ"]])
    assert ler_csv(windows)[2] == "cp1252"


# --- anonimizar.py ----------------------------------------------------------

def test_identifica_tipo_pelas_colunas():
    assert an.tipo_do_arquivo(CABECALHO_ESTEIRA) == "esteira"
    assert an.tipo_do_arquivo(["", "? Nome", "CPF", "Supervisor(es)", "Status"]) == "operadores"
    assert an.tipo_do_arquivo(["Descrição", "Valor R$", "Qtde. Lig", "Minutos"]) == "outro"


def test_esteira_troca_cpf_por_codigo_e_apaga_dados_do_cliente(tmp_path):
    origem = tmp_path / "entrada" / "esteira.csv"
    origem.parent.mkdir()
    escrever_csv(origem, [
        CABECALHO_ESTEIRA,
        ["1", "19/10/2026 10:00", "Pago", "12345678901", "FULANO DE TAL", "9876543210",
         "AGENTE UM", "EQUIPE X", "1.500,00", "10987654321", "cliente ligou"],
        ["2", "19/10/2026 11:00", "Andamento", "123.456.789-01", "FULANO DE TAL", "9876543210",
         "AGENTE DOIS", "EQUIPE X", "800,00", "", ""],
        ["3", "19/10/2026 12:00", "Pendente", "55566677788", "CICLANA", "1112223334",
         "AGENTE UM", "EQUIPE X", "300,00", "", ""],
    ])
    original = origem.read_bytes()

    resumo = an.anonimizar_arquivo(origem, tmp_path / "anonimizados", {})
    cabecalho, linhas, _ = ler_csv(resumo["destino"])
    coluna = {nome: i for i, nome in enumerate(cabecalho)}

    # O mesmo CPF (com ou sem pontuação) vira o mesmo código
    assert [l[coluna["CPF"]] for l in linhas] == ["CLIENTE-000001", "CLIENTE-000001", "CLIENTE-000002"]
    # Dados do cliente e do agente apagados
    for nome in ["Proposta", "Nome", "Beneficio", "CPF Agente", "Acompanhamento"]:
        assert all(l[coluna[nome]] == "" for l in linhas), nome
    # O que o relatório usa continua igual
    assert [l[coluna["Valor Contrato"]] for l in linhas] == ["1.500,00", "800,00", "300,00"]
    assert [l[coluna["Agente"]] for l in linhas] == ["AGENTE UM", "AGENTE DOIS", "AGENTE UM"]
    assert cabecalho == CABECALHO_ESTEIRA            # mesma estrutura de colunas
    assert resumo["suspeitos"] == {}
    assert origem.read_bytes() == original             # original não foi alterado


def test_operadores_mantem_nome_e_supervisor_e_apaga_cpf_e_email(tmp_path):
    origem = tmp_path / "OPERADORES.csv"
    escrever_csv(origem, [
        ["", "Cód. Usuário", "? Nome", "Login", "CPF", "E-mail", "Supervisor(es)", "Status"],
        ["", "1", "OPERADORA TESTE", "OPER.TESTE", "12345678901", "teste@exemplo.com", "(1) SUPERVISOR X", "ATIVO"],
    ], "cp1252")

    resumo = an.anonimizar_arquivo(origem, tmp_path / "anonimizados", {})
    _, linhas, codificacao = ler_csv(resumo["destino"])

    assert linhas == [["", "1", "OPERADORA TESTE", "", "", "", "(1) SUPERVISOR X", "ATIVO"]]
    assert codificacao == "cp1252"                     # mesma codificação do original


def test_tarifacao_e_copiada_sem_mudanca(tmp_path):
    origem = tmp_path / "4NET.csv"
    linhas = [["Descrição", "Valor R$", "Qtde. Lig", "Minutos"],
              ["<LIGAÇÕES SEM DERIVAÇÃO P/ OPERADOR>", "123,55", "7060", "633,0333"],
              ["OPERADORA TESTE", "1,3", "25", "32,1667"]]
    escrever_csv(origem, linhas, "cp1252")

    resumo = an.anonimizar_arquivo(origem, tmp_path / "anonimizados", {})

    assert resumo["tipo"] == "outro"
    assert ler_csv(resumo["destino"])[:2] == (linhas[0], linhas[1:])


def test_conferencia_final_avisa_cpf_ou_email_que_sobrou():
    achados = an.procurar_dados_sensiveis(
        ["Obs", "Valor"],
        [["CPF 123.456.789-01", "10,50"], ["mande para a@b.com", "7060"], ["ok", "17011"]],
    )
    assert achados == {"Obs": 2}


def test_nunca_grava_por_cima_do_original(tmp_path):
    origem = tmp_path / "leads.csv"
    escrever_csv(origem, [["Data", "Operador"], ["19/10/2026", "X"]])
    try:
        an.anonimizar_arquivo(origem, tmp_path, {})
    except ValueError as erro:
        assert "original" in str(erro)
    else:
        raise AssertionError("deveria ter recusado gravar por cima do original")
