"""
Anonimiza as exportações antes de usá-las no desenvolvimento ou de enviá-las a alguém.

RODE SEMPRE NO SEU COMPUTADOR. Os arquivos originais nunca são alterados:
cada arquivo de entrada/ gera uma cópia anonimizada em anonimizados/.

O que é feito em cada tipo de arquivo (identificado pelas colunas, não pelo nome):
  - Esteira do Vanguard (tem CPF, Etapa e Agente):
      CPF do cliente vira um código (CLIENTE-000001). O mesmo CPF recebe sempre
      o mesmo código, então dá para contar clientes únicos.
      Só ficam as colunas da lista COLUNAS_ESTEIRA; as outras (nome do cliente,
      benefício, salário, contrato, CPF do agente etc.) ficam vazias.
  - Cadastro de operadores do Argus (tem Nome, Supervisor(es) e Status):
      Só ficam as colunas da lista COLUNAS_OPERADORES; CPF, e-mail, RE e ramal ficam vazios.
  - Demais arquivos (desempenho, leads, tarifação): não têm dados de cliente
      e são copiados como estão.

No final, todos os arquivos gerados são conferidos: se aparecer algo com
cara de CPF (11 dígitos) ou e-mail, o script avisa.

Como usar:
  python anonimizar.py                       -> anonimiza todos os CSVs de entrada/
  python anonimizar.py arquivo1.csv arq2.csv -> anonimiza só esses arquivos
"""

import csv
import re
import sys
from pathlib import Path

from leitura import ler_csv, limpar_cabecalho

PASTA_PROJETO = Path(__file__).resolve().parent
PASTA_ENTRADA = PASTA_PROJETO / "entrada"
PASTA_ANONIMIZADOS = PASTA_PROJETO / "anonimizados"

# Colunas da esteira que podem ser mantidas (nomes já padronizados, sem acento)
COLUNAS_ESTEIRA = {
    "DATA/HORA", "DATA PAGAMENTO", "DATA RETORNO", "DATA CANCELAMENTO",
    "ETAPA", "STATUS", "BANCO", "ORGAO", "PRODUTO", "TABELA", "PARCELA", "PRAZO",
    "VALOR CONTRATO", "VL. LIBERADA", "AGENTE", "EQUIPE",
    "ULTIMA ATUALIZACAO", "ATUALIZADO POR", "DIGITADOR", "USUARIO-DIGITADOR",
}

# Colunas do cadastro de operadores que podem ser mantidas
COLUNAS_OPERADORES = {
    "COD. USUARIO", "NOME", "[CAMPANHA] GRUPO ASSOCIADO", "SUPERVISOR(ES)",
    "CAMPANHAS VINCULADAS", "STATUS",
}

# Padrões que NÃO podem sobrar nos arquivos anonimizados
PADRAO_CPF = re.compile(r"(?<!\d)(\d{3}\.?\d{3}\.?\d{3}-?\d{2})(?!\d)")
PADRAO_EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")


def tipo_do_arquivo(cabecalho):
    """Descobre o tipo da exportação pelas colunas."""
    colunas = {limpar_cabecalho(c) for c in cabecalho}
    if {"CPF", "ETAPA", "AGENTE"} <= colunas:
        return "esteira"
    if {"NOME", "SUPERVISOR(ES)", "STATUS"} <= colunas:
        return "operadores"
    return "outro"


def anonimizar_linhas(cabecalho, linhas, tipo, codigos_cpf):
    """
    Devolve as linhas anonimizadas.

    `codigos_cpf` é um dicionário {cpf: código} compartilhado entre os arquivos,
    para o mesmo CPF ter o mesmo código em todos eles.
    """
    if tipo == "outro":
        return [list(linha) for linha in linhas]

    permitidas = COLUNAS_ESTEIRA if tipo == "esteira" else COLUNAS_OPERADORES
    colunas = [limpar_cabecalho(c) for c in cabecalho]
    novas = []
    for linha in linhas:
        nova = []
        for coluna, valor in zip(colunas, linha):
            if tipo == "esteira" and coluna == "CPF":
                cpf = re.sub(r"\D", "", valor)  # só os dígitos
                if cpf:
                    if cpf not in codigos_cpf:
                        codigos_cpf[cpf] = f"CLIENTE-{len(codigos_cpf) + 1:06d}"
                    nova.append(codigos_cpf[cpf])
                else:
                    nova.append("")
            elif coluna in permitidas:
                nova.append(valor)
            else:
                nova.append("")  # coluna sensível ou desnecessária: fica vazia
        # Linhas com mais células que o cabeçalho: o excesso é descartado
        novas.append(nova)
    return novas


def procurar_dados_sensiveis(cabecalho, linhas):
    """
    Procura CPFs e e-mails que tenham sobrado.

    Devolve {nome_da_coluna: quantidade}, sem mostrar os valores encontrados.
    """
    achados = {}
    for linha in linhas:
        for coluna, valor in zip(cabecalho, linha):
            if PADRAO_CPF.search(valor) or PADRAO_EMAIL.search(valor):
                achados[coluna] = achados.get(coluna, 0) + 1
    return achados


def anonimizar_arquivo(origem, pasta_destino, codigos_cpf):
    """Anonimiza um CSV e grava a cópia em pasta_destino. Devolve um resumo."""
    origem = Path(origem)
    destino = Path(pasta_destino) / origem.name
    if destino.resolve() == origem.resolve():
        raise ValueError(f"{origem.name}: o destino seria o próprio arquivo original.")

    cabecalho, linhas, codificacao = ler_csv(origem)
    tipo = tipo_do_arquivo(cabecalho)
    novas = anonimizar_linhas(cabecalho, linhas, tipo, codigos_cpf)

    # Grava com a mesma codificação e separador do original
    destino.parent.mkdir(parents=True, exist_ok=True)
    with open(destino, "w", encoding=codificacao, newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";", quoting=csv.QUOTE_ALL)
        escritor.writerow(cabecalho)
        escritor.writerows(novas)

    return {
        "arquivo": origem.name,
        "tipo": tipo,
        "linhas": len(novas),
        "destino": destino,
        "suspeitos": procurar_dados_sensiveis(cabecalho, novas),
    }


def main(argumentos=None):
    argumentos = sys.argv[1:] if argumentos is None else argumentos
    if argumentos:
        arquivos = [Path(a) for a in argumentos]
    else:
        arquivos = sorted(PASTA_ENTRADA.glob("*.csv"))
        outros = [p.name for p in PASTA_ENTRADA.glob("*") if p.is_file() and p.suffix.lower() != ".csv"]
        if outros:
            print(f"Ignorados (exporte em CSV para anonimizar): {', '.join(outros)}")
    if not arquivos:
        print(f"Nenhum CSV encontrado em {PASTA_ENTRADA}.")
        return 1

    codigos_cpf = {}
    tem_problema = False
    print("\nANONIMIZAÇÃO")
    for caminho in arquivos:
        resumo = anonimizar_arquivo(caminho, PASTA_ANONIMIZADOS, codigos_cpf)
        print(f"- {resumo['arquivo']}: {resumo['tipo']}, {resumo['linhas']} linhas -> {resumo['destino']}")
        if resumo["suspeitos"]:
            tem_problema = True
            for coluna, qtd in resumo["suspeitos"].items():
                print(f"    ATENÇÃO: {qtd} valor(es) com cara de CPF/e-mail na coluna '{coluna}'")
    print(f"\nCPFs substituídos por código: {len(codigos_cpf)}")
    if tem_problema:
        print("ATENÇÃO: confira os avisos acima antes de compartilhar os arquivos anonimizados.")
        return 2
    print("Conferência final: nenhum CPF ou e-mail encontrado nos arquivos gerados.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
