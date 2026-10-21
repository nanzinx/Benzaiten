"""
Regras comuns para a esteira do Vanguard, usadas pelo relatório da URA e pelo ranking.

Todo dia entram DUAS exportações da esteira em entrada/:
  - a de "data de andamento" (propostas em andamento);
  - a de "data de pagamento" (propostas pagas no mês).
Uma mesma proposta pode aparecer nas duas. Para não contar em dobro, as linhas
são juntadas e cada proposta fica uma vez só.
"""

from leitura import normalizar_nome

# Produtos que contam como "margem/cartão" (ficam fora da portabilidade)
PRODUTOS_CARTAO = {"MARGEM LIVRE", "CARTAO BENEFICIO", "CARTAO"}


def chave_da_proposta(linha):
    """Número que identifica a proposta (coluna Proposta; se vazia, Numero-Proposta)."""
    return str(linha.get("PROPOSTA") or linha.get("NUMERO-PROPOSTA") or "").strip()


def remover_propostas_repetidas(linhas):
    """
    Deixa cada proposta uma vez só e devolve (linhas_sem_repetir, quantas_eram_repetidas).

    Se a mesma proposta vier com etapas diferentes (por exemplo, "Andamento" numa
    exportação e "Pago" na outra), fica a linha "Pago", que é a mais recente.
    Linhas sem número de proposta são mantidas como estão.
    """
    escolhidas = {}
    sem_numero = []
    repetidas = 0
    for linha in linhas:
        chave = chave_da_proposta(linha)
        if not chave:
            sem_numero.append(linha)
            continue
        if chave in escolhidas:
            repetidas += 1
            if normalizar_nome(linha.get("ETAPA")) == "PAGO":
                escolhidas[chave] = linha
        else:
            escolhidas[chave] = linha
    return list(escolhidas.values()) + sem_numero, repetidas


def nome_do_agente(texto):
    """'FULANO DE TAL - COMERCIAL' -> 'FULANO DE TAL' (tira o que vem depois de ' - ')."""
    return " ".join(str(texto or "").split(" - ")[0].split())


def eh_cartao(linha):
    """True se o produto da proposta for margem livre ou cartão."""
    return normalizar_nome(linha.get("PRODUTO")) in PRODUTOS_CARTAO
