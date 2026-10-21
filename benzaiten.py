"""
BENZAITEN — rotina diária completa em um só comando.

O que acontece, em ordem:
  1. Prepara os arquivos de configuração que ainda não existem (feriados e metas).
  2. Confere a pasta entrada/ e diz o que está faltando.
  3. Sorteia a escala de bases de HOJE (se hoje for dia útil e ainda não houver sorteio).
  4. Gera o relatório da URA do último dia útil.
  5. Gera o ranking de contratos pagos (vendedores e supervisores).
  6. Se tudo deu certo, arquiva os arquivos usados em entrada/processados/AAAA-MM-DD/.

Se uma etapa der erro, as outras continuam e o resumo final mostra o que falhou.

Como usar (ou dê dois cliques em BENZAITEN.bat):
  python benzaiten.py
  python benzaiten.py --ausentes "FULANO, BELTRANO"   -> ausentes na escala de hoje
  python benzaiten.py --data 17/10/2026               -> relatório de outro dia
  python benzaiten.py --nao-arquivar                  -> deixa os arquivos em entrada/
  python benzaiten.py --sem-escala                    -> sem a escala
  python benzaiten.py --sem-relatorio                 -> sem o relatório da URA
  python benzaiten.py --sem-ranking                   -> sem o ranking de vendas
  python benzaiten.py --sem-api                       -> não baixa nada do Argus
"""

import argparse
import sys
from datetime import date, datetime

import calendario
import escala_bases
import ranking_vendas
import relatorio_ura
from leitura import converter_data

LINHA = "=" * 64


def titulo(texto):
    print(f"\n{LINHA}\n  {texto}\n{LINHA}")


def preparar_configuracao():
    """Cria os modelos de config/feriados.xlsx e config/metas.xlsx, se faltarem."""
    if calendario.criar_modelo_feriados(calendario.ARQUIVO_FERIADOS):
        print(f"- Criado {calendario.ARQUIVO_FERIADOS} (anote feriados locais e pontos facultativos).")
    if relatorio_ura.criar_modelo_metas(relatorio_ura.ARQUIVO_METAS):
        print(f"- Criado {relatorio_ura.ARQUIVO_METAS} (preencha para ativar o semáforo de metas).")


def conferir_entrada(dia_relatorio):
    """
    Mostra o que há em entrada/ para o relatório e devolve a lista de pendências.

    Desempenho e leads podem faltar se o token do Argus estiver configurado,
    porque o relatório baixa pela API.
    """
    pasta = relatorio_ura.PASTA_ENTRADA
    pendencias = []
    if not pasta.exists():
        pasta.mkdir(parents=True)
        print(f"- Pasta {pasta} criada (estava faltando).")
    entradas, _ = relatorio_ura.ler_entradas(pasta)

    def do_dia(tipo):
        return sum(1 for _, linhas in entradas[tipo]
                   if any(converter_data(l.get("DATA")) == dia_relatorio for l in linhas))

    esteiras = len(entradas["esteira"])
    tarifacoes = len(entradas["tarifacao"])
    print(f"  {'OK' if esteiras == 2 else '!!' if esteiras else '--'}  "
          f"Esteiras do Vanguard (andamento + pagamento): {esteiras} arquivo(s)")
    if not esteiras:
        pendencias.append("exportar as esteiras do Vanguard (data de andamento e data de pagamento) para entrada/")
    elif esteiras == 1:
        pendencias.append("só há 1 esteira em entrada/: o ranking precisa da de andamento E da de pagamento")
    print(f"  {'OK' if tarifacoes == 2 else '!!'}  Tarifações (VENDITORE + 4NET): {tarifacoes} arquivo(s)")
    if tarifacoes < 2:
        pendencias.append("exportar as tarifações VENDITORE e 4NET para entrada/")
    elif tarifacoes > 2:
        pendencias.append("há mais de 2 tarifações em entrada/: apague as de outros dias")

    tem_token = bool(_token_configurado())
    for tipo, nome in [("desempenho", "Desempenho (TMA/TMO)"), ("leads", "Leads da URA")]:
        if do_dia(tipo):
            print(f"  OK  {nome} de {dia_relatorio:%d/%m/%Y}: em entrada/")
        elif tem_token:
            print(f"  ->  {nome} de {dia_relatorio:%d/%m/%Y}: será baixado da API do Argus")
        else:
            print(f"  --  {nome} de {dia_relatorio:%d/%m/%Y}: faltando (e sem token do Argus)")
            pendencias.append(f"{nome.lower()} de {dia_relatorio:%d/%m/%Y}")
    return pendencias


def _token_configurado():
    """True se houver token do Argus (sem mostrar o token)."""
    import argus_api

    try:
        return argus_api.ler_token(argus_api.ARQUIVO_TOKEN)
    except argus_api.ErroArgus:
        return None


def rodar_escala(hoje, ausentes):
    """Sorteia a escala de hoje, se fizer sentido. Devolve uma frase de resumo."""
    feriados = calendario.ler_feriados(calendario.ARQUIVO_FERIADOS)
    if not calendario.eh_dia_util(hoje, feriados):
        motivo = calendario.eh_feriado(hoje, feriados) or "fim de semana"
        print(f"Hoje ({hoje:%d/%m/%Y}) não é dia útil ({motivo}): escala não sorteada.")
        return "pulada (não é dia útil)"
    historico = escala_bases.ler_historico(escala_bases.ARQUIVO_HISTORICO)
    if any(r["data"] == hoje for r in historico):
        print(f"A escala de {hoje:%d/%m/%Y} já foi sorteada. Para refazer: "
              "python escala_bases.py --substituir")
        return "já estava sorteada"
    argumentos = ["--data", f"{hoje:%d/%m/%Y}"]
    if ausentes:
        argumentos += ["--ausentes", ausentes]
    escala_bases.main(argumentos)
    return "sorteada"


def rodar_relatorio(data, sem_api, usados):
    """Gera o relatório da URA. Devolve uma frase de resumo."""
    argumentos = ["--data", data] + (["--sem-api"] if sem_api else [])
    relatorio_ura.main(argumentos, usados)
    return "gerado"


def rodar_ranking(data, hoje, usados):
    """Gera o ranking de contratos pagos. Devolve uma frase de resumo."""
    ranking_vendas.main(["--data", data, "--hoje", f"{hoje:%d/%m/%Y}"], usados)
    return "gerado"


def main(argumentos=None):
    parser = argparse.ArgumentParser(description="BENZAITEN: rotina diária completa.")
    parser.add_argument("--ausentes", default="", help='ausentes na escala de hoje. Ex.: "FULANO, BELTRANO"')
    parser.add_argument("--data", help="dia do relatório dd/mm/aaaa (padrão: último dia útil)")
    parser.add_argument("--sem-api", action="store_true", help="não baixa nada do Argus")
    parser.add_argument("--nao-arquivar", action="store_true", help="deixa os arquivos usados em entrada/")
    parser.add_argument("--sem-escala", action="store_true", help="não sorteia a escala")
    parser.add_argument("--sem-relatorio", action="store_true", help="não gera o relatório da URA")
    parser.add_argument("--sem-ranking", action="store_true", help="não gera o ranking de vendas")
    parser.add_argument("--hoje", help=argparse.SUPPRESS)  # usado nos testes
    args = parser.parse_args(argumentos)

    hoje = datetime.strptime(args.hoje, "%d/%m/%Y").date() if args.hoje else date.today()
    if args.data:
        try:
            dia_relatorio = datetime.strptime(args.data, "%d/%m/%Y").date()
        except ValueError:
            print(f"ERRO: data inválida: {args.data}. Use dd/mm/aaaa, ex.: 17/10/2026")
            return 1
    else:
        dia_relatorio = calendario.dia_util_anterior(hoje, calendario.ler_feriados(calendario.ARQUIVO_FERIADOS))
        args.data = f"{dia_relatorio:%d/%m/%Y}"

    print(f"\n⛩  BENZAITEN — {hoje:%d/%m/%Y}  ·  relatório de {dia_relatorio:%d/%m/%Y}")
    preparar_configuracao()
    resumo = {}

    if not (args.sem_relatorio and args.sem_ranking):
        titulo("1. Conferindo a pasta entrada/")
        pendencias = conferir_entrada(dia_relatorio)
        for pendencia in pendencias:
            print(f"  PENDENTE: {pendencia}")

    usados = []  # arquivos de entrada/ usados pelo relatório e pelo ranking
    etapas = []
    if not args.sem_escala:
        etapas.append(("Escala de bases", lambda: rodar_escala(hoje, args.ausentes)))
    if not args.sem_relatorio:
        etapas.append(("Relatório da URA", lambda: rodar_relatorio(args.data, args.sem_api, usados)))
    if not args.sem_ranking:
        etapas.append(("Ranking de vendas", lambda: rodar_ranking(args.data, hoje, usados)))

    for numero, (nome, etapa) in enumerate(etapas, start=2):
        titulo(f"{numero}. {nome}")
        try:
            resumo[nome] = ("OK", etapa())
        except PermissionError as erro:
            print(f"ERRO: não consegui gravar {erro.filename}. Se estiver aberto no Excel, feche e rode de novo.")
            resumo[nome] = ("ERRO", "arquivo aberto no Excel")
        except (FileNotFoundError, ValueError) as erro:
            print(f"ERRO: {erro}")
            resumo[nome] = ("ERRO", str(erro).splitlines()[0])

    # Arquiva só se o relatório e o ranking deram certo (senão, dá para rodar de novo)
    if usados and not args.nao_arquivar:
        falhou = [nome for nome, (situacao, _) in resumo.items() if situacao == "ERRO" and nome != "Escala de bases"]
        if falhou:
            print(f"\nArquivos NÃO arquivados porque houve erro em: {', '.join(falhou)}. Corrija e rode de novo.")
            resumo["Arquivamento"] = ("--", "não feito (houve erro)")
        else:
            destino = relatorio_ura.arquivar(usados, relatorio_ura.PASTA_ENTRADA, dia_relatorio)
            print(f"\nArquivos usados movidos para: {destino}")
            resumo["Arquivamento"] = ("OK", f"entrada/processados/{dia_relatorio:%Y-%m-%d}")

    titulo("Resumo")
    for nome, (situacao, detalhe) in resumo.items():
        print(f"  [{situacao:^4}] {nome}: {detalhe}")
    return 1 if any(situacao == "ERRO" for situacao, _ in resumo.values()) else 0


if __name__ == "__main__":
    sys.exit(main())
