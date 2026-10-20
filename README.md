# Automações Benconsig

Scripts em Python para tarefas diárias do time:

| Script | O que faz | Situação |
|---|---|---|
| `escala_bases.py` | Sorteia as bases do Ativo 2 entre os grupos dos supervisores | pronto |
| `anonimizar.py` | Tira CPFs e dados pessoais das exportações | pronto |
| `argus_api.py` | Baixa do Argus o desempenho (TMA/TME), os leads da URA e os operadores | pronto (falta testar com o token real) |
| `relatorio_ura.py` | Relatório de produtividade da URA | pronto (falta conferir as vendas com a esteira real) |

Tudo roda **no seu computador**. Nenhum dado é enviado para a internet.

> ⚠️ **Este repositório é público.** Nunca envie para o GitHub planilhas, CSVs ou
> qualquer arquivo das pastas `entrada/`, `anonimizados/`, `config/`, `saida/` e
> `historico/`. O arquivo `.gitignore` já bloqueia isso no `git`, mas **o botão
> "Upload files" do site do GitHub ignora o `.gitignore`**: não use.

---

## 1. Instalação (só na primeira vez)

1. Instale o Python 3.10 ou mais novo: <https://www.python.org/downloads/>
   (no Windows, marque a opção **"Add Python to PATH"** na instalação).
2. Baixe este projeto e abra um terminal dentro da pasta dele
   (no Windows: abra a pasta, clique na barra de endereço, digite `cmd` e Enter).
3. Instale as bibliotecas:

   ```
   pip install -r requirements.txt
   ```

   (ou, uma por uma: `pip install pandas openpyxl xlrd pytest`)

### Pastas

```
entrada/       exportações reais do dia (Argus e Vanguard)
anonimizados/  cópias sem dados de clientes (geradas pelo anonimizar.py)
config/        bases.xlsx (bases e pessoas do Ativo 2)
saida/         escalas e relatórios gerados
historico/     histórico acumulado das escalas
```

As pastas são criadas sozinhas quando o script precisa delas.

---

## 2. Escala de bases do Ativo 2

### Primeira vez: importar a planilha antiga

A pasta `config/` **não vem junto** com o projeto (ela guarda dados e fica fora do GitHub).
Crie a pasta e copie a planilha para dentro dela. No terminal do Windows (PowerShell):

```
mkdir config
copy "C:\caminho\onde\esta\Escala_Bases_Ativo2.xlsm" config\
```

(ou crie a pasta `config` pelo Explorador de Arquivos e arraste a planilha para dentro). Depois rode:

```
python escala_bases.py --importar-xlsm config/Escala_Bases_Ativo2.xlsm
```

Isso cria o `config/bases.xlsx` (com as bases e as pessoas da aba Cadastro) e
copia os sorteios antigos para `historico/historico_escalas.csv`.
A planilha `.xlsm` não é alterada e as macros não são executadas.

### Todo dia

```
python escala_bases.py
```

O resultado aparece no terminal e é salvo em:

- `saida/escala_AAAA-MM-DD.xlsx`: a escala do dia (e os últimos sorteios numa segunda aba);
- `historico/historico_escalas.csv`: uma linha por base por dia (abre no Excel).

### Opções

| Situação | Comando |
|---|---|
| Alguém faltou | `python escala_bases.py --ausentes "FULANO"` |
| Mais de um ausente | `python escala_bases.py --ausentes "FULANO, BELTRANO"` |
| Sortear outra data (dia esquecido) | `python escala_bases.py --data 21/10/2026` |
| Refazer um dia sem perguntar | `python escala_bases.py --substituir` |

Se o dia já tiver sorteio, o script pergunta antes de substituir.

### Regras do sorteio

- Cada base recebe **um responsável**.
- **Ninguém repete a mesma base dos últimos 2 sorteios.** (Conta sorteios, não dias:
  segunda-feira olha sexta e quinta.) Para mudar, altere `JANELA_SORTEIOS` no
  início do `escala_bases.py`. Com 4 bases e 4 pessoas, 3 deixaria só uma combinação
  possível, e o "sorteio" viraria um rodízio fixo.
- **Menos pessoas que bases** (ex.: alguém ausente): algumas pessoas ficam com
  mais de uma base, com diferença máxima de 1 entre elas.
- **Mais pessoas que bases**: alguém fica sem base e aparece um aviso.
  Cadastre novas bases em `config/bases.xlsx` se precisar.
- Se não houver combinação possível sem repetir, a regra é afrouxada
  (só o último sorteio, depois nenhum) e um aviso aparece.

### Incluir ou tirar bases e pessoas

Abra `config/bases.xlsx`:

- aba **Bases**, coluna `Base`: uma base por linha;
- aba **Pessoas**, coluna `Pessoa`: uma pessoa (grupo do supervisor) por linha.

Linhas vazias são ignoradas.

---

## 3. Anonimizar exportações

Use antes de compartilhar qualquer exportação com alguém (inclusive com uma IA).

1. Coloque os CSVs exportados em `entrada/`.
2. Rode:

   ```
   python anonimizar.py
   ```

3. As cópias anonimizadas ficam em `anonimizados/`, com o mesmo nome. Os originais
   não são alterados.

O que acontece com cada arquivo (o tipo é reconhecido pelas colunas):

| Arquivo | O que muda |
|---|---|
| Esteira do Vanguard | CPF vira `CLIENTE-000001` (o mesmo CPF recebe sempre o mesmo código). Nome do cliente, benefício, salário, proposta, contrato, CPF do agente e observações ficam vazios. Datas, etapa, agente, equipe e valores continuam. |
| Operadores do Argus | CPF, e-mail, login, RE e ramal ficam vazios. Nome, supervisor, grupo e status continuam. |
| Desempenho, leads e tarifação | Copiados sem mudança (não têm dados de clientes). |

No final o script confere os arquivos gerados. Se aparecer
**ATENÇÃO: ... com cara de CPF/e-mail**, abra o arquivo e veja a coluna indicada
antes de compartilhar.

---

## 4. Relatório de produtividade da URA

### Todo dia

1. Exporte para `entrada\`:
   - a **esteira** do Vanguard;
   - a **tarifação** da VENDITORE e a da 4NET do dia (Argus > Tarifação > Painel de Visualização).

   Apague os arquivos do dia anterior: as tarifações não têm data e seriam somadas.
2. Rode:

   ```
   python relatorio_ura.py
   ```

   O desempenho (TMA/TMO), os leads da URA e o cadastro de operadores são baixados
   sozinhos pela API do Argus (precisa do `config\argus_token.txt`, veja a seção 5).
   Se preferir exportar à mão, coloque os arquivos em `entrada\` e use `--sem-api`.
3. Abra `saida\relatorio_produtividade_AAAA-MM-DD.xlsx`.

A data é a do **último dia útil** (na segunda-feira, a sexta). Para outro dia:
`python relatorio_ura.py --data 17/10/2026`.

### Como o relatório é calculado

| Coluna | Cálculo |
|---|---|
| TMA / TMO | colunas TMA e TME do desempenho do Argus. Na equipe: **soma** dos operadores |
| Leads atendidos/URA | leads do grupo RECEPTIVO - URA |
| Custo total | custo do operador na tarifação (VENDITORE + 4NET) + parte do custo dos robôs da URA (`OPER VIRT...`), dividido na proporção dos leads |
| Vendas geral | soma do Valor Contrato das etapas Andamento, Pendente e Pago, com Data/Hora no dia do relatório |
| Cliente efetivados | CPFs diferentes com etapa Andamento ou Pago |
| Custo por CPF | custo total ÷ clientes efetivados |
| Produção/lead | vendas geral ÷ leads |

Ordem: de quem **mais vendeu** para quem menos vendeu (empate: mais leads); depois o
total de cada equipe, também por vendas; depois TOTAL OPERADORES, a linha
LIGAÇÕES SEM DERIVAÇÃO P/ OPERADOR e o TOTAL GERAL (que inclui essa linha).
Só aparecem operadores com leads, custo, vendas ou clientes.

As regras ficam no início do `relatorio_ura.py` (etapas, prefixos dos robôs etc.).

### Cadastro e apelidos (`config\operadores.xlsx`)

- Aba **Operadores**: atualizada sozinha a cada exportação de operadores do Argus.
  A equipe é o primeiro nome do supervisor. Quem não está ATIVO é ignorado.
- Aba **Apelidos**: preenchida por você quando o nome é diferente entre os sistemas.
  Coluna A = nome como aparece no sistema; coluna B = nome oficial (o do Argus).
  Ex.: `JOANA SOUSA` (como aparece na tarifação) → `JOANNA SOUSA` (como está no Argus).
  Acentos, maiúsculas e espaços extras já são ignorados; não há ligação automática
  por semelhança.

No final, o terminal lista os **nomes sem vínculo** (para você decidir se é apelido)
e os ex-colaboradores ignorados. A aba **Conferência** do relatório traz os arquivos
lidos, a conferência dos custos e esses avisos.

---

## 5. Baixar os dados do Argus pela API

### Primeira vez: guardar o token

1. No Argus, abra **Configurações > Desenvolvedor > Acesso API** e copie a chave.
2. Crie o arquivo `config\argus_token.txt` e cole só a chave dentro dele.

O token é uma senha: ele fica só no seu computador (o `.gitignore` impede que vá
para o GitHub). Não mande esse arquivo para ninguém.

Para conferir se está tudo certo, rode uma vez:

```
python argus_api.py --diagnostico
```

Ele mostra os grupos operacionais e os **nomes** dos campos que o Argus devolve
(sem mostrar os dados). Se aparecer `NENHUM` ou `None` em alguma linha, me mande essa saída.

### Se aparecer "Não consegui conectar ao Argus"

- O script já tenta sozinho um modo compatível com servidores mais antigos
  (o certificado do Argus continua sendo verificado).
- Se mesmo assim não conectar, o endereço da API da Benconsig pode ser outro.
  Pergunte ao suporte do Argus qual é e escreva-o (uma linha só) em
  `config\argus_url.txt`, por exemplo: `https://argus4com.com.br:23001/apiargus`.

### Todo dia

```
python argus_api.py
```

Baixa os dados do **último dia útil** (ontem; na segunda-feira, a sexta) e grava em `entrada\`:

| Arquivo | Conteúdo |
|---|---|
| `argus_desempenho_AAAA-MM-DD.csv` | Data, Operador, Supervisor, TME, TMA |
| `argus_leads_ura_AAAA-MM-DD.csv` | Data, Operador, Supervisor, Leads. Atend. (grupo RECEPTIVO - URA) |
| `argus_operadores_AAAA-MM-DD.csv` | Cadastro de operadores (sem CPF, e-mail, login, RE e ramal) |

Para outro dia (por exemplo, depois de um feriado): `python argus_api.py --data 17/10/2026`.

O Argus devolve várias linhas por operador (uma por sessão/grupo). O script junta
em uma linha por operador: TMA e TME são a média das linhas, ponderada pela
quantidade de atendimentos; os leads da URA são a soma dos atendimentos no grupo
RECEPTIVO - URA.

Para conferir os números, `python argus_api.py --bruto` grava as linhas exatamente
como vieram do Argus (`argus_bruto_*.csv`).

> A **tarifação** (custos VENDITORE e 4NET) **não existe na API** do Argus.
> Continue exportando à mão: Argus > Tarifação > Painel de Visualização, para `entrada\`.

---

## 6. Testes

```
python -m pytest
```

Os testes usam só dados **inventados**, criados dentro dos próprios testes.
Todos devem aparecer como `passed`.
