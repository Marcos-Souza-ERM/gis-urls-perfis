# `main.py` — conferir o Plano de Amostragem contra os resultados

## Em poucas palavras

Imagine três listas diferentes:

1. o que **deveria** ser coletado/analisado (Plano de Amostragem);
2. o que a equipe registrou no **checklist**;
3. o que chegou nos arquivos de **resultado do laboratório** (EQuIS/EDD).

Este script põe essas listas lado a lado e cria um Excel para ajudar a encontrar diferenças. Ele não coleta dados nem corrige as fontes.

Os caminhos, sites ativos, nomes de arquivos, abas EQuIS e cabeçalhos do PA são configurados no grupo `MAIN_` do [`.env`](../.env). Os valores iniciais representam Braskem BA/terceira campanha. O layout dos checklists e o esquema interno do EQuIS continuam específicos do código.

## Entradas que o script espera

1. **Planilha de plano de amostragem** — caminho definido na variável `file`. O script percorre todas as abas cujo nome contém `_PA`. Em cada aba espera colunas equivalentes a `Grupos`, `Poços` e `Parâmetros` (aceita também o cabeçalho `Parâmetros ` com espaço no final, que renomeia).
2. **Checklists** — todos os arquivos encontrados na pasta `folder_checkl`. O código espera um layout determinado: usa a linha de índice 3 do DataFrame como cabeçalho (com a leitura padrão do pandas, normalmente corresponde à quinta linha física do Excel) e colunas como `Poço`, `Data`, `Poço foi amostrado?` e `Status do Poço`.
3. **Exportações EQuIS** — arquivos Excel mapeados no dicionário `sites`, dentro de `folder`. São lidas as abas `EFW2FSample`, `EFW2LabSMP`, `EFW2LabTST`, `EFW2LabRES` e `EFW2LabBCH`.

Os caminhos atuais são configuráveis no `.env` e, quando relativos, partem da raiz do projeto. As entradas de amostragem precisam ser fornecidas; nomes válidos no `.env` não garantem que os arquivos existam.

## Um exemplo antes e depois

Imagine uma linha da aba `_PA` assim:

| Grupos | Poços | Parâmetros |
|---|---|---|
| Metais | POCO-01, POCO-02 | Arsênio, Chumbo |

O programa transforma essa linha em quatro registros:

| Site (aba de origem) | Ponto | Grupo | Parâmetro |
|---|---|---|---|
| nome da aba | POCO-01 | Metais | Arsênio |
| nome da aba | POCO-01 | Metais | Chumbo |
| nome da aba | POCO-02 | Metais | Arsênio |
| nome da aba | POCO-02 | Metais | Chumbo |

Isso é necessário porque comparar fica mais simples quando cada linha representa **um ponto e um parâmetro**, em vez de várias listas separadas por vírgulas dentro de uma célula.

Depois, o script procura esses pontos e parâmetros nas outras fontes. Se o PA diz que POCO-01 deve ter Arsênio, mas não aparece resultado desse parâmetro para a amostra daquele ponto, essa combinação pode aparecer em `Análises ausentes`.

## Como ler algumas partes do Python

- `import pandas as pd`: carrega pandas e usa o apelido curto `pd`.
- `pd.read_excel(...)`: lê uma planilha Excel e a transforma em uma tabela para o programa.
- `df.loc[...]`: acessa linhas/colunas da tabela.
- `for ... in ...`: repete o bloco para cada aba, linha, ponto ou parâmetro.
- `data[...]`: em vários lugares, filtra a tabela para manter só as linhas que atendem a uma condição.
- `pd.concat([...])`: empilha linhas de tabelas, como juntar várias listas uma embaixo da outra.
- `pd.merge(...)`: combina tabelas que compartilham um identificador, como uma busca de correspondência.
- `to_excel(...)`: grava uma tabela em um arquivo Excel.

Os nomes `df`, `data`, `aba_sample` e `aba_result` são variáveis: cada uma guarda temporariamente uma tabela ou um resultado intermediário. A palavra “aba” nos nomes representa uma planilha/worksheet dentro de um arquivo Excel.

## Etapas do processamento

### 1. Lê e expande o plano

O script abre o Excel e seleciona as abas com `_PA` no nome. Para cada linha:

- obtém o grupo em `Grupos`;
- separa a lista de poços de `Poços` por vírgula, removendo espaços, quebras de linha e aspas;
- separa a lista de parâmetros de `Parâmetros` por vírgula;
- cria um registro para cada combinação de poço e parâmetro.

Por exemplo, uma linha que contenha dois poços e três parâmetros é expandida em até seis linhas na tabela resultante: cada poço é combinado com cada parâmetro. O grupo e o nome da aba de origem acompanham cada linha.

O nome da aba é guardado na coluna `Site`. Mais adiante, o script usa essa informação para decidir qual arquivo EQuIS consultar.

### 2. Padroniza nomes

Há substituições textuais para corrigir nomes conhecidos de parâmetros, incluindo variações de pontuação, espaços e erros de digitação. Algumas vírgulas internas de nomes químicos (como `1,2-...`) são trocadas temporariamente por pontos para que a divisão por vírgulas não separe partes de um mesmo nome; depois os pontos são convertidos de volta em vírgulas.

Também são aplicadas correções finais a alguns nomes de parâmetros e um mapeamento dos nomes das abas/sites para rótulos como `PE-1_PA` e `PE-2_PA`.

### 3. Seleciona a campanha/sites ativos

O dicionário `sites` é carregado de `MAIN_EQUIS_FILES`. Nos valores iniciais, somente `PE-1_PA` e `PE-2_PA` ficam ativos. A tabela do PA é filtrada para esses sites e salva no caminho `MAIN_PLAN_OUTPUT_FILE`, inicialmente:

`Plano de Amostragem Braskem BA - 3 campanha.xlsx`

O caminho de saída relativo parte da raiz do projeto, não do diretório do terminal.

Antes disso, o script mostra uma pausa interativa (`input`) para lembrar que a campanha deve ser conferida. É necessário responder ao prompt para a execução continuar.

**Para escolher sites:** inclua/remova pares no JSON `MAIN_EQUIS_FILES` e ajuste os checklists correspondentes em `MAIN_CHECKLIST_FILES`. Não precisa editar dicionários no Python.

### 4. Consolida os checklists

O script percorre todos os arquivos da pasta de checklist, tenta ajustar as primeiras colunas vazias e usa a linha de índice 3 do DataFrame como cabeçalho. Em seguida, combina os checklists em uma tabela e tenta associar cada arquivo ao site por meio do dicionário `sites_check`.

O dicionário ativo mapeia nomes de arquivos de PE-1 e PE-2. O mapeamento tem que corresponder exatamente aos nomes encontrados na pasta. Embora o script carregue essa tabela, uma falha nessa etapa pode impedir a geração da planilha final.

### 5. Compara o PA com EQuIS/EDD

Para cada site presente no PA filtrado, são abertas as abas EQuIS mencionadas acima. A comparação examina:

- **Pontos extras**: locais que aparecem nos samples EQuIS, mas não no PA do site;
- **Pontos ausentes**: pontos previstos no PA que não foram encontrados na aba `EFW2FSample`;
- **Parâmetros/análises extras**: resultados na aba `EFW2LabRES` que não aparecem no PA;
- **Parâmetros/análises ausentes**: parâmetros do PA sem resultado no EDD para as amostras do ponto;
- **Pontos sem amostra**: pontos do PA sem amostra correspondente, identificados como não presentes no EDD;
- **Não coletado sem justificativa**: pontos indicados como amostrados no checklist, mas sem resultados associados no EDD.

O script também usa a coluna `Status do Poço` do checklist como possível justificativa para pontos marcados como não amostrados. Essa associação exige que os nomes dos sites e dos poços coincidam entre checklist, PA e EQuIS.

Uma forma de pensar nas comparações:

| Pergunta feita pelo script | Resultado provável |
|---|---|
| “Este ponto existe no plano e na aba de amostras?” | Se não existir no EQuIS, entra em pontos ausentes. |
| “Este resultado químico estava previsto para este ponto?” | Se não estava, entra em análises extras. |
| “Este parâmetro previsto aparece nos resultados da amostra?” | Se não aparece, entra em análises ausentes. |
| “O checklist diz que coletou, mas há resultado no EDD para o ponto?” | Se o checklist diz sim e não há resultado, pode entrar em não coletado sem justificativa. |

Essas verificações dependem de identificadores iguais. `POCO-01` e `Poco-01`, por exemplo, podem ser tratados como nomes diferentes pelo código.

### 6. Gera o relatório

O caminho `MAIN_REPORT_OUTPUT_FILE` define o relatório, inicialmente `Relação de parâmetros Braskem BA - 2023T1.xlsx` na raiz do projeto. As abas criadas incluem:

| Aba | Interpretação |
|---|---|
| `Resumo` | Texto explicativo para quem analisa o relatório. |
| `Análises extras` | Resultados do EDD não previstos para o ponto no PA. |
| `Análises ausentes` | Parâmetros esperados no PA sem resultado localizado. |
| `Pontos extras` / `Pontos ausentes` | Diferenças de pontos entre PA e EQuIS. |
| `Nomes não encontrados no PA` / `Nomes não encontrados no edd` | Diferenças de nomes de parâmetros entre as fontes. |
| `Pontos extras geral` / `Pontos ausentes geral` | Comparação agregada dos pontos. |
| `Parâmetros extras geral` / `Parâmetros ausentes geral` | Comparação agregada dos parâmetros. |
| `Não coletado sem justificativa` | Amostragens indicadas como feitas no checklist, sem resultado localizado no compilado. |
| `Compilado EDDs` | Resultados reunidos dos arquivos EQuIS. |
| `Plano de amostragem` | PA normalizado usado na comparação. |
| `Check list` | Checklists consolidados usados na comparação. |

Algumas abas são deduplicadas antes da gravação. O script colore certas abas de vermelho e, ao final, abre a pasta de trabalho atual no navegador/gerenciador padrão; ele não abre automaticamente o Excel final pelo nome.

## Como interpretar “ausente” e “extra”

- **Ausente**: previsto no plano, mas não encontrado na fonte de resultados consultada. Pode ser coleta não realizada, atraso/incompletude do EDD, ou divergência de identificador/nome.
- **Extra**: encontrado no EDD, mas não localizado no plano usado pelo script.
- **Sem justificativa**: o script não encontrou uma justificativa de checklist associada para aquele caso. A equipe deve confirmar o contexto operacional antes de concluir que há uma falha.

Essas categorias são sinalizações para revisão, não conclusões automáticas sobre a causa.

Também é importante separar **“ponto ausente”** de **“análise ausente”**:

- ponto ausente: o próprio local não foi encontrado na lista de amostras;
- análise ausente: o local pode existir, mas falta um parâmetro/resultado associado a uma amostra.

## Configuração antes de executar

Confira no `.env`:

1. `MAIN_PLAN_FILE`: caminho do plano correto.
2. `MAIN_EQUIS_DIR` e `MAIN_EQUIS_FILES`: pasta, sites ativos e arquivos EQuIS.
3. `MAIN_CHECKLIST_DIR` e `MAIN_CHECKLIST_FILES`: pasta e nomes dos checklists.
4. `MAIN_SITE_RENAME`: mapeamento dos nomes de site.
5. `MAIN_GROUP_COLUMN`, `MAIN_POINTS_COLUMN`, `MAIN_PARAMETERS_COLUMN`: cabeçalhos do PA.
6. `MAIN_*_SHEET` e `MAIN_PLAN_SHEET_FILTER`: abas que serão lidas.
7. `MAIN_PLAN_OUTPUT_FILE` e `MAIN_REPORT_OUTPUT_FILE`: caminhos das saídas.

## Roteiro para conferir se a execução deu certo

1. Veja se o programa conseguiu abrir o plano, os checklists e os arquivos EQuIS sem erro.
2. Confira se `Plano de Amostragem Braskem BA - 3 campanha.xlsx` foi criado. Ele é uma versão achatada/normalizada do PA.
3. Confira se `Relação de parâmetros Braskem BA - 2023T1.xlsx` foi criado e se tem as abas indicadas acima.
4. Escolha alguns pontos conhecidos e compare-os manualmente entre o PA, o checklist e o EDD.
5. Investigue registros extras/ausentes antes de tratá-los como erros confirmados.

## Limitações/alertas do código atual

- O fluxo depende da campanha e dos layouts configurados. Os caminhos podem ser relativos ou absolutos.
- A seleção inicial inclui PE-1 e PE-2; outros sites podem ser incluídos em `MAIN_EQUIS_FILES`.
- O código assume que `Grupos`, `Poços`, `Parâmetros` e os cabeçalhos do checklist existem e têm valores em formatos esperados.
- Na verificação do formato do checklist há uma referência a `oi`, que não está definida. Se as colunas esperadas não forem encontradas, o script pode terminar com `NameError`.
- A variável de índice usada para recortar o checklist é reutilizada depois do laço que remove colunas. Isso torna a linha inicial dependente do estado do laço, em vez de ser uma constante explícita.
- Algumas comparações de ausências gerais usam a tabela `data` completa e não necessariamente o recorte do site corrente. Ao investigar diferenças por site, confira a aba e valide manualmente os registros.
- O teste que procura “não coletado sem justificativa” usa resultados compilados no nível do ponto/site; não garante que cada amostra ou cada parâmetro esperado esteja completo.
- As abas SMP, TST e BCH são lidas, mas não são usadas nas etapas mostradas do relatório.
- A grafia exata dos nomes de pontos e parâmetros influencia os resultados; as correções de texto cobrem apenas casos explicitamente listados no script.
