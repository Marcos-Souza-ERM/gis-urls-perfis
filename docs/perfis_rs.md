# `perfis_rs.py` — encontrar perfis, conferir diferenças e criar links

## Em poucas palavras

Este script também trabalha com duas listas — controle GIS e arquivos disponíveis —, mas faz uma checagem extra: mostra tanto registros do GIS sem arquivo correspondente quanto arquivos sem registro GIS correspondente.

Depois do controle, tenta gerar links curtos para os perfis encontrados.

Apesar do sufixo `_rs`, o valor inicial é `PERFIS_RS_SITE=BA` no [`.env`](../.env). O arquivo contém ramificações para BA, AL, AC, Tupy e RS, mas cada site depende de nomes de colunas e regras de caminho que precisam ser conferidos.

## Entradas e saídas

### Entradas configuradas

- Pasta de entrada: `PERFIS_RS_INPUT_DIR`, inicialmente `inputs`.
- Pasta de saída: `PERFIS_RS_OUTPUT_DIR`, inicialmente `outputs`.
- Planilha de controle:
  - BA: `BRASKEM_BA_PerfisConstrutivos_Controle.xlsx`
  - AL: `BRASKEM_AL_PerfisConstrutivos_Controle.xlsx`
  - AC: `ACELEN-BA Poços com perfis construtivos.xlsx`
  - Tupy: `TupyMaua_PerfisConstrutivos_Controle.xlsx`
  - RS: `BRK-RS-Controle perfis das locations.xlsx`; nesse caso, o código lê todas as abas e as concatena.
- Exportação do BI com nome `perfil {site}.xlsx`, por exemplo `perfil BA.xlsx`.

A exportação espera os cabeçalhos definidos em `PERFIS_RS_LINK_NAME_COLUMN` e `PERFIS_RS_LINK_PATH_COLUMN`, inicialmente `Name` e `Folder Path`. O controle usa `PERFIS_RS_LOCATION_COLUMN` e `PERFIS_RS_SITE_COLUMN`, inicialmente `ID Automação` e `Site`. O mapa de nomes de controles é `PERFIS_RS_CONTROL_FILES`; a lista exportada é `PERFIS_RS_LINKS_FILE`.

### Saídas

- Se encontrar diferenças, grava `QC - Encurtamento de links {site}.xlsx` em `outputs`, com abas para itens do controle GIS não encontrados na lista de links e itens da lista de links não encontrados no controle GIS.
- Grava `Links encurtados {site}-{data}.xlsx` em `outputs`, incluindo as colunas originais do GIS e os dados/links associados.
- Tenta abrir as planilhas geradas.

O diretório `outputs` precisa existir antes da execução; o script não o cria.

## Exemplo simples

Imagine esta lista do GIS:

| Site | ID Automação |
|---|---|
| BA | POCO-10 |
| BA | POCO-11 |

E esta lista exportada de arquivos:

| Site extraído da pasta | Nome do arquivo |
|---|---|
| BA | POCO-10.pdf |
| BA | POCO-12.pdf |

Depois de retirar a extensão, o script compara etiquetas formadas por site + nome:

- POCO-10 aparece nas duas listas: é uma correspondência; o script tentará gerar o link.
- POCO-11 aparece apenas no GIS: vai para o relatório de itens do GIS sem arquivo encontrado.
- POCO-12 aparece apenas na exportação: vai para o relatório de arquivos sem item GIS correspondente.

O relatório ajuda a localizar diferenças; não decide se POCO-11 deveria ter um arquivo ou se POCO-12 deveria ser cadastrado no GIS.

## Nomes para reconhecer no arquivo Python

- `gis`: tabela de controle carregada do Excel.
- `links`: tabela com os nomes e caminhos dos arquivos.
- `site`: configuração que escolhe qual cliente/local e qual regra de extração será usada.
- `gis_location_column`: nome da coluna com o identificador do local na tabela GIS.
- `key`: etiqueta calculada para procurar correspondência nos dois lados.
- `not_in_sharepoint`: linhas do GIS que não tiveram chave equivalente na lista de arquivos.
- `not_in_gis_file`: arquivos cuja chave não aparece na tabela GIS.
- `link antigo`: endereço original montado com caminho + nome do arquivo.
- `link`: resultado retornado pelo encurtador TinyURL.

## Fluxo detalhado

### 1. Seleciona a planilha de controle

O bloco baseado em `site` escolhe qual planilha carregar. No caso RS, cada aba do Excel é lida e combinada em uma única tabela. O resto do script trabalha sobre o DataFrame `gis`.

### 2. Prepara nomes dos arquivos e locais

Na tabela de links:

- remove espaços externos do campo `Name`;
- cria `Links_File_Name` removendo os quatro últimos caracteres do nome;
- substitui `|` por `I`.

No controle GIS, faz a mesma substituição no identificador `ID Automação` e remove espaços externos. Também remove espaços externos no campo `Site`.

### 3. Descobre o site a partir do caminho

As regras dependem do valor de `site`:

- **BA**: pega o nome da pasta pai de `Folder Path` e corrige alguns nomes como `PE1` para `PE-1`, `PVC1` para `PVC` e `UA-3` para `UA3`.
- **AC**: procura a parte do caminho depois de `01. Perfis Construtivos & Sondagens/` e corrige nomes específicos de subpastas.
- **RS**: tenta obter site e subárea a partir da parte do caminho que segue `Boring Logs - WebGIS`.
- **AL/Tupy e demais casos**: usa uma posição fixa do caminho separado por `/`.

Os caminhos exportados precisam seguir o padrão esperado pela regra do site. O script imprime os valores de site encontrados para ajudar a verificar o mapeamento.

### 4. Monta chaves de cruzamento

Para BA, AC e o caso padrão, a chave geralmente concatena site e nome do local/arquivo. Para RS, a chave inclui site, subárea e nome do local. A concatenação não adiciona separadores entre os componentes; os padrões dos identificadores precisam evitar ambiguidades.

O objetivo é identificar se cada item do GIS tem arquivo de perfil correspondente e vice-versa.

O código não “adivinha” que dois nomes são iguais. Por exemplo, se de um lado a chave ficou `BAPOCO10` e do outro `BAPO-10`, são textos diferentes e não haverá correspondência. As regras de normalização precisam produzir o mesmo texto em ambos os lados.

### 5. Gera o relatório de QC

O script calcula:

- `not_in_sharepoint`: registros do controle GIS cuja chave não está na lista de links;
- `not_in_gis_file`: registros da lista de links cuja chave não está no controle GIS.

Se qualquer lista tiver conteúdo, grava as duas em abas do relatório QC. Esse relatório mostra divergências nos dois sentidos; ele não altera nem sincroniza as fontes.

### 6. Associa, normaliza e encurta as URLs

Faz um left merge do controle GIS com os campos `Name`, `Folder Path`, `Links_File_Name` e `key` da lista de links. Para cada correspondência, concatena `Folder Path` e `Name` para montar o link original.

A função `normalize_url` separa os componentes da URL e codifica caracteres especiais do caminho e da query string antes do envio ao TinyURL. O script tenta encurtar cada URL; em caso de falha, espera 5 segundos e tenta de novo, depois espera mais 10 segundos antes da última tentativa. Uma falha final é propagada e pode interromper o restante do processamento.

O texto resultante é escrito como `<a href="URL_CURTA">Disponível</a>`. Isso representa marcação HTML em uma célula; não é uma fórmula nativa do Excel.

## Como adaptar para outro site

Antes de mudar `site`, confira em conjunto:

1. Se o bloco de leitura escolhe o arquivo de controle correto.
2. Se `gis_location_column` e `site_column` existem na planilha desse site.
3. Se `perfil {site}.xlsx` está na pasta de entrada e tem as colunas esperadas.
4. Se a regra de extração do site/subárea corresponde ao caminho real exportado.
5. Se a chave formada em ambos os lados identifica o mesmo local sem colisões.
6. Se `outputs` existe e o usuário tem permissão para gravar nela.
7. Se a rede permite chamadas ao TinyURL.

Preencha as opções no grupo `PERFIS_RS_` do `.env`. Não basta trocar o site: alguns ramos usam colunas diferentes ou regras específicas de caminho. Os marcadores RS/AC também podem ser definidos em `PERFIS_RS_RS_PATH_MARKER` e `PERFIS_RS_AC_PATH_MARKER`. As saídas são configuradas em `PERFIS_RS_QC_FILE` e `PERFIS_RS_OUTPUT_FILE`; os intervalos de repetição, em `LINK_RETRY_WAIT_1` e `LINK_RETRY_WAIT_2`.

## O que a função `normalize_url` faz?

No começo do arquivo há uma função que recebe uma URL e separa partes como protocolo, domínio, caminho e parâmetros. Ela codifica caracteres especiais no caminho e na query para que espaços ou caracteres acentuados não deixem a URL malformada. Depois reúne as partes em uma URL normalizada.

Exemplo conceitual: um espaço no caminho pode precisar ser codificado como `%20`. A função prepara a URL; quem efetivamente gera o endereço curto é a biblioteca `pyshorteners` consultando o serviço TinyURL.

## Limitações/alertas do código atual

- Caminhos relativos partem da raiz do projeto. As entradas e a pasta de saída precisam existir.
- `PERFIS_RS_SUBAREA_COLUMN` pode ficar vazio nos demais sites, mas é obrigatório para RS. O programa interrompe com uma mensagem de configuração se RS for escolhido sem preencher esse cabeçalho.
- No ramo RS, a extração assume uma estrutura específica de `Folder Path`; nomes/pastas fora do padrão podem gerar chaves incorretas.
- As colunas-chave são concatenadas sem delimitadores. Identificadores diferentes podem teoricamente produzir o mesmo texto concatenado.
- A remoção fixa dos quatro últimos caracteres pressupõe que todos os nomes tenham a mesma extensão de quatro caracteres, como `.pdf`.
- O arquivo QC só é escrito se houver pelo menos uma divergência; ausência do relatório pode significar que não houve divergências, mas também é importante confirmar que o script concluiu normalmente.
- O hyperlink final é armazenado como texto HTML, não como hyperlink nativo do Excel. Verifique a forma como a aplicação de destino interpreta esse conteúdo.
- Se o TinyURL falhar na terceira tentativa, não há tratamento local para continuar os demais registros nem arquivo de falhas parciais.
- Algumas importações/variáveis e comentários representam tentativas antigas ou não são usados no processamento atual.

## Checklist mental de execução

Antes de confiar na planilha:

1. O `site` escolhido corresponde às entradas corretas?
2. Os sites que o código imprimiu para GIS e BI fazem sentido?
3. O relatório QC contém diferenças que precisam ser resolvidas?
4. Os registros com link têm o arquivo e a URL esperados?
5. O script chegou à mensagem de processo finalizado?

Se houver erro antes do final, não assuma que o Excel de links está completo; o encurtamento pode ter parado no meio do processamento.
