# `perfis.py` — encontrar o link de perfil para cada registro GIS

## Em poucas palavras

Imagine que existem duas listas:

- uma lista oficial do GIS dizendo quais poços/perfis devem existir;
- uma lista exportada do BI/SharePoint dizendo quais arquivos estão disponíveis e onde ficam.

Este script tenta descobrir qual arquivo corresponde a cada registro do GIS. Depois cria uma URL curta para o arquivo e grava o resultado em outro Excel.

É a versão mais antiga e mais simples do fluxo. A versão `perfis_rs.py` tem tratamento adicional para mais configurações de site e para registros sem correspondência dos dois lados.

## Entradas esperadas

O script usa uma pasta definida em `folder` e nela espera:

1. Uma planilha de controle, escolhida pela variável `site`:
   - BA: `BRASKEM_BA_PerfisConstrutivos_Controle.xlsx`
   - AL: `BRASKEM_AL_PerfisConstrutivos_Controle.xlsx`
   - Tupy: `TupyMaua_PerfisConstrutivos_Controle.xlsx`
2. Uma exportação de arquivos com nome `perfil {site}.xlsx`, por exemplo `perfil AL.xlsx`. Espera colunas como `Name` e `Folder Path`.

O valor inicial é `PERFIS_SITE=AL` no [`.env`](../.env). `PERFIS_INPUT_DIR` define a pasta de entrada, `PERFIS_OUTPUT_DIR` a de saída, `PERFIS_CONTROL_FILES` o mapa de arquivos por site e `PERFIS_LINKS_FILE` a lista exportada. Campos `PERFIS_*_COLUMN` configuram os cabeçalhos, convertidos pelo script para os nomes internos.

A planilha de controle precisa ter campos de identificação do local e site. Se não tiver `File_Name`, o script usa `SYS_LOC_CODE` como substituto. O código também espera colunas como `Site`, `GlobalID`, `ObjectID` e `SYS_LOC_CODE` em trechos posteriores.

## Exemplo de como encontra uma correspondência

Imagine que a planilha GIS tenha:

| Site | File_Name |
|---|---|
| AL | POCO-10 |

e a exportação de arquivos tenha:

| Name | Folder Path |
|---|---|
| POCO-10.pdf | .../AL/ |

O script tira `.pdf` do nome, calcula o site com base na pasta e cria uma chave para cada lado. Se as chaves forem iguais, considera que encontrou o arquivo para aquele registro GIS.

Em termos simples, a **chave** é uma etiqueta montada com os campos de site e nome do local. O `merge` procura linhas que tenham a mesma etiqueta, como uma procura/PROCV entre tabelas.

## O caminho dos dados no programa

1. **Escolhe o site** — a variável `site` decide qual planilha de controle será lida. Atualmente está configurada como `AL`.
2. **Lê duas tabelas** — `gis` guarda o controle, e `links` guarda a exportação dos arquivos.
3. **Arruma nomes** — remove espaços ou caracteres específicos para tentar deixar os nomes comparáveis.
4. **Cria a chave** — combina site e nome do arquivo/local em cada tabela.
5. **Junta as tabelas** — o `pd.merge(..., 'left', ...)` mantém todos os registros GIS e acrescenta os campos do arquivo quando há correspondência.
6. **Cria a URL** — combina o caminho da pasta e o nome do arquivo e pede ao TinyURL para gerar um endereço curto.
7. **Salva o Excel** — inclui o resultado e tenta abrir os arquivos gerados.

Alguns nomes para reconhecer no código:

- `gis`: tabela de controle do GIS.
- `links`: lista de arquivos exportada.
- `key`: etiqueta usada para relacionar as duas tabelas.
- `link antigo`: endereço original montado pelo script.
- `link`: endereço encurtado.
- `folder`: pasta que o código usa para ler e salvar arquivos.

## Etapas do processamento

### 1. Carrega e normaliza os nomes

O controle do GIS e a exportação de links são carregados com pandas. O script:

- remove espaços do nome do arquivo `Name`;
- usa o nome sem os quatro últimos caracteres como `File_Name` (na prática, costuma remover uma extensão como `.pdf`);
- troca `|` por `I` no nome de arquivo dos dois lados;
- extrai um rótulo de site do caminho da exportação;
- remove espaços e hífens nos nomes de site para comparar os dois lados.

Essas normalizações são simples e dependem de um padrão consistente de nome de arquivo e estrutura de pastas.

### 2. Cria uma chave e faz o cruzamento

As chaves são:

- da exportação: `site + File_Name`;
- do controle GIS: `Site + File_Name`.

O `merge` é do tipo left: mantém todos os registros do GIS e acrescenta os dados de link quando acha a mesma chave. Pense nele como: “comece pela lista GIS e procure, para cada linha, um arquivo correspondente”. Portanto, a planilha resultante é orientada pelo conteúdo do controle GIS.

### 3. Produz uma lista de controle

Os registros do GIS sem correspondência são selecionados e gravados em:

`Pontos da planilha de controle não encontrados no sharepoint.xlsx`

O script tenta abrir esse arquivo. Esse relatório inclui os campos `GlobalID`, `ObjectID`, `SYS_LOC_CODE`, nome de arquivo e site.

Apesar do nome do relatório mencionar SharePoint, a entrada concreta usada pelo código é a planilha `perfil {site}.xlsx` exportada/listada previamente.

### 4. Cria e encurta os links

Para cada registro que teve correspondência, monta `link antigo` juntando `Folder Path` e `Name`. Em seguida, chama `pyshorteners` com o serviço TinyURL.

Se uma tentativa falhar, o código espera 5 segundos e tenta novamente; se falhar outra vez, espera 10 segundos e tenta uma terceira vez. Se a última tentativa falhar, a exceção não é tratada e a execução para.

Os links encurtados são colocados como texto HTML, por exemplo `<a href="...">Disponível</a>`. `PERFIS_OUTPUT_FILE` e `PERFIS_QC_FILE` definem as saídas dentro de `PERFIS_OUTPUT_DIR`, inicialmente a mesma pasta das entradas. Os intervalos de repetição são configurados em `LINK_RETRY_WAIT_1` e `LINK_RETRY_WAIT_2`.

## O que conferir antes de executar

- Se `site` corresponde ao nome do arquivo de controle e ao arquivo `perfil {site}.xlsx`.
- Se o caminho `folder` existe e contém todos os arquivos.
- Se o último diretório relevante de `Folder Path` representa o site esperado. O código usa a penúltima parte do caminho separado por `/`.
- Se os identificadores e nomes dos arquivos são consistentes entre o controle e a exportação.
- Se há acesso à internet para o serviço TinyURL.
- Se o arquivo de saída já existe e pode ser substituído.

Essas opções são preenchidas no grupo `PERFIS_` do `.env`, não mais editando caminhos no Python. Caminhos relativos partem da raiz do projeto. As pastas de saída precisam existir.

## Como saber se a correspondência foi encontrada?

Depois do cruzamento, os arquivos que não foram localizados têm campos como `File_Name_y` vazios (o `_y` é criado porque há colunas de mesmo nome nos dois lados do cruzamento). O script grava esses registros em uma planilha de controle.

Se a planilha final tiver muitas linhas sem `link`, investigue primeiro:

1. se o nome do arquivo bate com o identificador do poço;
2. se o script extraiu o nome de site correto de `Folder Path`;
3. se espaços, hífens ou o caractere `|` foram tratados da mesma forma nos dois arquivos.

## Limitações/alertas do código atual

- Não existe no script um relatório equivalente para arquivos da exportação que não tenham correspondência no GIS; ele lista principalmente registros do GIS sem link encontrado.
- O caminho `link antigo` é montado concatenando caminho e nome exatamente como estão. Se `Folder Path` não terminar com separador, o endereço pode ficar malformado.
- O código trata qualquer falha na primeira chamada com uma repetição, mas não registra um relatório de falhas por linha. A última falha interrompe o script.
- A mensagem de “nenhuma correspondência” executa `oi`, uma variável inexistente. Nesse caso, a execução termina com `NameError` em vez de sair com uma mensagem controlada.
- O conteúdo `<a href="...">Disponível</a>` é uma string HTML dentro do Excel; o código não cria uma fórmula de hyperlink nativa do Excel. Confira no destino se esse formato é interpretado como link clicável.
- A extensão é removida fixamente pelos quatro últimos caracteres; nomes com extensão diferente ou sem extensão podem ser tratados incorretamente.
- `math` e algumas variáveis/imports não são usados no fluxo atual.

## Diferença para `perfis_rs.py`

`perfis.py` foi feito para um conjunto menor de opções e só destaca os registros do GIS sem correspondência. [`perfis_rs.md`](./perfis_rs.md) descreve a outra versão, que inclui uma comparação nos dois sentidos e regras para mais sites. Isso não significa que seja seguro trocar de um arquivo para o outro sem conferir as colunas e caminhos: as configurações não são idênticas.
