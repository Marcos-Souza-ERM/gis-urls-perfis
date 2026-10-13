# GIS URL Perfis — guia do projeto

Automação em Python para relacionar registros de perfis construtivos a URLs e comparar planos de amostragem com outras fontes de dados. Os programas leem planilhas Excel, organizam ou comparam seus registros e geram novas planilhas para revisão.

**O projeto é um conjunto de scripts independentes, não uma aplicação web.** Não há servidor, interface gráfica própria, banco de dados ou comando que execute todos os fluxos automaticamente.

## Sumário

- [Programas e responsabilidades](#programas-e-responsabilidades)
- [Origem dos dados e escopo](#origem-dos-dados-e-escopo)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Pré-requisitos e instalação](#pré-requisitos-e-instalação)
- [Configuração de ambiente](#configuração-de-ambiente)
- [Processamento de perfis](#processamento-de-perfis)
- [Comparação de planos de amostragem](#comparação-de-planos-de-amostragem)
- [Conferência do plano contra o laboratório](#conferência-do-plano-contra-o-laboratório)
- [Execução](#execução)
- [Testes e validação](#testes-e-validação)
- [Limitações e solução de problemas](#limitações-e-solução-de-problemas)
- [Glossário e documentação detalhada](#glossário-e-documentação-detalhada)

## Programas e responsabilidades

| Arquivo | Responsabilidade | Entradas | Saídas |
|---|---|---|---|
| [`perfis_rs.py`](./src/perfis_rs.py) | Relacionar locais do GIS aos arquivos de perfis, identificar diferenças nos dois sentidos e encurtar URLs. | Controle GIS e lista de arquivos exportada do BI. | Excel de links e relatório QC, quando houver diferenças. |
| [`perfis.py`](./src/perfis.py) | Executar outra versão do relacionamento de perfis, com layout e normalização próprios. | Controle GIS e lista de arquivos. | Excel de links e lista de registros GIS sem correspondência. |
| [`pa_check.py`](./src/pa_check.py) | Comparar duas versões de um plano de amostragem. | Dois planos com site, ponto e parâmetro. | Excel com as duas versões e observações de diferenças. |
| [`main.py`](./src/main.py) | Expandir o plano e compará-lo com dados de amostras/resultados e checklists. | Plano de amostragem, arquivos EQuIS/EDD e checklists. | Plano normalizado e relatório de análises, pontos e parâmetros extras/ausentes. |
| [`config.py`](./src/config.py) | Carregar e validar configurações de ambiente usadas pelos scripts. | `.env` na raiz e variáveis do sistema/terminal. | Valores de configuração; não gera relatórios. |
| [`tests/test_environment.py`](./tests/test_environment.py) | Verificar configurações e executar os fluxos com dados fictícios. | Planilhas temporárias criadas pelos testes. | Resultado dos testes no terminal. |

Escolha o programa de acordo com a tarefa. **Não é necessário executar os quatro scripts em sequência.** Em especial, o nome `main.py` não significa que ele coordena os outros programas.

## Origem dos dados e escopo

Os arquivos de entrada precisam ser obtidos previamente pela equipe responsável e disponibilizados em uma pasta acessível ao computador.

| Fonte | Informação fornecida |
|---|---|
| Equipe GIS | Planilha de controle dos locais/perfis. |
| Exportação BI/PBIX | Lista de nomes e caminhos dos arquivos de perfis no SharePoint, conforme os comentários do código. |
| Responsáveis pelo plano | Planos de amostragem a comparar ou conferir. |
| Equipe de campo/laboratório | Checklists e arquivos EQuIS/EDD da campanha. |

O projeto **não implementa a exportação dessas fontes**, autenticação no SharePoint, consulta direta ao GIS ou obtenção automática de dados do laboratório.

```text
Obter/exportar as planilhas
           |
Disponibilizar os arquivos em uma pasta
           |
Configurar caminhos, arquivos e colunas no .env
           |
Executar o script correspondente
           |
Revisar as planilhas geradas
```

### O que os scripts de perfis não fazem

- Não baixam nem abrem PDFs, imagens ou documentos Word.
- Não extraem o conteúdo dos documentos de perfis.
- Não percorrem o SharePoint para descobrir arquivos.
- Não comprovam que o documento existe, está correto ou é acessível ao usuário.
- Não atualizam automaticamente o GIS ou os arquivos de origem.

Eles leem **informações sobre os documentos em um Excel**. O conteúdo de `Folder Path` e `Name` é usado para construir a URL do documento. O serviço TinyURL recebe essa URL e devolve uma URL curta; isso não concede acesso ao documento nem substitui as permissões do SharePoint.

## Estrutura do projeto

```text
gis-url-perfis\
├── .env                         Configuração local, ignorada pelo Git
├── .env.example                 Modelo comentado de configuração
├── .gitignore                   Exclusões de arquivos locais
├── requirements.txt             Dependências Python
├── README.md                     Guia principal
├── src\
│   ├── __init__.py               Pacote dos programas
│   ├── config.py                 Leitura e validação de configuração
│   ├── perfis_rs.py              Perfis com QC nos dois sentidos
│   ├── perfis.py                 Outra versão do fluxo de perfis
│   ├── pa_check.py               Comparação de dois planos
│   └── main.py                   Conferência de amostragem
├── inputs\                      Entradas locais, se assim configuradas
├── outputs\                     Resultados locais, se assim configurados
├── old_runs\                    Arquivos de execuções anteriores
├── tests\
│   └── test_environment.py
└── docs\                        Guias detalhados de cada fluxo
```

As pastas de dados não são obrigatórias com esses nomes: os caminhos podem ser alterados no ambiente. Arquivos disponíveis em uma cópia local do projeto podem não acompanhar uma nova cópia do repositório. Resultados antigos não demonstram que uma nova execução foi concluída.

As pastas `inputs`, `outputs` e `old_runs` são preservadas no Git por arquivos `.gitkeep`, mas seu conteúdo real não é versionado. Planilhas `.xlsx`, `.xls` e `.xlsm`, configurações locais, ambientes virtuais e arquivos do editor são ignorados. Os testes geram dados fictícios em pastas temporárias, sem depender das planilhas privadas.

## Pré-requisitos e instalação

- Python instalado. A configuração atual foi validada em Python 3.12; não há uma matriz de certificação de versões.
- Permissões de leitura nas entradas e de escrita nas saídas.
- Planilhas compatíveis com o fluxo escolhido.
- Acesso de rede ao TinyURL para executar o encurtamento real.
- Autorização da organização para enviar URLs ao serviço externo.

O Microsoft Excel não é necessário para o Python processar as planilhas. Um leitor de Excel é útil para revisar as saídas.

As dependências estão em [`requirements.txt`](./requirements.txt): `pandas`, `openpyxl`, `pyshorteners` e `python-dotenv`. O pandas está limitado à versão principal 2 por compatibilidade com as chamadas atuais dos scripts.

### Instalação no Windows / PowerShell

Abra o terminal na raiz do projeto e execute:

```powershell
py --version
py -m venv .venv-teste
.\.venv-teste\Scripts\python.exe -m pip install -r .\requirements.txt
```

Execute cada etapa somente se a anterior terminar sem erro. Se `py` não estiver disponível, instale/configure Python antes de continuar. Use um ambiente criado no seu computador; ambientes virtuais copiados de outro computador podem conter caminhos inválidos.

Para conferir os imports sem iniciar o processamento:

```powershell
.\.venv-teste\Scripts\python.exe -c "import pandas, openpyxl, pyshorteners, dotenv; print('Dependências disponíveis')"
```

Esse comando valida somente a disponibilidade das bibliotecas.

## Configuração de ambiente

Os quatro programas carregam automaticamente o arquivo `.env` localizado na raiz do projeto. Se ele ainda não existir, copie o [modelo comentado](./.env.example):

```powershell
Copy-Item -LiteralPath .\.env.example -Destination .\.env
```

**Não substitua um `.env` já configurado sem revisar seus valores.**

| Prefixo | Configurações usadas por |
|---|---|
| `PERFIS_RS_` | `perfis_rs.py` |
| `PERFIS_` | `perfis.py` |
| `PA_CHECK_` | `pa_check.py` |
| `MAIN_` | `main.py` |
| `LINK_RETRY_` | Ambos os scripts de perfis. |

Configure o grupo do programa que será executado. Variáveis de processamento, como tabelas em memória e contadores, continuam no Python.

### Caminhos e nomes de arquivos

Exemplo para usar uma pasta fora do projeto:

```dotenv
PERFIS_RS_INPUT_DIR='C:\Dados\Perfis'
PERFIS_RS_OUTPUT_DIR='C:\Dados\Resultados'
PERFIS_RS_SITE=BA
```

- Caminhos relativos, como `inputs`, são resolvidos a partir da raiz do projeto, não da pasta atual do terminal.
- Use aspas simples em caminhos Windows. Não é necessário acrescentar barra ao final da pasta.
- As entradas e pastas de saída precisam existir; o `.env` não cria arquivos nem diretórios.
- Os scripts leem caminhos locais ou de rede acessíveis pelo sistema. Uma URL de navegador do SharePoint não substitui o caminho da pasta de entrada.
- Variáveis definidas no sistema/terminal têm prioridade sobre valores do `.env`.
- Depois de alterar o arquivo, salve e execute o programa novamente.

Nos scripts de perfis, a pasta e o nome do arquivo são configurações separadas:

```dotenv
PERFIS_RS_CONTROL_FILES='{"BA":"BRASKEM_BA_PerfisConstrutivos_Controle.xlsx"}'
PERFIS_RS_LINKS_FILE='perfil {site}.xlsx'
```

O primeiro valor é um mapa JSON de site para nome de arquivo. O segundo é um modelo: com `PERFIS_RS_SITE=BA`, o programa procura `perfil BA.xlsx` na pasta de entrada.

Mapas devem ser objetos JSON válidos, em uma única linha, com chaves e valores não vazios. Os modelos de nomes de saída dos perfis aceitam `{site}` e `{date}`; a data é preenchida no formato `AAAA-MM-DD`.

O `.env` é ignorado pelo Git. Compartilhe [`.env.example`](./.env.example), não configurações locais ou credenciais.

## Processamento de perfis

### Entradas e colunas de `perfis_rs.py`

O script carrega as tabelas Excel, mas usa principalmente estas colunas:

| Entrada | Cabeçalho inicial | Variável de ambiente | Finalidade |
|---|---|---|---|
| Controle GIS | `ID Automação` | `PERFIS_RS_LOCATION_COLUMN` | Identificar o local e procurar o arquivo correspondente. |
| Controle GIS | `Site` | `PERFIS_RS_SITE_COLUMN` | Identificar a unidade e compor a chave de correspondência. |
| Controle GIS | Cabeçalho a informar | `PERFIS_RS_SUBAREA_COLUMN` | Complementar a chave para RS; obrigatório apenas nesse site. |
| Lista de arquivos | `Name` | `PERFIS_RS_LINK_NAME_COLUMN` | Obter o nome do documento. |
| Lista de arquivos | `Folder Path` | `PERFIS_RS_LINK_PATH_COLUMN` | Obter o endereço da pasta e extrair site/subárea conforme a regra selecionada. |

Os demais campos do controle são mantidos na tabela resultante, mas não determinam a correspondência. Campos da exportação como `Extension`, `Date accessed` e `Date modified` não são usados para selecionar o documento ou encurtar URLs.

O modelo suporta `BA`, `AL`, `AC`, `Tupy` e `RS`. O valor configurado seleciona tanto o arquivo de controle quanto a regra de interpretação dos caminhos:

- **BA:** extrai a pasta pai e aplica substituições específicas nos nomes das unidades.
- **AC:** usa o marcador configurado em `PERFIS_RS_AC_PATH_MARKER`.
- **RS:** usa `PERFIS_RS_RS_PATH_MARKER`, site e subárea; combina todas as abas do controle.
- **AL/Tupy:** usam uma posição fixa do caminho separado por `/`; confira se isso corresponde à estrutura real da exportação.

Trocar apenas o site não garante compatibilidade com uma nova planilha. Confira cabeçalhos, arquivos e estrutura dos caminhos.

### Exemplo de correspondência

Controle GIS:

| Site | ID Automação |
|---|---|
| PE-1 | POCO-01 |

Lista de arquivos:

| Name | Folder Path |
|---|---|
| POCO-01.pdf | https://example.com/PE1/ |

Na regra BA, o programa:

1. Remove os quatro últimos caracteres de `POCO-01.pdf`, obtendo `POCO-01`.
2. Extrai `PE1` do caminho e normaliza para `PE-1`.
3. Compara a chave site + identificador nas duas tabelas.
4. Monta `https://example.com/PE1/POCO-01.pdf`.
5. Normaliza a URL e solicita seu encurtamento ao TinyURL.
6. Grava o resultado em um novo Excel.

O exemplo é fictício. O script não abre o PDF.

### Diferenças da versão `perfis.py`

Essa versão aceita `BA`, `AL` e `Tupy`, mas usa um layout diferente:

| Campo interno | Configuração | Uso |
|---|---|---|
| `File_Name` | `PERFIS_FILE_COLUMN` | Identificador usado na chave. |
| `SYS_LOC_CODE` | `PERFIS_LOCATION_COLUMN` | Substituto quando a coluna `File_Name` não existe e campo do relatório. |
| `Site` | `PERFIS_SITE_COLUMN` | Unidade usada na chave. |
| `GlobalID` / `ObjectID` | `PERFIS_GLOBAL_ID_COLUMN` / `PERFIS_OBJECT_ID_COLUMN` | Campos do relatório de registros sem correspondência. |
| `Name` / `Folder Path` | `PERFIS_LINK_NAME_COLUMN` / `PERFIS_LINK_PATH_COLUMN` | Nome e endereço do documento. |

Remove todos os espaços do nome do arquivo; a versão RS remove apenas espaços nas extremidades. As regras de normalização de site também diferem. Por isso, as versões não são intercambiáveis sem conferir as entradas.

### Saídas e encurtamento

- A versão RS gera QC nos dois sentidos: registros GIS sem arquivo e arquivos sem registro GIS. O QC só é gravado quando há divergências.
- A versão antiga grava uma lista de registros GIS sem correspondência.
- As duas versões fazem um cruzamento que mantém os registros GIS e adiciona os dados dos arquivos encontrados.
- As saídas de links são configuradas em `PERFIS_RS_OUTPUT_FILE` ou `PERFIS_OUTPUT_FILE`, dentro da pasta de saída correspondente.
- O código grava `link antigo` e `link`. O último contém texto HTML como `<a href="URL_CURTA">Disponível</a>`, **não um hyperlink nativo do Excel**.
- Há até três tentativas por URL. `LINK_RETRY_WAIT_1` e `LINK_RETRY_WAIT_2` configuram as esperas antes da segunda e terceira tentativas. A última falha pode interromper o processamento.

## Comparação de planos de amostragem

[`pa_check.py`](./src/pa_check.py) compara duas planilhas informadas em `PA_CHECK_YAEL_FILE` e `PA_CHECK_TAISE_FILE`.

| Configuração | Cabeçalho inicial |
|---|---|
| `PA_CHECK_SITE_COLUMN` | `Site` |
| `PA_CHECK_POINT_COLUMN` | `Ponto` |
| `PA_CHECK_PARAMETER_COLUMN` | `Parâmetro` |

As configurações de cabeçalho se aplicam aos dois arquivos. O programa compara:

- A combinação site + ponto + parâmetro.
- Pontos isoladamente, sem considerar o site.
- Parâmetros isoladamente, sem considerar o ponto ou site.

Exemplo: se um plano pede Arsênio em P-01 e o outro pede Chumbo em P-01, o ponto existe nos dois, mas as combinações são diferentes.

O resultado é salvo em `PA_CHECK_OUTPUT_FILE`, com abas `PA Yael` e `PA Taise`. O programa não escolhe a versão correta, não cria um plano aprovado e não consulta dados do laboratório.

## Conferência do plano contra o laboratório

[`main.py`](./src/main.py) cruza três fontes: planejamento, registros de campo e dados de amostras/resultados.

| Fonte | Configuração principal | Estrutura esperada |
|---|---|---|
| Plano | `MAIN_PLAN_FILE` | Abas selecionadas por `MAIN_PLAN_SHEET_FILTER`; cabeçalhos configurados em `MAIN_GROUP_COLUMN`, `MAIN_POINTS_COLUMN`, `MAIN_PARAMETERS_COLUMN`. |
| EQuIS | `MAIN_EQUIS_DIR` e `MAIN_EQUIS_FILES` | Um arquivo por site ativo, contendo as abas configuradas. |
| Checklists | `MAIN_CHECKLIST_DIR` e `MAIN_CHECKLIST_FILES` | Layout específico de cabeçalhos e campos de coleta; todos os arquivos da pasta são lidos. |

As abas EQuIS iniciais são `EFW2FSample`, `EFW2LabSMP`, `EFW2LabTST`, `EFW2LabRES` e `EFW2LabBCH`. As colunas utilizadas incluem:

- Amostras: `sys_sample_code`, `sys_loc_code`, `sample_date`.
- Resultados: `sys_sample_code`, `chemical_name`.
- Checklists, após o ajuste de cabeçalho: `Poço`, `Data`, `Poço foi amostrado?`, `Status do Poço`.

Esses campos internos e o layout dos checklists não são todos configuráveis pelo ambiente.

O programa expande listas de pontos e parâmetros em combinações individuais. Dois pontos e dois parâmetros, por exemplo, geram quatro registros de planejamento. Depois procura diferenças em relação ao EDD e associa possíveis justificativas de não coleta.

As saídas são `MAIN_PLAN_OUTPUT_FILE` e `MAIN_REPORT_OUTPUT_FILE`. O relatório inclui análises/pontos/parâmetros extras e ausentes, compilado dos EDDs, plano e checklists. **Ausência no relatório é uma sinalização para investigação**, não comprovação automática de uma falha operacional.

## Execução

Após instalar dependências, configurar o ambiente e disponibilizar as entradas, escolha **um** comando:

```powershell
# Relacionar perfis e gerar QC nos dois sentidos
.\.venv-teste\Scripts\python.exe -m src.perfis_rs

# Outra versao do relacionamento de perfis
.\.venv-teste\Scripts\python.exe -m src.perfis

# Comparar dois planos
.\.venv-teste\Scripts\python.exe -m src.pa_check

# Conferir plano, checklists e EDD
.\.venv-teste\Scripts\python.exe -m src.main
```

Execute da raiz do projeto com `-m` conforme os comandos acima, não abrindo diretamente os arquivos de `src`. Em `main.py`, pressione Enter após conferir o aviso de campanha. Os scripts podem tentar abrir as saídas ou a pasta de trabalho no aplicativo padrão.

Antes da execução:

1. Confirme site, campanha, arquivos, cabeçalhos e abas.
2. Use cópias das entradas na primeira execução.
3. Garanta que a pasta de saída exista e esteja acessível.
4. Feche arquivos de saída abertos no Excel.
5. Verifique se arquivos existentes podem ser substituídos.
6. Para perfis, confirme a autorização para enviar URLs ao TinyURL.

Não há opção de simulação nos comandos de produção. Para uma validação sem serviços externos, use os testes automatizados.

## Testes e validação

Execute:

```powershell
.\.venv-teste\Scripts\python.exe -m unittest discover -s tests -v
```

Os testes usam planilhas fictícias temporárias, simulam o TinyURL e evitam abrir o navegador. Cobrem configurações obrigatórias, mapas JSON, caminhos relativos, prioridade de ambiente, templates, intervalos de repetição e os quatro fluxos, incluindo os sites da versão RS.

**Não validam disponibilidade do TinyURL, permissões do SharePoint ou consistência dos dados reais.** Os testes não precisam de `.env` local: usam `.env.example` e configurações temporárias.

Para validar dados reais, confira amostras conhecidas:

| Cenário | Resultado esperado |
|---|---|
| Local GIS e documento com a mesma chave | Dados do documento associados ao local e link gerado. |
| Local apenas no GIS | Registro sem correspondência no QC. |
| Documento apenas na exportação | Divergência na versão RS. |
| Combinação presente em apenas um plano | Observação de diferença na comparação dos planos. |
| Parâmetro planejado sem resultado na amostra | Análise ausente na conferência EDD. |

Um arquivo gerado ou um código de saída zero não bastam para comprovar correção. Confira conteúdo, quantidade de registros, vínculos entre locais e documentos e se a execução chegou ao fim.

## Limitações e solução de problemas

### Limitações conhecidas

- A identificação do arquivo remove os quatro últimos caracteres do nome. Extensões como `.pdf` se ajustam a essa regra; `.xlsx`, `.docx` e nomes sem extensão podem produzir identificadores incorretos.
- As chaves dos perfis concatenam campos sem separadores. Valores distintos podem formar a mesma chave; duplicidades também podem multiplicar linhas no cruzamento.
- Nomes, acentos, maiúsculas/minúsculas, espaços e organização das pastas influenciam a correspondência.
- `Folder Path` precisa permitir a concatenação com `Name`, normalmente terminando em `/`.
- URLs curtas não garantem existência, conteúdo correto ou acesso ao documento.
- Não há salvamento progressivo dos links; uma falha final do encurtador pode impedir a geração do Excel final.
- Na versão RS, ausência total de correspondências encerra o programa com mensagem e código zero, sem gerar o Excel final de links. Na versão antiga, esse caso ainda referencia uma variável inexistente (`oi`).
- A leitura de checklists e algumas comparações de `main.py` são específicas do layout original e têm limitações descritas no guia detalhado.
- O conjunto atual de dependências atende às entradas `.xlsx`. Não trate isso como suporte genérico a CSV, PDF ou ao formato Excel antigo `.xls`.

### Erros frequentes

| Sintoma | O que verificar |
|---|---|
| `ModuleNotFoundError` | Ambiente Python usado pelo comando e instalação de `requirements.txt`. |
| `ValueError` citando uma variável | Presença da configuração, valor permitido, JSON, template ou intervalo de espera. |
| `FileNotFoundError` | Pasta, nome, extensão e acesso ao arquivo/pasta de saída. |
| `KeyError` | Cabeçalho exato e layout da entrada; configuração de coluna no ambiente. |
| Erro de aba não encontrada | Nomes das abas EQuIS e configurações `MAIN_*_SHEET`. |
| `PermissionError` | Permissão de escrita ou Excel de saída aberto. |
| Muitos registros sem link | Identificadores, sites, estrutura de `Folder Path`, extensão e duplicidades. |
| Erro no TinyURL | Conectividade, política de rede, URL enviada ou disponibilidade do serviço. |
| `NameError` citando `oi` | Trecho antigo de sinalização de erro; investigue formato/correspondências no guia do script. |

## Glossário e documentação detalhada

- **GIS:** sistema de informações geográficas; neste fluxo, origem do controle de locais/perfis.
- **PA:** plano de amostragem, com pontos e parâmetros planejados.
- **EDD/EQuIS:** arquivos e tabelas de intercâmbio usados para dados de amostras e resultados.
- **DataFrame:** tabela em memória manipulada pelo pandas.
- **Chave:** texto calculado para identificar correspondências entre tabelas.
- **Merge:** operação de associação de tabelas por uma chave comum.
- **QC:** controle de qualidade; destaca divergências para revisão.

Guias por programa:

- [Perfis com QC e regras por site](./docs/perfis_rs.md).
- [Outra versão do processamento de perfis](./docs/perfis.md).
- [Comparação entre dois planos](./docs/pa_check.md).
- [Conferência do plano, checklists e EDD](./docs/main.md).

Consulte [`.env.example`](./.env.example) para a referência comentada dos valores de configuração.
