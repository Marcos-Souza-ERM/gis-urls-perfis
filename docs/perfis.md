# `perfis.py` — associação de perfis a registros GIS

## Finalidade

O módulo `src.perfis` associa registros de uma planilha de controle GIS a arquivos listados em uma exportação do BI/SharePoint. Para as correspondências encontradas, constrói uma URL e solicita ao TinyURL um endereço curto.

Esta é a versão alternativa/mais antiga do fluxo de perfis. Ela atende BA, AL e Tupy e tem regras de comparação diferentes de `perfis_rs.py`. Não é um encurtador genérico: o vínculo entre o controle GIS e a listagem de arquivos é parte obrigatória do processamento.

## Entradas

O fluxo requer as duas planilhas abaixo dentro de `PERFIS_INPUT_DIR`:

1. **Controle GIS**, selecionado por site em `PERFIS_CONTROL_FILES`. Os valores de referência incluem:
   - BA: `BRASKEM_BA_PerfisConstrutivos_Controle.xlsx`;
   - AL: `BRASKEM_AL_PerfisConstrutivos_Controle.xlsx`;
   - Tupy: `TupyMaua_PerfisConstrutivos_Controle.xlsx`.
2. **Exportação/listagem dos documentos**, configurada em `PERFIS_LINKS_FILE`. O modelo padrão é `perfil {site}.xlsx`; por exemplo, `perfil AL.xlsx`.

A planilha de controle precisa conter colunas equivalentes às configuradas em `PERFIS_FILE_COLUMN`, `PERFIS_LOCATION_COLUMN` e `PERFIS_SITE_COLUMN`. Se o campo de arquivo `PERFIS_FILE_COLUMN` não existir, o script usa o identificador de local como alternativa. Também são utilizadas as colunas configuradas para `GlobalID` e `ObjectID` no relatório de QC.

A exportação de arquivos precisa conter as colunas configuradas em `PERFIS_LINK_NAME_COLUMN` e `PERFIS_LINK_PATH_COLUMN` (por padrão, `Name` e `Folder Path`). O nome do arquivo identifica o perfil; o caminho da pasta é usado para compor o endereço de origem.

## Configuração

As variáveis são carregadas de `.env` a partir da raiz do repositório:

| Variável | Uso |
|---|---|
| `PERFIS_SITE` | Site selecionado (`BA`, `AL` ou `Tupy`). |
| `PERFIS_INPUT_DIR` / `PERFIS_OUTPUT_DIR` | Pastas de entrada e saída. |
| `PERFIS_CONTROL_FILES` | Mapa JSON dos arquivos de controle por site. |
| `PERFIS_LINKS_FILE` | Nome/modelo da planilha exportada; aceita `{site}`. |
| `PERFIS_*_COLUMN` | Cabeçalhos de controle GIS e de exportação. |
| `PERFIS_QC_FILE` / `PERFIS_OUTPUT_FILE` | Nomes dos relatórios; aceitam `{site}` e `{date}`. |
| `LINK_RETRY_WAIT_1` / `LINK_RETRY_WAIT_2` | Intervalos, em segundos, entre tentativas ao TinyURL. |

Os mapas são objetos JSON em uma única linha. Caminhos relativos partem da raiz do projeto. As pastas configuradas precisam existir antes da execução.

## Lógica de associação

1. Lê o controle GIS e a exportação do BI.
2. Renomeia os cabeçalhos configurados para os nomes internos usados pelo script.
3. Remove espaços e normaliza caracteres selecionados (`|` é substituído por `I`).
4. Remove os quatro últimos caracteres do nome exportado para obter o identificador do arquivo.
5. Extrai o site do segmento esperado de `Folder Path` e forma uma chave com site e identificador do local/arquivo.
6. Faz uma junção do tipo left, mantendo os registros GIS e acrescentando os dados do documento correspondente.
7. Gera uma planilha QC de registros GIS sem correspondência, constrói as URLs para os documentos associados e tenta encurtá-las.
8. Grava a planilha final.

O caminho exportado deve seguir a estrutura esperada pelo código, que usa o segmento de pasta pai para identificar o site. A chave é baseada em texto: pequenas diferenças de nome podem impedir a correspondência.

## Arquivos gerados

- **QC** (`PERFIS_QC_FILE`): registros GIS sem documento correspondente.
- **Resultado** (`PERFIS_OUTPUT_FILE`): dados GIS unidos às informações do documento e ao link encurtado, quando disponível.

Os arquivos são gravados em `PERFIS_OUTPUT_DIR`. O script pode tentar abrir os arquivos gerados por meio do aplicativo padrão do sistema. O valor de link no Excel é armazenado como texto HTML (`<a href="...">Disponível</a>`), não como fórmula nativa de hyperlink.

## Execução

Configure `.env`, confirme os dois arquivos de entrada e execute a partir da raiz:

```powershell
python -m src.perfis
```

O TinyURL exige conectividade externa. Em caso de falha, o código tenta novamente com os intervalos configurados; uma falha final pode interromper o processamento.

## Limitações e diagnóstico

- O script gera uma lista de controle para registros GIS sem correspondência, mas não produz um relatório equivalente para cada arquivo extra da exportação.
- A extensão é removida assumindo os quatro últimos caracteres do nome, normalmente `.pdf`.
- O código combina o caminho e o nome do documento conforme os valores recebidos. Verifique se essa concatenação forma uma URL válida.
- A extração do site e a normalização dependem do padrão dos arquivos da equipe; não são uma descoberta automática de pastas do SharePoint.
- Verifique a planilha QC, os campos sem link e os identificadores usados na chave antes de interpretar o resultado.
- Use `perfis_rs.py` somente se as colunas e regras específicas desse fluxo corresponderem às suas fontes.

## Referência

Consulte o [README principal](../README.md) para requisitos gerais e o arquivo [`.env.example`](../.env.example) para valores de configuração de referência.
