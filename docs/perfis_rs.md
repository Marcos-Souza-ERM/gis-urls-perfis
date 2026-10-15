# `perfis_rs.py` — associação de perfis e controle de divergências

## Finalidade

O módulo `src.perfis_rs` cruza uma planilha de controle GIS com uma exportação/listagem de arquivos de perfis. Ele identifica registros sem correspondência nos dois sentidos e tenta encurtar os endereços dos documentos encontrados.

Apesar do sufixo `_rs`, o código inclui regras para BA, AL, AC, Tupy e RS. A regra de site determina como o script interpreta as pastas e cria as chaves de associação; alterar apenas a sigla não garante compatibilidade com uma nova fonte.

## Entradas

As duas planilhas devem estar disponíveis em `PERFIS_RS_INPUT_DIR`:

1. **Controle GIS**, selecionado pelo mapa `PERFIS_RS_CONTROL_FILES`. Exemplos de nomes configurados:
   - BA: `BRASKEM_BA_PerfisConstrutivos_Controle.xlsx`;
   - AL: `BRASKEM_AL_PerfisConstrutivos_Controle.xlsx`;
   - AC: `ACELEN-BA Poços com perfis construtivos.xlsx`;
   - Tupy: `TupyMaua_PerfisConstrutivos_Controle.xlsx`;
   - RS: `BRK-RS-Controle perfis das locations.xlsx`.
2. **Exportação de arquivos do BI/SharePoint**, selecionada em `PERFIS_RS_LINKS_FILE`. O modelo configurado é `perfil {site}.xlsx`, por exemplo, `perfil BA.xlsx`.

Para RS, o script lê todas as abas da planilha de controle e as combina. Nos demais casos, lê a planilha de controle diretamente.

### Colunas necessárias

- No controle GIS: identificador do local (`PERFIS_RS_LOCATION_COLUMN`) e site (`PERFIS_RS_SITE_COLUMN`). RS também requer a coluna de subárea configurada em `PERFIS_RS_SUBAREA_COLUMN`.
- Na exportação: nome de arquivo (`PERFIS_RS_LINK_NAME_COLUMN`, padrão `Name`) e caminho de pasta (`PERFIS_RS_LINK_PATH_COLUMN`, padrão `Folder Path`).

Os cabeçalhos devem corresponder exatamente às colunas do Excel, inclusive espaços, acentos e pontuação. A exportação de links é uma lista preparada previamente; o script não consulta o SharePoint nem baixa os documentos.

## Configuração

O módulo carrega configurações de `.env`:

| Variável | Uso |
|---|---|
| `PERFIS_RS_SITE` | Site selecionado: `BA`, `AL`, `AC`, `Tupy` ou `RS`. |
| `PERFIS_RS_INPUT_DIR` / `PERFIS_RS_OUTPUT_DIR` | Pastas de entrada e saída. |
| `PERFIS_RS_CONTROL_FILES` | Mapa JSON de controles por site. |
| `PERFIS_RS_LINKS_FILE` | Nome/modelo da exportação; aceita `{site}`. |
| `PERFIS_RS_LOCATION_COLUMN`, `PERFIS_RS_SITE_COLUMN`, `PERFIS_RS_SUBAREA_COLUMN` | Cabeçalhos GIS. A subárea é obrigatória para RS. |
| `PERFIS_RS_LINK_NAME_COLUMN`, `PERFIS_RS_LINK_PATH_COLUMN` | Cabeçalhos da exportação do BI. |
| `PERFIS_RS_RS_PATH_MARKER`, `PERFIS_RS_AC_PATH_MARKER`, `PERFIS_RS_AL_PATH_MARKER` | Marcadores de caminho utilizados nas regras correspondentes. |
| `PERFIS_RS_QC_FILE`, `PERFIS_RS_OUTPUT_FILE` | Nomes dos resultados; aceitam `{site}` e `{date}`. |
| `LINK_RETRY_WAIT_1`, `LINK_RETRY_WAIT_2` | Esperas, em segundos, antes da segunda e da terceira tentativa ao TinyURL. |

Mapas devem ser objetos JSON válidos em uma única linha. Caminhos relativos partem da raiz do projeto. As pastas de entrada e saída precisam existir.

## Regras de associação por site

O script normaliza alguns caracteres e forma uma chave a partir do site, do identificador do local/arquivo e, para RS, também da subárea. A extração do site depende da estrutura de `Folder Path`:

- **BA**: usa o segmento de pasta esperado e contém correções de nomes como `PE1`/`PE-1`, `PVC1`/`PVC` e variações de `UA`.
- **AC**: localiza o trecho após `PERFIS_RS_AC_PATH_MARKER` e aplica correções específicas de nomes de pasta.
- **AL**: localiza o trecho após `PERFIS_RS_AL_PATH_MARKER` e usa o primeiro segmento subsequente como site.
- **RS**: procura o trecho posterior a `PERFIS_RS_RS_PATH_MARKER`, divide site e subárea quando aplicável e inclui ambos na chave.
- **Tupy**: usa a regra padrão de extração por posição do caminho.

Revise esses marcadores e teste a extração com valores reais autorizados antes de processar uma campanha. As chaves são concatenadas sem separador entre alguns componentes; os dados precisam ser consistentes e não ambíguos.

## Processamento e saídas

1. Lê o controle GIS e a exportação de documentos.
2. Normaliza nomes e identificadores e calcula a chave de cada fonte.
3. Identifica os registros GIS não encontrados na exportação e os arquivos da exportação não encontrados no GIS.
4. Se houver diferenças, grava `PERFIS_RS_QC_FILE` com duas abas: itens do GIS sem correspondência e itens da exportação sem correspondência.
5. Associa os arquivos ao controle GIS, forma o endereço original a partir do caminho e do nome e solicita ao TinyURL um link curto.
6. Grava o resultado em `PERFIS_RS_OUTPUT_FILE`.

O resultado inclui colunas de origem, dados associados e campos de link. O endereço encurtado é representado como texto HTML (`<a href="...">Disponível</a>`), não como hyperlink nativo do Excel. O script pode abrir relatórios/resultados com o aplicativo padrão do sistema.

Se não houver nenhuma URL associada, o script encerra sem gerar o Excel final de links. Se houver divergências, o QC é condicional; portanto, a ausência desse arquivo só indica que não houve divergências se a execução também terminou normalmente.

## Execução

Configure o site, as colunas, os arquivos e os diretórios em `.env`. Execute a partir da raiz:

```powershell
python -m src.perfis_rs
```

O encurtamento requer conexão com TinyURL e aprovação organizacional para enviar as URLs ao serviço externo. Uma falha na última tentativa pode interromper o processamento; confirme que a mensagem de conclusão foi exibida e que o arquivo de saída está íntegro.

## Interpretação e limitações

- A aba QC sinaliza diferenças entre duas listagens; não confirma que o documento deveria ou não existir.
- A planilha exportada precisa conter URLs/pastas que possam ser usadas para compor um endereço válido. O código não cria links de compartilhamento nem concede permissões no SharePoint.
- A remoção da extensão assume que os quatro últimos caracteres do nome do arquivo correspondem à extensão, normalmente `.pdf`.
- A normalização cobre somente regras explícitas no código; diferenças adicionais de grafia, espaços ou estrutura de diretório podem impedir o cruzamento.
- Uma falha do TinyURL pode interromper a geração; verifique o resultado antes de utilizá-lo.
- Valide manualmente uma amostra de correspondências, os valores de site extraídos e os registros do QC.

## Referência

Veja o [README principal](../README.md) para requisitos, instalação e testes. A lista das opções iniciais está em [`.env.example`](../.env.example).
