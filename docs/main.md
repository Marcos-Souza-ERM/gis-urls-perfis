# `main.py` — conferência do plano de amostragem

## Finalidade

O módulo `src.main` normaliza um Plano de Amostragem (PA) e compara os pontos e parâmetros previstos com checklists de campo e resultados exportados do EQuIS/EDD. O relatório destaca diferenças para análise da equipe; não altera as fontes nem determina automaticamente a causa de uma divergência.

O script é específico ao formato configurado para a campanha. Revise os caminhos, arquivos, abas e cabeçalhos antes de cada execução.

## Fontes de entrada

Todas as entradas são arquivos Excel preparados previamente.

### Plano de amostragem

- Caminho: `MAIN_PLAN_FILE`.
- O script processa somente abas cujo nome contém o valor de `MAIN_PLAN_SHEET_FILTER` (padrão: `_PA`).
- Em cada aba são esperadas colunas equivalentes às configuradas em `MAIN_GROUP_COLUMN`, `MAIN_POINTS_COLUMN` e `MAIN_PARAMETERS_COLUMN`. Os nomes internos utilizados pelo processamento são `Grupos`, `Poços` e `Parâmetros`.
- A grafia histórica `Parâmetros `, com espaço final, também é tratada.
- A célula de pontos e a de parâmetros podem conter listas separadas por vírgulas; cada combinação ponto/parâmetro vira um registro.

### Exportações EQuIS/EDD

- Diretório: `MAIN_EQUIS_DIR`.
- Mapa de arquivos por site: `MAIN_EQUIS_FILES`, como objeto JSON.
- Abas configuráveis: `MAIN_SAMPLE_SHEET`, `MAIN_SMP_SHEET`, `MAIN_TEST_SHEET`, `MAIN_RESULT_SHEET` e `MAIN_BATCH_SHEET`.
- A comparação usa principalmente amostras e resultados. O script também tenta ler as abas SMP, TST e BCH indicadas na configuração; elas precisam existir mesmo quando não são usadas diretamente na consolidação do relatório.

### Checklists de campo

- Diretório: `MAIN_CHECKLIST_DIR`.
- `MAIN_CHECKLIST_FILES` associa os nomes exatos dos arquivos aos sites normalizados.
- O formato esperado inclui cabeçalhos como `Poço`, `Data`, `Poço foi amostrado?` e `Status do Poço`. O script usa a linha configurada implicitamente pelo layout atual do arquivo; não é uma importação genérica de checklists.

## Configuração

As variáveis são carregadas de `.env` pela raiz do projeto. Caminhos relativos também partem da raiz.

| Variável | Uso |
|---|---|
| `MAIN_PLAN_FILE` | Plano de entrada. |
| `MAIN_EQUIS_DIR` / `MAIN_EQUIS_FILES` | Diretório e mapa dos arquivos EQuIS por site. |
| `MAIN_CHECKLIST_DIR` / `MAIN_CHECKLIST_FILES` | Diretório e mapa dos checklists por site. |
| `MAIN_PLAN_SHEET_FILTER` | Texto que identifica abas do PA a processar. |
| `MAIN_GROUP_COLUMN`, `MAIN_POINTS_COLUMN`, `MAIN_PARAMETERS_COLUMN` | Cabeçalhos de grupo, pontos e parâmetros na planilha do PA. |
| `MAIN_SITE_RENAME` | Mapeamento de nomes de abas/sites para os identificadores usados na comparação. |
| `MAIN_SAMPLE_SHEET`, `MAIN_SMP_SHEET`, `MAIN_TEST_SHEET`, `MAIN_RESULT_SHEET`, `MAIN_BATCH_SHEET` | Nomes de abas nas exportações EQuIS. |
| `MAIN_PLAN_OUTPUT_FILE` | Destino do plano normalizado. |
| `MAIN_REPORT_OUTPUT_FILE` | Destino do relatório de comparação. |

Os caminhos de saída e os diretórios correspondentes precisam ser graváveis. Confira se os sites ativos nos mapas representam a campanha que será analisada; os valores do `.env.example` são apenas uma referência de configuração.

## Processamento

1. Lê as abas do plano que correspondem ao filtro configurado.
2. Renomeia os cabeçalhos configurados para os nomes internos e separa listas de pontos e parâmetros em registros individuais.
3. Aplica as correções textuais de nomes existentes no script e normaliza o nome do site por meio de `MAIN_SITE_RENAME`.
4. Solicita confirmação interativa para lembrar a revisão da campanha.
5. Filtra o plano para os sites presentes em `MAIN_EQUIS_FILES` e grava o plano normalizado.
6. Lê e consolida os checklists do diretório configurado.
7. Para cada site, lê as abas EQuIS selecionadas e compara pontos e parâmetros com o plano e o checklist.
8. Grava o relatório final no caminho `MAIN_REPORT_OUTPUT_FILE`.

O nome da aba do plano é inicialmente usado como identificador do site. A correspondência entre esse nome, o mapeamento `MAIN_SITE_RENAME`, as chaves do mapa EQuIS e os arquivos de checklist deve ser consistente.

## Compatibilidade conhecida

Na versão atual do código, uma chamada de gravação do relatório passa o nome da aba de saída de forma posicional para `DataFrame.to_excel`. Com o pandas instalado na validação, isso causa `TypeError` durante a geração do relatório e impede que o fluxo conclua normalmente. A suíte automatizada reproduz o problema. A documentação dos resultados abaixo descreve o formato pretendido do relatório; não representa uma garantia de geração até que a chamada seja corrigida no código.

## Interpretação dos resultados

O relatório pode incluir as seguintes abas:

| Aba/indicador | Interpretação |
|---|---|
| `Análises extras` | Resultados do EDD que não foram encontrados entre os parâmetros previstos para o ponto. |
| `Análises ausentes` | Parâmetros previstos que não foram localizados nos resultados relacionados à amostra/ponto. |
| `Pontos extras` / `Pontos ausentes` | Diferenças de pontos entre o PA e a fonte EQuIS consultada. |
| `Parâmetros extras geral` / `Parâmetros ausentes geral` | Comparações agregadas de parâmetros. |
| `Não coletado sem justificativa` | Registros marcados como amostrados no checklist para os quais o processamento não localizou resultados correspondentes. |
| `Compilado EDDs` | Resultados EDD reunidos pelo script. |
| `Plano de amostragem` | PA normalizado usado na comparação. |
| `Check list` | Dados de checklist consolidados. |

O script pode incluir outras abas de apoio, como diferenças de nomes e pontos em nível geral. Consulte o conteúdo do arquivo gerado e valide amostras representativas contra as fontes originais.

- **Ausente** não identifica a causa: pode indicar coleta não realizada, resultado ainda não exportado, incompatibilidade de identificador ou grafia diferente.
- **Extra** significa que o dado foi encontrado na fonte comparada, mas não no recorte do PA considerado.
- Um ponto presente no EQuIS não implica que todos os parâmetros previstos tenham resultado correspondente.
- As comparações dependem de nomes de pontos, sites e parâmetros consistentes; normalizações existentes cobrem apenas casos específicos.

## Execução

Na raiz do repositório, com `.env` configurado:

```powershell
python -m src.main
```

O programa pausa para confirmação da campanha antes de prosseguir. Responda ao prompt no terminal. Confirme a criação tanto do plano normalizado quanto do relatório configurado; uma execução interrompida não deve ser considerada completa.

## Limitações e cuidados

- O formato do checklist e o esquema das exportações EQuIS são específicos ao código atual.
- Os arquivos de entrada não são obtidos automaticamente do SharePoint nem do laboratório.
- Ajustes de grafia são explícitos no código; variações não previstas podem aparecer como diferenças.
- A lista de sites ativos depende de `MAIN_EQUIS_FILES`; atualizar somente um caminho não garante que o restante dos mapas corresponda à campanha.
- As abas SMP, TST e BCH são lidas, mas não participam diretamente das comparações principais descritas acima.
- Verifique os resultados manualmente com os responsáveis antes de tomar decisões operacionais.

## Referência

Consulte o [README principal](../README.md) para instalação, configuração geral, testes e estrutura do repositório. As variáveis de exemplo estão em [`.env.example`](../.env.example).
