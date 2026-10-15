# GIS URL Perfis

## Nome do Projeto

**GIS URL Perfis** — ferramentas Python para relacionar registros de perfis construtivos a documentos listados em exportações do SharePoint, criar URLs curtas e conferir planos de amostragem com outras fontes tabulares.

## Visão Geral

O repositório reúne quatro scripts independentes de processamento de arquivos Excel. Cada script atende a uma atividade específica; não existe um servidor, interface web, banco de dados ou programa único que execute todos os fluxos.

| Script | Finalidade | Entradas principais | Resultado |
|---|---|---|---|
| [`src/perfis_rs.py`](./src/perfis_rs.py) | Relacionar registros GIS a arquivos de perfis e conferir divergências nos dois sentidos. Apesar do nome, contém regras para BA, AL, AC, Tupy e RS. | Controle GIS e exportação de arquivos `perfil {site}.xlsx`. | Excel com os registros associados e URLs curtas; relatório QC quando há divergências. |
| [`src/perfis.py`](./src/perfis.py) | Executar uma versão alternativa e mais antiga da associação de perfis. | Controle GIS e exportação de arquivos `perfil {site}.xlsx`. | Excel com links e lista de registros GIS sem correspondência. |
| [`src/pa_check.py`](./src/pa_check.py) | Comparar duas versões de um plano de amostragem. | Dois planos com colunas de site, ponto e parâmetro. | Excel com as duas versões e indicações de diferenças. |
| [`src/main.py`](./src/main.py) | Conferir o plano de amostragem contra checklists e resultados EQuIS/EDD. | Plano, checklists e exportações EQuIS. | Plano normalizado e relatório de pontos e análises extras ou ausentes. |

Os arquivos de entrada devem ser exportados ou preparados antes da execução. O projeto não autentica no SharePoint, não navega pelo site para localizar arquivos e não baixa documentos. Nos fluxos de perfis, o arquivo `perfil {sigla}.xlsx` é uma planilha de listagem/exportação, não o próprio PDF do perfil.

## Objetivo

Reduzir tarefas manuais de cruzamento de dados, padronizar relatórios de controle e facilitar a identificação de inconsistências entre:

- registros do controle GIS e a listagem de documentos de perfis;
- duas versões de um plano de amostragem;
- o plano de amostragem, checklists de campo e resultados do laboratório.

Os relatórios são instrumentos de conferência. Uma divergência não determina, por si só, qual fonte está correta nem a causa do problema.

## Tecnologias Utilizadas

| Tecnologia | Uso |
|---|---|
| Python | Execução dos scripts e transformação dos dados. |
| pandas | Leitura, normalização, comparação e gravação de tabelas Excel. |
| openpyxl | Leitura e gravação de arquivos `.xlsx` por meio do pandas. |
| pyshorteners | Solicitação de URLs curtas ao serviço TinyURL nos fluxos de perfis. |
| python-dotenv | Carregamento das configurações locais a partir do arquivo `.env`. |
| unittest | Testes automatizados incluídos em `tests/test_environment.py`. |

As dependências de execução estão em [`requirements.txt`](./requirements.txt). O pandas está restrito à versão principal 2 (`pandas>=2,<3`).

## Arquitetura

```text
Arquivos Excel preparados pelo usuário
             |
             v
      Configuração .env
             |
             v
   Script escolhido em src/
      |             |
      |             +--> TinyURL (somente nos fluxos de perfis)
      v
 Planilhas de resultado e relatórios QC
```

- [`src/config.py`](./src/config.py) centraliza a leitura de `.env`, a validação de opções e a resolução dos caminhos relativos à raiz do repositório.
- Os scripts são executáveis como módulos (`python -m src.<modulo>`) e usam configurações separadas por prefixos: `PERFIS_RS_`, `PERFIS_`, `PA_CHECK_` e `MAIN_`.
- `inputs/` e `outputs/` são pastas convencionais. Os caminhos reais são configuráveis e as pastas necessárias precisam existir antes da execução.
- O acesso ao SharePoint é uma etapa externa: o usuário precisa obter as exportações e ter permissão para ler os arquivos de origem. O encurtamento envia as URLs resultantes ao TinyURL; confirme a autorização organizacional antes de usar esse serviço com endereços internos.

## Pré-requisitos

1. Python instalado. O projeto não declara uma versão mínima ou uma matriz oficial de compatibilidade; a configuração existente registra validação em Python 3.12.
2. Dependências instaladas a partir de `requirements.txt`.
3. Acesso às fontes necessárias: permissão no SharePoint para obter as exportações, acesso local/de rede às planilhas e permissão de escrita na pasta de saída.
4. Arquivos Excel com os nomes, abas e cabeçalhos esperados pelo fluxo selecionado.
5. Conexão de rede ao TinyURL e autorização para enviar URLs ao serviço, quando executar `perfis.py` ou `perfis_rs.py`.

O Microsoft Excel não é necessário para executar o processamento. É necessário um aplicativo compatível para abrir e revisar os arquivos produzidos.

## Instalação

No Windows, abra o PowerShell na raiz do repositório:

```powershell
py --version
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r .\requirements.txt
```

Se o comando `py` não estiver disponível, instale o Python e configure o PATH ou use o executável `python` correspondente à instalação.

## Configuração

1. Crie sua configuração local copiando o modelo:

   ```powershell
   Copy-Item -LiteralPath .\.env.example -Destination .\.env
   ```

2. Edite somente as variáveis do fluxo que pretende executar.
3. Confirme os caminhos, nomes de arquivos, abas e cabeçalhos com as planilhas reais.
4. Crie as pastas de entrada e saída necessárias. Os scripts não criam automaticamente os diretórios configurados.

O arquivo `.env` não deve ser versionado nem compartilhado com dados ou informações confidenciais. O `.env.example` contém valores de referência; não substitui a configuração da campanha ou do ambiente real. Variáveis já definidas no sistema têm prioridade sobre as do `.env`.

### Entradas e configuração dos fluxos

**Perfis (`perfis_rs.py` e `perfis.py`)**

- São necessárias duas planilhas: o controle GIS e a exportação/listagem de arquivos do BI/SharePoint.
- A listagem deve estar no formato `perfil {sigla}.xlsx`, conforme `PERFIS_RS_LINKS_FILE` ou `PERFIS_LINKS_FILE`; por exemplo, `perfil AL.xlsx`. O marcador `{site}` é substituído pelo valor configurado para o site.
- O controle GIS é selecionado nos mapas `PERFIS_RS_CONTROL_FILES` ou `PERFIS_CONTROL_FILES`.
- Confira os cabeçalhos configurados para identificador, site, nome do documento e caminho da pasta. Em `perfis_rs.py`, RS também exige a coluna de subárea.
- A planilha de exportação precisa conter os nomes e caminhos utilizáveis para associar os registros. Um caminho de pasta local ou um endereço que não seja uma URL acessível pode não produzir um link TinyURL válido.

**Comparação de planos (`pa_check.py`)**

- São necessários dois arquivos Excel configurados em `PA_CHECK_YAEL_FILE` e `PA_CHECK_TAISE_FILE`.
- Ambos precisam conter as colunas configuradas em `PA_CHECK_SITE_COLUMN`, `PA_CHECK_POINT_COLUMN` e `PA_CHECK_PARAMETER_COLUMN`.
- `PA_CHECK_OUTPUT_FILE` define o arquivo de resultado.

**Conferência de campanha (`main.py`)**

- `MAIN_PLAN_FILE`: plano com abas cujo nome contém `MAIN_PLAN_SHEET_FILTER` (por padrão, `_PA`).
- `MAIN_EQUIS_DIR` e `MAIN_EQUIS_FILES`: diretório e mapa dos arquivos EQuIS associados aos sites ativos. São esperadas as abas configuradas em `MAIN_SAMPLE_SHEET`, `MAIN_SMP_SHEET`, `MAIN_TEST_SHEET`, `MAIN_RESULT_SHEET` e `MAIN_BATCH_SHEET`.
- `MAIN_CHECKLIST_DIR` e `MAIN_CHECKLIST_FILES`: diretório e mapa de arquivos de checklist. O layout precisa corresponder ao formato esperado pelo script.
- Configure os cabeçalhos do plano, o mapeamento de nomes de site e os caminhos de saída `MAIN_PLAN_OUTPUT_FILE` e `MAIN_REPORT_OUTPUT_FILE`.

Os caminhos relativos são resolvidos a partir da raiz do repositório, não do diretório atual do terminal. No formato `.env`, use aspas simples para caminhos Windows que contenham espaços ou caracteres especiais. Os mapas de arquivos são objetos JSON em uma única linha.

## Como Executar

Execute apenas o módulo que corresponde à tarefa, a partir da raiz do repositório:

```powershell
.\.venv\Scripts\python.exe -m src.perfis_rs
.\.venv\Scripts\python.exe -m src.perfis
.\.venv\Scripts\python.exe -m src.pa_check
.\.venv\Scripts\python.exe -m src.main
```

Cada linha é uma alternativa; não execute todos os fluxos em sequência sem necessidade. O módulo `main` apresenta uma confirmação interativa sobre a campanha. Os scripts de perfis podem tentar abrir relatórios ou resultados usando o aplicativo padrão do sistema.

Antes de uma execução real:

- confirme que as planilhas não estão abertas com bloqueio de gravação;
- verifique se os arquivos correspondem à campanha e ao site selecionados;
- confirme se os diretórios de saída existem e se resultados anteriores podem ser substituídos;
- revise QC e divergências antes de utilizar os dados em um processo operacional.

## Como Testar

Execute a suíte automatizada a partir da raiz do projeto:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Os testes validam a configuração e executam fluxos com planilhas temporárias e dados fictícios. Chamadas ao navegador, pausas e chamadas externas ao encurtador são isoladas por mocks; portanto, a suíte não comprova acesso real ao SharePoint, permissões de rede, funcionamento do TinyURL nem compatibilidade com arquivos de uma campanha específica.

## Estrutura do Projeto

```text
gis-url-perfis/
├── .env.example              Modelo de configuração local
├── README.md                 Guia principal do projeto
├── requirements.txt          Dependências Python
├── src/
│   ├── config.py             Leitura e validação de configuração
│   ├── main.py               Conferência de campanha/EDD/checklists
│   ├── pa_check.py           Comparação de dois planos
│   ├── perfis.py             Fluxo alternativo de associação de perfis
│   └── perfis_rs.py          Fluxo de perfis com QC nos dois sentidos
├── tests/
│   └── test_environment.py   Testes de configuração e processamento
├── docs/
│   ├── main.md               Guia do fluxo main
│   ├── pa_check.md           Guia da comparação de planos
│   ├── perfis.md             Guia do fluxo perfis
│   └── perfis_rs.md          Guia do fluxo perfis_rs
├── inputs/                   Local convencional para entradas
├── outputs/                  Local convencional para saídas
└── old_runs/                 Local convencional para arquivos de execuções anteriores
```

As pastas de dados podem estar vazias em uma cópia limpa do repositório. As planilhas operacionais não devem ser presumidas como presentes ou atualizadas só porque existem em uma cópia local.

## Fluxo de Desenvolvimento

1. Atualize ou crie uma branch a partir da versão de trabalho acordada pela equipe.
2. Confirme o layout das fontes e configure as variáveis adequadas em `.env`.
3. Faça alterações pequenas e específicas no script e no guia correspondente em `docs/`.
4. Execute os testes automatizados e revise o diff.
5. Faça uma execução controlada com arquivos de teste autorizados, quando a mudança afetar um fluxo de negócio.
6. Revise os arquivos gerados e os registros QC antes de divulgar os resultados.

Não inclua planilhas reais, exportações do SharePoint, credenciais ou o `.env` em commits.

## CI/CD

Não há configuração de CI/CD versionada neste repositório. A validação disponível é a suíte `unittest`, executada localmente conforme [Como Testar](#como-testar). Qualquer pipeline futuro deve usar dados fictícios ou autorizados e não deve publicar arquivos de entrada ou resultados confidenciais.

## Documentação

- [`docs/main.md`](./docs/main.md): conferência do plano de amostragem com checklists e EQuIS/EDD.
- [`docs/pa_check.md`](./docs/pa_check.md): comparação de duas versões de um plano.
- [`docs/perfis.md`](./docs/perfis.md): associação de perfis usando o fluxo alternativo.
- [`docs/perfis_rs.md`](./docs/perfis_rs.md): associação de perfis com regras por site e QC nos dois sentidos.
- [`.env.example`](./.env.example): referência das variáveis de configuração disponíveis.

## Contribuição

Contribuições devem preservar os fluxos já utilizados e acompanhar mudanças de comportamento com testes e documentação:

1. Descreva o problema e o fluxo afetado.
2. Evite incluir dados operacionais ou identificáveis nos exemplos e testes.
3. Mantenha os nomes de configuração e os layouts documentados sincronizados com o código.
4. Execute a suíte de testes antes de abrir a revisão.
5. Informe claramente requisitos externos, como permissões de SharePoint, rede ou aprovação para uso do TinyURL.

## Troubleshooting

| Sintoma | Verificações recomendadas |
|---|---|
| `Configure ... no .env` | Confirme que `.env` existe na raiz, que a variável do script foi preenchida e que nenhuma variável de ambiente vazia está sobrescrevendo o valor. |
| Arquivo ou diretório não encontrado | Revise o caminho, o nome exato, o site selecionado e a existência do diretório. Caminhos relativos partem da raiz do projeto. |
| `Worksheet ... not found` ou erro de coluna | Compare os nomes das abas e cabeçalhos com a configuração e com as expectativas descritas no guia do fluxo. Espaços e acentos importam. |
| `NDFrame.to_excel() takes 2 positional arguments ...` | Há chamadas de gravação incompatíveis com a API do pandas nas versões atuais em `main.py` e `pa_check.py`. A suíte de testes também falha nesses dois fluxos até que as chamadas sejam corrigidas no código. |
| Nenhuma correspondência no fluxo de perfis | Confira se a exportação `perfil {site}.xlsx` é a do site correto, se `Name`/`Folder Path` estão corretos e se as regras de extração do site correspondem ao caminho exportado. |
| Falha ao encurtar URLs | Confirme acesso de rede ao TinyURL e validade/autorização do endereço de origem. Uma falha na tentativa final pode interromper o processamento. |
| Saída ausente ou incompleta | Verifique permissões de escrita, se o arquivo está aberto, se o script terminou normalmente e se houve erro durante a leitura ou o encurtamento. |
| `main.py` aguarda no terminal | Responda ao prompt de confirmação da campanha para continuar. |
| Relatório QC de perfis não foi criado | Em `perfis_rs.py`, o relatório é gerado quando há divergências; confirme também que o programa terminou normalmente. |

Consulte o guia específico do script antes de interpretar divergências. Os relatórios sinalizam diferenças; cabe à equipe responsável investigar e validar a fonte correta.

## Licença

Não foi encontrado um arquivo de licença neste repositório. Os termos de uso e redistribuição precisam ser definidos pelo titular do projeto antes de tratar o código como software licenciado para terceiros.
