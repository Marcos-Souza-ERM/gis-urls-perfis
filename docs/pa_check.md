# `pa_check.py` — comparação entre dois planos de amostragem

## Finalidade

O módulo `src.pa_check` compara duas planilhas de Plano de Amostragem (PA) e sinaliza diferenças entre combinações de site, ponto e parâmetro. O resultado mantém as linhas de cada plano em abas separadas e acrescenta colunas de controle.

O script não escolhe qual plano é o correto, não mescla as versões em um plano aprovado e não compara resultados de laboratório.

## Entradas

São necessários dois arquivos Excel:

- `PA_CHECK_YAEL_FILE`: caminho da primeira versão, identificada no relatório como `PA Yael`.
- `PA_CHECK_TAISE_FILE`: caminho da segunda versão, identificada no relatório como `PA Taise`.

Ambos os arquivos precisam conter as três colunas correspondentes aos cabeçalhos configurados em:

- `PA_CHECK_SITE_COLUMN` (padrão `Site`);
- `PA_CHECK_POINT_COLUMN` (padrão `Ponto`);
- `PA_CHECK_PARAMETER_COLUMN` (padrão `Parâmetro`).

Os nomes de cabeçalho são configuráveis, mas os dois arquivos devem usar os valores configurados. As planilhas precisam estar disponíveis localmente ou em um caminho de rede acessível ao usuário.

## Como a comparação funciona

O script renomeia as colunas de entrada para os nomes internos `Site`, `Ponto` e `Parâmetro` e constrói uma chave textual com os três valores, separados por hífens.

São feitas três verificações:

1. **Combinação completa**: identifica combinações site + ponto + parâmetro existentes em uma planilha e ausentes na outra.
2. **Ponto**: identifica valores da coluna `Ponto` que não aparecem em nenhum registro do outro arquivo.
3. **Parâmetro geral**: identifica valores de `Parâmetro` que não aparecem em nenhum registro do outro arquivo, sem considerar site ou ponto.

Exemplo: se ambos os planos contêm `Site A` e `P-01`, mas um contém `Arsênio` e o outro `Chumbo`, o ponto existe nos dois; as combinações de parâmetros são diferentes. Essa distinção é importante ao interpretar as colunas.

## Saída

`PA_CHECK_OUTPUT_FILE` define o destino do arquivo Excel. Caminhos relativos são resolvidos a partir da raiz do repositório. O resultado contém:

- aba `PA Yael`, com os registros da primeira entrada;
- aba `PA Taise`, com os registros da segunda entrada;
- chave de comparação e colunas de observação, incluindo `Parâmetros ausentes`, `Pontos ausentes` e `Parâmetro ausentes geral`.

As descrições nas colunas indicam a origem do registro e a lista em que a combinação não foi encontrada. Uma célula vazia nessas colunas significa que aquela verificação não marcou a linha como diferença.

## Configuração e execução

1. Configure os caminhos dos dois planos e da saída em `.env`.
2. Confirme os três cabeçalhos nos dois arquivos.
3. Execute na raiz do repositório:

   ```powershell
   python -m src.pa_check
   ```

O arquivo de saída e sua pasta precisam estar graváveis. Se já existir um arquivo no mesmo caminho, ele poderá ser substituído.

## Compatibilidade conhecida

Na versão atual do código, os nomes das abas são passados como argumentos posicionais para `DataFrame.to_excel`. Com o pandas instalado na validação, a API exige esse argumento por nome e a execução termina com `TypeError` ao gravar o arquivo. A suíte automatizada reproduz o problema. Portanto, o formato de saída descrito acima é o resultado pretendido, mas o script precisa de uma correção no código para concluí-lo com essa versão do pandas.

## Interpretação e limitações

- A comparação é literal. Maiúsculas/minúsculas, espaços, acentos, pontuação e grafia podem produzir diferenças.
- A chave concatena os campos com hífens; padronize os dados de origem e revise manualmente eventuais combinações ambíguas.
- As comparações de ponto e parâmetro são globais: não consideram site/ponto, respectivamente. Um identificador ou parâmetro presente em outro site ainda conta como existente nessa verificação isolada.
- Uma diferença indica uma combinação não localizada na outra planilha; não prova que o dado de origem esteja incorreto.
- O script não remove duplicidades nem normaliza automaticamente os valores das colunas.

## Referência

Para visão geral, instalação e testes, consulte o [README principal](../README.md). O modelo das variáveis de configuração está em [`.env.example`](../.env.example).
