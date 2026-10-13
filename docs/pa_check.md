# `pa_check.py` — comparar dois Planos de Amostragem

## Em poucas palavras

Este script responde à pergunta: **“O que está diferente entre o PA da Yael e o PA da Taise?”**

Ele compara as duas planilhas e grava cada versão em uma aba de um novo Excel, com colunas que descrevem diferenças encontradas.

Ele não combina os planos em um único plano aprovado, nem decide qual versão está correta. Seu papel é destacar diferenças para revisão.

## Entradas

Os caminhos são definidos pelas variáveis `pa_yael` e `pa_taise`. Cada arquivo deve ser um Excel que contenha pelo menos as colunas:

- `Site`
- `Ponto`
- `Parâmetro`

Configure `PA_CHECK_YAEL_FILE` e `PA_CHECK_TAISE_FILE` no [`.env`](../.env). Os cabeçalhos de entrada podem ser alterados em `PA_CHECK_SITE_COLUMN`, `PA_CHECK_POINT_COLUMN` e `PA_CHECK_PARAMETER_COLUMN`; o script os converte para os nomes internos acima. As planilhas precisam ser fornecidas.

## Como funciona

1. `pandas.read_excel` carrega cada arquivo em uma tabela: `df_y` para Yael e `df_t` para Taise.
2. O script cria em cada tabela uma coluna `key` concatenando `Site`, `Ponto` e `Parâmetro`, separados por hífens.
3. Compara as chaves entre as planilhas:
   - uma chave presente no PA de Taise e ausente no de Yael recebe uma mensagem em `Parâmetros ausentes`;
   - uma chave presente no PA de Yael e ausente no de Taise recebe a mensagem inversa.
4. Faz também comparações separadas de pontos e parâmetros:
   - `Pontos ausentes`: ponto presente em um arquivo, mas não encontrado em nenhum lugar no outro;
   - `Parâmetro ausentes geral`: parâmetro que não aparece em nenhum lugar do outro arquivo.
5. Grava `df_y` e `df_t` nas abas `PA Yael` e `PA Taise` do arquivo `Planos de amostragem.xlsx`.

O caminho de saída é definido em `PA_CHECK_OUTPUT_FILE`. Caminhos relativos são resolvidos a partir da raiz do projeto.

## Exemplo sem termos técnicos

Suponha que o PA de Taise tenha:

| Site | Ponto | Parâmetro |
|---|---|---|
| Planta A | P-01 | Arsênio |

e o de Yael tenha:

| Site | Ponto | Parâmetro |
|---|---|---|
| Planta A | P-01 | Chumbo |

O script monta uma etiqueta única para cada linha:

- `Planta A-P-01-Arsênio`
- `Planta A-P-01-Chumbo`

Como as etiquetas não são iguais, a combinação de Arsênio será marcada como existente em um PA e ausente no outro. Mas `P-01` existe nos dois; portanto, ele não é um ponto ausente.

## O que fazem as principais variáveis?

- `df_y`: cópia em memória da planilha de Yael.
- `df_t`: cópia em memória da planilha de Taise.
- `key`: etiqueta formada com os três campos `Site`, `Ponto` e `Parâmetro`.
- `Parâmetros ausentes`, `Pontos ausentes` e `Parâmetro ausentes geral`: colunas de observação que o script acrescenta às tabelas antes de salvar.

O código usa expressões como `~df_y.key.isin(...)`. Em linguagem simples: “se a chave desta linha **não** (`~`) estiver (`isin`) na lista do outro arquivo, marque a linha”.

## O que as colunas de diferença significam

- **`Parâmetros ausentes`** compara a combinação completa site + ponto + parâmetro. Uma diferença nessa coluna quer dizer que aquela combinação não existe no outro PA.
- **`Pontos ausentes`** compara apenas o texto do ponto, sem considerar site ou parâmetro.
- **`Parâmetro ausentes geral`** compara apenas o texto do parâmetro, sem considerar site ou ponto.

Um registro pode ser marcado como diferente na combinação completa, mesmo que seu ponto e seu parâmetro existam separadamente no outro plano, mas em outra combinação.

## Exemplo simples

Considere estas linhas:

| Site | Ponto | Parâmetro |
|---|---|---|
| Unidade A | P-01 | Arsênio |

e, no outro arquivo:

| Site | Ponto | Parâmetro |
|---|---|---|
| Unidade A | P-01 | Chumbo |

O ponto `P-01` existe nos dois arquivos, então não deve ser marcado como ponto ausente. Porém, a combinação `Unidade A-P-01-Arsênio` não existe no segundo arquivo; ela será sinalizada como diferença de parâmetro/combinação.

## Cuidados antes de usar o resultado

- A comparação é literal. Diferenças de maiúsculas/minúsculas, espaços, acentos, hífens ou grafia podem aparecer como divergências mesmo se as pessoas considerarem os valores equivalentes.
- A concatenação pressupõe que as colunas usadas sejam texto. Valores vazios ou tipos diferentes podem produzir chaves inesperadas ou erros.
- A comparação de ponto e parâmetro ignora o site. Se o mesmo identificador for reutilizado em sites diferentes, as colunas de comparação separada podem não refletir uma diferença por site.
- A coluna `Parâmetros ausentes` é inicializada com texto vazio; as demais observações não preenchidas podem conter valores ausentes. Células em branco não significam que a comparação não ocorreu.
- O script escreve colunas novas e um arquivo existente no caminho de saída pode ser substituído.
- `os` é importado, mas não é usado no código atual.

## O que este script não faz

- Não escolhe qual plano é o correto.
- Não cria um plano combinado.
- Não compara checklists nem resultados de laboratório.
- Não padroniza automaticamente grafias equivalentes (como espaços ou diferenças de maiúsculas/minúsculas).

Use o arquivo gerado como uma lista de pontos para revisar com as pessoas responsáveis pelos planos.
