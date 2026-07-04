# Spec — Estilos, seleção e posições

## Alvo de `aplicarEstilos()` (API fluente)

`StyleManager.aplicarEstilos()` escolhe o alvo nesta ordem:

1. **Dado** `selecionar().todaPlanilha()`, **Então** o estilo vale para a
   planilha inteira.
2. **Dado** `selecionar().intervalo("A1","C3")`, **Então** vale para o intervalo.
3. **Dado** que dados acabaram de ser inseridos (última linha inserida), **Então**
   vale para a **linha** recém-inserida.
4. **Dado** `selecionar().celula("A1")` sem inserção posterior, **Então** vale
   apenas para **A1**.
5. **Dado** nada selecionado, **Então** no-op.

Após aplicar, o estado de seleção é resetado.

### Regressão coberta

- **Dado** `celula("A1").inserirDados("x")` e depois `celula("A5").inserirDados("y")`,
  **Quando** `celula("A1").aplicarEstilos().aplicarNegrito()`, **Então** **A1**
  fica em negrito e **A5 não** (a seleção nova reseta o "último índice inserido").

## `aplicarTodasAsBordas()` / `bordas(intervalo)`

- **Dado** um intervalo com células **sem nenhuma borda**, **Quando**
  `aplicarTodasAsBordas()` (facade: `bordas(intervalo)`), **Então** todas as
  células não vazias do intervalo recebem borda `THIN` em todos os lados.
- **Dado** uma célula que já tem uma borda `THICK` em algum lado, **Quando**
  a mesma operação é aplicada, **Então** essa célula é **ignorada** (a borda
  espessa não é rebaixada para fina).
- Regressão: antes desta sessão, a checagem de "já tem borda espessa" estava
  invertida e testava "não tem borda nenhuma" — por isso a operação
  silenciosamente não fazia nada no caso comum (célula sem borda prévia). Ver
  `EstiloCelulaTest.deveAplicarBordasFinasEmCelulasSemBorda`.

## `todasAsBordasEmTudo()` / `contornarTudo()`

- **Dado** dados em `A1:B2`, **Quando** `todasAsBordasEmTudo()`, **Então** a área
  usada recebe bordas (externas espessas, internas finas) — sem usar área fixa.
- **Dado** planilha vazia, **Então** não lança exceção (no-op).

## `PosicaoConverter` — posições e limites

- `"A1"` → `[0, 0]`; `"B2"` → `[1, 1]`; `"AA10"` → `[26, 9]`.
- `converterColuna`: `"A"`→0, `"Z"`→25, `"AA"`→26, `"XFD"`→16383.
- `converterIndice` é o inverso de `converterColuna`.
- Aceita minúsculas (`"a1"` == `"A1"`).
- **Lança `PosicaoInvalidaException`** para: nulo/vazio, formato inválido
  (`"1A"`, `"ABC"`), linha `0`, linha `> 1.048.576`, coluna `> XFD`.

## Operações de coluna (`ManipuladorPlanilha`)

Sobre a planilha `H1 H2 H3` / `a1 b1 c1` (colunas A,B,C):

- `moverColuna("A","C")` → `H2 H3 H1` (A vai para C, demais deslocam).
- `removerColuna("B")` → `H1 H3` (C desloca para B).
- `limparColuna("B")` → B fica vazia, A e C intactas.
- `inserirColunaEntre("A","B")` → coluna vazia entra em B, empurrando as demais.
- `inserirColunaEntre` exige colunas **adjacentes**, senão lança
  `IllegalArgumentException`.
