# Spec — Coerção de tipos (texto x número)

Ao inserir dados via `escrever(...)`, `inserirDados(...)` ou
`escreverLinha/Coluna/Tabela`, a biblioteca decide se o valor vira **número** ou
**texto**. Regras implementadas em `InsersorDeDados.deveManterComoTexto`.

## Vira NÚMERO

- **Dado** `"10"`, **Então** número `10`.
- **Dado** `"3.5"`, **Então** número `3.5`.
- **Dado** `"-400"`, **Então** número `-400`.
- **Dado** `"1.2e3"`, **Então** número `1200`.
- **Dado** `"0"` ou `"0.5"`, **Então** número (zero sozinho / decimal são ok).

## Permanece TEXTO

- **Dado** `"007"` (zero à esquerda), **Então** texto `"007"`.
- **Dado** `"01310-100"` (CEP com traço), **Então** texto.
- **Dado** `"NaN"` / `"Infinity"`, **Então** texto (não são números "de verdade").
- **Dado** um inteiro com **16+ dígitos** (CPF, CNPJ, cartão), **Então** texto
  (evita perda de precisão do `double`).
- **Dado** qualquer palavra (`"Produto X"`), **Então** texto.

## Forçar texto sempre

- **Quando** `escreverTexto("A1", "7")`, **Então** A1 é texto `"7"` — **nunca**
  converte, mesmo sendo número limpo. Use para CPF/CNPJ/telefone/códigos.

## Racional

Números de quantidade/valor devem somar e formatar como número. Já
identificadores (documento, CEP, código) não são "quantidades" — convertê-los
quebraria zeros à esquerda e a precisão. A regra é conservadora: na dúvida entre
quebrar um cálculo legítimo e preservar um identificador raro, o padrão converte
números limpos e curtos; para identificadores, o usuário usa `escreverTexto`.
