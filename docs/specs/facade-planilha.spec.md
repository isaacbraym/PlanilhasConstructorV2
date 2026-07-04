# Spec — Facade `Planilha`

A classe `com.abnote.planilhas.Planilha` é a API amigável. Todo método (exceto
`fechar`/`close`) devolve `Planilha` para encadeamento, salvo `estilo(...)`,
`avancado()` e `workbook()`.

## Criação e abas

- **Dado** nada, **Quando** `Planilha.nova("Vendas")`, **Então** existe uma
  planilha com a aba "Vendas" ativa.
- **Dado** um arquivo `.xlsx` existente, **Quando** `Planilha.abrir(caminho)`,
  **Então** o conteúdo é carregado, a primeira aba fica ativa e pode ser editado
  com os mesmos comandos; arquivo inexistente/ inválido lança `ArquivoException`.
- **Quando** `novaAba("Resumo")`, **Então** uma nova aba "Resumo" é criada e
  passa a ser a ativa.
- **Quando** `duplicarAba("Cópia")`, **Então** a aba ativa é copiada (conteúdo e
  estilos) para "Cópia", que passa a ser a ativa, e o workbook tem +1 aba.
- **Quando** `irParaAba("X")` e "X" não existe, **Então** lança exceção.

## Escrita

- **Quando** `escrever("A1", 10)`, **Então** A1 é numérico `10`.
- **Quando** `escrever("A1", "Nome")`, **Então** A1 é texto `"Nome"`.
- **Quando** `escreverTexto("A1", "007")`, **Então** A1 é **texto** `"007"`
  (zero preservado).
- **Quando** `escreverLinha("A1", "a", "b", "c")`, **Então** A1/B1/C1 recebem
  a/b/c.
- **Quando** `escreverColuna("A1", 1, 2, 3)`, **Então** A1/A2/A3 recebem 1/2/3.
- **Dado** uma linha já preenchida, **Quando** `adicionarLinha("x", "y")`,
  **Então** os valores vão para a **próxima** linha vazia, a partir da coluna A.
- **Quando** `escreverTabela("A1", [[a,b],[1,2]])`, **Então** a matriz é escrita
  a partir de A1.
- **Quando** `escreverData("A1", data)`, **Então** A1 é uma data formatada
  `dd/MM/aaaa` (`DateUtil.isCellDateFormatted` verdadeiro); `escreverDataHora`
  usa `dd/MM/aaaa HH:mm`.

## Fórmulas

- **Dado** `A1..A3 = 10,20,30`, **Quando** `somar("A4", "A1:A3")`, **Então**
  A4 contém a fórmula `SUM(A1:A3)` e avalia `60`.
- `media`, `contar`, `minimo`, `maximo` seguem o mesmo padrão com AVERAGE,
  COUNT, MIN, MAX.
- **Quando** `seEntao("D2", "C2>100", "Alto", "Baixo")`, **Então** D2 contém
  `IF(C2>100,"Alto","Baixo")`.
- **Quando** `multiplicar("D2", "B2", "C2")`, **Então** D2 contém `B2*C2`;
  `subtrair`/`dividir` geram `B2-C2` / `B2/C2`.
- **Quando** `formula("D2", "=B2*C2")`, **Então** D2 recebe a fórmula `B2*C2`
  (o "=" inicial é opcional); fórmula inválida lança `FormulaException`.
- **Dado** `preencherColuna("D", 2, 3, "B{}*C{}")`, **Então** D2 = `B2*C2` e
  D3 = `B3*C3` (o `{}` vira o número da linha).

## Formatos

- **Quando** `formatarComoMoeda("B2")` ou `formatarComoContabil("B2")`,
  **Então** as células numéricas da coluna B (a partir de B2) usam formato com
  `R$`.
- **Quando** `formatarComoTexto("B2")`, **Então** os números viram texto (sem
  `.0` para inteiros).

## Colunas e linhas

- `moverColuna(de, para)`, `removerColuna(c)`, `limparColuna(c)`,
  `inserirColunaEntre(esq, dir)` operam na coluna indicada (ver
  `estilos-e-selecao` para índices).
- **Quando** `duplicarColuna("A", "D")`, **Então** conteúdo e estilo da coluna A
  são copiados para a coluna D.
- **Quando** `duplicarLinha(1, 5)`, **Então** a linha 1 é copiada para a linha 5.

## Aparência

- `negrito`, `italico`, `corDoTexto`, `corDeFundo`, `fonte`, `tamanhoDaFonte`
  aceitam **célula** (`A1`) ou **intervalo** (`A1:C1`).
- `centralizar`, `bordas`, `mesclar` operam em **intervalo** (célula única é
  tratada como intervalo de 1 célula).
- `contornarTudo`, `removerLinhasDeGrade`, `ajustarColunas`,
  `congelarPrimeiraLinha`, `filtrosNoCabecalho` operam na aba ativa.
- **Regra**: estilizar deve ser feito **depois** de escrever — estilos de fonte
  só afetam células que já existem.

## Buscar e filtrar linhas

Comparação por texto; números casam sem `.0` (ex.: `10` casa com `"10"`).

- **Dado** cabeçalho + `Ana/SP`, `Bia/RJ`, `Cid/SP`, **Quando**
  `buscarLinhas("B", "SP")`, **Então** retorna `[2, 4]` (números 1-based).
- **Quando** `contarLinhasOnde("B", "SP")`, **Então** retorna `2`.
- **Quando** `copiarLinhasParaAba("B", "SP", "SoSP")`, **Então** a aba "SoSP" é
  criada (se preciso) com as linhas correspondentes; a origem **não** muda.
- **Quando** `removerLinhasOnde("B", "SP")`, **Então** as linhas somem e as de
  baixo sobem.
- **Quando** `moverLinhasParaAba("B", "SP", "Arquivo")`, **Então** copia para a
  aba e remove da origem.

## Salvar e recursos

- **Quando** `salvar("caminho.xlsx")`, **Então** um arquivo `.xlsx` válido é
  gravado e pode ser reaberto.
- **Quando** usada em `try (Planilha p = ...)`, **Então** os recursos são
  liberados ao final; fora disso, chame `fechar()`.

## Escape hatch

- `avancado()` devolve `IPlanilha` (API fluente completa).
- `workbook()` devolve o `Workbook` do Apache POI.
- `estilo(pos)` devolve `EstiloCelula` para encadear vários estilos de uma vez.
