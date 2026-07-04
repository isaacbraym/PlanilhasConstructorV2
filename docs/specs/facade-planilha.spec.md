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
- **Quando** `procurarValor("E1", "D1", "A2:B3", 2)`, **Então** E1 contém
  `VLOOKUP(D1,A2:B3,2,FALSE)`; `procurarValorNaAba(..., "Produtos", ...)`
  qualifica o intervalo como `'Produtos'!...` (PROCV com correspondência exata).

## Formatos

- **Quando** `formatarComoMoeda("B2")` ou `formatarComoContabil("B2")`,
  **Então** as células numéricas da coluna B (a partir de B2) usam formato com
  `R$`.
- **Quando** `formatarComoTexto("B2")`, **Então** os números viram texto (sem
  `.0` para inteiros).
- **Quando** `formatarComoPorcentagem("B2")`, **Então** a coluna usa formato de
  porcentagem (`0,15` exibe `15%`).
- `larguraColuna("A", n)` define a largura (≈ n caracteres, 0–255);
  `alturaLinha(l, p)` define a altura em pontos; `congelar(linhas, colunas)`
  fixa linhas/colunas ao rolar.

## Formatação condicional

- **Quando** `realcarSeMaiorQue(intervalo, valor, cor)`, **Então** é criada uma
  regra de formatação condicional `GT` (maior que) sobre o intervalo, com cor de
  fundo igual à informada; `realcarSeMenorQue`/`realcarSeEntre` seguem o mesmo
  padrão com `LT`/`BETWEEN`.
- **Quando** `realcarSeIgual(intervalo, valor, cor)`, **Então** a regra usa
  `EQUAL`; se `valor` for `Number`, a fórmula usa o número puro; se for outro
  tipo (texto), a fórmula envolve o valor em aspas duplas (sintaxe de fórmula
  do Excel).
- **Quando** `escalaDeCores(intervalo)`, **Então** é criada uma regra de escala
  de 3 cores (vermelho/amarelo/verde) com limiares MIN / PERCENTIL 50 / MAX.
- Todas essas regras assumem planilha `.xlsx` (XSSF) — coerente com o resto da
  biblioteca, que só cria/edita workbooks XSSF.

## Lista suspensa (dropdown)

- **Quando** `listaSuspensa(intervalo, "A", "B", "C")`, **Então** uma validação
  de dados do tipo lista é criada sobre o intervalo, com exatamente essas
  opções, seta do menu visível (`getSuppressDropDownArrow()` verdadeiro) e uma
  caixa de erro amigável para valores fora da lista.
- Sem opções → `DadosInvalidosException`. Soma dos textos das opções > 255
  caracteres → `DadosInvalidosException` (limite real do Excel para listas com
  valores fixos), sugerindo `listaSuspensaDoIntervalo`.
- **Quando** `listaSuspensaDoIntervalo(intervaloDestino, "F2:F5")`, **Então** a
  restrição referencia a fórmula `$F$2:$F$5` (convertida para referência
  absoluta automaticamente); se o intervalo já contiver `!` (ex.:
  `"Opcoes!$A$1:$A$3"`), é usado exatamente como informado, sem conversão.

## Gráficos

- **Quando** `graficoDeBarras(titulo, categorias, valores, celula)`, **Então**
  um gráfico de barras verticais (`BarDirection.COL`) é desenhado ancorado na
  célula informada, com uma série cujo título é `titulo`, cujas categorias vêm
  do intervalo `categorias` e valores do intervalo `valores`.
- `graficoDeLinha` segue o mesmo contrato com tipo `LINE` (cria eixos de
  categoria e de valor, como as barras).
- `graficoDePizza` **não** cria eixos (`chart.getAxes()` vazio) e ativa cores
  variadas por fatia (`setVaryColors(true)`).
- Todos sobrevivem a salvar em disco e reabrir (round-trip OOXML verificado).

## Imagens

- **Quando** `inserirImagem(celula, caminho)`, **Então** a imagem é inserida
  ancorada na célula, no tamanho natural (pixels reais do arquivo convertidos
  para EMU via DPI padrão).
- **Quando** `inserirImagem(celula, caminho, escala)`, **Então** a imagem é
  inserida redimensionada por esse fator em relação ao tamanho natural (ex.:
  `0.5` = metade, `2.0` = dobro).
- Aceita apenas `.png`/`.jpg`/`.jpeg` (senão `DadosInvalidosException`);
  arquivo inexistente/ilegível → `ArquivoException`.

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

## Ordenar

- **Dado** cabeçalho + `Carlos/30`, `Ana/25`, `Bia/40`, **Quando**
  `ordenarPorCrescente("A")`, **Então** a ordem dos dados vira `Ana`, `Bia`,
  `Carlos`, com a coluna Idade viajando junto e o cabeçalho intacto.
- Números ordenam numericamente; `ordenarPorDecrescente` inverte;
  `ordenarPor(coluna, crescente, linhaInicial)` permite ordenar sem cabeçalho.

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

## Validação de entrada (número/data com limites)

- **Quando** `validarNumeroEntre(intervalo, min, max)`, **Então** a restrição
  criada é do tipo `DECIMAL`, operador `BETWEEN`, com `formula1`/`formula2`
  iguais aos limites; `validarInteiroEntre` é igual mas tipo `INTEGER`.
- **Quando** `validarDataEntre(intervalo, minimo, maximo)` (com `LocalDate`),
  **Então** a restrição é do tipo `DATE` (formato interno `yyyy-MM-dd`).
- Todas mostram uma caixa de erro amigável para valores fora do limite e
  sobrevivem a salvar em disco e reabrir.

## Impressão

- `orientacaoPaisagem()`/`orientacaoRetrato()` alternam
  `PrintSetup.getLandscape()`.
- `areaDeImpressao(intervalo)` registra a área em `workbook.getPrintArea(indiceAba)`,
  convertendo para referência absoluta automaticamente.
- `ajustarImpressaoEmPaginas(larguraPaginas, alturaPaginas)` ativa
  `sheet.getFitToPage()` e define `PrintSetup.getFitWidth()`/`getFitHeight()`.
- Todos sobrevivem a salvar em disco e reabrir (round-trip verificado).

## Proteção

- **Dado** nada, **Então** `sheet.getProtect()` é falso por padrão.
- **Quando** `protegerPlanilha(senha)`, **Então** `sheet.getProtect()` vira
  verdadeiro; toda célula sem tratamento prévio continua **travada** (padrão
  do Excel).
- **Quando** `desbloquearCelulas(intervalo)` **antes** de `protegerPlanilha`,
  **Então** as células do intervalo ficam com `CellStyle.getLocked() == false`
  e continuam editáveis mesmo com a planilha protegida; células fora do
  intervalo **não são afetadas** (o estilo é clonado antes de destravar, nunca
  mutado — evita destravar acidentalmente outras células que compartilham o
  mesmo `CellStyle` padrão). Cria a célula se ela ainda não existir.

## Salvar e recursos

- **Quando** `salvar("caminho.xlsx")`, **Então** um arquivo `.xlsx` válido é
  gravado e pode ser reaberto.
- **Quando** usada em `try (Planilha p = ...)`, **Então** os recursos são
  liberados ao final; fora disso, chame `fechar()`.

## Escape hatch

- `avancado()` devolve `IPlanilha` (API fluente completa).
- `workbook()` devolve o `Workbook` do Apache POI.
- `estilo(pos)` devolve `EstiloCelula` para encadear vários estilos de uma vez.
