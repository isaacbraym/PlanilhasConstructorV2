# Spec — Facade `Planilha`

A classe `com.abnote.planilhas.Planilha` é a API amigável. Todo método (exceto
`fechar`/`close`) devolve `Planilha` para encadeamento, salvo `estilo(...)`,
`avancado()` e `workbook()`.

## Criação e abas

- **Dado** nada, **Quando** `Planilha.nova("Vendas")`, **Então** existe uma
  planilha com a aba "Vendas" ativa.
- **Quando** `Planilha.duplicarArquivo(origem, destino)`, **Então** o arquivo
  em `destino` é uma cópia byte-a-byte de `origem` (todas as abas, fórmulas,
  estilos, gráficos, imagens) e os dois passam a ser independentes — editar um
  não afeta o outro. Origem inexistente, origem/destino nulos ou vazios, e
  caminho inválido → `ArquivoException` (não `NullPointerException`/
  `InvalidPathException`). Não precisa abrir a planilha (é uma cópia direta do
  arquivo, mais simples e robusta do que reconstruir o workbook via POI).
- **Dado** um arquivo `.xlsx` existente, **Quando** `Planilha.abrir(caminho)`,
  **Então** o conteúdo é carregado, a primeira aba fica ativa e pode ser editado
  com os mesmos comandos; arquivo inexistente/ inválido lança `ArquivoException`.
- **Dado** um arquivo `.xls` (formato antigo, HSSF), **Quando**
  `Planilha.abrir(caminho)`, **Então** lança `ArquivoException` com mensagem
  orientando a converter para `.xlsx` — **nunca** `ClassCastException` (vários
  métodos da facade fazem cast interno para `XSSFSheet`, que quebraria sem
  essa checagem prévia em `PlanilhaBase.abrirPlanilha`).
- **Quando** `novaAba("Resumo")`, **Então** uma nova aba "Resumo" é criada e
  passa a ser a ativa.
- **Quando** `duplicarAba("Cópia")`, **Então** a aba ativa é copiada (conteúdo e
  estilos) para "Cópia", que passa a ser a ativa, e o workbook tem +1 aba.
- **Quando** `duplicarAba(nome)` recebe um nome já existente ou inválido para o
  Excel, **Então** lança `IllegalArgumentException` antes de clonar; o workbook
  permanece com as mesmas abas, sem cópia parcial tipo `"Original (2)"`.
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

## Leitura (o "inverso" de escrever)

- **Quando** `ler(celula)`, **Então** devolve `String` para célula de texto,
  `Double` para numérica, `Boolean` para booleana, `LocalDateTime` para
  célula formatada como data/hora (via `DateUtil.isCellDateFormatted`), ou
  `null` para célula vazia/inexistente. Fórmulas são avaliadas antes: o
  resultado segue a mesma regra de tipos.
- **Quando** `lerTexto(celula)`, **Então** devolve o texto **formatado como
  aparece no Excel** (via `DataFormatter`, respeitando o formato numérico da
  célula — ex.: moeda vira `"R$ 1.234,56"`); `""` se a célula não existir.
- `lerNumero`/`lerData` são atalhos sobre `ler(...)`: devolvem `null` se o
  valor lido não for do tipo esperado (não lançam exceção).
- **Quando** `lerTabela(celulaCabecalho)`, **Então** devolve as linhas de
  dados (sem o cabeçalho) como `List<List<Object>>`, usando a mesma detecção
  de largura (pelo cabeçalho) e altura (pela primeira coluna) de
  `adicionarTotais`; tabela sem dados → lista vazia.
- **Quando** `contarLinhasPreenchidas(coluna)`, **Então** devolve quantas
  linhas têm célula não vazia nessa coluna (delega para
  `IPlanilha.getNumeroDeLinhas`, já existente na API fluente).
- Nenhum método de leitura cria linha/célula na planilha (ao contrário de
  `escrever`, que cria a célula se faltar).

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
- **Quando** `definirNome("Precos", "B2:B100")`, **Então** um `Name` é
  registrado no workbook com `refersToFormula` igual a
  `'AbaAtual'!$B$2:$B$100` (referência absoluta, qualificada com a aba atual).
  O nome funciona dentro de `formula(...)` e `procurarValor`/
  `procurarValorNaAba` (ex.: `formula("D1", "SUM(Precos)")` avalia
  corretamente), e também em `somar`/`media`/`contar`/`minimo`/`maximo`
  (ex.: `somar("D1", "Precos")` gera `SUM(Precos)` e avalia corretamente).
- **Quando** o nome não segue as regras do Excel (começa com dígito, tem
  espaço, ou é igual a uma referência de célula como `"A1"`/`"$B$2"`, ou é
  vazio), **Então** lança `DadosInvalidosException` com a causa original —
  nunca a `IllegalArgumentException` crua do POI.

## Formatos

- **Quando** `formatarComoMoeda("B2")` ou `formatarComoContabil("B2")`,
  **Então** as células numéricas da coluna B (a partir de B2) usam formato com
  `R$`.
- **Quando** `formatarComoTexto("B2")`, **Então** os números viram texto (sem
  `.0` para inteiros).
- **Quando** `formatarComoPorcentagem("B2")`, **Então** a coluna usa formato de
  porcentagem (`0,15` exibe `15%`).
- **Quando** `formatarComoPersonalizado(intervalo, formatoExcel)`, **Então** a
  coluna recebe exatamente esse formato (`CellStyle.getDataFormatString()`
  igual ao informado); o mesmo formato pedido duas vezes reaproveita o mesmo
  `CellStyle` (cache por string de formato em `FormatosDeCelula`). A sintaxe
  do formato **não é validada** — é responsabilidade de quem chama.
- `larguraColuna("A", n)` define a largura (≈ n caracteres, 0–255);
  `alturaLinha(l, p)` define a altura em pontos; `congelar(linhas, colunas)`
  fixa linhas/colunas ao rolar.

## adicionarTotais (linha de totais automática)

- **Dado** um cabeçalho em `A1` e dados nas linhas 2-3, **Quando**
  `adicionarTotais("A1")`, **Então** a linha 4 recebe, para cada coluna cujos
  valores existentes são **todos numéricos**, uma fórmula `SUM(...)` cobrindo
  exatamente as linhas de dados detectadas; a **primeira** coluna não numérica
  encontrada recebe o texto `"Total"` (as demais colunas não numéricas ficam
  em branco).
- Celulas com formula entram nessa decisao pelo resultado calculado: uma coluna
  com `D2=B2*C2` e `D3=B3*C3` recebe `SUM(D2:D3)` se ambas avaliarem como
  numero, e esse contrato deve sobreviver a salvar/reabrir o `.xlsx`.
- A largura da tabela é detectada pelo cabeçalho (para de contar colunas no
  primeiro cabeçalho vazio); a altura é detectada pela primeira coluna (para
  na primeira linha em branco nessa coluna, a partir do início dos dados).
- Uma coluna com qualquer célula não numérica (mesmo que as outras sejam
  números) **não** recebe soma.
- Tabela sem nenhuma linha de dados → nenhuma linha de totais é criada
  (no-op).

## colarComoValores (congelar fórmulas em valores)

- **Dado** uma célula com fórmula, **Quando** `colarComoValores(intervalo)`,
  **Então** a célula passa a ter o **valor calculado** (número, texto ou
  booleano, conforme o resultado) e deixa de ser do tipo `FORMULA`.
- Células sem fórmula dentro do intervalo não são alteradas.
- `colarComoValores()` sem argumento aplica a toda a área usada da aba atual;
  em aba vazia é no-op.
- Internamente é preciso chamar `Cell.removeFormula()` **antes** de
  `setCellValue(...)` — sozinho, `setCellValue` numa célula `FORMULA` só
  atualiza o valor em cache, sem trocar o tipo da célula (comportamento do
  POI verificado empiricamente, não documentado).

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
- **Quando** `barrasDeDados(intervalo, cor)`, **Então** é criada uma regra
  `DataBarFormatting` com limiares MIN/MAX e cor igual à informada.
- **Quando** `iconesSemaforo(intervalo)`, **Então** é criada uma regra
  `IconMultiStateFormatting` com `IconSet.GYR_3_TRAFFIC_LIGHTS` e limiares
  percentuais padrão `0`/`33`/`66`.
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
- **Quando** `listaSuspensaDoIntervalo(intervaloDestino, "Apoio", "F2:F5")`,
  **Então** a restrição referencia a fórmula `'Apoio'!$F$2:$F$5`, validando
  que a aba existe antes de criar a lista. Nomes de aba com espaços são
  protegidos por aspas simples e sobrevivem a round-trip (`XSSFWorkbook`
  fresco após salvar/reabrir).
- Referências já absolutas (ex.: `"$F$2:$F$5"`) continuam exatamente assim; a
  facade não duplica cifrões ao montar listas suspensas ou nomes definidos.

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
- **Quando** `graficoDeBarras(titulo, abaCategorias, categorias, abaValores,
  valores, celula)` (sobrecarga de 6 argumentos), **Então** categorias e
  valores são lidos de `abaCategorias`/`abaValores` (que podem ser diferentes
  entre si e diferentes da aba atual, onde o gráfico é ancorado). Confirmado
  empiricamente que `Series.getCategoryData().getDataRangeReference()` traz o
  nome da aba de origem (ex.: `"Produtos!$A$2:$A$4"`), não da aba do gráfico.
  Antes desta sobrecarga existir, montar um gráfico assim silenciosamente lia
  dado errado (intervalo resolvido contra a aba atual, sem erro nenhum) — por
  isso a versão de 4 argumentos permanece restrita a uma única aba (a atual);
  cross-sheet exige a sobrecarga explícita. Aba inexistente lança
  `IllegalArgumentException`. Mesmo contrato para `graficoDePizza`/
  `graficoDeLinha`.

## Comentários

- **Quando** `comentario(celula, texto)`, **Então** a célula ganha um
  `Comment` (nota) cujo `getString().getString()` é igual ao texto informado;
  sobrevive a salvar em disco e reabrir.

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
- Se a celula copiada contiver formula, referencias relativas acompanham a nova
  posicao como no Excel: `A2*2` em `B2` vira `A5*2` ao duplicar a linha para 5,
  e `A2+B2` em `C2` vira `B2+C2` ao duplicar a coluna C para D. Partes
  absolutas (`$A$2`, `B$2`, `$A2`) permanecem fixas no eixo correspondente.
  Esse contrato deve sobreviver a salvar e reabrir o `.xlsx`.

## Aparência

- `negrito`, `italico`, `corDoTexto`, `corDeFundo`, `fonte`, `tamanhoDaFonte`
  aceitam **célula** (`A1`) ou **intervalo** (`A1:C1`).
- `centralizar`, `bordas`, `mesclar` operam em **intervalo** (célula única é
  tratada como intervalo de 1 célula).
- `contornarTudo`, `removerLinhasDeGrade`, `ajustarColunas`,
  `congelarPrimeiraLinha`, `filtrosNoCabecalho` operam na aba ativa.
- **Regra**: estilizar deve ser feito **depois** de escrever — estilos de fonte
  só afetam células que já existem.
- **Quando** `desmesclar(intervalo)` sobre um intervalo **exatamente igual**
  a um já mesclado com `mesclar`, **Então** a mesclagem é desfeita
  (`getNumMergedRegions()` diminui em 1). Se o intervalo não estiver mesclado
  (ou não corresponder exatamente a uma região mesclada existente), é no-op —
  não lança exceção nem afeta outras mesclagens.

## Cor da aba

- **Quando** `corDaAba(cor)`, **Então** `XSSFSheet.getTabColor()` da aba ativa
  tem o RGB exato de `cor` (`CorEnum.getRed/getGreen/getBlue`).

## Ordenar

- **Dado** cabeçalho + `Carlos/30`, `Ana/25`, `Bia/40`, **Quando**
  `ordenarPorCrescente("A")`, **Então** a ordem dos dados vira `Ana`, `Bia`,
  `Carlos`, com a coluna Idade viajando junto e o cabeçalho intacto.
- Números ordenam numericamente; `ordenarPorDecrescente` inverte;
  `ordenarPor(coluna, crescente, linhaInicial)` permite ordenar sem cabeçalho.
- Se a coluna de ordenação contiver fórmulas, a comparação usa o resultado
  calculado da fórmula (ex.: `B2*2` com resultado `4` ordena como número `4`,
  não pelo texto `"B2*2"`). Ao reescrever linhas, referências relativas dentro
  das fórmulas acompanham a nova posição da linha (`B3*2` vira `B2*2` se a
  linha for para a linha 2); o contrato deve sobreviver a salvar e reabrir o
  `.xlsx`.
- Atributos da linha acompanham os dados ordenados: altura personalizada,
  `Row.getZeroHeight()` (linha oculta) e `RowStyle` viajam junto com a linha.
  Uma posição que recebe uma linha sem atributos personalizados volta ao padrão
  da planilha, sem herdar altura/ocultação da linha que ocupava aquela posição
  antes da ordenação.

## Buscar e filtrar linhas

Comparação por texto; números casam sem `.0` (ex.: `10` casa com `"10"`).
Fórmulas são avaliadas antes da comparação, então uma célula com `=A2*2` e
resultado `20` casa com `"20"`, não com o texto `"A2*2"` da fórmula.

- **Dado** cabeçalho + `Ana/SP`, `Bia/RJ`, `Cid/SP`, **Quando**
  `buscarLinhas("B", "SP")`, **Então** retorna `[2, 4]` (números 1-based).
- **Quando** `contarLinhasOnde("B", "SP")`, **Então** retorna `2`.
- **Quando** `copiarLinhasParaAba("B", "SP", "SoSP")`, **Então** a aba "SoSP" é
  criada (se preciso) com as linhas correspondentes; a origem **não** muda.
  Formulas copiadas sao ajustadas para a nova linha/aba de destino (ex.:
  `A2*2` na origem vira `A1*2` quando a linha 2 e colada na primeira linha da
  aba "SoSP").
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

## Ocultar / exibir

- `ocultarLinha(n)`/`exibirLinha(n)` alternam `Row.getZeroHeight()`.
- `ocultarColuna(c)`/`exibirColuna(c)` alternam `Sheet.isColumnHidden(indice)`.
- `ocultarAba(nome)`/`exibirAba(nome)` alternam `Workbook.isSheetHidden(indice)`.
- **Regra de segurança**: `ocultarAba` recusa (lança `DadosInvalidosException`)
  ocultar a **única aba visível** do arquivo — verificado empiricamente que o
  POI permite silenciosamente criar um arquivo sem nenhuma aba visível, que o
  Excel não consegue abrir.
- Aba inexistente em `ocultarAba`/`exibirAba` → `IllegalArgumentException`
  (mesma convenção já usada por `selecionarSheet`, não a hierarquia de
  `PlanilhaException` — consistência com o comportamento pré-existente).

## Impressão

- `orientacaoPaisagem()`/`orientacaoRetrato()` alternam
  `PrintSetup.getLandscape()`.
- `areaDeImpressao(intervalo)` registra a área em `workbook.getPrintArea(indiceAba)`,
  convertendo para referência absoluta automaticamente.
- `ajustarImpressaoEmPaginas(larguraPaginas, alturaPaginas)` ativa
  `sheet.getFitToPage()` e define `PrintSetup.getFitWidth()`/`getFitHeight()`.
- `margensDeImpressao(superiorCm, inferiorCm, esquerdaCm, direitaCm)` recebe
  centímetros, converte para polegadas (unidade nativa do OOXML/POI), grava
  `PageMargin.TOP`/`BOTTOM`/`LEFT`/`RIGHT` e persiste no
  arquivo salvo. Margem negativa, infinita ou `NaN` lança
  `DadosInvalidosException`.
- `cabecalhoDeImpressao(centro)`/`cabecalhoDeImpressao(esquerda, centro, direita)`
  escrevem em `sheet.getHeader()` (Left/Center/Right). Sobrecarga de 1 argumento
  preenche só o centro, deixando esquerda/direita vazios (`""`, nunca `null`).
- `rodapeDeImpressao(centro)`/`rodapeDeImpressao(esquerda, centro, direita)` —
  mesma ideia para `sheet.getFooter()`.
- Ambos traduzem marcadores amigáveis via `CodigosDeImpressao.traduzir`:
  `{pagina}`→`&P`, `{total}`→`&N`, `{data}`→`&D`, `{hora}`→`&T`,
  `{arquivo}`→`&F`, `{aba}`→`&A` — a sintaxe nativa do Excel (`&P`/`&N`/...)
  não é descobrível sem documentação, então o usuário nunca precisa vê-la.
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

## Agrupamento (outline)

- **Quando** `agruparLinhas(2, 4)`, **Então** as linhas 2-4 (1-based) recebem
  `outlineLevel == 1` no arquivo salvo; linhas fora do grupo permanecem em
  `outlineLevel == 0`; **nenhuma linha fica oculta** (`getZeroHeight() ==
  false` em todas). `agruparColunas("B","D")` é o equivalente para colunas
  (`Sheet.getColumnOutlineLevel`).
- **Deliberadamente não há** um parâmetro "recolhido"/"colapsado": um teste
  empírico mostrou que `Sheet.setRowGroupCollapsed(...)` marca **todas** as
  linhas da planilha como ocultas (`hidden="true"` no XML), não só as do
  grupo — um risco real de esconder dados do usuário sem ele perceber. A
  facade só cria a estrutura de agrupamento; expandir/recolher fica a cargo
  do próprio usuário clicando no "+"/"-" do Excel.

## Salvar e recursos

- **Quando** `salvar("caminho.xlsx")`, **Então** um arquivo `.xlsx` válido é
  gravado e pode ser reaberto.
- **Quando** usada em `try (Planilha p = ...)`, **Então** os recursos são
  liberados ao final; fora disso, chame `fechar()`.

## Carga / milhares de linhas

- **Dado** uma planilha com 3.000 linhas de dados, **Quando** a facade escreve
  linhas, aplica `somar`, `congelarPrimeiraLinha`, `filtrosNoCabecalho`,
  salva e reabre com `XSSFWorkbook` fresco, **Então** a última linha, filtros,
  freeze pane e fórmulas `SUM` persistem e avaliam corretamente.
- **Dado** uma planilha com 2.000 linhas, **Quando** a facade ordena, busca e
  copia linhas por critério para outra aba, **Então** a ordenação numérica, a
  contagem de linhas encontradas e a aba de destino ficam consistentes.
- Esses testes usam `assertTimeout` generoso (30s por cenário) como guarda
  contra regressão grosseira de carga, não como benchmark rígido.

## Regressões internas que protegem a facade

- Linhas de total/resumo em cálculos legados devem preservar células já
  existentes fora da área calculada. Helpers internos devem reutilizar a linha
  quando ela já existe, em vez de chamar `Sheet.createRow(indice)` às cegas.
- Manipulação de colunas deve copiar/colar tipos reais do POI sem chamar
  `Cell.setCellType(CellType.FORMULA)`; fórmulas devem ser recriadas com
  `setCellFormula(...)`, erros com `setCellErrorValue(...)`, e células em
  branco com `setBlank()`. O recorte temporário deve preservar valor e estilo
  para texto, número, booleano, fórmula, erro e branco.
- Logs internos de modificação devem imprimir as ações conhecidas de coluna
  (deslocamento, remoção, inserção vazia e limpeza), incluir colunas
  deslocadas quando houver, ignorar tipos futuros desconhecidos e limpar a fila
  após `exibirLogs()`.
- Inserção delimitada da API fluente (`inserirDados(texto, delimitador)`,
  `inserirDados(lista, delimitador)` e `inserirDadosArquivo`) deve preservar
  campos vazios finais; `"A,B,"` precisa criar a terceira célula vazia em vez
  de descartá-la pelo comportamento padrão de `String.split(...)`. Quando a
  entrada for `List<String>` com delimitador não vazio, cada item da lista é uma
  linha delimitada. Dados nulos lançam `DadosInvalidosException`, não
  `NullPointerException`.
- Caminhos básicos de erro da API fluente devem permanecer amigáveis:
  `salvar(" ")` e `abrirPlanilha(" ")` lançam `ArquivoException`, criar aba
  duplicada lança `IllegalArgumentException`, `inserirFiltros()` antes de criar
  sheet lança `IllegalStateException`, e `inserirFiltros()` em sheet vazia é
  no-op sem gravar `<autoFilter>` no OOXML.
- Estilos internos devem respeitar o alvo selecionado (célula, linha inteira,
  intervalo ou planilha inteira). Cores em hexadecimal inválido devem lançar
  `IllegalArgumentException`, e `centralizarERedimensionarTudo` deve avaliar
  fórmulas antes de redimensionar colunas.
- Atributos combinados de fonte (nome, tamanho, itálico, sublinhado, tachado e
  cor RGB) devem sobreviver a salvar e reabrir o arquivo com um `XSSFWorkbook`
  novo; a cobertura não deve depender apenas do estado em memória.

## Escape hatch

- `avancado()` devolve `IPlanilha` (API fluente completa).
- `workbook()` devolve o `Workbook` do Apache POI.
- `estilo(pos)` devolve `EstiloCelula` para encadear vários estilos de uma vez.
