# Arquitetura — ProjetoPlanilha3

## Visão em camadas

```
┌───────────────────────────────────────────────────────────────────┐
│  Planilha (facade)        ← API amigável p/ quem não programa      │
│  escrever / somar / negrito / duplicar / ordenar / gráfico / ...   │
│   ├─ delega parte para IPlanilha (API fluente, abaixo)             │
│   └─ delega parte direto para utils/graficos/imagens (sem passar   │
│      pela API fluente — ver "Duas rotas de delegação" abaixo)      │
└───────────────┬─────────────────────────────────────────────────┘
                │
        ┌───────┴────────────────────────────┐
        │                                     │
┌───────▼─────────────────────────┐  ┌────────▼──────────────────────┐
│ IPlanilha (PlanilhaBase/Xlsx)    │  │ utils / graficos / imagens     │
│  ← API fluente, seleciona-e-age  │  │  ← utilitários "sheet-level"   │
│  ├─ selecionar() → SelecaoManager│  │  FiltroDeLinhas, OrdenadorDeLinhas│
│  ├─ converter()  → ConversaoManager│ CopiadorDeCelulas, FormatosDeCelula│
│  ├─ formula()    → FormulaBuilder│  │  FormatacaoCondicionalHelper,   │
│  ├─ aplicarEstilos() → StyleManager│ ListaSuspensaHelper, GraficoHelper,│
│  ├─ manipularPlanilha() → ...    │  │  ImagemHelper                   │
│  └─ dados → DataManipulator      │  └────────┬────────────────────────┘
└───────┬──────────────────────────┘           │
        │                                      │
        └──────────────┬───────────────────────┘
                        │ usa
        ┌───────────────▼─────────────────────────────────────┐
        │  Apache POI (XSSFWorkbook, Sheet, Row, Cell, ...)    │
        └───────────────────────────────────────────────────────┘
```

### Duas rotas de delegação (por que a facade não usa só `IPlanilha`)

A API fluente (`IPlanilha`) foi desenhada em torno da máquina de estados de
seleção (`selecionar().celula(X)` → age). Isso funciona bem para dados/estilo
de célula, mas várias features adicionadas depois (ordenar, buscar/filtrar,
formatação condicional, dropdown, gráfico, imagem) são operações de **sheet
inteira** ou **intervalo explícito**, sem fazer sentido nessa máquina de
estados. Por isso a facade `Planilha` tem uma segunda rota: para essas
features, ela chama um utilitário estático em `utils/`, `graficos/` ou
`imagens/` **diretamente**, passando `XSSFSheet` + índices já convertidos via
`PosicaoConverter`. Isso mantém `IPlanilha` coeso (não vira um "todo-poderoso")
e mantém a facade fina (a lógica pesada mora no utilitário, testável isolado).

Regra prática: **nem toda feature da facade passa por `IPlanilha`** — isso é
intencional, não uma inconsistência a "corrigir".

## A máquina de estados de seleção

A API fluente funciona em dois passos: **selecionar** e depois **agir**.

- `PositionManager` guarda a seleção atual: célula única (`posicaoDefinida`),
  intervalo (`intervaloDefinida`) ou toda a planilha (`todaPlanilhaDefinida`).
- `DataManipulator` guarda a célula selecionada (`linha/colunaSelecionadaAtual`)
  e o "último índice inserido" (`ultimoIndice...Inserido`).
- Ao chamar `selecionar().celula(X)`, `DataManipulator.naCelula` **reseta** o
  último índice inserido — assim `aplicarEstilos()` mira a célula selecionada e
  não uma linha antiga (regressão coberta por `EstiloSelecaoTest`).

### Como `aplicarEstilos()` decide o alvo (`StyleManager`)

Ordem de prioridade:
1. `todaPlanilha` → planilha inteira;
2. `intervalo` → o intervalo;
3. última inserção (`ultimoIndice != -1`) → a **linha** recém-inserida;
4. célula selecionada (`posicaoDefinida`) → aquela **célula**;
5. nada → no-op.

Ao final, `resetarPosicao()` limpa o estado para não vazar para a próxima ação.

## Conversão de posições (`PosicaoConverter`)

- `"A1"` → `[coluna=0, linha=0]` (ambos 0-based).
- Valida limites do Excel: linhas 1..1.048.576, colunas até `XFD` (16.384).
- Erros viram `PosicaoInvalidaException` (unchecked) com mensagem amigável.

## Inserção de dados e coerção de tipo (`InsersorDeDados`)

`definirValorCelula` decide número x texto:
- Texto quando: não é número "limpo" (regex `NUMERO_ESTRITO`), tem zero à
  esquerda (`007`) ou é inteiro com 16+ dígitos (perde precisão como `double`).
- Caso contrário, vira número.

A facade adiciona `escreverTexto(...)`, que **sempre** grava texto (bypass da
coerção), ideal para CEP/CPF/telefone.

## Estilos (`EstiloCelula` + `estilos/estilos/`)

`EstiloCelula` é a fachada; delega para helpers especializados
(`BoldStyle`, `Fontes`, `BackGroundColor`, `AlinhamentoStyle`, `CenterStyle`,
`BorderStyleHelper`). Um `styleCache` (Map) evita recriar `CellStyle`
idênticos — Excel tem limite de estilos por workbook.

Convenção dos helpers para o alvo:
- `isRange == true` → itera o intervalo `[start..end]`.
- `rowIndex != -1, columnIndex == -1` → a linha inteira.
- `rowIndex != -1, columnIndex != -1` → a célula específica.

Coberta por testes diretos em `EstiloCelulaTest` (não só via integração).

**Pegadinha corrigida**: `BorderStyleHelper.verificarBordaEspessa` checa se a
célula **já tem** uma borda `THICK` (para não rebaixá-la ao aplicar bordas
finas). Antes desta sessão a checagem estava invertida (`NONE` em vez de
`THICK`), então `bordas(intervalo)` nunca aplicava borda no caso comum (célula
sem nenhuma borda). Se mexer em `aplicarTodasAsBordas`, rode
`EstiloCelulaTest.deveAplicarBordasFinasEmCelulasSemBorda` primeiro.

## Ordenação (`utils/OrdenadorDeLinhas`)

Captura cada linha do intervalo (célula a célula: tipo, valor e `CellStyle`),
ordena a lista capturada por uma chave (número antes de texto; alfabético
ignorando caixa; vazio por último) e reescreve as linhas na nova ordem — a
formatação "viaja" junto com o conteúdo porque é capturada e reaplicada, não
porque a célula em si se move.

## Formatação condicional (`utils/FormatacaoCondicionalHelper`)

Usa `XSSFSheet.getSheetConditionalFormatting()` +
`XSSFConditionalFormattingRule` + `XSSFPatternFormatting`. Para cor de fundo em
XSSF (diferente do quirk clássico do HSSF), o caminho certo é
`setFillForegroundColor(Color)` + `setFillPattern(SOLID_FOREGROUND)` — **não**
`setFillBackgroundColor` (isso é para HSSF). Escala de cores usa
`createConditionalFormattingColorScaleRule()` + `ColorScaleFormatting` com 3
`Color` e 3 `ConditionalFormattingThreshold` (MIN/PERCENTILE 50/MAX).

## Lista suspensa (`utils/ListaSuspensaHelper`)

Usa `XSSFDataValidationHelper` (`createExplicitListConstraint` ou
`createFormulaListConstraint`) + `sheet.addValidationData(...)`.

**Pegadinha do POI (não é bug nosso)**: em XSSF,
`validation.setSuppressDropDownArrow(true)` é o que **exibe** a seta do menu —
o nome sugere o oposto. Não trocar para `false`.

O limite de 255 caracteres somados nas opções (Excel) é validado na facade
antes de chegar ao POI (`Planilha.validarOpcoesDaLista`), com mensagem
sugerindo `listaSuspensaDoIntervalo` como alternativa sem esse limite.

## Gráficos (`graficos/GraficoHelper`)

Usa a API `XDDFChart` (não a antiga `HSSFChart`/JFreeChart). Fluxo:
`sheet.createDrawingPatriarch()` → `drawing.createAnchor(...)` →
`drawing.createChart(anchor)` → `chart.createCategoryAxis`/`createValueAxis`
(pizza usa `null, null` — não tem eixos) → `XDDFDataSourcesFactory.fromString/
NumericCellRange` → `chart.createData(ChartTypes.X, ...)` →
`data.addSeries(...)` → `chart.plot(data)`. Barras usam `BarDirection.COL`
(verticais). Verificado com round-trip real (salvar em disco + reabrir).

## Imagens (`imagens/ImagemHelper`)

Usa `workbook.addPicture(bytes, tipo)` + `drawing.createPicture(anchor, idx)`.

**Pegadinha do POI (verificada empiricamente, não documentada)**:
`XSSFPicture.resize(double escala)` escala o tamanho **atual** da âncora, não
o tamanho natural da imagem. Numa âncora recém-criada (tamanho zero),
`resize(escala)` sozinho resulta em um retângulo zerado. É preciso chamar
`resize()` (sem argumento, fixa o tamanho natural) **antes** de
`resize(escala)`. Ver `ImagemHelper.inserir(..., escala)`.

## Exceções (todas unchecked)

```
PlanilhaException (RuntimeException)
 ├─ FormulaException         (getFormulaTentada)
 ├─ PosicaoInvalidaException (getPosicaoInvalida)
 ├─ DadosInvalidosException  (getDadoInvalido)
 └─ ArquivoException         (getCaminhoArquivo; encapsula IOException)
```

Motivo de serem unchecked: manter a API fluente limpa, sem `throws` em toda
assinatura. O usuário pode capturá-las opcionalmente.

## Proteção (`utils/ProtecaoHelper`)

Toda célula do Excel nasce com `CellStyle.getLocked() == true` — a trava só
tem efeito quando a planilha é protegida (`sheet.protectSheet(senha)`).
`ProtecaoHelper.desbloquearIntervalo` clona o `CellStyle` de cada célula antes
de destravar. **Nunca mutar o estilo original diretamente**: células sem
formatação própria compartilham o mesmo `CellStyle` padrão do workbook (índice
0) — mutar esse estilo destravaria (ou alteraria) **todas** as células que o
usam, não só as do intervalo. Mesmo cuidado de sempre `workbook.createCellStyle()`
+ `cloneStyleFrom(...)` já seguido pelos helpers de `estilos/estilos/`.

## Logging

`LoggerUtil` configura `java.util.logging` com um `ColorFormatter` custom
escrevendo em `System.out`. O aviso "Log4j2 could not find a logging
implementation" vem de dependência transitiva do POI e é inofensivo.
