# Arquitetura — ProjetoPlanilha3

## Visão em camadas

```
┌─────────────────────────────────────────────────────────────┐
│  Planilha (facade)          ← API amigável p/ quem não programa│
│  escrever / somar / negrito / duplicar / salvar ...           │
└───────────────┬─────────────────────────────────────────────┘
                │ delega
┌───────────────▼─────────────────────────────────────────────┐
│  IPlanilha (PlanilhaBase → PlanilhaXlsx)   ← API fluente      │
│   ├─ selecionar()  → SelecaoManager                          │
│   ├─ converter()   → ConversaoManager → Conversores          │
│   ├─ formula()     → FormulaBuilder                          │
│   ├─ aplicarEstilos() → StyleManager → EstiloCelula          │
│   ├─ manipularPlanilha() → ManipuladorPlanilha(Helper)       │
│   └─ dados          → DataManipulator → InsersorDeDados      │
└───────────────┬─────────────────────────────────────────────┘
                │ usa
┌───────────────▼─────────────────────────────────────────────┐
│  Apache POI (XSSFWorkbook, Sheet, Row, Cell, CellStyle)      │
└─────────────────────────────────────────────────────────────┘
```

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

## Logging

`LoggerUtil` configura `java.util.logging` com um `ColorFormatter` custom
escrevendo em `System.out`. O aviso "Log4j2 could not find a logging
implementation" vem de dependência transitiva do POI e é inofensivo.
