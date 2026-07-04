# AGENTS.md — ProjetoPlanilha3

Guia para agentes de IA (Claude Code, Codex e outros) que trabalham **dentro**
de `ProjetoPlanilha3`. Leia por completo antes de editar. Este arquivo é a
fonte de verdade sobre o estado atual; o `docs/historico/CONTEXTO_PROJETO.md` é
histórico e pode estar desatualizado.

## 0. HANDOFF ENTRE AGENTES (leia isto primeiro)

> **Se você é um agente assumindo esta sessão pela primeira vez** (ex.: Codex
> retomando depois do Claude esgotar a sessão), leia esta seção inteira antes
> de tocar em qualquer arquivo. Ela é **atualizada a cada lote de trabalho** —
> confie nela mais do que na sua intuição sobre "o que deve faltar".

**Última atualização:** ver `git log -1` (esta seção é editada a cada commit
relevante; o commit mais recente sempre reflete o estado real).

**Como confirmar o estado exato agora:**
```bash
cd ProjetoPlanilha3
git log --oneline -15        # histórico recente = trilha de auditoria dos lotes
mvn clean test                # deve terminar em BUILD SUCCESS, veja "Tests run: N"
```
Se `git log` ou o número de testes divergirem do que está escrito abaixo, **o
git/código manda** — esta seção pode estar um lote atrasada em relação ao commit
mais recente se o agente anterior foi interrompido no meio (ex.: sem tokens).

### O que já está pronto (não re-implementar)

Facade `com.abnote.planilhas.Planilha` já cobre: criar/abrir/salvar; escrever
(célula/linha/coluna/tabela/append/texto-forçado/data/data-hora); fórmulas
(soma/média/contar/mínimo/máximo/se-então/multiplicar/subtrair/dividir/fórmula
livre/preencherColuna/PROCV); ordenar linhas; buscar/copiar/mover/remover linhas
por valor; formatos (moeda/contábil/número/texto/data/porcentagem); colunas e
linhas (mover/remover/limpar/inserir/duplicar); estilos (negrito, cores, bordas,
mesclar, largura/altura, congelar N, filtros, autoajuste). Veja a seção 7 para
a lista completa e `docs/specs/facade-planilha.spec.md` para o contrato exato.

### Sessão em andamento (autônoma, iniciada por pedido do usuário em 2026-07-04)

O usuário pediu para continuar **de forma autônoma**, com atenção especial a:
testes unitários para tudo, coesão de classes/interfaces (cada método no lugar
certo), documentar tudo neste arquivo para o Codex retomar sem fricção, e manter
o foco em "chamadas simples e intuitivas" para quem não programa. Ordem de
trabalho desta sessão (lotes E-I), **marque aqui o que já foi feito**:

- [x] **Refatoração de coesão** — `Planilha.java` (1154→1093 linhas) teve a
  cópia de célula e o cache de estilos de formato extraídos para
  `utils/CopiadorDeCelulas` e `utils/FormatosDeCelula` (cada um com testes
  diretos). Sem mudança de comportamento (99 testes, mesma cobertura +7 novos).
- [x] **Lote E** — Formatação condicional: `realcarSeMaiorQue`/`MenorQue`/
  `Entre`/`Igual` + `escalaDeCores` (semáforo), via novo
  `utils/FormatacaoCondicionalHelper`. Testes validam a estrutura POI real
  (operador, fórmula, cor, thresholds), não só "não lança exceção".
- [ ] **Lote F** — Listas suspensas / dropdown (validação de dados).
- [ ] **Lote G** — Gráficos (barras/pizza/linha) via `XDDFChart`.
- [ ] **Lote H** — Inserir imagem/logo na planilha.
- [ ] **Lote I** — Cobertura de testes diretos para `EstiloCelula` e helpers de
  estilo (`estilos/estilos/*`), hoje só testados indiretamente.
- [ ] **Revisão final** — atualizar roadmap/README/specs com tudo entregue e o
  que ainda falta para "o construtor de planilhas perfeito".

**APIs do Apache POI já confirmadas via `javap` nesta sessão** (não precisa
reconferir, os nomes/assinaturas abaixo estão corretos para POI 5.2.5):
- Formatação condicional: `sheet.getSheetConditionalFormatting()` →
  `createConditionalFormattingRule(byte operador, String formula)` (operador via
  `org.apache.poi.ss.usermodel.ComparisonOperator.GT/LT/EQUAL/BETWEEN`) →
  `rule.createPatternFormatting()` → cast para preencher cor: em XSSF use
  `setFillForegroundColor(org.apache.poi.ss.usermodel.Color)` (aceita
  `XSSFColor`) + `setFillPattern(PatternFormatting.SOLID_FOREGROUND)` (**não**
  é o quirk antigo do HSSF de usar `FillBackgroundColor`). Escala de cores:
  `sheetCF.createConditionalFormattingColorScaleRule()` →
  `rule.getColorScaleFormatting()` → `setNumControlPoints(3)` +
  `setColors(Color[])` + `createThreshold()` com
  `RangeType.MIN/PERCENTILE/MAX`.
- Dropdown: `new XSSFDataValidationHelper(sheet)` (ou
  `sheet.getDataValidationHelper()`) → `createExplicitListConstraint(String[])`
  ou `createFormulaListConstraint(String)` → `helper.createValidation(constraint,
  new CellRangeAddressList(r1,r2,c1,c2))` → **`validation.setSuppressDropDownArrow(true)`
  é o que EXIBE a seta em XSSF** (nome contraintuitivo, é o comportamento real
  documentado do POI — não trocar para `false`) → `sheet.addValidationData(validation)`.
- Gráficos: `sheet.createDrawingPatriarch()` (retorna `XSSFDrawing`) →
  `drawing.createAnchor(dx1,dy1,dx2,dy2,col1,row1,col2,row2)` →
  `drawing.createChart(anchor)` (retorna `XSSFChart extends XDDFChart`) →
  `chart.createCategoryAxis(AxisPosition.BOTTOM)` /
  `chart.createValueAxis(AxisPosition.LEFT)` →
  `XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress)` /
  `.fromNumericCellRange(...)` → `chart.createData(ChartTypes.BAR/PIE/LINE,
  catAxis, valAxis)` (para **PIE**, passar `null, null` nos eixos — pizza não
  usa eixos) → `data.addSeries(categorias, valores)` → `chart.plot(data)`. Para
  barras, cast para `XDDFBarChartData` e `setBarDirection(BarDirection.COL)`
  para barras verticais (mais intuitivo para leigos que barras horizontais).
- Imagens: `workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG/JPEG)` →
  `drawing.createPicture(anchor, indice)` → `picture.resize()` (tamanho natural)
  ou `picture.resize(escala)`.
- `CellRangeAddress.valueOf("A1:B2")` aceita célula única também (`"A1"`).

### Se você (próximo agente) ficar sem tokens no meio de um lote

1. **Não deixe testes vermelhos parados.** Se um lote está pela metade e vai
   faltar contexto, prefira reverter a mudança incompleta (`git checkout --
   <arquivo>`) a deixar `mvn clean test` quebrado — o próximo agente precisa
   herdar um build verde.
2. Se já commitou algo parcial, deixe uma linha clara aqui em "Sessão em
   andamento" dizendo exatamente o que falta naquele lote específico.
3. Sempre rode `mvn clean test` e confirme BUILD SUCCESS antes de finalizar seu
   turno, mesmo que não tenha terminado todos os lotes planejados.

## 1. O que é

Biblioteca Java (wrapper sobre Apache POI) para criar planilhas `.xlsx` com uma
API simples. Público-alvo: **pessoas que quase não programam**. A prioridade é
que as chamadas sejam amigáveis e em português.

Duas camadas de API:

1. **Facade `com.abnote.planilhas.Planilha`** — a porta de entrada recomendada.
   Verbos simples (`escrever`, `somar`, `duplicar`, `mover`, `negrito`,
   `salvar`) que devolvem `Planilha` para encadear. **Toda feature nova voltada
   ao usuário final deve aparecer aqui**, mesmo que a lógica viva nas camadas de
   baixo.
2. **API fluente** (`IPlanilha` → `PlanilhaXlsx`/`PlanilhaBase`) — poderosa e
   verbosa, baseada em uma máquina de estados de seleção. A facade é construída
   por cima dela.

## 2. Stack

| Item | Valor |
|---|---|
| Linguagem | Java 8 (`maven.compiler.source/target = 1.8`) — **não subir sem pedir** |
| Build | Maven (`mvn clean test`) |
| Dependência | Apache POI 5.2.5 |
| Testes | JUnit 5.10.1 (+ Mockito disponível, pouco usado) |
| Estado dos testes | **66 testes, todos verdes** |

Não é Spring. **Não** introduzir Spring, Lombok, Jakarta Validation nem
dependências novas sem confirmar com o usuário.

## 3. Mapa de pacotes (`src/main/java/com/abnote/planilhas/`)

```
Planilha.java          → FACADE amigável (ponto de entrada recomendado)
calculos/              → Calculos, Conversores (soma, formatos numéricos/moeda/texto)
estilos/               → EstiloCelula (fachada de estilos) + estilos/ (helpers) + util/
examples/              → ExemploFacade (facade), ExemploFormulas (API fluente)
exceptions/            → PlanilhaException + FormulaException, PosicaoInvalidaException,
                         DadosInvalidosException, ArquivoException (todas UNCHECKED)
formulas/              → FormulaBuilder (11 fórmulas)
impl/                  → PlanilhaBase, PlanilhaXlsx, DataManipulator, StyleManager,
                         ConversaoManager, SelecaoManager
interfaces/            → contratos públicos (IPlanilha, ISelecao, IConversao, ...)
utils/                 → PosicaoConverter, PositionManager, InsersorDeDados,
                         ManipuladorPlanilha(Helper), LoggerUtil,
                         FiltroDeLinhas, OrdenadorDeLinhas, CopiadorDeCelulas,
                         FormatosDeCelula, ...
```

Detalhes em [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).
Comportamento esperado (specs) em [`docs/specs/`](docs/specs/).

## 4. Regras de comportamento que NÃO podem regredir

Há testes cobrindo cada item — rode `mvn clean test` após qualquer mudança.

1. **Coerção numérica segura** (`InsersorDeDados.deveManterComoTexto`): valores
   com zero à esquerda (`007`), não-números (`NaN`) e inteiros com 16+ dígitos
   (CPF/CNPJ/cartão) são preservados como **texto**. Números limpos viram número.
2. **Direcionamento de `aplicarEstilos()`**: após `selecionar().celula(X)` sem
   inserir, estiliza **aquela** célula; após inserir uma lista/linha, estiliza a
   **linha** inteira; `selecionar().intervalo(...)` estiliza o intervalo. A
   seleção nova reseta o "último índice inserido" (ver `DataManipulator.naCelula`).
3. **Limites do Excel**: `PosicaoConverter` valida 1..1.048.576 linhas e até a
   coluna `XFD` (16.384), lançando `PosicaoInvalidaException`.
4. **Formatos**: `emContabil`/`emMoeda` produzem formato com `R$`; `emTexto`
   preserva a representação (sem `.0` para inteiros).

## 5. Convenções de código (obrigatórias)

- **Idioma**: nomes de métodos/variáveis e comentários em **português BR**.
- **JavaDoc**: em métodos **públicos** — especialmente na facade `Planilha`,
  que é a API que o usuário final lê. Inclua exemplos curtos quando ajudar.
- **Exceções unchecked**: seguir a hierarquia de `exceptions/`. Nunca criar
  `extends Exception` (checked). `salvar()` **não** declara `throws IOException`.
- **Tamanho**: ≤ 25 linhas por método; complexidade ciclomática ≤ 15.
- **Imutabilidade**: `final` em parâmetros/campos quando possível; evitar `null`.
- **Injeção por construtor** nas classes internas (padrão já usado).
- **Nomes específicos**: proibido `total`, `soma`, `lista`, `resultado`, `valor`,
  `temp`, `aux` soltos — use nomes que digam o que são no escopo.
- **Estilo**: Google Java Style.

### Edição cirúrgica

1. Não renomear/refatorar código fora da tarefa atual.
2. Não adicionar imports/dependências desnecessários.
3. Mudança de assinatura pública → avisar claramente antes.
4. Toda I/O com tratamento de exceção (try-with-resources onde couber).
5. Em dúvida sobre requisito, **pergunte antes de implementar**.

## 6. Fluxo de trabalho esperado

1. Entenda o pedido e localize a camada certa (facade vs. interna).
2. Se for feature de usuário, exponha na facade `Planilha`.
3. Escreva/atualize **testes** (JUnit 5) — o projeto preza cobertura.
4. `mvn clean test` deve ficar **verde** antes de concluir.
5. Atualize `README.md` e as specs quando o comportamento público mudar.

## 7. Roadmap aberto (candidatos, não obrigatórios)

- Aumentar cobertura de `EstiloCelula`/helpers de estilo.
- Validação de dados (listas suspensas/dropdowns) e gráficos, se houver demanda.

Já entregue:
- Aritmética por fórmula (`FormulaBuilder.personalizada` + facade
  `multiplicar`/`subtrair`/`dividir`/`formula`/`preencherColuna`).
- PROCV amigável: `procurarValor` / `procurarValorNaAba` (VLOOKUP).
- Busca/filtro de linhas na facade (`buscarLinhas`, `contarLinhasOnde`,
  `copiarLinhasParaAba`, `moverLinhasParaAba`, `removerLinhasOnde`) via
  `utils/FiltroDeLinhas`. A interface órfã `IBuscaDados` foi **removida**
  (design POI-leaking incompatível com o objetivo amigável).
- Ordenar linhas (`ordenarPorCrescente`/`Decrescente`/`ordenarPor`) via
  `utils/OrdenadorDeLinhas`.
- Datas (`escreverData`/`escreverDataHora`/`formatarComoData`), porcentagem e
  dimensões (`larguraColuna`/`alturaLinha`/`congelar`).
- Abrir/editar arquivos existentes: `Planilha.abrir(caminho)` +
  `IPlanilhaBasica.abrirPlanilha` (via `WorkbookFactory`).
- Formatação condicional: `realcarSeMaiorQue`/`MenorQue`/`Entre`/`Igual` +
  `escalaDeCores`, via `utils/FormatacaoCondicionalHelper`.
