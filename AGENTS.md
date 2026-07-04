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
mesclar, largura/altura, congelar N, filtros, autoajuste); **formatação
condicional** (realçar/escala de cores); **lista suspensa** (dropdown, fixa ou
de intervalo); **gráficos** (barras/pizza/linha); **imagens/logo**. Veja a
seção 7 para o que ainda falta e `docs/specs/facade-planilha.spec.md` para o
contrato exato de cada comando.

### Sessão autônoma de 2026-07-04 (lotes E-I) — CONCLUÍDA

O usuário pediu para continuar **de forma autônoma**, com atenção especial a:
testes unitários para tudo, coesão de classes/interfaces (cada método no lugar
certo), documentar tudo neste arquivo para o Codex retomar sem fricção, e manter
o foco em "chamadas simples e intuitivas" para quem não programa. Os lotes
E-I abaixo foram todos concluídos, testados e commitados/enviados ao GitHub
(repo `isaacbraym/PlanilhasConstructorV2`, branch `main`). **Se você está
retomando depois disso**: a lista abaixo é só histórico desta sessão — o que
fazer a seguir está na seção 7 (roadmap priorizado). Se encontrar um `[ ]`
(não marcado) abaixo, é sinal de que a sessão foi interrompida no meio desse
item específico — confira `git log` e `mvn clean test` antes de continuar.

- [x] **Refatoração de coesão** — `Planilha.java` (1154→1093 linhas) teve a
  cópia de célula e o cache de estilos de formato extraídos para
  `utils/CopiadorDeCelulas` e `utils/FormatosDeCelula` (cada um com testes
  diretos). Sem mudança de comportamento (99 testes, mesma cobertura +7 novos).
- [x] **Lote E** — Formatação condicional: `realcarSeMaiorQue`/`MenorQue`/
  `Entre`/`Igual` + `escalaDeCores` (semáforo), via novo
  `utils/FormatacaoCondicionalHelper`. Testes validam a estrutura POI real
  (operador, fórmula, cor, thresholds), não só "não lança exceção".
- [x] **Lote F** — Listas suspensas: `listaSuspensa` (opções fixas, valida
  limite real do Excel de 255 caracteres somados) e `listaSuspensaDoIntervalo`
  (opções vindas de um intervalo, sem esse limite), via novo
  `utils/ListaSuspensaHelper`.
- [x] **Lote G** — Gráficos: `graficoDeBarras`/`graficoDePizza`/`graficoDeLinha`
  via novo `graficos/GraficoHelper` (pacote novo, 1 classe, mesmo padrão de
  `formulas/`). Verificado com round-trip real (salvar em disco + reabrir com
  workbook novo) além dos testes de estrutura em memória.
- [x] **Lote H** — Inserir imagem/logo: `inserirImagem(celula, caminho)` e
  overload com `escala`, via novo `imagens/ImagemHelper`. **Bug real
  encontrado e corrigido pelos testes**: `XSSFPicture.resize(double escala)`
  escala o tamanho ATUAL da âncora, não o tamanho natural da imagem — numa
  âncora recém-criada (tamanho zero), `resize(3.0)` sozinho dava um retângulo
  zerado. Corrigido chamando `resize()` (fixa o tamanho natural) e **depois**
  `resize(escala)`. Ver comentário em `ImagemHelper.inserir(...)`.
- [x] **Lote I** — `EstiloCelulaTest` (19 testes): alvo único/intervalo/
  planilha inteira, fontes, alinhamento, cores, bordas, centralização.
  **Bug real encontrado e corrigido**: `BorderStyleHelper.verificarBordaEspessa`
  tinha a lógica invertida — `bordas(intervalo)`/`aplicarTodasAsBordas()` nunca
  aplicava borda fina em células sem borda prévia (o caso comum). Corrigido
  para pular apenas células que já têm borda **espessa** (preservando-a). Ver
  regra 5 na seção 4 acima.
- [x] **Revisão final** — `docs/ARCHITECTURE.md` ganhou seções para cada
  feature nova (ordenação, formatação condicional, dropdown, gráficos,
  imagens) + diagrama atualizado explicando a "segunda rota" de delegação da
  facade (direto para `utils/`/`graficos/`/`imagens`, sem passar por
  `IPlanilha`). Roadmap da seção 7 reescrito com prioridades claras para quem
  continuar (impressão e proteção de planilha no topo).

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
  `drawing.createPicture(anchor, indice)` → `picture.resize()` (tamanho natural).
  **Cuidado**: `picture.resize(escala)` sozinho **não funciona** numa âncora
  recém-criada — ele escala o tamanho ATUAL da âncora (que começa zerado), não
  o tamanho natural da imagem. Sempre chame `resize()` **antes** de
  `resize(escala)` (confirmado empiricamente com um debug script, não está
  documentado no javadoc do POI). Bug real pego pelos testes desta sessão.
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
| Estado dos testes | **135 testes, todos verdes** (ver seção 0 para o número mais atual) |

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
graficos/              → GraficoHelper (barras/pizza/linha via XDDFChart)
imagens/               → ImagemHelper (inserir PNG/JPEG ancorado numa célula)
impl/                  → PlanilhaBase, PlanilhaXlsx, DataManipulator, StyleManager,
                         ConversaoManager, SelecaoManager
interfaces/            → contratos públicos (IPlanilha, ISelecao, IConversao, ...)
utils/                 → PosicaoConverter, PositionManager, InsersorDeDados,
                         ManipuladorPlanilha(Helper), LoggerUtil,
                         FiltroDeLinhas, OrdenadorDeLinhas, CopiadorDeCelulas,
                         FormatosDeCelula, FormatacaoCondicionalHelper,
                         ListaSuspensaHelper, ...
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
5. **`bordas(intervalo)` aplica borda fina mesmo sem borda prévia**
   (`BorderStyleHelper.verificarBordaEspessa`): só pula células que **já têm**
   uma borda **espessa** (para não rebaixá-la); antes desta sessão a checagem
   estava invertida (pulava células **sem nenhuma** borda, ou seja, o caso
   comum) e `bordas(...)` silenciosamente não fazia nada na maioria dos casos.

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

## 7. Roadmap — o que falta para "o construtor de planilhas perfeito"

### Já entregue (não re-implementar; ver README.md para a lista completa de comandos)

Criar/abrir/salvar planilhas; escrever (célula/linha/coluna/tabela/append/
texto-forçado/data/data-hora); fórmulas (agregação/aritmética/fórmula livre/
preencherColuna/PROCV); ordenar linhas; buscar/copiar/mover/remover linhas por
valor; formatos (moeda/contábil/número/texto/data/porcentagem); colunas e
linhas (mover/remover/limpar/inserir/duplicar); estilos (fonte, cor, borda,
mesclar, largura/altura, congelar N, filtros, autoajuste); formatação
condicional (realçar/escala de cores); lista suspensa (opções fixas ou de
intervalo); gráficos (barras/pizza/linha); imagens/logo.

Dois bugs reais foram encontrados e corrigidos por testes nesta sessão:
direcionamento de `aplicarEstilos()` em célula única, e
`BorderStyleHelper.verificarBordaEspessa` (bordas finas nunca eram aplicadas em
células sem borda prévia). Ver seção 4 para os detalhes que não podem regredir.

### Prioridade alta (maior valor prático, ainda não coberto)

1. **Configuração de impressão** — orientação retrato/paisagem, área de
   impressão, ajustar para caber em N páginas, margens, cabeçalho/rodapé de
   impressão. API POI: `sheet.getPrintSetup()` (`setLandscape`, `setFitToPage`,
   `setFitWidth`/`setFitHeight`), `sheet.setPrintArea(...)` (ou
   `workbook.setPrintArea(sheetIndex, ...)`), `sheet.getMargin`/`setMargin`.
   Zero cobertura hoje — quem monta uma planilha para **imprimir** (relatório,
   nota, orçamento) não tem nenhum comando amigável para isso.
2. **Proteção de planilha/células** — travar células de fórmula/estrutura e
   deixar só as de entrada editáveis (par natural de `listaSuspensa`, para
   formulários). API POI: `CellStyle.setLocked(boolean)` (todas as células são
   `locked=true` por padrão, mesmo sem proteção ativa) +
   `sheet.protectSheet(senha)` (senha pode ser vazia/null para só travar sem
   pedir senha — checar comportamento exato do POI antes de assumir).
3. **Validação numérica/de data na lista suspensa** — hoje
   `utils/ListaSuspensaHelper` só cobre `createExplicitListConstraint`/
   `createFormulaListConstraint`; `XSSFDataValidationHelper` também tem
   `createIntegerConstraint`/`createDecimalConstraint`/`createDateConstraint`
   prontos para uso (só faltou expor). Ex.: `validarNumeroEntre(intervalo, min,
   max)`, `validarDataEntre(intervalo, min, max)`.

### Prioridade média

4. **Nomes de intervalo (named ranges)** — `definirNome("Produtos", "A2:A100")`
   e então usar `"Produtos"` em `procurarValor`/`somar`/`listaSuspensaDoIntervalo`
   em vez do range cru. Mais legível para quem não programa. API POI:
   `workbook.createName()` + `Name.setNameName`/`setRefersToFormula`.
5. **Mais tipos de formatação condicional** — barras de dados (`DataBarFormatting`)
   e conjuntos de ícones (`IconMultiStateFormatting`), complementando
   `escalaDeCores`. `FormatacaoCondicionalHelper` já tem a estrutura pronta
   para crescer nessa direção.
6. **Comentários em células** (notas) — `sheet.createDrawingPatriarch()` +
   `Drawing.createCellComment(anchor)`; útil para explicar uma fórmula ou
   instrução num formulário.
7. **`adicionarTotais()` de alto nível** — método que soma automaticamente
   todas as colunas numéricas de uma tabela e insere uma linha "Total"
   formatada, sem o usuário precisar chamar `somar()` coluna por coluna.

### Prioridade baixa / nice-to-have

- Duplicar planilha inteira para outro arquivo (hoje só duplica aba dentro do
  mesmo workbook).
- Testes de performance/carga com planilhas de milhares de linhas (confiança,
  não funcionalidade nova).
- Exportar para CSV (fora do escopo original — esta lib foca em `.xlsx`;
  avaliar com o usuário antes de assumir que é desejado).

### Convenção para continuar

Cada item novo deve seguir o padrão já estabelecido: lógica pesada num
utilitário estático (`utils/`, ou pacote novo de 1 classe como `graficos/`/
`imagens/` se for um subsistema distinto), facade só delega e documenta,
teste dedicado validando a **estrutura real do POI** (não só "não lança
exceção" — ver `docs/skills-codex/README.md` ou a skill `planilha-testes`),
e atualização de README + spec + este arquivo (seção 0 e 7) a cada lote.
