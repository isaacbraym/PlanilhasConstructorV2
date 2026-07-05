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
de intervalo na mesma/outra aba, e validação numérica/de data); **nomes de intervalo**
(`definirNome`, com nome inválido virando `DadosInvalidosException` amigável);
**`adicionarTotais()`** automático; **gráficos** (barras/pizza/linha, com
suporte a categorias/valores em abas diferentes da do gráfico); **imagens/
logo**; **comentários** em células; **impressão** (orientação/área/ajustar em
N páginas/cabeçalho e rodapé com marcadores amigáveis); **proteção** de
planilha e desbloqueio de células (formulários); **leitura** (`ler`/
`lerTexto`/`lerNumero`/`lerData`/`lerTabela`/`contarLinhasPreenchidas` — o
"inverso" de escrever, sem precisar de `workbook()`); **ocultar/exibir**
(linha/coluna/aba); **desmesclar**; **cor da aba**; **formato numérico
personalizado** (escape hatch). Veja a seção 7 para o que ainda falta e
`docs/specs/facade-planilha.spec.md` para o contrato exato de cada comando.

### Sessão de "fechar vazamentos de POI" (2026-07-04, após reanálise crítica) — CONCLUÍDA

O usuário pediu uma reanálise crítica sob a lente "o objetivo é o usuário
**nunca** precisar tocar em Apache POI para criações avançadas". Achado
principal: a facade tinha **zero API de leitura** — para ler qualquer valor
de volta era preciso `planilha.workbook().getSheetAt(0).getRow(...)...` (POI
puro). Lista de gaps identificados, **marque aqui o que já foi feito**:

- [x] **API de leitura**: `ler`/`lerTexto`/`lerNumero`/`lerData`/`lerTabela`/
  `contarLinhasPreenchidas`, via novos `utils/LeitorDeCelulas` (célula única,
  reaproveitado por tudo) e `utils/LeitorDeTabela` (reaproveita a detecção de
  largura/altura de `TotalizadorDeTabela`, que teve 2 métodos promovidos de
  `private` para pacote-privado para isso). `lerTexto` usa `DataFormatter`
  para devolver o valor **como aparece no Excel** (moeda, data etc.), `ler`
  devolve o tipo Java natural. Nenhum método de leitura cria célula (ao
  contrário dos de escrita). Total: 176 testes verdes.
- [x] **Ocultar/exibir linha, coluna e aba**: `ocultarLinha`/`exibirLinha`,
  `ocultarColuna`/`exibirColuna`, `ocultarAba`/`exibirAba`. **Achado de
  segurança**: testei empiricamente e o POI permite ocultar a única aba
  visível de um workbook sem erro — produzindo um arquivo que o Excel não
  consegue abrir. `ocultarAba` agora recusa isso com `DadosInvalidosException`
  amigável (contando abas visíveis antes de ocultar). Aba inexistente lança
  `IllegalArgumentException` (mesma convenção de `selecionarSheet`, não a
  hierarquia `PlanilhaException`). Total: 181 testes verdes.
- [x] **Desmesclar células**: `desmesclar(intervalo)` — procura a região
  mesclada com o range **exatamente igual** e remove; no-op se não achar
  (não lança exceção). Total: 183 testes verdes.
- [x] **Cor da aba**: `corDaAba(CorEnum)` via `XSSFSheet.setTabColor`. Total:
  184 testes verdes.
- [x] **Formato numérico personalizado**:
  `formatarComoPersonalizado(intervalo, formatoExcel)`, via novo método
  `personalizado(String)` em `utils/FormatosDeCelula` (cache por string de
  formato, mesmo padrão de `data()`/`porcentagem()`). Sintaxe do formato não é
  validada (Excel/POI não valida nesse ponto). Total: 186 testes verdes.
- [x] **Cabeçalho/rodapé de impressão com texto**: `cabecalhoDeImpressao`/
  `rodapeDeImpressao`, cada um com sobrecarga de 1 argumento (só centro) e de
  3 (esquerda/centro/direita), via `sheet.getHeader()`/`getFooter()`. Novo
  `utils/CodigosDeImpressao.traduzir` converte marcadores amigáveis
  (`{pagina}`, `{total}`, `{data}`, `{hora}`, `{arquivo}`, `{aba}`) para os
  códigos nativos do Excel (`&P`, `&N`, `&D`, `&T`, `&F`, `&A`) — a sintaxe
  `&`-prefixada não é descobrível sem doc, então o usuário nunca escreve ela
  diretamente. Round-trip verificado (salva e reabre com XSSFWorkbook fresco).
  Total: 190 testes verdes.
- [x] **`definirNome` não vaza mais `IllegalArgumentException` do POI**:
  `setNameName` agora está em try/catch, relançando `DadosInvalidosException`
  (com a exceção original como causa) para os 5 casos empíricos confirmados —
  nome começando com dígito, com espaço, igual a uma referência de célula
  (`"A1"`, `"$B$2"`) ou vazio. Teste parametrizado (`@ValueSource`) cobre os 5
  casos em `NomeDeIntervaloFacadeTest`. Total: 195 testes verdes.
- [x] **Suporte de verdade a gráfico entre abas**: `graficoDeBarras`/
  `graficoDePizza`/`graficoDeLinha` ganharam uma sobrecarga de 6 argumentos —
  `(titulo, abaCategorias, intervaloCategorias, abaValores, intervaloValores,
  celulaSuperiorEsquerda)` — que busca categorias e valores em abas diferentes
  entre si (e diferentes da aba onde o gráfico é desenhado). Escolhida em vez
  da opção "só validar e barrar": o cenário de dashboard resumindo dados de
  outras abas é comum o bastante para merecer suporte real, não só um erro
  melhor. Internamente, `GraficoHelper` ganhou overloads aceitando duas
  `XSSFSheet` (uma por fonte de dados) — a versão de 1 sheet virou um atalho
  que chama a de 2 sheets com a mesma folha duas vezes. Aba inexistente lança
  `IllegalArgumentException` (via `xssfDaAba`, mesma convenção de
  `indiceDaAba`/`selecionarSheet`). Verificado empiricamente que
  `getDataRangeReference()` do XDDF traz o nome da aba corretamente
  (`"Produtos!$A$2:$A$4"`) e que os valores resolvem certo. Total: 197 testes
  verdes.
- [x] **Nomes definidos nas fórmulas prontas de agregação**: `somar`/`media`/
  `contar`/`minimo`/`maximo` agora aceitam nomes criados por `definirNome`
  (ex.: `definirNome("Precos", "B2:B4").somar("D1", "Precos")` gera
  `SUM(Precos)` e avalia corretamente). Investigação empírica confirmou que
  `formula("D1", "SUM(Precos)")` já funcionava e que o bloqueio real era só a
  validação estrita de `FormulaBuilder.validarRange`; ela agora aceita ranges
  de células ou nomes existentes no `Workbook`, mantendo `FormulaException`
  amigável para texto solto não definido. Total: 198 testes verdes.
- [x] **Margens de impressão**: `margensDeImpressao(superiorCm, inferiorCm,
  esquerdaCm, direitaCm)` expõe margens em centímetros (mais natural em pt-BR)
  e converte para polegadas no novo `utils/MargensDeImpressaoHelper`, que é a
  unidade nativa de `Sheet.setMargin(PageMargin, ...)`. Investigação empírica
  confirmou round-trip em OOXML (`2,54 cm` → `1.0` polegada ao reabrir).
  Teste dedicado salva/reabre com `XSSFWorkbook` fresco e valida as quatro
  margens reais, além de recusar margem negativa com `DadosInvalidosException`.
  Total: 200 testes verdes.
- [x] **Formatação condicional avançada**: `barrasDeDados(intervalo, cor)` e
  `iconesSemaforo(intervalo)` fecham o roadmap de data bars/ícones sem expor
  `DataBarFormatting`/`IconSet` ao usuário. Debug empírico confirmou que o POI
  cria barras com limiares MIN/MAX e ícones semáforo com percentuais 0/33/66,
  ambos sobrevivendo a round-trip OOXML. Testes dedicados salvam/reabrem com
  `XSSFWorkbook` fresco e validam a estrutura real (`DataBarFormatting`,
  `IconMultiStateFormatting`, cor e thresholds). Total: 202 testes verdes.
- [x] **Lista suspensa com opções em outra aba**:
  `listaSuspensaDoIntervalo(intervaloDestino, abaOpcoes, intervaloOpcoes)`
  monta a fórmula qualificada (`'Apoio'!$F$2:$F$5`) sem exigir que o usuário
  conheça sintaxe de Excel/POI. Investigação empírica confirmou round-trip
  OOXML para aba com espaço no nome. Extraído `utils/ReferenciasExcel` para
  centralizar referências absolutas/qualificadas e corrigir um bug silencioso:
  intervalos já absolutos (`$F$2:$F$5`) não podem virar `$$F$$2`. Testes
  dedicados cobrem lista em outra aba, aba inexistente, intervalo já absoluto
  e `definirNome` com intervalo absoluto. Total: 206 testes verdes.
- [x] **Testes de carga com milhares de linhas**: novo `CargaFacadeTest`
  cobre dois cenários de confiança operacional, sem virar benchmark rígido:
  (1) 3.000 linhas escritas, fórmulas `SUM`, filtros, congelamento, salvar e
  reabrir com `XSSFWorkbook` fresco; (2) 2.000 linhas ordenadas, buscadas e
  copiadas para outra aba. Cada cenário usa `assertTimeout` generoso de 30s
  para pegar regressão grosseira sem criar teste frágil. Teste focal passou em
  ~5s localmente; `mvn clean test` confirmou BUILD SUCCESS. Total: 208 testes
  verdes.
- [x] **Bug de preservação de linhas em cálculos legados**:
  `Calculos.multiplicarColunasComTexto` usava `sheet.createRow(linhaSoma)` na
  linha de total. Se essa linha já existisse por conter dados em outra coluna,
  o POI podia substituir a linha e apagar células fora do cálculo. Corrigido
  com `obterOuCriarLinha(...)` e testes em `CalculosTest` cobrindo preservação
  de células existentes tanto em soma com texto quanto em multiplicação com
  texto. `mvn clean test` confirmou BUILD SUCCESS. Total: 210 testes verdes.
- [x] **Cobertura estrutural de estilos**: `EstiloCelulaTest` ganhou cenários
  que faltavam para `CenterStyle`, `Fontes` e `BackGroundColor`: estilo em
  linha inteira (`columnIndex == -1`), rejeição de hexadecimal inválido e
  `centralizarERedimensionarTudo` com fórmula avaliada antes do `autoSize`.
  Teste focal verde; `mvn clean test` confirmou BUILD SUCCESS. Total: 213
  testes verdes.
- [x] **Round-trip de fontes combinadas**: `EstiloCelulaTest` ganhou validação
  salvando e reabrindo `.xlsx` com `XSSFWorkbook` fresco para uma fonte com
  nome, tamanho, itálico, sublinhado, tachado e cor RGB. A investigação
  empírica em `target/debug` confirmou a estrutura antes do teste permanente.
  Também foi removido código morto em `Fontes` (`fontMatchesAttributes` e
  comparadores privados não chamados), reduzindo área enganosa de cobertura.
  Total: 214 testes verdes.
- [x] **Bug em mover/colar coluna com fórmula corrigido**:
  `ManipuladorPlanilhaHelper` chamava `Cell.setCellType(CellType.FORMULA)`
  antes de `setCellFormula(...)`, operação ilegal no POI 5.2.5. Uma coluna com
  fórmula deslocada por `moverColuna` podia quebrar com
  `IllegalArgumentException`. Corrigido para usar o setter correto de cada
  tipo (fórmula, erro, branco etc.) sem forçar `setCellType`, e coberto em
  `ManipuladorPlanilhaTest` com fluxo público e recorte temporário preservando
  texto, número, booleano, fórmula, erro, branco e estilo. Total: 217 testes
  verdes.
- [x] **Cobertura de logs de modificação**: novo `LogsDeModificadoresTest`
  captura `System.out` e cobre deslocamento, remoção, inserção vazia, limpeza,
  colunas deslocadas, tipo futuro desconhecido (ignorado) e limpeza da fila
  após `exibirLogs()`. Isso fecha o ponto fraco histórico de log/auditoria no
  JaCoCo sem mudar API pública. Total: 220 testes verdes.
- [x] **Bug de campos finais vazios em dados delimitados corrigido**:
  `InsersorDeDados` usava `String.split(...)` sem limite, descartando campos
  vazios no fim (`"A,B,"` virava só duas células). Corrigido com helper que
  preserva vazios finais para delimitadores não vazios e mantém o comportamento
  antigo para delimitador vazio. `PlanilhaXlsxTest` cobre string delimitada e
  arquivo texto com célula final vazia real. Total: 222 testes verdes.
- [x] **Caminhos de erro/no-op da API fluente básica cobertos**:
  `PlanilhaXlsxTest` agora cobre `salvar(" ")` e `abrirPlanilha(" ")` com
  `ArquivoException`, criação de aba duplicada com `IllegalArgumentException`,
  `inserirFiltros()` antes de criar sheet com `IllegalStateException`, e
  `inserirFiltros()` em sheet vazia como no-op sem `<autoFilter>` no OOXML.
  Total: 227 testes verdes.
- [x] **Bug de lista delimitada na API fluente corrigido**:
  `InsersorDeDados.inserirDados(List<String>, delimitador)` ignorava o
  delimitador e gravava linhas como textos únicos (`"A,B,"`) na mesma linha.
  Debug empírico contra `target/classes` confirmou o estado real das células.
  Agora cada item da lista é tratado como uma linha delimitada, preservando
  campos vazios finais, e entradas `null` lançam `DadosInvalidosException`
  amigável. `PlanilhaXlsxTest` cobre round-trip com `XSSFWorkbook` novo.
  Total: 229 testes verdes.
- [x] **Bug de duplicação parcial de aba corrigido**:
  `Planilha.duplicarAba("NomeExistente")` chamava `Workbook.cloneSheet(...)`
  antes de validar o novo nome; quando `setSheetName(...)` falhava, o workbook
  ficava com uma aba parcial tipo `"Original (2)"`. Debug empírico confirmou a
  mutação. Agora `WorkbookUtil.validateSheetName(...)` e a checagem de nome
  existente rodam antes do clone. `PlanilhaFacadeTest` cobre nome duplicado e
  nome inválido, garantindo que a lista real de abas do POI não muda.
  Total: 231 testes verdes.
- [x] **Erros de `duplicarArquivo` padronizados**:
  `Planilha.duplicarArquivo` já prometia `ArquivoException`, mas caminhos
  `null`/vazios ou inválidos podiam escapar como `NullPointerException`/
  `InvalidPathException` antes da cópia. Agora origem e destino são validados
  antes de `Paths.get(...)`, e caminho inválido é encapsulado em
  `ArquivoException`. `DuplicarArquivoFacadeTest` cobre origem inexistente,
  origem/destino obrigatórios e caminho inválido. Total: 233 testes verdes.
- [x] **Busca por resultado de fórmula corrigida**:
  `FiltroDeLinhas.valorComoTexto` comparava células `FORMULA` pelo texto da
  fórmula (`"A2*2"`), então `buscarLinhas("B", "20")` não encontrava uma célula
  cujo resultado exibido era `20`. Debug empírico confirmou o comportamento.
  Agora a busca usa `FormulaEvaluator` para comparar pelo resultado calculado,
  com fallback para o texto da fórmula se a avaliação falhar. `BuscaFacadeTest`
  cobre busca e cópia de linhas por resultado de fórmula. Total: 234 testes
  verdes.
- [x] **Ordenação por resultado de fórmula corrigida**:
  `OrdenadorDeLinhas` também tratava células `FORMULA` pelo texto interno da
  fórmula, então `ordenarPorCrescente("C")` em uma coluna `B{}*2` preservava a
  ordem física (`B2*2`, `B3*2`, `B4*2`) em vez de usar os resultados exibidos.
  Debug empírico revelou ainda que, quando uma linha com fórmula mudava de
  posição, a fórmula podia continuar apontando para a linha antiga. Agora a
  ordenação avalia fórmulas com `FormulaEvaluator` e, ao reescrever as linhas,
  ajusta referências relativas com `FormulaShifter.createForRowCopy`.
  `OrdenarFacadeTest` cobre a estrutura real do POI e round-trip com
  `XSSFWorkbook` fresco. Total: 235 testes verdes.
- [x] **Ordenação preserva atributos da linha**:
  Debug empírico mostrou que `OrdenadorDeLinhas` movia as células, mas deixava
  atributos da linha física (altura personalizada e `zeroHeight` de linha
  oculta) presos no índice antigo. Assim, ordenar podia fazer o conteúdo subir
  enquanto a altura/ocultação ficava em outra linha. Agora a captura da linha
  inclui altura, ocultação e estilo de linha (`RowStyle`), e a reescrita limpa
  atributos residuais quando a linha capturada era vazia. `OrdenarFacadeTest`
  cobre o caso com round-trip OOXML. Total: 236 testes verdes.
- [x] **Copia de formulas ao duplicar linhas/colunas corrigida**:
  Debug empirico mostrou que `duplicarLinha`, `duplicarColuna` e
  `copiarLinhasParaAba` copiavam formulas pelo texto original, deixando
  referencias relativas presas na origem (`A2*2` continuava `A2*2` ao ir para
  a linha 5). Novo `utils/AjustadorDeFormulas` centraliza `FormulaParser`/
  `FormulaShifter` e faz a copia se comportar como o Excel: referencias
  relativas acompanham linha/coluna, partes absolutas (`$A$2`, `B$2`, `$A2`)
  ficam preservadas. `PlanilhaFacadeTest` cobre o fluxo publico com
  round-trip OOXML, `CopiadorDeCelulasTest` cobre o utilitario direto e
  `BuscaFacadeTest` valida copia de linhas filtradas para outra aba. Total:
  239 testes verdes.
- [x] **`adicionarTotais()` soma colunas de formulas numericas**:
  Debug empirico mostrou que uma coluna "Total" preenchida com formulas
  numericas (`D2=B2*C2`, `D3=B3*C3`) era ignorada por `adicionarTotais()`,
  porque `TotalizadorDeTabela` aceitava apenas `CellType.NUMERIC`. Agora o
  helper usa `FormulaEvaluator` para classificar formulas pelo resultado:
  formulas numericas recebem `SUM(...)`, formulas/textos nao numericos continuam
  sem soma. `TotalizadorFacadeTest` cobre o fluxo publico com round-trip OOXML.
  Total: 240 testes verdes.

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
- [x] **Lote J** (fora do plano original E-I, feito na sequência por ser o
  topo do roadmap) — Impressão: `orientacaoPaisagem`/`Retrato`,
  `areaDeImpressao`, `ajustarImpressaoEmPaginas`. Proteção:
  `protegerPlanilha(senha)` + `desbloquearCelulas(intervalo)` via novo
  `utils/ProtecaoHelper` (clona `CellStyle` antes de destravar — testado que
  destravar uma célula não afeta outras que compartilham o estilo padrão).
  Verificado com round-trip real (salvar + reabrir). Total: 142 testes verdes.
- [x] **Lote K** — `validarNumeroEntre`/`validarInteiroEntre`/`validarDataEntre`
  via novo `utils/ValidacaoDeEntradaHelper` (`DataValidation` do tipo
  DECIMAL/INTEGER/DATE, sem menu de opções — por isso não entrou em
  `ListaSuspensaHelper`). Seção da facade renomeada de "LISTA SUSPENSA
  (DROPDOWN)" para "VALIDAÇÃO DE DADOS (LISTA SUSPENSA / LIMITES)" para
  acomodar os dois conceitos coerentemente. Total: 146 testes verdes.
- [x] **Lote L** — `definirNome(nome, intervalo)` (named range via
  `Workbook.createName()`), qualificado com a aba atual e referência absoluta.
  Testado funcionando dentro de `formula(...)` e `procurarValor` com avaliação
  real (`SUM(Precos)`, `VLOOKUP(D1,Tabela,2,FALSE)`). **Limitação conhecida e
  documentada**: não funciona ainda com `somar`/`media`/etc. (regex estrita do
  `FormulaBuilder`). Total: 149 testes verdes.
- [x] **Lote M** — `adicionarTotais(celulaCabecalho)` via novo
  `utils/TotalizadorDeTabela`: detecta a tabela (largura pelo cabeçalho, altura
  pela primeira coluna), soma via fórmula `SUM` cada coluna 100% numérica, e
  escreve "Total" na primeira coluna não numérica. Casos de borda testados:
  tabela sem dados (no-op), coluna com texto misturado (ignorada), coluna fora
  do cabeçalho (não tocada). Total: 153 testes verdes.
- [x] **Lote N** — `comentario(celula, texto)` via novo
  `utils/ComentarioHelper` (nota/balão do Excel). Verificado com round-trip
  real (salvar + reabrir). Total: 155 testes verdes.

**Com isso, TODA a lista de "prioridade alta" e "prioridade média" do roadmap
original desta seção foi entregue** (impressão, proteção, validação de
entrada, nomes de intervalo, adicionarTotais, comentários). Resta só
"formatação condicional avançada" (data bars/ícones, prioridade média,
adiada) e os itens de "prioridade baixa" (ver lista atualizada abaixo).

### Sessão de melhorias pós-review (após o usuário pedir uma lista de
sugestões e mandar "pode ir pra cima") — em andamento

- [x] **Bug crítico corrigido**: `Planilha.abrir()` prometia suporte a `.xls`
  mas 6 métodos da facade fazem cast `(XSSFSheet)` internamente — abrir um
  `.xls` de verdade gerava `ClassCastException` cru. Corrigido em
  `PlanilhaBase.abrirPlanilha` (checagem `instanceof XSSFWorkbook` +
  `ArquivoException` amigável). Ver regra 6 na seção 4. Testado com um `.xls`
  real gerado via `HSSFWorkbook` em `AbrirFacadeTest`. Total: 156 testes.
- [x] **CI**: `.github/workflows/ci.yml` roda `mvn clean test` em JDK 8 e 17
  (matrix) a cada push/PR na `main`, com upload dos relatórios de teste como
  artefato. Badge no topo do README. **Verificado com `gh run watch`**: os
  dois jobs passaram (run #28717716179) — confirma de verdade, em JDK 8 real,
  a promessa de compatibilidade Java 8 do projeto.
- [x] **JaCoCo**: `jacoco-maven-plugin` 0.8.11 no `pom.xml`, `report` atado à
  fase `test` — `mvn clean test` já gera
  `target/site/jacoco/index.html`/`jacoco.csv` sem passo extra. **Baseline
  medida em 2026-07-04: ~67,5% de linhas / 68,5% de instruções / 45% de
  branches.** Achado imediato: `utils/RowIteratorUtil` estava em **0%** porque
  era **código morto de verdade** (nunca referenciado em `src/main`, testes
  ou docs) — removido, mesmo precedente do `IBuscaDados` (retomada inicial).
  Pontos fracos daquele baseline (`Fontes`, `ManipuladorPlanilhaHelper` e
  `LogsDeModificadores`) já receberam cobertura estrutural em lotes posteriores.
  Para continuar a frente de cobertura, consulte o `jacoco.csv` atual em vez de
  confiar nessa fotografia histórica.
- [x] **"Colar como valores"**: `colarComoValores(intervalo)` +
  `colarComoValores()` (toda a área usada), via novo
  `utils/ColarComoValoresHelper`. **Bug real encontrado e corrigido pelos
  testes**: `Cell.setCellValue(...)` sozinho não remove a fórmula de uma
  célula `FORMULA` (só atualiza o valor em cache, tipo continua `FORMULA`) —
  é preciso `celula.removeFormula()` antes. Ver regra 7 na seção 4. Total:
  161 testes verdes.
- [x] **Duplicar planilha inteira para outro arquivo**:
  `Planilha.duplicarArquivo(origem, destino)` (método estático, como
  `nova`/`abrir`) — decidido fazer via **cópia direta do arquivo**
  (`Files.copy`) em vez de reconstruir o workbook via POI: mais simples, mais
  robusto (garante byte-a-byte idêntico, sem risco de o POI não
  round-tripar algum recurso avançado como gráfico/imagem/comentário), e
  cobre o caso de uso real sem precisar nem abrir a planilha. Nota: isso já
  era parcialmente possível via `planilha.salvar(a).salvar(b)`
  (encadeável) — o método novo existe para o caso de já ter um arquivo salvo
  e querer duplicá-lo sem reabrir. Total: 163 testes verdes.
- [x] **Cookbook de receitas prontas**: 3 novos exemplos executáveis em
  `examples/` — `ExemploControleFinanceiro` (totais automáticos + destaque
  condicional + gráfico), `ExemploFormularioProtegido` (dropdown + validação
  de data/número + proteção), `ExemploRelatorioParaImpressao` (impressão em 1
  página + `colarComoValores` antes de compartilhar). **Todos rodados de
  verdade** (não só compilados) e o `.xlsx` gerado por cada um foi reaberto e
  verificado (gráfico presente, proteção/validação corretas, área de
  impressão e conversão de fórmula em valor). README ganhou uma seção
  "Receitas prontas" com tabela indexando os 5 exemplos.
- [x] **Agrupamento**: `agruparLinhas(inicio, fim)` e `agruparColunas(de,
  para)` (outline via `groupRow`/`groupColumn`). **Achado de segurança**:
  `setRowGroupCollapsed` esconde a planilha inteira, não só o grupo — por
  isso **não** foi exposto um parâmetro de colapso automático. Ver regra 8 na
  seção 4 para os detalhes completos da investigação. Total: 165 testes.
- [x] **JitPack + CHANGELOG.md**: `CHANGELOG.md` criado (Keep a Changelog);
  `pom.xml` bump `2.0.0-SNAPSHOT` → `2.1.0` (minor, sem breaking changes);
  tag `v2.1.0` criada e enviada; README com instruções de instalação via
  JitPack (+ badge) como opção 1. **Build do JitPack verificado de verdade**
  via `curl` (a ferramenta `WebFetch` do agente leva 403 do JitPack — usar
  `curl -A "Mozilla/5.0" ...` do Bash funciona): `status: "ok"`, `isTag:
  true`, `private: false`, e o `.jar` foi baixado com sucesso (137KB, HTTP
  200) em
  `https://jitpack.io/com/github/isaacbraym/PlanilhasConstructorV2/v2.1.0/PlanilhasConstructorV2-v2.1.0.jar`.
  **Com isso, TODOS os itens da lista de sugestões pós-review foram
  entregues.**

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
| Estado dos testes | **240 testes, todos verdes** (ver seção 0 para o número mais atual) |

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
                         AjustadorDeFormulas,
                         FormatosDeCelula, FormatacaoCondicionalHelper,
                         ListaSuspensaHelper, ProtecaoHelper,
                         ValidacaoDeEntradaHelper, TotalizadorDeTabela,
                         ComentarioHelper, ColarComoValoresHelper,
                         LeitorDeCelulas, LeitorDeTabela, ...
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
6. **`Planilha.abrir()` só aceita `.xlsx`**: `PlanilhaBase.abrirPlanilha` checa
   `instanceof XSSFWorkbook` logo após `WorkbookFactory.create(...)` e lança
   `ArquivoException` amigável para `.xls`. **Não remover essa checagem** —
   `realcarSeX`/`listaSuspensa`/`validarXEntre`/`graficoDeX`/`inserirImagem`/
   `comentario` fazem cast interno `(XSSFSheet)` que lançaria
   `ClassCastException` cru sem ela.
7. **`Cell.setCellValue()` não remove fórmula sozinho**: para converter uma
   célula `FORMULA` num valor fixo, é preciso chamar `celula.removeFormula()`
   **antes** de `setCellValue(...)` — sem isso, o tipo da célula continua
   `FORMULA` (só o valor em cache muda). Ver `ColarComoValoresHelper`.
8. **`Sheet.setRowGroupCollapsed(...)` é PERIGOSO — não usar sem investigar
   mais.** Testado empiricamente (grupo em linhas 2-4, colapso via a linha
   índice 4 e também via índice 0): o XML resultante marcou **TODAS** as 6
   linhas da planilha como `hidden="true"` (não só as do grupo). Por isso
   `agruparLinhas`/`agruparColunas` da facade **não** têm parâmetro de
   colapso — só criam a estrutura de outline (`outlineLevel`), sem esconder
   nada. Se algum dia quiser reativar recolhimento automático, reproduza o
   teste em `AgrupamentoFacadeTest` antes (talvez precise de
   `sheet.setRowSumsBelow`/`setRowSumsRight` explícitos, ou seja mesmo um bug
   da versão 5.2.5 do POI — não assumir que corrigir sozinho é trivial).
9. **`Workbook.setSheetHidden` não impede ocultar a única aba visível** — o
   POI deixa criar um arquivo sem nenhuma aba visível (Excel não abre). A
   facade valida isso em `ocultarAba` (conta abas visíveis antes de ocultar);
   se implementar outro caminho para ocultar aba, replicar a checagem.
10. **`Sheet.createRow(indice)` não deve ser usado às cegas em linha que pode
    existir** — ao criar linhas de total/resumo, use "obter ou criar" primeiro.
    Bug real corrigido em `Calculos.multiplicarColunasComTexto`: uma linha de
    total que já tivesse células em outras colunas podia ser substituída,
    apagando esses dados.
11. **Nunca chame `Cell.setCellType(CellType.FORMULA)` ao copiar fórmulas** —
    no POI 5.2.5 isso lança `IllegalArgumentException`; use
    `setCellFormula(...)` diretamente. O mesmo padrão vale para tipos especiais
    em cópias manuais: erro com `setCellErrorValue(...)`, branco com
    `setBlank()`. Bug real corrigido em `ManipuladorPlanilhaHelper` ao mover
    colunas com fórmula.
12. **`String.split(delimitador)` descarta campos vazios finais** — para
    importação/inserção delimitada, use o helper de `InsersorDeDados` que chama
    `split(..., -1)` quando o delimitador não é vazio. Sem isso, `"A,B,"`
    perde a terceira célula e desloca a estrutura importada. Para
    `List<String>` com delimitador, cada item da lista deve ser uma linha
    delimitada, não uma célula única.
13. **`cloneSheet` deve vir depois da validação do nome da nova aba** — se
    `duplicarAba` clonar primeiro e só depois chamar `setSheetName`, uma falha
    por nome duplicado/inválido deixa uma aba parcial no workbook. Valide com
    `WorkbookUtil.validateSheetName` e cheque duplicidade antes de clonar.
14. **Caminhos de arquivo públicos devem virar `ArquivoException` amigável** —
    valide `null`/vazio antes de `Paths.get(...)` e capture
    `InvalidPathException` quando o comando promete erro de arquivo (ex.:
    `Planilha.duplicarArquivo`).
15. **Busca/filtro deve comparar fórmula pelo resultado calculado** — para
    `buscarLinhas`/`copiarLinhasParaAba`/`moverLinhasParaAba`/
    `removerLinhasOnde`, uma célula `FORMULA` deve casar com o valor exibido ao
    usuário, não com o texto interno da fórmula; use `FormulaEvaluator` com
    fallback seguro.
16. **Ordenação por fórmula deve usar o resultado calculado e ajustar
    referências relativas** — em `OrdenadorDeLinhas`, a chave de uma célula
    `FORMULA` deve vir de `FormulaEvaluator`, não de `getCellFormula()`. Ao
    mover a linha, fórmulas relativas como `B3*2` precisam virar `B2*2` se a
    linha for reescrita na linha 2; use `FormulaShifter.createForRowCopy` e
    valide com round-trip OOXML.
17. **Ordenação deve mover atributos da linha junto com as células** — altura
    personalizada, linha oculta (`Row.getZeroHeight()`) e `RowStyle` fazem parte
    da linha que está sendo ordenada. Ao reescrever em outra posição, limpe
    atributos residuais quando a linha capturada era vazia para evitar que a
    posição antiga "empreste" altura/ocultação à linha errada.
18. **Copias de celulas com formula devem ajustar referencias relativas** —
    `CopiadorDeCelulas` alimenta `duplicarLinha`, `duplicarColuna`,
    `copiarLinhasParaAba` e outros fluxos de copia. Ao copiar formula, use
    `AjustadorDeFormulas` para deslocar linha/coluna como o Excel faria, sem
    mexer nas partes absolutas (`$A$2`, `B$2`, `$A2`). Nao volte a copiar
    `getCellFormula()` diretamente.
19. **`adicionarTotais()` deve tratar formula numerica como coluna numerica** —
    uma coluna calculada (`B2*C2`, `SUM(...)`, etc.) tambem e dado numerico para
    o usuario. Em `TotalizadorDeTabela`, avalie formulas com `FormulaEvaluator`
    antes de decidir se a coluna recebe `SUM(...)`; texto, erro ou booleano
    continuam barrando a soma da coluna.

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
intervalo na mesma/outra aba); gráficos (barras/pizza/linha); imagens/logo; configuração de
impressão (orientação/área/ajustar em N páginas); proteção de planilha e
desbloqueio de células para formulários.
Copias de linhas/colunas/celulas com formulas ajustam referencias relativas
como o Excel, inclusive ao copiar linhas filtradas para outra aba.

Dois bugs reais foram encontrados e corrigidos por testes nesta sessão:
direcionamento de `aplicarEstilos()` em célula única, e
`BorderStyleHelper.verificarBordaEspessa` (bordas finas nunca eram aplicadas em
células sem borda prévia). Ver seção 4 para os detalhes que não podem regredir.

### Prioridade alta (maior valor prático, ainda não coberto)

1. ~~**Configuração de impressão**~~ — **ENTREGUE**: `orientacaoPaisagem`/
   `orientacaoRetrato`, `areaDeImpressao`, `ajustarImpressaoEmPaginas`,
   `cabecalhoDeImpressao`/`rodapeDeImpressao` (com marcadores amigáveis tipo
   `{pagina}`/`{total}`) e `margensDeImpressao` em centímetros.
2. ~~**Proteção de planilha/células**~~ — **ENTREGUE** nesta sessão:
   `protegerPlanilha(senha)` + `desbloquearCelulas(intervalo)`, via novo
   `utils/ProtecaoHelper` (clona `CellStyle` antes de destravar — nunca muta o
   estilo compartilhado). Ver `docs/ARCHITECTURE.md`.
3. ~~**Validação numérica/de data**~~ — **ENTREGUE** nesta sessão:
   `validarNumeroEntre`/`validarInteiroEntre`/`validarDataEntre`, via novo
   `utils/ValidacaoDeEntradaHelper` (classe separada de `ListaSuspensaHelper`
   — mesmo padrão de `DataValidation`, mas sem menu de opções, então não fazia
   sentido morar na mesma classe).

### Prioridade média

4. ~~**Nomes de intervalo (named ranges)**~~ — **ENTREGUE** nesta sessão:
   `definirNome(nome, intervalo)`. Funciona com `formula(...)` e
   `procurarValor`/`procurarValorNaAba` (validado com avaliação real da
   fórmula), e também com `somar`/`media`/`contar`/`minimo`/`maximo` quando o
   nome já existe no `Workbook`.
5. ~~**Mais tipos de formatação condicional**~~ — **ENTREGUE**: barras de dados
   (`barrasDeDados`) e ícones de semáforo (`iconesSemaforo`), via
   `FormatacaoCondicionalHelper`, com round-trip OOXML.
6. ~~**Comentários em células**~~ — **ENTREGUE** nesta sessão:
   `comentario(celula, texto)` via novo `utils/ComentarioHelper`.
7. ~~**`adicionarTotais()` de alto nível**~~ — **ENTREGUE** nesta sessão: soma
   automaticamente cada coluna numérica de uma tabela via novo
   `utils/TotalizadorDeTabela` (detecta largura pelo cabeçalho, altura pela
   primeira coluna, soma também colunas com fórmulas numéricas e ignora colunas
   com qualquer célula não numérica).

### Prioridade baixa / nice-to-have

- ~~Testes de performance/carga com planilhas de milhares de linhas~~ —
  **ENTREGUE**: `CargaFacadeTest` cobre escrita/salvar/reabrir 3.000 linhas e
  ordenação/busca/cópia com 2.000 linhas.
- Exportar para CSV (fora do escopo original — esta lib foca em `.xlsx`;
  avaliar com o usuário antes de assumir que é desejado).

### Convenção para continuar

Cada item novo deve seguir o padrão já estabelecido: lógica pesada num
utilitário estático (`utils/`, ou pacote novo de 1 classe como `graficos/`/
`imagens/` se for um subsistema distinto), facade só delega e documenta,
teste dedicado validando a **estrutura real do POI** (não só "não lança
exceção" — ver `docs/skills-codex/README.md` ou a skill `planilha-testes`),
e atualização de README + spec + este arquivo (seção 0 e 7) a cada lote.
