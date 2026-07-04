# ProjetoPlanilha3 — Contexto para Agente de IA

> Documento gerado para orientar um agente de IA (Claude Code) que vai atuar neste
> repositório. Contém arquitetura, convenções, estado atual e regras de trabalho.
> **Leia este arquivo por completo antes de fazer qualquer alteração.**

---

> ## ⚠️ ATUALIZAÇÃO (retomada do projeto)
>
> Este documento é **histórico**. A fonte de verdade atual passou a ser
> `ProjetoPlanilha3/AGENTS.md` (+ `CLAUDE.md`), com `README.md`,
> `docs/ARCHITECTURE.md`, `docs/specs/` e skills em `.claude/skills/` e
> `docs/skills-codex/`.
>
> Principais mudanças desde este texto:
> - **Nova facade amigável `com.abnote.planilhas.Planilha`** — ponto de entrada
>   para quem não programa (`escrever`, `somar`, `duplicar`, `mover`, `negrito`,
>   `salvar`...). É a API recomendada; a fluente continua disponível.
> - **Bugs corrigidos:** direcionamento de `aplicarEstilos()` em célula
>   selecionada; coerção numérica preservando CEP/CPF (zeros à esquerda);
>   `emContabil`/`emMoeda` com `R$` e `emTexto` implementados; `todasAsBordasEmTudo`
>   sem área fixa `A1:Z100`; validação de limites do Excel em `PosicaoConverter`.
> - **Testes:** de 28 para **66**, todos verdes (`mvn clean test`).
> - `emTexto`/`emMoeda` **não** lançam mais `UnsupportedOperationException`.
>
> As seções abaixo refletem o estado anterior e podem divergir do código atual.

---

## 1. Visão Geral

**ProjetoPlanilha3** é uma biblioteca Java que atua como *wrapper* sobre o Apache POI,
simplificando drasticamente a criação e manipulação de planilhas Excel (`.xlsx`).

- **Problema que resolve**: Apache POI é poderoso, mas verboso (criar uma célula
  simples com estilo pode levar 10+ linhas). A biblioteca oferece uma API fluente
  que reduz isso a 2-3 linhas.
- **Autor/mantenedor**: Isaac (usuário deste projeto).
- **Status atual**: em refatoração ativa para a versão **2.0.0**, com foco em
  produção (production-readiness), não apenas em features novas.
- **Filosofia adotada**: breaking changes são aceitáveis quando resultam em
  arquitetura mais limpa. Não há compromisso de manter retrocompatibilidade com
  a API 1.x.

---

## 2. Stack Técnica

| Item | Valor |
|---|---|
| Linguagem | Java 8 (`maven.compiler.source/target = 1.8`) |
| Build | Maven |
| Dependência principal | Apache POI 5.2.5 (`poi` + `poi-ooxml`) |
| Testes | JUnit 5.10.1 (Jupiter) |
| Mocking | Mockito 5.8.0 + mockito-junit-jupiter (declarado no pom, ainda pouco usado) |
| IDE do usuário | Eclipse (workspace em `C:\Users\Dell\Desktop\eclipse-workspace\ProjetoPlanilha3`) |
| Versionamento | Git + GitHub (branch `main`) |
| groupId/artifactId | `com.projetoplanilha:ProjetoPlanilha3:2.0.0-SNAPSHOT` |

**Atenção**: o projeto é Java 8 puro, **não é Spring Boot**. Não introduzir
anotações Spring (`@Autowired`, `@Service`, etc.), Jakarta Validation, Lombok ou
qualquer dependência não declarada no `pom.xml` sem confirmar com o usuário antes.

---

## 3. Estrutura de Pacotes (`src/main/java/com/abnote/planilhas/`)

```
com.abnote.planilhas/
├── calculos/          → Calculos.java, Conversores.java (lógica de cálculo/conversão)
├── estilos/           → EstiloCelula.java (fachada de estilos)
│   ├── estilos/       → BoldStyle, BackGroundColor, BorderStyleHelper, CenterStyle,
│   │                    Fontes, AlinhamentoStyle, CorEnum, FonteEnum (9 classes)
│   └── util/          → CellIteratorUtil.java
├── examples/          → (pasta criada, ainda sem exemplos — pendente)
├── exceptions/        → Hierarquia de exceções customizadas (ver seção 5)
├── formulas/          → FormulaBuilder.java (implementa IFormulas)
├── impl/              → Implementações concretas:
│                          PlanilhaBase, PlanilhaXlsx, DataManipulator,
│                          StyleManager, ConversaoManager, SelecaoManager
├── interfaces/        → Contratos públicos (ver seção 4)
└── utils/             → InsersorDeDados, ManipuladorPlanilhaHelper,
                           PositionManager, PosicaoConverter, LoggerUtil,
                           ManipuladorPlanilha
```

**Testes** (`src/test/java/com/abnote/planilhas/`):
```
calculos/CalculosTest.java        (3 testes)
formulas/FormulasTest.java        (18 testes)
impl/PlanilhaXlsxTest.java        (7 testes)
test/TestePlanilha.java           (classe de teste manual/legado)
```
Total confirmado: **28 testes, todos passando (BUILD SUCCESS)**.

Há também um pacote legado fora de `com.abnote.planilhas` —
`br.com.abnote.scp2.base.valid.leitordeplanilhas.util` — que **não faz parte**
deste projeto conceitualmente e não deve ser tocado a menos que solicitado.

---

## 4. API Pública (Interfaces)

A API foi reorganizada na v2.0.0 em **namespaces semânticos**, substituindo métodos
soltos por agrupadores descobríveis via autocomplete.

### `IPlanilhaBasica` (base, `extends AutoCloseable`)
Operações fundamentais: `criarPlanilha`, `criarSheet`, `selecionarSheet`, `salvar`,
`setDiretorioSaida`/`getDiretorioSaida`, `obterWorkbook`, `emTodaAPlanilha`,
`ultimaLinha`, `manipularPlanilha()`, `formula()` (→ `IFormulas`), `close()`.

### `IPlanilha` (interface principal, `extends IPlanilhaBasica, IEstilos, IManipulacaoDados`)
Adiciona: `inserirFiltros()`, `getNumeroDeLinhas`, `getNumeroDeColunasNaLinha`,
`naUltimaLinha`, e os dois novos pontos de entrada:
```java
IConversao converter();   // operações de formatação/conversão de célula
ISelecao selecionar();    // operações de posicionamento
```

### `ISelecao` (novo namespace de seleção)
```java
IPlanilha celula(String posicao);
IPlanilha intervalo(String posicaoInicial, String posicaoFinal);
IPlanilha todaPlanilha();
IManipulacaoDados ultimaLinha(String coluna);
```
Uso: `planilha.selecionar().celula("A1").inserirDados("Teste");`

### `IConversao` (novo namespace de conversão)
```java
IPlanilha emNumero(String posicaoInicial);
IPlanilha emContabil(String posicaoInicial);
IPlanilha emTexto(String posicaoInicial);   // lança UnsupportedOperationException — planejado p/ v2.1.0
IPlanilha emMoeda(String posicaoInicial);   // lança UnsupportedOperationException — planejado p/ v2.1.0
```
Uso: `planilha.converter().emContabil("A1");`

### `IFormulas` (builder fluente de fórmulas, via `IPlanilhaBasica.formula()`)
11 fórmulas suportadas:
```java
soma, media, contar, minimo, maximo          // matemáticas
seEntao (IF), contarSe (COUNTIF), somarSe (SUMIF)  // condicionais
concatenar, hoje (TODAY), agora (NOW)        // adicionais
aplicar()                                     // finaliza e grava na célula
```
Uso:
```java
planilha.selecionar().celula("D10").formula().soma("A1:A9").aplicar();
```
**Pré-requisito**: é necessário selecionar uma célula (`selecionar().celula(...)`)
antes de chamar `.formula()`, senão lança `IllegalStateException`
("Nenhuma célula foi selecionada").

### `IManipulacaoDados`
Inserção e manipulação de dados: `inserirDados` (String/List/Object+delimitador),
`inserir(String/int/double)`, `mesclarCelulas`, `somarColunaComTexto`,
`multiplicarColunasComTexto`, `aplicarEstilos()` (→ `EstiloCelula`),
`naUltimaLinha`. Os métodos antigos `naCelula(...)` / `noIntervalo(...)`
foram comentados/removidos desta interface em favor de `selecionar()`.

### `IEstilos` / `EstiloCelula`
Fachada de estilos: negrito, cor de fonte/fundo, alinhamento, bordas, mesclagem,
redimensionamento de colunas, remoção de linhas de grade, etc.

### `IBuscaDados`
**Interface vazia — pendente de análise.** Decisão em aberto: implementar busca
básica por critério ou remover a interface (ver seção 7, item pendente).

---

## 5. Hierarquia de Exceções Customizadas (criada nesta refatoração)

**Decisão arquitetural**: todas as exceções são **unchecked** (`extends RuntimeException`),
não `Exception` (checked). Motivo: manter a API fluente limpa, sem poluir todas as
assinaturas de método com `throws`. O usuário pode capturá-las opcionalmente.

```
PlanilhaException (RuntimeException) — classe base
  ├─ FormulaException          — campo extra: getFormulaTentada()
  ├─ PosicaoInvalidaException  — campo extra: getPosicaoInvalida()
  ├─ DadosInvalidosException   — campo extra: getDadoInvalido()
  └─ ArquivoException          — campo extra: getCaminhoArquivo()
                                  (tem construtor mensagem+caminho+causa, usado
                                  para encapsular IOException preservando a causa raiz)
```

**Status de integração**:
- `FormulaBuilder` → lança `FormulaException` em `validarRange`, `validarCondicao`
  (validarCriterio ainda pendente de confirmação).
- `ConversaoManager` / `SelecaoManager` → lançam `PosicaoInvalidaException` em
  `validarPosicao`/`validarColuna`.
- `PlanilhaBase.salvar()` / `InsersorDeDados` → lançam `ArquivoException`,
  encapsulando `IOException` como causa raiz.
- **Assinatura de `salvar()`**: foi alterada para **não declarar mais**
  `throws IOException` (a exceção agora é unchecked via `ArquivoException`).
  ⚠️ Isso é uma mudança de assinatura de método já aplicada — o agente deve
  saber disso para não reintroduzir o `throws IOException` por engano.

**Pendência conhecida**: os testes em `FormulasTest` foram atualizados para
esperar `FormulaException` no lugar de `IllegalArgumentException`/`IllegalStateException`
antigos — mas os `.txt` de project knowledge ainda mostram a versão antiga dos
testes (esperando `IllegalArgumentException`). **Confiar no código real do
repositório, não nos `.txt`, em caso de divergência.**

---

## 6. Estado Atual (última execução confirmada)

```
mvn clean test
Tests run: 28, Failures: 0, Errors: 0, Skipped: 0
BUILD SUCCESS
```

- ✅ API reorganizada (`converter()` / `selecionar()`) funcionando e testada.
- ✅ Hierarquia de exceções criada, integrada e testada.
- ✅ Projeto compilando limpo (44 arquivos fonte principais).
- ✅ Código sincronizado no GitHub (branch `main`), incluindo `.gitignore`
  criado para excluir metadata do Eclipse (`.metadata/`, `.settings/`,
  `.classpath`, `.project`) e outros projetos não relacionados que compartilham
  o mesmo workspace do usuário.

**Aviso ambiental**: o warning `Log4j2 could not find a logging implementation`
aparece nos testes — o projeto usa `java.util.logging` via `LoggerUtil`
customizado, não Log4j2. Provavelmente uma dependência transitiva do POI
tentando usar Log4j2; não é um erro funcional, mas pode ser investigado se
incomodar nos logs de CI.

---

## 7. Roadmap / Pendências (ordem sugerida)

**Prioridade alta:**
1. Criar `ExemploFormulas.java` demonstrando as 11 fórmulas (pacote `examples/` já existe, vazio).
2. Completar JavaDoc faltante em `PositionManager`, `PosicaoConverter` e métodos privados importantes.
3. Criar 3-5 testes de integração (cenários reais combinando features: criar → inserir → estilizar → fórmula → salvar).
4. Analisar `IBuscaDados` (interface vazia) — decidir entre implementar busca básica ou remover.
5. Implementar validações de limites do Excel `.xlsx`:
   - Máximo 1.048.576 linhas
   - Máximo 16.384 colunas (até `XFD`)
   - Deve lançar `PosicaoInvalidaException` (já existe, só falta aplicar nos pontos certos).

**Prioridade média:**
6. Aumentar cobertura de testes para 85%+ (faltam testes diretos para `ManipuladorPlanilhaHelper`, `PosicaoConverter`, `PositionManager`).
7. Criar exemplos `.java` cobrindo todos os casos do `README.md`.
8. Implementar `emTexto()` e `emMoeda()` em `IConversao` (hoje lançam `UnsupportedOperationException`, planejadas para v2.1.0 — confirmar com o usuário se antecipar para 2.0.0).

**Necessário levantar (análise, não implementação direta):**
- Code smells (duplicação, métodos longos).
- Violações SOLID.
- Gargalos de performance.
- Pontos sem validação de input.

---

## 8. Convenções de Código (OBRIGATÓRIAS)

- **Idioma**: nomes de métodos/variáveis e comentários/JavaDoc em **português
  brasileiro** (ex.: `processarPagamento`, `buscarUsuario`). JavaDoc só em
  métodos públicos e **apenas quando solicitado explicitamente**.
- **Tamanho**: máximo 25 linhas por método; complexidade ciclomática ≤ 15.
- **Imutabilidade**: preferir `final` em parâmetros/campos; usar `Optional` em
  vez de retornar `null`.
- **Injeção**: sempre via construtor (nunca campo/field injection) — mesmo
  fora de contexto Spring, manter o princípio para as classes internas do
  projeto (ex.: `ConversaoManager`, `SelecaoManager` já seguem isso).
- **Nomes específicos**: proibido usar nomes genéricos como `total`, `soma`,
  `lista`, `resultado`, `valor`, `temp`, `aux` sem prefixo/sufixo que os torne
  únicos no escopo.
- **Estilo**: Google Java Style Guide.
- **Exceções unchecked**: seguir o padrão já estabelecido na seção 5 — não
  criar exceções checked (`extends Exception`) neste projeto.

### Regras de Edição Cirúrgica (para o agente)
1. **Nunca** renomear variáveis/métodos/classes existentes sem pedido explícito.
2. **Nunca** adicionar imports desnecessários.
3. **Nunca** refatorar código não mencionado na tarefa atual.
4. **Nunca** mudar assinatura de método sem avisar claramente antes.
5. Em caso de incerteza sobre requisito: **perguntar antes de implementar**.
6. Mudanças devem ser o **conjunto mínimo necessário** para resolver o problema.
7. Preferir O(n) ou O(n log n); usar `StringBuilder` em loops de concatenação.
8. Toda operação de I/O precisa de tratamento de exceção (try-with-resources
   quando aplicável — o projeto já usa isso em `PlanilhaBase.salvar()`).

---

## 9. Git / GitHub

- Repositório já inicializado e sincronizado (branch `main`).
- **Atenção crítica**: o workspace do Eclipse do usuário
  (`eclipse-workspace/`) contém **múltiplos outros projetos não relacionados**
  (ex.: `ProjetoBitLifeCloneV1`, `ControleMouseVirtual`, `GerenciamentoFinanceiro`,
  etc.) no mesmo diretório pai. O `.gitignore` do `ProjetoPlanilha3` já foi
  ajustado para não subir metadata do Eclipse (`.metadata/`, `.classpath`,
  `.project`, `.settings/`), mas **qualquer comando `git add` deve ser
  escopado para dentro da pasta do próprio `ProjetoPlanilha3`** — nunca usar
  `git add ..` ou comandos que alcancem o diretório pai.
- Convenção de commit usada até agora: mensagem descritiva em português,
  cabeçalho curto + lista de mudanças (estilo *conventional-ish*, não
  estritamente Conventional Commits).

---

## 10. Observação sobre a Memória deste Projeto (Claude.ai)

Os arquivos `.txt` anexados a este projeto no Claude.ai (`impl.txt`,
`interfaces.txt`, `formulas.txt`, etc.) são **snapshots gerados manualmente**
pelo usuário em 28/12/2025 e **ficaram desatualizados** após a criação do
pacote `exceptions/` e da nova API `selecionar()/converter()` (28-29/12).
Um agente trabalhando diretamente no repositório (Claude Code) deve **sempre
ler o código-fonte real em disco**, não confiar nesses `.txt` para o estado
atual — eles servem apenas como contexto histórico complementar aqui no chat.

Recomenda-se ao usuário regenerar esses `.txt` após cada lote grande de
mudanças, para manter o contexto do Claude.ai alinhado com o repositório.
