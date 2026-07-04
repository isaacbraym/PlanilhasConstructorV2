# AGENTS.md — ProjetoPlanilha3

Guia para agentes de IA (Claude Code, Codex e outros) que trabalham **dentro**
de `ProjetoPlanilha3`. Leia por completo antes de editar. Este arquivo é a
fonte de verdade sobre o estado atual; o `docs/historico/CONTEXTO_PROJETO.md` é
histórico e pode estar desatualizado.

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
                         ManipuladorPlanilha(Helper), LoggerUtil, ...
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

- `IBuscaDados` está **sem implementação** — decidir implementar busca básica
  (buscar/mover/copiar/remover linhas) ou remover a interface.
- Fórmula de multiplicação/subtração no `FormulaBuilder` (hoje só há soma e
  agregações; produto existe só via `multiplicarColunasComTexto`).
- Abrir/editar planilhas existentes na facade (hoje ela foca em **criar**).
- Aumentar cobertura de `EstiloCelula`/helpers de estilo.
