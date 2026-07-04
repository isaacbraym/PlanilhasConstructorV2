# Skills para Codex (e outros agentes sem `.claude/skills`)

O Codex lê automaticamente o [`AGENTS.md`](../../AGENTS.md) na raiz do projeto —
ele já contém as regras principais. Este diretório complementa com **playbooks**
prontos para colar no chat quando quiser acionar um comportamento específico
(equivalentes às skills do Claude Code em `.claude/skills/`).

Como usar: copie o bloco do playbook desejado e cole junto do seu pedido.

---

## Playbook 1 — Criar uma planilha com a facade

> Contexto: use a facade `com.abnote.planilhas.Planilha` (não Apache POI cru).
> Comece por `try (Planilha planilha = Planilha.nova("Aba")) { ... salvar(...); }`
> — a facade fecha sozinha, sem `throws`. Estilize **depois** de escrever. Para
> CEP/CPF/telefone use `escreverTexto` (preserva zero à esquerda). Verbos
> disponíveis: escrever, escreverTexto, escreverLinha/Coluna/Tabela,
> adicionarLinha, somar/media/contar/minimo/maximo/seEntao, formatarComo*,
> moverColuna/removerColuna/limparColuna/inserirColunaEntre/duplicarColuna/
> duplicarLinha, negrito/italico/corDoTexto/corDeFundo/centralizar/fonte/
> tamanhoDaFonte/bordas/mesclar/contornarTudo/removerLinhasDeGrade/
> ajustarColunas/congelarPrimeiraLinha/filtrosNoCabecalho, salvar/salvarNaPasta.
> Contrato completo: `docs/specs/facade-planilha.spec.md`.
> **Tarefa:** <descreva a planilha que você quer>.

---

## Playbook 2 — Adicionar uma feature à biblioteca

> Contexto: implemente a feature de ponta a ponta seguindo o AGENTS.md.
> 1) Se for de usuário final, exponha na facade `Planilha` (método curto em
> pt-BR, JavaDoc com exemplo, retornando `this`); a lógica pesada vai no manager/
> util certo e a facade só delega. 2) Reutilize `PosicaoConverter` e respeite os
> limites do Excel. 3) Escreva teste JUnit 5 relendo a célula via
> `planilha.workbook()` (molde: `PlanilhaFacadeTest`). 4) Atualize a tabela do
> `README.md` e, se mudar comportamento, a spec em `docs/specs/`. 5) Rode
> `mvn clean test` e só conclua com BUILD SUCCESS. Não adicione dependências,
> Spring ou Lombok; não suba a versão do Java. Preserve os invariantes com teste
> (coerção texto/número, alvo de `aplicarEstilos`, limites de posição, `salvar`
> sem `throws IOException`).
> **Tarefa:** <descreva a feature>.

---

## Playbook 3 — Escrever testes

> Contexto: JUnit 5 + Apache POI. Pacote de teste espelha o de produção.
> Use `@DisplayName` em pt-BR e ARRANGE/ACT/ASSERT. Use `@TempDir` para arquivos
> (nunca caminhos fixos). Asserja **relendo a célula**: tipo via
> `cell.getCellType()`, fórmula via `getCellFormula()` (+ `FormulaEvaluator` para
> avaliar), negrito via `((XSSFCellStyle) cell.getCellStyle()).getFont().getBold()`,
> formato via `getCellStyle().getDataFormatString()`, mesclagem via
> `sheet.getNumMergedRegions()`. Exceções com `assertThrows(...)`. Feche
> workbooks. Finalize com `mvn clean test` verde.
> **Tarefa:** <o que testar>.

---

## Regras de ouro (valem para qualquer agente)

- API do usuário final = facade `Planilha`. Só desça para a API fluente/POI
  quando a facade não cobrir.
- Português BR em nomes e comentários; JavaDoc nos métodos públicos.
- Exceções **unchecked** (hierarquia em `exceptions/`).
- `mvn clean test` verde antes de concluir. Sem dependências novas.
