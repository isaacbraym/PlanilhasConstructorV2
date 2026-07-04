---
name: planilha-nova-feature
description: Adicionar uma nova capacidade à biblioteca ProjetoPlanilha3 de ponta a ponta (facade + camada interna + teste + docs). Use quando o usuário pedir para "adicionar/implementar/criar um comando/método/feature" na biblioteca de planilhas, ou expor algo novo na facade Planilha.
---

# Skill: adicionar feature ao ProjetoPlanilha3

Fluxo padrão para introduzir uma nova capacidade sem quebrar nada. Sempre
termine com `mvn clean test` verde.

## Passo 1 — Decidir a camada

- **Feature de usuário final** → precisa aparecer na facade
  `com.abnote.planilhas.Planilha` como um verbo simples que devolve `Planilha`.
- **Lógica de baixo nível** → implemente no manager/util certo
  (`DataManipulator`, `StyleManager`, `ConversaoManager`, `Calculos`,
  `Conversores`, `ManipuladorPlanilha`, `FormulaBuilder`) e apenas **delegue** a
  partir da facade.

## Passo 2 — Implementar

1. Adicione o método na camada interna (respeite ≤25 linhas, `final` em
   parâmetros, nomes específicos em pt-BR, injeção por construtor).
2. Exponha na facade `Planilha` um método curto, com **JavaDoc em pt-BR** e um
   mini-exemplo. Deve retornar `this` (`Planilha`) para encadear, salvo se
   devolver um objeto de configuração (ex.: `EstiloCelula`).
3. Reutilize `PosicaoConverter` para posições e respeite os limites do Excel.

## Passo 3 — Testar (obrigatório)

- Crie/atualize um teste JUnit 5. Para código da facade, use
  `src/test/java/com/abnote/planilhas/PlanilhaFacadeTest.java` como molde:
  escreva com a facade e **releia a célula** via `planilha.workbook()`.
- Cubra o caminho feliz e ao menos um caso de borda/erro.
- Veja a skill `planilha-testes` para padrões.

## Passo 4 — Documentar

- Adicione o comando na tabela do `README.md` (seção "Guia de comandos").
- Se mudar comportamento observável, atualize a spec em `docs/specs/`.
- Se for um verbo novo importante, cite no `AGENTS.md` (seção 3/4) se afetar
  regras que não podem regredir.

## Passo 5 — Verificar

```bash
mvn clean test
```

Só conclua com **BUILD SUCCESS** e todos os testes verdes. Não bump de versão
nem novas dependências sem pedido explícito.

## Invariantes a preservar (têm teste!)

- Coerção texto/número (`docs/specs/coercao-de-tipos.spec.md`).
- Alvo de `aplicarEstilos()` e reset de seleção (`docs/specs/estilos-e-selecao.spec.md`).
- Validação de limites do Excel em `PosicaoConverter`.
- `salvar()` não declara `throws IOException` (exceção é unchecked).
