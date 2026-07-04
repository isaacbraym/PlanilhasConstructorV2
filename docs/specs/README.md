# Specs de comportamento

Estas specs descrevem **o que a biblioteca deve fazer** (contrato observável),
em estilo Dado/Quando/Então. Servem como referência para humanos e agentes de
IA e como checklist ao evoluir a API. Cada regra aqui tem (ou deveria ter) um
teste correspondente em `src/test/`.

| Spec | Assunto | Testes relacionados |
|---|---|---|
| [facade-planilha.spec.md](facade-planilha.spec.md) | API amigável `Planilha` (todos os comandos) | `PlanilhaFacadeTest`, `AritmeticaFacadeTest`, `BuscaFacadeTest`, `AbrirFacadeTest`, `DataFacadeTest`, `OrdenarFacadeTest`, `DimensoesFacadeTest`, `FormatacaoCondicionalFacadeTest`, `ListaSuspensaFacadeTest`, `GraficoFacadeTest`, `ImagemFacadeTest` |
| [coercao-de-tipos.spec.md](coercao-de-tipos.spec.md) | Texto x número ao escrever | `CoercaoNumericaTest` |
| [estilos-e-selecao.spec.md](estilos-e-selecao.spec.md) | Alvo de `aplicarEstilos()`, posições e estilos diretos | `EstiloSelecaoTest`, `EstiloCelulaTest`, `PosicaoConverterTest` |

## Convenção

- **Dado** um estado inicial, **Quando** uma ação, **Então** o resultado.
- Posições no formato Excel: coluna (letras) + linha (número), ex.: `A1`, `AA10`.
- Toda regra deve ser verificável por um teste automatizado.
