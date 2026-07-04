---
name: planilha-facade
description: Gerar código Java que cria planilhas Excel (.xlsx) com a facade amigável com.abnote.planilhas.Planilha deste projeto. Use quando o usuário pedir para "criar/montar/gerar uma planilha", "fazer um relatório em Excel", uma tabela, lista, controle financeiro ou similar usando esta biblioteca.
---

# Skill: usar a facade `Planilha`

Quando o usuário pedir para criar uma planilha usando este projeto, gere código
com a facade `com.abnote.planilhas.Planilha` (não com Apache POI cru nem com a
API fluente interna, a menos que peçam algo avançado).

## Molde base (sempre comece por aqui)

```java
import com.abnote.planilhas.Planilha;
import com.abnote.planilhas.estilos.estilos.CorEnum;

public class Gerar {
    public static void main(String[] args) {
        try (Planilha planilha = Planilha.nova("NomeDaAba")) {
            // ... comandos ...
            planilha.salvar("C:/tmp/arquivo.xlsx");
        }
    }
}
```

`try (Planilha ...)` não obriga `throws` — a facade fecha sozinha.

## Comandos disponíveis (use SOMENTE estes na facade)

- Abas: `nova`, `abrir`, `novaAba`, `irParaAba`, `duplicarAba`.
- Escrita: `escrever`, `escreverTexto`, `escreverLinha`, `escreverColuna`,
  `adicionarLinha`, `escreverTabela`, `escreverData`, `escreverDataHora`.
- Fórmulas: `somar`, `media`, `contar`, `minimo`, `maximo`, `seEntao`,
  `multiplicar`, `subtrair`, `dividir`, `formula(celula, expr)`,
  `preencherColuna(coluna, ini, fim, "B{}*C{}")`, `procurarValor`,
  `procurarValorNaAba` (PROCV/VLOOKUP).
- Ordenar: `ordenarPorCrescente`, `ordenarPorDecrescente`, `ordenarPor`.
- Busca/filtro: `buscarLinhas`, `contarLinhasOnde`, `copiarLinhasParaAba`,
  `moverLinhasParaAba`, `removerLinhasOnde`.
- Formatação condicional: `realcarSeMaiorQue`, `realcarSeMenorQue`,
  `realcarSeEntre`, `realcarSeIgual`, `escalaDeCores`.
- Lista suspensa: `listaSuspensa(intervalo, opcoes...)`,
  `listaSuspensaDoIntervalo(intervaloDestino, intervaloOpcoes)`.
- Validação de limites (sem menu): `validarNumeroEntre`/`validarInteiroEntre(
  intervalo, min, max)`, `validarDataEntre(intervalo, LocalDate min, max)`.
- Gráficos: `graficoDeBarras`/`graficoDePizza`/`graficoDeLinha(titulo,
  intervaloCategorias, intervaloValores, celulaSuperiorEsquerda)`.
- Imagens: `inserirImagem(celula, caminho)` ou `inserirImagem(celula, caminho,
  escala)` — aceita `.png`/`.jpg`/`.jpeg`.
- Impressão: `orientacaoPaisagem`/`orientacaoRetrato`, `areaDeImpressao`,
  `ajustarImpressaoEmPaginas(larguraPaginas, alturaPaginas)`.
- Proteção: `desbloquearCelulas(intervalo)` (chame antes) +
  `protegerPlanilha(senha)` — para formulários com campos editáveis.
- Formatos: `formatarComoMoeda`, `formatarComoContabil`, `formatarComoNumero`,
  `formatarComoTexto`, `formatarComoData`, `formatarComoPorcentagem`.
- Colunas/linhas: `moverColuna`, `removerColuna`, `limparColuna`,
  `inserirColunaEntre`, `duplicarColuna`, `duplicarLinha`.
- Estilos/layout: `negrito`, `italico`, `corDoTexto`, `corDeFundo`, `centralizar`,
  `fonte`, `tamanhoDaFonte`, `bordas`, `mesclar`, `contornarTudo`,
  `removerLinhasDeGrade`, `ajustarColunas`, `larguraColuna`, `alturaLinha`,
  `congelarPrimeiraLinha`, `congelar`, `filtrosNoCabecalho`.
- Salvar: `salvar`, `salvarNaPasta`.
- Avançado: `avancado()` → `IPlanilha`; `workbook()` → POI; `estilo(pos)`.

Detalhes/contrato: `docs/specs/facade-planilha.spec.md`.

## Regras que evitam erros comuns

1. **Estilize DEPOIS de escrever.** Estilos de fonte só afetam células que já
   existem. Ordem: escrever → formatar → estilizar → salvar.
2. **CEP/CPF/CNPJ/telefone/código:** use `escreverTexto(...)`, nunca `escrever`,
   para não perder zeros à esquerda.
3. **Posições** são Excel: coluna(letra)+linha(número), ex.: `A1`; intervalo com
   `:`, ex.: `A1:C1`.
4. Para cálculo por linha (ex.: Total = Preço × Qtd) use `multiplicar`,
   `subtrair`, `dividir` ou `preencherColuna("D", 2, 10, "B{}*C{}")`. Para
   fórmulas do Excel arbitrárias, `formula("D2", "B2*C2")` (o "=" é opcional).
5. Não adicione dependências nem use Spring/Lombok.

## Exemplo de referência

Veja `src/main/java/com/abnote/planilhas/examples/ExemploFacade.java`.
Se o usuário pedir algo novo e recorrente, considere a skill
`planilha-nova-feature` para adicionar o comando à facade.
