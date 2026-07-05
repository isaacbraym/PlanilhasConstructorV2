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

- Abas: `nova`, `abrir`, `novaAba`, `irParaAba`, `duplicarAba`,
  `Planilha.duplicarArquivo(origem, destino)` (estático, copia o arquivo
  inteiro).
- Escrita: `escrever`, `escreverTexto`, `escreverLinha`, `escreverColuna`,
  `adicionarLinha`, `escreverTabela`, `escreverData`, `escreverDataHora`.
- Leitura (o inverso da escrita — **use isto em vez de `workbook()` cru para
  ler valores de volta**): `ler`, `lerTexto`, `lerNumero`, `lerData`,
  `lerTabela`, `contarLinhasPreenchidas`.
- Fórmulas: `somar`, `media`, `contar`, `minimo`, `maximo`, `seEntao`,
  `multiplicar`, `subtrair`, `dividir`, `formula(celula, expr)`,
  `preencherColuna(coluna, ini, fim, "B{}*C{}")`, `procurarValor`,
  `procurarValorNaAba` (PROCV/VLOOKUP).
- Ordenar: `ordenarPorCrescente`, `ordenarPorDecrescente`, `ordenarPor`
  (fórmulas ordenam pelo resultado calculado e referências relativas acompanham
  a nova linha; altura/ocultação/estilo de linha também viajam com os dados).
- Busca/filtro: `buscarLinhas`, `contarLinhasOnde`, `copiarLinhasParaAba`,
  `moverLinhasParaAba`, `removerLinhasOnde`.
- Formatação condicional: `realcarSeMaiorQue`, `realcarSeMenorQue`,
  `realcarSeEntre`, `realcarSeIgual`, `escalaDeCores`,
  `barrasDeDados(intervalo, cor)`, `iconesSemaforo(intervalo)`.
- Lista suspensa: `listaSuspensa(intervalo, opcoes...)`,
  `listaSuspensaDoIntervalo(intervaloDestino, intervaloOpcoes)` ou
  `listaSuspensaDoIntervalo(intervaloDestino, abaOpcoes, intervaloOpcoes)`.
- Validação de limites (sem menu): `validarNumeroEntre`/`validarInteiroEntre(
  intervalo, min, max)`, `validarDataEntre(intervalo, LocalDate min, max)`.
- Nomes de intervalo: `definirNome(nome, intervalo)` — use com `somar`/
  `media`/`contar`/`minimo`/`maximo`, `formula(...)` e
  `procurarValor`/`procurarValorNaAba`; erros de nome/intervalo nao deixam
  `Name` parcial no workbook.
- Totais automáticos: `adicionarTotais(celulaCabecalho)` — detecta a tabela e
  soma sozinho cada coluna numerica, incluindo colunas com formulas numericas,
  sem precisar chamar `somar()` por coluna.
- Comentários: `comentario(celula, texto)` — nota/balão do Excel.
- Colar como valores: `colarComoValores(intervalo)` ou `colarComoValores()`
  (toda a área usada) — remove fórmulas, congela os valores calculados.
- Gráficos: `graficoDeBarras`/`graficoDePizza`/`graficoDeLinha(titulo,
  intervaloCategorias, intervaloValores, celulaSuperiorEsquerda)`. Sobrecarga
  de 6 argumentos — `(titulo, abaCategorias, intervaloCategorias, abaValores,
  intervaloValores, celulaSuperiorEsquerda)` — busca categorias e valores em
  abas diferentes entre si (ex.: um "Dashboard" resumindo dados de outras
  abas); o gráfico em si sempre fica na aba atual.
- Imagens: `inserirImagem(celula, caminho)` ou `inserirImagem(celula, caminho,
  escala)` — aceita `.png`/`.jpg`/`.jpeg`.
- Impressão: `orientacaoPaisagem`/`orientacaoRetrato`, `areaDeImpressao`,
  `ajustarImpressaoEmPaginas(larguraPaginas, alturaPaginas)`,
  `margensDeImpressao(superiorCm, inferiorCm, esquerdaCm, direitaCm)`,
  `cabecalhoDeImpressao(centro)`/`cabecalhoDeImpressao(esquerda, centro,
  direita)`, `rodapeDeImpressao(centro)`/`rodapeDeImpressao(esquerda, centro,
  direita)` — aceitam marcadores `{pagina}`, `{total}`, `{data}`, `{hora}`,
  `{arquivo}`, `{aba}`.
- Proteção: `desbloquearCelulas(intervalo)` (chame antes) +
  `protegerPlanilha(senha)` — para formulários com campos editáveis.
- Formatos: `formatarComoMoeda`, `formatarComoContabil`, `formatarComoNumero`,
  `formatarComoTexto`, `formatarComoData`, `formatarComoPorcentagem`,
  `formatarComoPersonalizado(intervalo, formatoExcel)` (qualquer formato não
  coberto pelos outros).
- Colunas/linhas: `moverColuna`, `removerColuna`, `limparColuna`,
  `inserirColunaEntre`, `duplicarColuna`, `duplicarLinha` (formulas relativas
  acompanham a nova linha/coluna como no Excel).
- Estilos/layout: `negrito`, `italico`, `corDoTexto`, `corDeFundo`, `centralizar`,
  `fonte`, `tamanhoDaFonte`, `bordas`, `mesclar`, `desmesclar`, `contornarTudo`,
  `removerLinhasDeGrade`, `ajustarColunas`, `larguraColuna`, `alturaLinha`,
  `congelarPrimeiraLinha`, `congelar`, `filtrosNoCabecalho`,
  `agruparLinhas(inicio, fim)`, `agruparColunas(de, para)` (outline — sem
  opção de colapso automático, ver AGENTS.md regra 8 do porquê).
- Ocultar/exibir: `ocultarLinha`/`exibirLinha`, `ocultarColuna`/
  `exibirColuna`, `ocultarAba`/`exibirAba` (recusa ocultar a única aba
  visível).
- Cor da aba: `corDaAba(CorEnum)`.
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
6. No PowerShell, coloque argumentos Maven `-D...` entre aspas. Ex.:
   `mvn "-Dtest=ListaSuspensaFacadeTest" test`; sem aspas, argumentos com ponto
   (como `-Dmdep.outputFile=...`) podem ser quebrados pelo shell.
7. Ao mexer em abas da facade (`novaAba`, `irParaAba`, `duplicarAba`), rode
   `mvn "-Dtest=PlanilhaFacadeTest" test`; ele protege duplicação sem mutação
   parcial quando o nome da cópia já existe ou é inválido.
8. Ao mexer em `Planilha.duplicarArquivo`, rode
   `mvn "-Dtest=DuplicarArquivoFacadeTest" test`; ele protege cópia
   independente e erros amigáveis para origem/destino ausentes ou inválidos.
9. Ao mexer em busca/filtro (`buscarLinhas`, `copiarLinhasParaAba`,
   `moverLinhasParaAba`, `removerLinhasOnde`) ou `FiltroDeLinhas`, rode
   `mvn "-Dtest=BuscaFacadeTest" test`; ele protege comparação por números e
   por resultados avaliados de fórmulas.
10. Ao mexer em ordenação (`ordenarPorCrescente`, `ordenarPorDecrescente`,
   `ordenarPor`) ou `OrdenadorDeLinhas`, rode
   `mvn "-Dtest=OrdenarFacadeTest" test`; ele protege ordenação por texto,
   números e resultado de fórmulas, incluindo ajuste de referências relativas
   e preservação de atributos da linha.
11. Ao mexer em escrita massiva, ordenação, filtros ou cópia de linhas, rode
   `mvn "-Dtest=CargaFacadeTest" test` além do teste focal da feature.
12. Ao mexer em `CopiadorDeCelulas`, `AjustadorDeFormulas`, `duplicarLinha`,
   `duplicarColuna` ou `copiarLinhasParaAba`, rode
   `mvn "-Dtest=CopiadorDeCelulasTest,PlanilhaFacadeTest,BuscaFacadeTest" test`;
   ele protege ajuste de formulas relativas, partes absolutas e round-trip da
   facade.
13. Ao mexer em `adicionarTotais` ou `TotalizadorDeTabela`, rode
   `mvn "-Dtest=TotalizadorFacadeTest" test`; ele protege soma de colunas
   numericas, colunas com formulas numericas e round-trip OOXML.
14. Ao mexer em `definirNome` ou referencias de nomes de intervalo, rode
   `mvn "-Dtest=NomeDeIntervaloFacadeTest" test`; ele protege nomes validos,
   agregacoes prontas/PROCV e ausencia de `Name` parcial em erros.
15. Ao mexer em linhas de total/resumo da API fluente legada, rode
   `mvn "-Dtest=CalculosTest" test` para proteger contra perda de células em
   linhas já existentes.
16. Ao mexer em mover/remover/limpar/inserir coluna, rode
   `mvn "-Dtest=ManipuladorPlanilhaTest" test`; ele protege fórmulas contra o
   bug real de `Cell.setCellType(CellType.FORMULA)` e valida preservação de
   tipo/estilo no recorte de colunas.
17. Ao mexer em `logAlteracoes()` ou `LogsDeModificadores`, rode
   `mvn "-Dtest=LogsDeModificadoresTest" test`; ele captura `System.out` e
   confirma que a fila interna é limpa após exibir.
18. Ao mexer em inserção delimitada, importação de arquivo texto ou
   `InsersorDeDados`, rode `mvn "-Dtest=PlanilhaXlsxTest,CoercaoNumericaTest" test`;
   isso protege campos vazios finais (`"A,B,"`) em string/lista/arquivo,
   coerção segura de CPF/CEP, dados nulos amigáveis e caminhos básicos de
   erro/no-op da API fluente.
19. Ao mexer em fonte, cores, bordas, alinhamento ou autoajuste, rode
   `mvn "-Dtest=EstiloCelulaTest" test`; ele também salva/reabre fonte
   combinada para proteger a serialização OOXML de nome, tamanho, cor e
   atributos.

## Exemplos de referência (cookbook)

Veja `src/main/java/com/abnote/planilhas/examples/`: `ExemploFacade` (tour
geral), `ExemploControleFinanceiro` (totais + destaque + gráfico),
`ExemploFormularioProtegido` (dropdown + validação + proteção),
`ExemploRelatorioParaImpressao` (impressão + colar como valores). Prefira
adaptar um desses ao pedido do usuário em vez de começar do zero.
Se o usuário pedir algo novo e recorrente, considere a skill
`planilha-nova-feature` para adicionar o comando à facade.
