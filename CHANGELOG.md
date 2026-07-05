# Changelog

Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/).
Este projeto segue [versionamento semântico](https://semver.org/lang/pt-BR/).

## [2.2.0] — 2026-07-04

Reanálise crítica sob a lente "o usuário nunca deveria precisar tocar em
Apache POI para criações avançadas". Fecha os 8 gaps de vazamento de POI
identificados.

### Adicionado
- **Leitura** (o "inverso" de escrever, faltava por completo): `ler`/
  `lerTexto`/`lerNumero`/`lerData`/`lerTabela`/`contarLinhasPreenchidas`, via
  novos `utils/LeitorDeCelulas` e `utils/LeitorDeTabela`.
- **Ocultar/exibir**: `ocultarLinha`/`exibirLinha`, `ocultarColuna`/
  `exibirColuna`, `ocultarAba`/`exibirAba` (recusa ocultar a única aba
  visível).
- **Desmesclar células**: `desmesclar(intervalo)`.
- **Cor da aba**: `corDaAba(CorEnum)`.
- **Formato numérico personalizado** (escape hatch): `formatarComoPersonalizado(intervalo, formatoExcel)`.
- **Cabeçalho/rodapé de impressão**: `cabecalhoDeImpressao`/`rodapeDeImpressao`,
  com marcadores amigáveis (`{pagina}`, `{total}`, `{data}`, `{hora}`,
  `{arquivo}`, `{aba}`) em vez da sintaxe nativa `&P`/`&N`/etc.
- **Gráfico entre abas**: sobrecarga de 6 argumentos em `graficoDeBarras`/
  `graficoDePizza`/`graficoDeLinha` para categorias e valores vindos de abas
  diferentes (ex.: um "Dashboard" resumindo dados de outras abas).
- Suíte de testes: 186 → 197, todos verdes.

### Corrigido
- `definirNome` vazava `IllegalArgumentException` crua do POI para nomes de
  intervalo inválidos — agora vira `DadosInvalidosException` amigável.
- Gráfico com categorias/valores logicamente de abas diferentes montava
  silenciosamente com dado errado (sem erro nenhum), porque tudo resolvia
  contra a aba atual — agora tem suporte real via a sobrecarga de 6 argumentos.

## [2.1.0] — 2026-07-04

### Adicionado
- **Fórmulas**: aritmética (`multiplicar`/`subtrair`/`dividir`/`formula` livre/
  `preencherColuna`), PROCV (`procurarValor`/`procurarValorNaAba`), nomes de
  intervalo (`definirNome`), totais automáticos (`adicionarTotais`), colar
  como valores (`colarComoValores`).
- **Organização de dados**: ordenar linhas (`ordenarPorCrescente`/
  `Decrescente`/`ordenarPor`), buscar/copiar/mover/remover linhas por valor,
  agrupamento de linhas/colunas (`agruparLinhas`/`agruparColunas`).
- **Formatos**: data/hora (`escreverData`/`escreverDataHora`/
  `formatarComoData`), porcentagem, largura/altura de coluna/linha, congelar
  N linhas/colunas.
- **Validação e proteção**: lista suspensa (`listaSuspensa`/
  `listaSuspensaDoIntervalo`), validação numérica/de data
  (`validarNumeroEntre`/`validarInteiroEntre`/`validarDataEntre`), proteção de
  planilha e desbloqueio de células para formulários.
- **Visual**: formatação condicional (`realcarSeMaiorQue`/`MenorQue`/`Entre`/
  `Igual`, `escalaDeCores`), gráficos (`graficoDeBarras`/`Pizza`/`Linha`),
  imagens/logo (`inserirImagem`), comentários em células (`comentario`).
- **Arquivos**: abrir arquivo existente (`Planilha.abrir`), duplicar arquivo
  inteiro (`Planilha.duplicarArquivo`), configuração de impressão
  (`orientacaoPaisagem`/`Retrato`, `areaDeImpressao`,
  `ajustarImpressaoEmPaginas`).
- Cookbook de receitas prontas em `examples/`: `ExemploControleFinanceiro`,
  `ExemploFormularioProtegido`, `ExemploRelatorioParaImpressao`.
- CI (GitHub Actions, JDK 8 + 17) e cobertura de testes (JaCoCo).
- Suíte de testes: 28 → 165, todos verdes.

### Corrigido
- `aplicarEstilos()` estilizava a última linha inserida em vez da célula
  selecionada.
- Coerção numérica quebrava CEP/CPF/CNPJ (zeros à esquerda, números longos).
- `emContabil`/`emMoeda` sem `R$`; `emTexto`/`emMoeda` lançavam
  `UnsupportedOperationException`.
- `BorderStyleHelper` tinha a checagem de "borda espessa" invertida —
  `bordas(intervalo)` não aplicava borda em células sem borda prévia (o caso
  comum).
- `Planilha.abrir()` lançava `ClassCastException` cru ao abrir um `.xls`
  (formato antigo) — agora rejeita com `ArquivoException` amigável.
- `colarComoValores` não convertia de fato a célula (faltava
  `Cell.removeFormula()` antes de `setCellValue`).

### Removido
- Interface órfã `IBuscaDados` (sem implementação, vazava tipos do POI).
- Classe utilitária morta `RowIteratorUtil` (0% de uso, encontrada via
  cobertura JaCoCo).

## [2.0.0] — retomada do projeto

- Facade amigável `com.abnote.planilhas.Planilha` como ponto de entrada
  recomendado, por cima da API fluente existente (`IPlanilha`).
- Reorganização da API fluente em namespaces semânticos (`selecionar()`,
  `converter()`, `formula()`).
- Hierarquia de exceções unchecked (`PlanilhaException` e subclasses).
- README, `AGENTS.md`/`CLAUDE.md`, specs de comportamento e skills para
  agentes de IA.
