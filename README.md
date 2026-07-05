# ProjetoPlanilha3 — Construtor de Planilhas

[![CI](https://github.com/isaacbraym/PlanilhasConstructorV2/actions/workflows/ci.yml/badge.svg)](https://github.com/isaacbraym/PlanilhasConstructorV2/actions/workflows/ci.yml)
[![](https://jitpack.io/v/isaacbraym/PlanilhasConstructorV2.svg)](https://jitpack.io/#isaacbraym/PlanilhasConstructorV2)

> Uma biblioteca Java que deixa **criar planilhas Excel (`.xlsx`) fácil**, mesmo
> para quem quase não programa. Você escreve comandos em português — `escrever`,
> `somar`, `duplicar`, `mover`, `negrito` — e a biblioteca cuida do Apache POI
> por baixo.

```java
try (Planilha planilha = Planilha.nova("Vendas")) {
    planilha.escreverLinha("A1", "Produto", "Preço", "Quantidade")
            .negrito("A1:C1")
            .adicionarLinha("Caneta", 2.50, 100)
            .adicionarLinha("Caderno", 15.90, 30)
            .somar("C4", "C2:C3")
            .formatarComoMoeda("B2")
            .ajustarColunas()
            .salvar("C:/tmp/vendas.xlsx");
}
```

Pronto: isso gera um arquivo `vendas.xlsx` com cabeçalho em negrito, preços em
R$, total somado e colunas ajustadas.

---

## Índice

1. [Instalação](#instalação)
2. [Começando em 5 minutos](#começando-em-5-minutos)
3. [Conceitos básicos](#conceitos-básicos)
4. [Guia de comandos (facade `Planilha`)](#guia-de-comandos-facade-planilha)
5. [API avançada (fluente)](#api-avançada-fluente)
6. [Perguntas frequentes](#perguntas-frequentes)
7. [Para desenvolvedores](#para-desenvolvedores)

---

## Instalação

Requer **Java 8+** e **Maven**. O projeto usa Apache POI 5.2.5 (já declarado no
`pom.xml`).

### Opção 1 — via JitPack (mais fácil, sem clonar o repositório)

Adicione o repositório do JitPack e a dependência (troque `2.1.0` pela tag mais
recente, se houver uma mais nova):

```xml
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>

<dependency>
    <groupId>com.github.isaacbraym</groupId>
    <artifactId>PlanilhasConstructorV2</artifactId>
    <version>2.1.0</version>
</dependency>
```

### Opção 2 — clonando e instalando localmente

```bash
git clone https://github.com/isaacbraym/PlanilhasConstructorV2.git
cd PlanilhasConstructorV2/ProjetoPlanilha3
mvn clean install
```

```xml
<dependency>
    <groupId>com.projetoplanilha</groupId>
    <artifactId>ProjetoPlanilha3</artifactId>
    <version>2.1.0</version>
</dependency>
```

---

## Começando em 5 minutos

Todo trabalho começa com `Planilha.nova("NomeDaAba")` e termina com
`salvar(...)`. Entre os dois, você encaixa comandos:

```java
import com.abnote.planilhas.Planilha;

public class MeuPrograma {
    public static void main(String[] args) {
        try (Planilha planilha = Planilha.nova("Clientes")) {
            planilha.escreverLinha("A1", "Nome", "Cidade", "CEP")
                    .adicionarLinha("Maria", "São Paulo")
                    .escreverTexto("C2", "01310-100")   // CEP nunca vira número!
                    .negrito("A1:C1")
                    .salvar("C:/tmp/clientes.xlsx");
        }
    }
}
```

> **Dica:** use `try (Planilha planilha = ...)`. Assim o arquivo é fechado
> sozinho quando você termina.

---

## Conceitos básicos

- **Célula**: um quadradinho, identificado por **coluna + linha**, como `A1`,
  `B2`, `AA10`. Letra = coluna, número = linha.
- **Intervalo**: um retângulo de células, escrito com dois pontos: `A1:C10`.
- **Aba**: cada "página" da planilha. Você cria com `nova(...)` / `novaAba(...)`
  e troca com `irParaAba(...)`.
- **Encadeamento**: quase todo comando devolve a própria planilha, então você
  pode escrever `planilha.escrever(...).negrito(...).salvar(...)`.

### Texto x número (importante!)

Ao usar `escrever(...)`, a biblioteca é esperta:

| Você escreve            | Vira o quê        | Por quê                                        |
|-------------------------|-------------------|------------------------------------------------|
| `escrever("A1", 10)`    | número `10`       | é um número de verdade                         |
| `escrever("A1", "10")`  | número `10`       | texto que "é número" é convertido por conveniência |
| `escrever("A1", "007")` | **texto** `007`   | zero à esquerda é preservado (CEP, código...)  |
| `escreverTexto("A1","7")`| **texto** `7`    | `escreverTexto` **nunca** converte             |

Para CPF, CNPJ, telefone e códigos, prefira **`escreverTexto(...)`**.

---

## Guia de comandos (facade `Planilha`)

### Criar e navegar entre abas
| Comando | O que faz |
|---|---|
| `Planilha.nova("Vendas")` | cria a planilha com a primeira aba |
| `Planilha.abrir("dados.xlsx")` | abre um arquivo existente para editar |
| `novaAba("Resumo")` | cria uma aba nova e vai para ela |
| `irParaAba("Vendas")` | volta para uma aba existente |
| `duplicarAba("Cópia")` | duplica a aba atual (conteúdo + estilo) |
| `Planilha.duplicarArquivo("a.xlsx", "b.xlsx")` | copia o arquivo inteiro (todas as abas) para outro caminho |

### Escrever dados
| Comando | O que faz |
|---|---|
| `escrever("A1", valor)` | escreve um valor (número ou texto inteligente) |
| `escreverTexto("A1", "007")` | escreve **sempre** como texto |
| `escreverLinha("A1", v1, v2, ...)` | escreve valores lado a lado (horizontal) |
| `escreverColuna("A1", v1, v2, ...)` | escreve valores um embaixo do outro |
| `adicionarLinha(v1, v2, ...)` | acrescenta uma linha após a última preenchida |
| `escreverTabela("A1", listaDeLinhas)` | escreve uma tabela inteira |
| `escreverData("A1", LocalDate.of(2024,1,15))` | escreve uma data (dd/MM/aaaa) |
| `escreverDataHora("A1", LocalDateTime.now())` | escreve data e hora |

### Ler dados (o "inverso" de escrever)
| Comando | O que faz |
|---|---|
| `ler("A1")` | lê o valor como tipo Java natural (Double/String/Boolean/LocalDateTime/null) |
| `lerTexto("A1")` | lê formatado como aparece no Excel (ex.: "R$ 1.234,56") |
| `lerNumero("A1")` | lê como `Double` (`null` se não for número) |
| `lerData("A1")` | lê como `LocalDate` (`null` se não for data) |
| `lerTabela("A1")` | lê as linhas de dados de uma tabela (sem o cabeçalho), como `List<List<Object>>` |
| `contarLinhasPreenchidas("A")` | conta quantas linhas têm dado numa coluna |

Todos leem sem criar nada na planilha (diferente de `escrever`, que cria a
célula se ela não existir) — se a célula não existir, o retorno é `null`
(ou `""` em `lerTexto`, ou lista vazia em `lerTabela`).

### Cálculos (fórmulas do Excel)
| Comando | O que faz |
|---|---|
| `somar("C5", "C2:C4")` | soma um intervalo |
| `media("C5", "C2:C4")` | média |
| `contar("C5", "C2:C4")` | conta valores |
| `minimo(...)` / `maximo(...)` | menor / maior valor |
| `seEntao("D2", "C2>100", "Alto", "Baixo")` | condição se/então |
| `multiplicar("D2", "B2", "C2")` | D2 = B2 × C2 |
| `subtrair("D2", "B2", "C2")` / `dividir(...)` | subtração / divisão entre células |
| `formula("D2", "B2*C2")` | qualquer fórmula do Excel (o "=" é opcional) |
| `preencherColuna("D", 2, 10, "B{}*C{}")` | repete a fórmula por linha (`{}` = nº da linha) |
| `procurarValor("E2", "A2", "H2:J100", 3)` | PROCV: busca A2 na tabela e traz a coluna 3 |
| `procurarValorNaAba("B2", "A2", "Produtos", "A2:C100", 3)` | PROCV buscando em outra aba |
| `definirNome("Precos", "B2:B100")` | dá nome a um intervalo, para usar em `formula(...)`/`procurarValor(...)` |
| `adicionarTotais("A1")` | soma sozinho cada coluna numérica da tabela, logo abaixo do último dado |
| `colarComoValores("A1:D50")` | calcula as fórmulas do intervalo e as substitui pelo valor (útil antes de compartilhar o arquivo) |
| `colarComoValores()` | igual, mas em toda a área usada da aba |

### Formatar valores
| Comando | O que faz |
|---|---|
| `formatarComoMoeda("B2")` | R$ 1.234,56 |
| `formatarComoContabil("B2")` | padrão contábil (R$ alinhado) |
| `formatarComoNumero("B2")` | converte texto em número |
| `formatarComoTexto("B2")` | converte em texto puro |
| `formatarComoData("A2")` | formata a coluna como data (dd/MM/aaaa) |
| `formatarComoPorcentagem("B2")` | formata como % (0,15 → 15%) |
| `formatarComoPersonalizado("B2", "0.00 \"kg\"")` | qualquer formato numérico do Excel não coberto acima |

### Colunas e linhas
| Comando | O que faz |
|---|---|
| `moverColuna("C", "F")` | move uma coluna inteira |
| `removerColuna("I")` | remove uma coluna |
| `limparColuna("B")` | esvazia sem remover |
| `inserirColunaEntre("A", "B")` | insere coluna vazia entre duas |
| `duplicarColuna("A", "D")` | copia conteúdo de uma coluna |
| `duplicarLinha(1, 5)` | copia a linha 1 para a linha 5 |

### Ordenar
| Comando | O que faz |
|---|---|
| `ordenarPorCrescente("B")` | ordena por uma coluna (A→Z, menor→maior), mantendo o cabeçalho |
| `ordenarPorDecrescente("B")` | ordena ao contrário |
| `ordenarPor("B", true, 1)` | ordena sem cabeçalho (a partir da linha 1) |

### Buscar e filtrar linhas
| Comando | O que faz |
|---|---|
| `buscarLinhas("B", "SP")` | números das linhas onde a coluna B é "SP" |
| `contarLinhasOnde("B", "SP")` | quantas linhas têm B = "SP" |
| `copiarLinhasParaAba("B", "SP", "SoSP")` | copia essas linhas para outra aba |
| `moverLinhasParaAba("B", "SP", "Arquivo")` | copia e remove da origem |
| `removerLinhasOnde("B", "SP")` | apaga as linhas correspondentes |

### Formatação condicional (realçar células)
| Comando | O que faz |
|---|---|
| `realcarSeMaiorQue("B2:B20", 100, CorEnum.VERDE)` | pinta o fundo se o valor for maior que 100 |
| `realcarSeMenorQue("B2:B20", 0, CorEnum.VERMELHO_ESCURO)` | pinta se for menor |
| `realcarSeEntre("B2:B20", 10, 20, CorEnum.AMARELO)` | pinta se estiver no intervalo |
| `realcarSeIgual("B2:B20", "Atrasado", CorEnum.VERMELHO_ESCURO)` | pinta se for igual (texto ou número) |
| `escalaDeCores("B2:B20")` | escala vermelho→amarelo→verde (semáforo) |

### Lista suspensa (dropdown)
| Comando | O que faz |
|---|---|
| `listaSuspensa("C2:C50", "Pendente", "Pago", "Atrasado")` | menu com opções fixas |
| `listaSuspensaDoIntervalo("A2:A50", "F2:F5")` | menu com opções vindas de outra coluna (sem limite de 255 caracteres) |
| `validarNumeroEntre("B2:B50", 0, 100.5)` | só aceita números entre 0 e 100,5 |
| `validarInteiroEntre("C2:C50", 1, 10)` | só aceita inteiros entre 1 e 10 |
| `validarDataEntre("D2:D50", LocalDate.of(2024,1,1), LocalDate.of(2024,12,31))` | só aceita datas no intervalo |

### Gráficos
| Comando | O que faz |
|---|---|
| `graficoDeBarras("Vendas", "A2:A5", "B2:B5", "D2")` | gráfico de barras verticais |
| `graficoDePizza("Regiões", "A2:A5", "B2:B5", "D2")` | gráfico de pizza |
| `graficoDeLinha("Progresso", "A2:A5", "B2:B5", "D2")` | gráfico de linha |

Os três recebem: título, intervalo das categorias, intervalo dos valores e a
célula onde o canto superior esquerdo do gráfico deve ficar.

### Comentários
| Comando | O que faz |
|---|---|
| `comentario("B2", "Meta mensal de vendas")` | adiciona uma nota (balãozinho) à célula |

### Imagens
| Comando | O que faz |
|---|---|
| `inserirImagem("B2", "logo.png")` | insere a imagem no tamanho original |
| `inserirImagem("B2", "logo.png", 0.5)` | insere reduzida pela metade |

Aceita `.png`, `.jpg` e `.jpeg`.

### Aparência (estilos)
| Comando | O que faz |
|---|---|
| `negrito("A1:C1")` | negrito |
| `italico("A1")` | itálico |
| `corDoTexto("A1", CorEnum.AZUL)` | cor da letra |
| `corDeFundo("A1", CorEnum.AMARELO)` | cor de fundo |
| `centralizar("A1:C1")` | centraliza |
| `fonte("A1", "Arial")` / `tamanhoDaFonte("A1", 14)` | fonte e tamanho |
| `bordas("A1:C10")` | bordas em um intervalo |
| `mesclar("A1:C1")` | junta células |
| `desmesclar("A1:C1")` | desfaz a junção |
| `contornarTudo()` | bordas em toda a área usada |
| `removerLinhasDeGrade()` | tira o quadriculado cinza |
| `corDaAba(CorEnum.VERDE)` | muda a cor da etiqueta da aba |
| `ajustarColunas()` | ajusta a largura ao conteúdo |
| `larguraColuna("A", 20)` / `alturaLinha(1, 30)` | largura (≈ caracteres) / altura (pontos) |
| `congelarPrimeiraLinha()` / `congelar(1, 2)` | fixa cabeçalho / N linhas e colunas |
| `filtrosNoCabecalho()` | ativa filtros (setinhas) |
| `agruparLinhas(2, 4)` | agrupa linhas para recolher/expandir (relatórios hierárquicos) |
| `agruparColunas("B", "D")` | agrupa colunas para recolher/expandir |
| `ocultarLinha(3)` / `exibirLinha(3)` | esconde/reexibe uma linha |
| `ocultarColuna("C")` / `exibirColuna("C")` | esconde/reexibe uma coluna |
| `ocultarAba("Auxiliar")` / `exibirAba("Auxiliar")` | esconde/reexibe uma aba inteira |

### Impressão
| Comando | O que faz |
|---|---|
| `orientacaoPaisagem()` / `orientacaoRetrato()` | deitada / em pé |
| `areaDeImpressao("A1:F30")` | só essa área sai na impressão |
| `ajustarImpressaoEmPaginas(1, 1)` | encolhe para caber em N páginas (largura, altura) |
| `cabecalhoDeImpressao("Relatório Mensal")` | texto central no topo de cada página impressa |
| `cabecalhoDeImpressao("{arquivo}", "Relatório", "{data}")` | esquerda/centro/direita; aceita `{pagina}`, `{total}`, `{data}`, `{hora}`, `{arquivo}`, `{aba}` |
| `rodapeDeImpressao("Página {pagina} de {total}")` | mesma ideia, no rodapé |

### Proteção (formulários)
| Comando | O que faz |
|---|---|
| `desbloquearCelulas("B2:B10")` | mantém essas células editáveis mesmo protegida |
| `protegerPlanilha("senha")` | trava o resto (chame **depois** de desbloquear os campos) |

```java
// Formulário simples: rótulo travado, campo de entrada editável.
planilha.escreverLinha("A1", "Nome:", "")
        .desbloquearCelulas("B1:B1")
        .protegerPlanilha(""); // senha vazia: protege sem exigir senha
```

### Salvar
| Comando | O que faz |
|---|---|
| `salvar("C:/tmp/arquivo.xlsx")` | salva no caminho |
| `salvarNaPasta("C:/tmp", "arquivo.xlsx")` | salva em uma pasta |

### Receitas prontas (cookbook)

Exemplos completos e executáveis (`main()`) para os casos de uso mais comuns
— rode qualquer um deles e abra o arquivo gerado:

| Exemplo | O que mostra |
|---|---|
| [`ExemploFacade.java`](src/main/java/com/abnote/planilhas/examples/ExemploFacade.java) | tour geral: tabela, fórmula, moeda, estilo, filtros |
| [`ExemploControleFinanceiro.java`](src/main/java/com/abnote/planilhas/examples/ExemploControleFinanceiro.java) | orçamento mensal com totais automáticos, destaque de categorias estouradas e gráfico |
| [`ExemploFormularioProtegido.java`](src/main/java/com/abnote/planilhas/examples/ExemploFormularioProtegido.java) | cadastro com lista suspensa, validação de data/número e proteção (só os campos de entrada ficam editáveis) |
| [`ExemploRelatorioParaImpressao.java`](src/main/java/com/abnote/planilhas/examples/ExemploRelatorioParaImpressao.java) | relatório pronto para imprimir em 1 página (paisagem, área de impressão) e com valores congelados antes de compartilhar |
| [`ExemploFormulas.java`](src/main/java/com/abnote/planilhas/examples/ExemploFormulas.java) | as 11 fórmulas da API fluente avançada |

---

## API avançada (fluente)

Quem quiser mais controle pode usar a API completa, acessível por
`planilha.avancado()` (retorna `IPlanilha`) ou diretamente via `new PlanilhaXlsx()`:

```java
planilha.avancado()
        .selecionar().celula("D10").formula().soma("D2:D9").aplicar();

planilha.avancado()
        .selecionar().intervalo("A1", "C1")
        .aplicarEstilos().aplicarNegrito().corDeFundo(CorEnum.CINZA_CLARO);
```

- `selecionar()` → escolhe célula, intervalo, toda a planilha ou a última linha.
- `converter()` → `emNumero`, `emContabil`, `emMoeda`, `emTexto`.
- `formula()` → 11 fórmulas (SUM, AVERAGE, COUNT, MIN, MAX, IF, COUNTIF, SUMIF,
  CONCATENATE, TODAY, NOW).
- `aplicarEstilos()` → `EstiloCelula` com dezenas de estilos.
- `manipularPlanilha()` → mover/remover/limpar/inserir colunas.

Para o objeto cru do Apache POI: `planilha.workbook()`.

---

## Perguntas frequentes

**Meu CEP/CPF virou número e perdeu o zero.**
Use `escreverTexto(...)` em vez de `escrever(...)`, ou
`formatarComoTexto(...)` na coluna.

**Onde o arquivo é salvo?**
No caminho que você passar em `salvar(...)`. Use uma pasta que exista.

**Dá pra abrir uma planilha existente e editar?**
Sim! Use `Planilha.abrir("caminho.xlsx")` e edite com os mesmos comandos.
Só funciona com arquivos `.xlsx` — se o seu arquivo for `.xls` (formato
antigo), abra no Excel e salve novamente como `.xlsx` antes.

**Minha lista suspensa não aceita todas as opções que eu passei.**
O Excel tem um limite de 255 caracteres somados para listas com opções fixas
(`listaSuspensa`). Para listas maiores, use `listaSuspensaDoIntervalo`,
apontando para uma coluna com as opções — não tem esse limite.

**Preciso fechar a planilha?**
Se usar `try (Planilha planilha = ...)`, é automático. Senão, chame `fechar()`.

**Posso usar `definirNome` dentro de `somar`/`media`?**
Ainda não — essas fórmulas prontas exigem um intervalo de células (ex.:
`"B2:B100"`). Nomes funcionam com `formula(...)` (ex.:
`formula("D1", "SUM(Precos)")`) e com `procurarValor`/`procurarValorNaAba`.

---

## Para desenvolvedores

- **Arquitetura, convenções e regras de contribuição:** veja
  [`AGENTS.md`](AGENTS.md) e [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).
- **Comportamento esperado da API:** veja as specs em [`docs/specs/`](docs/specs/).
- **Rodar os testes:**
  ```bash
  mvn clean test
  ```
- **CI:** todo push/PR na `main` roda a suíte em JDK 8 e 17
  ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)).
- **Changelog:** veja [`CHANGELOG.md`](CHANGELOG.md) para o histórico de
  versões.
- **Cobertura de testes:** `mvn clean test` já gera um relatório JaCoCo em
  `target/site/jacoco/index.html`. Baseline em 2026-07-04: ~67% de linhas
  cobertas (156 testes). Pontos mais fracos: `Calculos`, `CenterStyle`,
  `Fontes`, `ManipuladorPlanilhaHelper` — candidatos ao roadmap de cobertura
  do `AGENTS.md`.
- **Skills para agentes de IA** (Claude Code e Codex): veja
  [`.claude/skills/`](.claude/skills/) e [`docs/skills-codex/`](docs/skills-codex/).

Licença e autoria: projeto de Isaac (ver `pom.xml`).
