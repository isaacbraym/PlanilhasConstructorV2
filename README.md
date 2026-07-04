# ProjetoPlanilha3 — Construtor de Planilhas

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

```bash
mvn clean install
```

Para usar em outro projeto Maven, adicione a dependência:

```xml
<dependency>
    <groupId>com.projetoplanilha</groupId>
    <artifactId>ProjetoPlanilha3</artifactId>
    <version>2.0.0-SNAPSHOT</version>
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

### Formatar valores
| Comando | O que faz |
|---|---|
| `formatarComoMoeda("B2")` | R$ 1.234,56 |
| `formatarComoContabil("B2")` | padrão contábil (R$ alinhado) |
| `formatarComoNumero("B2")` | converte texto em número |
| `formatarComoTexto("B2")` | converte em texto puro |
| `formatarComoData("A2")` | formata a coluna como data (dd/MM/aaaa) |
| `formatarComoPorcentagem("B2")` | formata como % (0,15 → 15%) |

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
| `contornarTudo()` | bordas em toda a área usada |
| `removerLinhasDeGrade()` | tira o quadriculado cinza |
| `ajustarColunas()` | ajusta a largura ao conteúdo |
| `larguraColuna("A", 20)` / `alturaLinha(1, 30)` | largura (≈ caracteres) / altura (pontos) |
| `congelarPrimeiraLinha()` / `congelar(1, 2)` | fixa cabeçalho / N linhas e colunas |
| `filtrosNoCabecalho()` | ativa filtros (setinhas) |

### Salvar
| Comando | O que faz |
|---|---|
| `salvar("C:/tmp/arquivo.xlsx")` | salva no caminho |
| `salvarNaPasta("C:/tmp", "arquivo.xlsx")` | salva em uma pasta |

Veja um exemplo completo em
[`ExemploFacade.java`](src/main/java/com/abnote/planilhas/examples/ExemploFacade.java).

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

**Preciso fechar a planilha?**
Se usar `try (Planilha planilha = ...)`, é automático. Senão, chame `fechar()`.

---

## Para desenvolvedores

- **Arquitetura, convenções e regras de contribuição:** veja
  [`AGENTS.md`](AGENTS.md) e [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).
- **Comportamento esperado da API:** veja as specs em [`docs/specs/`](docs/specs/).
- **Rodar os testes:**
  ```bash
  mvn clean test
  ```
- **Skills para agentes de IA** (Claude Code e Codex): veja
  [`.claude/skills/`](.claude/skills/) e [`docs/skills-codex/`](docs/skills-codex/).

Licença e autoria: projeto de Isaac (ver `pom.xml`).
