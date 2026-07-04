---
name: planilha-testes
description: Escrever ou corrigir testes JUnit 5 para o ProjetoPlanilha3 (biblioteca de planilhas Excel). Use quando o usuário pedir para "escrever testes", "aumentar cobertura", "testar" um método/classe desta biblioteca, ou quando um teste estiver falhando.
---

# Skill: testes no ProjetoPlanilha3

Padrões usados no projeto (JUnit 5, Apache POI). Rode com `mvn clean test`.

## Estrutura de um teste

- Pacote espelha o de produção (`src/test/java/com/abnote/planilhas/...`).
- Use `@DisplayName` em pt-BR e o padrão **ARRANGE / ACT / ASSERT**.
- Use `@TempDir` (JUnit) para arquivos, nunca caminhos fixos como `C:\opt\...`.
- Feche workbooks (`@AfterEach` com `workbook.close()`), ou use try-with-resources.

## Ler o resultado de volta

A forma mais confiável de asserir é **reler a célula** do workbook:

```java
try (Planilha planilha = Planilha.nova("T")) {
    planilha.escrever("A1", 10);
    Sheet sheet = planilha.workbook().getSheetAt(0);
    Cell a1 = sheet.getRow(0).getCell(0);
    assertEquals(CellType.NUMERIC, a1.getCellType());
    assertEquals(10.0, a1.getNumericCellValue(), 0.001);
}
```

Dicas de leitura:
- Tipo da célula: `cell.getCellType()` (`STRING`, `NUMERIC`, `FORMULA`, ...).
- Fórmula: `cell.getCellFormula()`; para avaliar:
  `wb.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell)`.
- Estilo/negrito (xlsx): `((XSSFCellStyle) cell.getCellStyle()).getFont().getBold()`.
- Formato numérico: `cell.getCellStyle().getDataFormatString()` (ex.: contém `R$`).
- Mesclagem: `sheet.getNumMergedRegions()`.

## Salvar e reabrir (teste de round-trip)

```java
String caminho = pasta.resolve("saida.xlsx").toString(); // @TempDir Path pasta
try (Planilha p = Planilha.nova("S")) { p.escrever("A1", "Oi").salvar(caminho); }
try (Workbook wb = new XSSFWorkbook(new File(caminho))) {
    assertEquals("Oi", wb.getSheet("S").getRow(0).getCell(0).getStringCellValue());
}
```

## Testar exceções

```java
assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao("A0"));
```

## Onde olhar como referência

- `PlanilhaFacadeTest` — facade ponta a ponta.
- `CoercaoNumericaTest` — texto x número.
- `EstiloSelecaoTest` — direcionamento de estilos.
- `ManipuladorPlanilhaTest` — operações de coluna.
- `PosicaoConverterTest` — conversões e limites do Excel.

Sempre finalize com `mvn clean test` e **BUILD SUCCESS**.
