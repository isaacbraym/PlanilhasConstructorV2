package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Testes de ordenação de linhas na facade.
 */
@DisplayName("Planilha — ordenar linhas")
class OrdenarFacadeTest {

	/** Cabeçalho + Carlos/30, Ana/25, Bia/40. */
	private Planilha cadastro() {
		Planilha planilha = Planilha.nova("Pessoas");
		planilha.escreverLinha("A1", "Nome", "Idade")
				.adicionarLinha("Carlos", 30)
				.adicionarLinha("Ana", 25)
				.adicionarLinha("Bia", 40);
		return planilha;
	}

	private String texto(Planilha planilha, int linha, int coluna) {
		return planilha.workbook().getSheetAt(0).getRow(linha).getCell(coluna).getStringCellValue();
	}

	private double numero(Planilha planilha, int linha, int coluna) {
		return planilha.workbook().getSheetAt(0).getRow(linha).getCell(coluna).getNumericCellValue();
	}

	@Test
	@DisplayName("ordenarPorCrescente por texto deve manter o cabeçalho e mover a linha inteira")
	void deveOrdenarPorTexto() throws Exception {
		try (Planilha planilha = cadastro()) {
			planilha.ordenarPorCrescente("A");

			assertEquals("Nome", texto(planilha, 0, 0), "Cabeçalho preservado");
			assertEquals("Ana", texto(planilha, 1, 0));
			assertEquals(25.0, numero(planilha, 1, 1), 0.001, "Idade viajou junto com o nome");
			assertEquals("Bia", texto(planilha, 2, 0));
			assertEquals("Carlos", texto(planilha, 3, 0));
		}
	}

	@Test
	@DisplayName("ordenarPorCrescente por número deve ordenar numericamente")
	void deveOrdenarPorNumeroCrescente() throws Exception {
		try (Planilha planilha = cadastro()) {
			planilha.ordenarPorCrescente("B");

			assertEquals("Ana", texto(planilha, 1, 0)); // 25
			assertEquals("Carlos", texto(planilha, 2, 0)); // 30
			assertEquals("Bia", texto(planilha, 3, 0)); // 40
		}
	}

	@Test
	@DisplayName("ordenarPorDecrescente deve inverter a ordem")
	void deveOrdenarDecrescente() throws Exception {
		try (Planilha planilha = cadastro()) {
			planilha.ordenarPorDecrescente("B");

			assertEquals(40.0, numero(planilha, 1, 1), 0.001);
			assertEquals(30.0, numero(planilha, 2, 1), 0.001);
			assertEquals(25.0, numero(planilha, 3, 1), 0.001);
		}
	}

	@Test
	@DisplayName("ordenarPor com linhaInicial=1 ordena tudo, sem cabeçalho")
	void deveOrdenarSemCabecalho() throws Exception {
		try (Planilha planilha = Planilha.nova("Nums")) {
			planilha.escreverColuna("A1", 3, 1, 2);
			planilha.ordenarPor("A", true, 1);

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals(1.0, sheet.getRow(0).getCell(0).getNumericCellValue(), 0.001);
			assertEquals(2.0, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001);
			assertEquals(3.0, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("ordenarPorCrescente por formula deve usar o resultado calculado e ajustar referencias")
	void deveOrdenarPorResultadoDeFormulaEAjustarReferencias(@TempDir Path tempDir) throws Exception {
		Path arquivo = tempDir.resolve("ordenado-formulas.xlsx");

		try (Planilha planilha = Planilha.nova("Dados")) {
			planilha.escreverLinha("A1", "Item", "Base", "Dobro")
					.escreverLinha("A2", "Grande", 100).formula("C2", "B2*2")
					.escreverLinha("A3", "Pequeno", 2).formula("C3", "B3*2")
					.escreverLinha("A4", "Medio", 30).formula("C4", "B4*2")
					.ordenarPorCrescente("C");

			assertLinhasOrdenadasPorFormula(planilha.workbook().getSheetAt(0));
			planilha.salvar(arquivo.toString());
		}

		try (InputStream entrada = Files.newInputStream(arquivo);
				XSSFWorkbook workbook = new XSSFWorkbook(entrada)) {
			assertLinhasOrdenadasPorFormula(workbook.getSheetAt(0));
		}
	}

	@Test
	@DisplayName("ordenar deve mover altura e ocultacao junto com a linha")
	void deveMoverAtributosDaLinhaAoOrdenar(@TempDir Path tempDir) throws Exception {
		Path arquivo = tempDir.resolve("ordenado-atributos-linha.xlsx");

		try (Planilha planilha = Planilha.nova("Dados")) {
			planilha.escreverLinha("A1", "Nome", "Valor")
					.escreverLinha("A2", "Baixo", 1)
					.escreverLinha("A3", "Alto", 9)
					.alturaLinha(3, 42)
					.ocultarLinha(3)
					.ordenarPorDecrescente("B");

			assertAtributosDeLinhaOrdenada(planilha.workbook().getSheetAt(0));
			planilha.salvar(arquivo.toString());
		}

		try (InputStream entrada = Files.newInputStream(arquivo);
				XSSFWorkbook workbook = new XSSFWorkbook(entrada)) {
			assertAtributosDeLinhaOrdenada(workbook.getSheetAt(0));
		}
	}

	private void assertLinhasOrdenadasPorFormula(Sheet sheet) {
		FormulaEvaluator avaliador = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		assertLinhaFormula(sheet, avaliador, 1, "Pequeno", 2.0, "B2*2", 4.0);
		assertLinhaFormula(sheet, avaliador, 2, "Medio", 30.0, "B3*2", 60.0);
		assertLinhaFormula(sheet, avaliador, 3, "Grande", 100.0, "B4*2", 200.0);
	}

	private void assertLinhaFormula(Sheet sheet, FormulaEvaluator avaliador, int indiceLinha, String item,
			double base, String formula, double resultado) {
		Cell celulaFormula = sheet.getRow(indiceLinha).getCell(2);
		assertEquals(item, sheet.getRow(indiceLinha).getCell(0).getStringCellValue());
		assertEquals(base, sheet.getRow(indiceLinha).getCell(1).getNumericCellValue(), 0.001);
		assertEquals(formula, celulaFormula.getCellFormula());
		assertEquals(resultado, avaliador.evaluate(celulaFormula).getNumberValue(), 0.001);
	}

	private void assertAtributosDeLinhaOrdenada(Sheet sheet) {
		assertEquals("Alto", sheet.getRow(1).getCell(0).getStringCellValue());
		assertEquals(42.0f, sheet.getRow(1).getHeightInPoints(), 0.1f);
		assertTrue(sheet.getRow(1).getZeroHeight());

		assertEquals("Baixo", sheet.getRow(2).getCell(0).getStringCellValue());
		assertEquals(sheet.getDefaultRowHeightInPoints(), sheet.getRow(2).getHeightInPoints(), 0.1f);
		assertFalse(sheet.getRow(2).getZeroHeight());
	}
}
