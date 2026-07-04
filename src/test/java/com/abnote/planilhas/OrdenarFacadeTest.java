package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

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
}
