package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@link ManipuladorPlanilha}: mover, remover, limpar e inserir
 * colunas — a área com aritmética de índices mais delicada da biblioteca.
 */
@DisplayName("ManipuladorPlanilha — operações com colunas")
class ManipuladorPlanilhaTest {

	private Workbook workbook;
	private Sheet sheet;

	@BeforeEach
	void setUp() {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("T");
		// Cabeçalho: H1 H2 H3 | Dados: a1 b1 c1
		criarLinha(0, "H1", "H2", "H3");
		criarLinha(1, "a1", "b1", "c1");
	}

	@AfterEach
	void tearDown() throws Exception {
		workbook.close();
	}

	private void criarLinha(int rowIndex, String... valores) {
		Row row = sheet.createRow(rowIndex);
		for (int i = 0; i < valores.length; i++) {
			row.createCell(i).setCellValue(valores[i]);
		}
	}

	private String texto(int rowIndex, int colIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			return null;
		}
		Cell cell = row.getCell(colIndex);
		return cell == null ? null : cell.getStringCellValue();
	}

	@Test
	@DisplayName("moverColuna deve reposicionar a coluna preservando os dados")
	void deveMoverColuna() {
		new ManipuladorPlanilha(sheet).moverColuna("A", "C");

		// A vai para C; B e C deslocam para a esquerda -> H2 H3 H1
		assertEquals("H2", texto(0, 0));
		assertEquals("H3", texto(0, 1));
		assertEquals("H1", texto(0, 2));
		assertEquals("a1", texto(1, 2));
	}

	@Test
	@DisplayName("removerColuna deve apagar a coluna e deslocar as seguintes")
	void deveRemoverColuna() {
		new ManipuladorPlanilha(sheet).removerColuna("B");

		assertEquals("H1", texto(0, 0));
		assertEquals("H3", texto(0, 1));
		assertEquals("c1", texto(1, 1));
	}

	@Test
	@DisplayName("limparColuna deve esvaziar a coluna sem deslocar as outras")
	void deveLimparColuna() {
		new ManipuladorPlanilha(sheet).limparColuna("B");

		Cell b0 = sheet.getRow(0).getCell(1);
		assertEquals(CellType.BLANK, b0.getCellType());
		// Vizinhas permanecem
		assertEquals("H1", texto(0, 0));
		assertEquals("H3", texto(0, 2));
	}

	@Test
	@DisplayName("inserirColunaVaziaEntre deve empurrar colunas para a direita")
	void deveInserirColunaVazia() {
		new ManipuladorPlanilha(sheet).inserirColunaVaziaEntre("A", "B");

		assertEquals("H1", texto(0, 0));
		Cell b0 = sheet.getRow(0).getCell(1);
		assertTrue(b0 == null || b0.getCellType() == CellType.BLANK, "Coluna B deve ficar vazia");
		assertEquals("H2", texto(0, 2));
		assertEquals("H3", texto(0, 3));
	}

	@Test
	@DisplayName("inserirColunaVaziaEntre deve exigir colunas adjacentes")
	void deveExigirAdjacencia() {
		assertThrows(IllegalArgumentException.class,
				() -> new ManipuladorPlanilha(sheet).inserirColunaVaziaEntre("A", "C"));
	}
}
