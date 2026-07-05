package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaError;
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
	@DisplayName("moverColuna deve deslocar fórmulas sem forçar setCellType(FORMULA)")
	void deveMoverColunaComFormula() {
		sheet.getRow(1).getCell(1).setCellFormula("C2*2");

		new ManipuladorPlanilha(sheet).moverColuna("A", "C");

		assertEquals(CellType.FORMULA, sheet.getRow(1).getCell(0).getCellType());
		assertEquals("C2*2", sheet.getRow(1).getCell(0).getCellFormula());
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

	@Test
	@DisplayName("copiar/colar coluna temporária deve preservar tipo, valor e estilo")
	void devePreservarTiposEEstiloAoCopiarEColarColunaTemporaria() {
		Sheet tipos = workbook.createSheet("Tipos");
		CellStyle formatoDecimal = workbook.createCellStyle();
		formatoDecimal.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

		tipos.createRow(0).createCell(0).setCellValue("Texto");
		Cell numerica = tipos.createRow(1).createCell(0);
		numerica.setCellValue(123.45);
		numerica.setCellStyle(formatoDecimal);
		tipos.createRow(2).createCell(0).setCellValue(true);
		tipos.createRow(3).createCell(0).setCellFormula("B1*2");
		tipos.createRow(4).createCell(0).setCellErrorValue(FormulaError.DIV0.getCode());
		tipos.createRow(5).createCell(0, CellType.BLANK);

		ManipuladorPlanilhaHelper helper = new ManipuladorPlanilhaHelper(tipos, 0);
		helper.colarColunaTemporaria(2, helper.copiarColuna(0));

		assertNull(tipos.getRow(0).getCell(0), "Coluna original deve ser removida após o recorte");
		assertEquals("Texto", tipos.getRow(0).getCell(2).getStringCellValue());
		assertEquals(123.45, tipos.getRow(1).getCell(2).getNumericCellValue(), 0.0001);
		assertEquals("0.00", tipos.getRow(1).getCell(2).getCellStyle().getDataFormatString());
		assertTrue(tipos.getRow(2).getCell(2).getBooleanCellValue());
		assertEquals("B1*2", tipos.getRow(3).getCell(2).getCellFormula());
		assertEquals(FormulaError.DIV0.getCode(), tipos.getRow(4).getCell(2).getErrorCellValue());
		assertEquals(CellType.BLANK, tipos.getRow(5).getCell(2).getCellType());
	}

	@Test
	@DisplayName("obterValorCelulaComoString deve cobrir tipos usados em cabeçalhos")
	void deveConverterTiposDeCelulaParaString() {
		Sheet tipos = workbook.createSheet("Strings");
		Row row = tipos.createRow(0);
		row.createCell(0).setCellValue("Nome");
		row.createCell(1).setCellValue(10);
		row.createCell(2).setCellValue(false);
		row.createCell(3).setCellFormula("B1*2");
		row.createCell(4).setCellErrorValue(FormulaError.VALUE.getCode());

		assertEquals("Nome", ManipuladorPlanilhaHelper.obterValorCelulaComoString(row.getCell(0)));
		assertEquals("10.0", ManipuladorPlanilhaHelper.obterValorCelulaComoString(row.getCell(1)));
		assertEquals("false", ManipuladorPlanilhaHelper.obterValorCelulaComoString(row.getCell(2)));
		assertEquals("B1*2", ManipuladorPlanilhaHelper.obterValorCelulaComoString(row.getCell(3)));
		assertEquals(Byte.toString(FormulaError.VALUE.getCode()),
				ManipuladorPlanilhaHelper.obterValorCelulaComoString(row.getCell(4)));
		assertNull(ManipuladorPlanilhaHelper.obterValorCelulaComoString(null));
	}
}
