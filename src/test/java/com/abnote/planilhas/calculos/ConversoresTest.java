package com.abnote.planilhas.calculos;

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
 * Testes de {@link Conversores}: número, contábil, moeda e texto.
 */
@DisplayName("Conversores — formatos de célula")
class ConversoresTest {

	private Workbook workbook;
	private Sheet sheet;

	@BeforeEach
	void setUp() {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("T");
	}

	@AfterEach
	void tearDown() throws Exception {
		workbook.close();
	}

	private void numero(int rowIndex, int colIndex, double valor) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		row.createCell(colIndex).setCellValue(valor);
	}

	@Test
	@DisplayName("converterEmNumero deve transformar String em número")
	void deveConverterEmNumero() {
		sheet.createRow(0).createCell(0).setCellValue("42.5");

		Conversores.converterEmNumero(sheet, "A1");

		Cell cell = sheet.getRow(0).getCell(0);
		assertEquals(CellType.NUMERIC, cell.getCellType());
		assertEquals(42.5, cell.getNumericCellValue(), 0.001);
	}

	@Test
	@DisplayName("converterEmContabil deve aplicar formato com R$")
	void deveAplicarContabilComReal() {
		numero(0, 0, 1000.0);
		numero(1, 0, 2000.0);

		Conversores.converterEmContabil(sheet, "A1", workbook);

		String formato = sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString();
		assertTrue(formato.contains("R$"), "Formato contábil deve conter 'R$': " + formato);
	}

	@Test
	@DisplayName("converterEmMoeda deve aplicar formato com R$")
	void deveAplicarMoedaComReal() {
		numero(0, 0, 1234.56);

		Conversores.converterEmMoeda(sheet, "A1", workbook);

		String formato = sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString();
		assertTrue(formato.contains("R$"), "Formato de moeda deve conter 'R$': " + formato);
	}

	@Test
	@DisplayName("converterEmTexto deve transformar número em texto sem casas decimais desnecessárias")
	void deveConverterEmTexto() {
		numero(0, 0, 1234.0);

		Conversores.converterEmTexto(sheet, "A1", workbook);

		Cell cell = sheet.getRow(0).getCell(0);
		assertEquals(CellType.STRING, cell.getCellType());
		assertEquals("1234", cell.getStringCellValue());
		assertEquals("@", cell.getCellStyle().getDataFormatString());
	}
}
