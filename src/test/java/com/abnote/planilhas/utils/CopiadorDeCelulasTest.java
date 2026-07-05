package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@link CopiadorDeCelulas}.
 */
@DisplayName("CopiadorDeCelulas — cópia de valor e estilo")
class CopiadorDeCelulasTest {

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

	private Cell celula(int linha, int coluna) {
		Row row = sheet.getRow(linha);
		if (row == null) {
			row = sheet.createRow(linha);
		}
		return row.getCell(coluna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	}

	@Test
	@DisplayName("Deve copiar valor de texto e estilo")
	void deveCopiarTexto() {
		Cell origem = celula(0, 0);
		origem.setCellValue("Olá");
		CellStyle estilo = workbook.createCellStyle();
		estilo.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		origem.setCellStyle(estilo);

		Cell destino = celula(1, 1);
		CopiadorDeCelulas.copiar(origem, destino);

		assertEquals("Olá", destino.getStringCellValue());
		assertEquals(estilo, destino.getCellStyle());
	}

	@Test
	@DisplayName("Deve copiar valor numérico, booleano e fórmula")
	void deveCopiarOutrosTipos() {
		Cell numerica = celula(0, 0);
		numerica.setCellValue(42.5);
		Cell destinoNumerica = celula(0, 1);
		CopiadorDeCelulas.copiar(numerica, destinoNumerica);
		assertEquals(42.5, destinoNumerica.getNumericCellValue(), 0.001);

		Cell booleana = celula(1, 0);
		booleana.setCellValue(true);
		Cell destinoBooleana = celula(1, 1);
		CopiadorDeCelulas.copiar(booleana, destinoBooleana);
		assertTrue(destinoBooleana.getBooleanCellValue());

		Cell formula = celula(2, 0);
		formula.setCellFormula("SUM(A1:A2)");
		Cell destinoFormula = celula(2, 1);
		CopiadorDeCelulas.copiar(formula, destinoFormula);
		assertEquals("SUM(B1:B2)", destinoFormula.getCellFormula());
	}

	@Test
	@DisplayName("Deve ajustar formula relativa ao copiar para outra linha")
	void deveAjustarFormulaRelativaAoCopiarLinha() {
		Cell formula = celula(1, 1);
		formula.setCellFormula("A2*2");
		Cell destinoFormula = celula(4, 1);

		CopiadorDeCelulas.copiar(formula, destinoFormula);

		assertEquals("A5*2", destinoFormula.getCellFormula());
	}

	@Test
	@DisplayName("Deve preservar partes absolutas ao ajustar formula copiada")
	void devePreservarReferenciasAbsolutasAoCopiarFormula() {
		Cell formula = celula(1, 1);
		formula.setCellFormula("$A$2+B$2+$A2+A2");
		Cell destinoFormula = celula(4, 2);

		CopiadorDeCelulas.copiar(formula, destinoFormula);

		assertEquals("$A$2+C$2+$A5+B5", destinoFormula.getCellFormula());
	}

	@Test
	@DisplayName("Deve deixar destino em branco quando origem está em branco")
	void deveCopiarCelulaEmBranco() {
		Cell origem = celula(0, 0); // criada em branco (CREATE_NULL_AS_BLANK)
		Cell destino = celula(0, 1);
		destino.setCellValue("valor antigo");

		CopiadorDeCelulas.copiar(origem, destino);

		assertEquals(CellType.BLANK, destino.getCellType());
	}
}
