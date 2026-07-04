package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@code colarComoValores} na facade.
 */
@DisplayName("Planilha — colarComoValores")
class ColarComoValoresFacadeTest {

	@Test
	@DisplayName("Deve substituir a fórmula pelo valor calculado dentro do intervalo")
	void deveSubstituirFormulaPeloValor() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", 10, 20, 30).somar("A4", "A1:A3").colarComoValores("A4:A4");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Cell celula = sheet.getRow(3).getCell(0);
			assertEquals(CellType.NUMERIC, celula.getCellType(), "Não deve mais ser fórmula");
			assertEquals(60.0, celula.getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("Deve preservar texto calculado por fórmula (ex.: CONCATENATE)")
	void devePreservarTextoCalculado() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", "Java").escrever("B1", "POI").formula("C1", "CONCATENATE(A1,B1)")
					.colarComoValores("C1:C1");

			Cell celula = planilha.workbook().getSheetAt(0).getRow(0).getCell(2);
			assertEquals(CellType.STRING, celula.getCellType());
			assertEquals("JavaPOI", celula.getStringCellValue());
		}
	}

	@Test
	@DisplayName("Célula sem fórmula dentro do intervalo não deve ser alterada")
	void deveIgnorarCelulaSemFormula() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", "Texto normal").colarComoValores("A1:A1");

			Cell celula = planilha.workbook().getSheetAt(0).getRow(0).getCell(0);
			assertEquals(CellType.STRING, celula.getCellType());
			assertEquals("Texto normal", celula.getStringCellValue());
		}
	}

	@Test
	@DisplayName("colarComoValores() sem argumento deve converter toda a área usada")
	void deveConverterAreaInteira() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", 5, 15).somar("A3", "A1:A2").formula("B1", "A1*2").colarComoValores();

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals(CellType.NUMERIC, sheet.getRow(2).getCell(0).getCellType());
			assertEquals(20.0, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.001);
			assertEquals(CellType.NUMERIC, sheet.getRow(0).getCell(1).getCellType());
			assertEquals(10.0, sheet.getRow(0).getCell(1).getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("colarComoValores() em aba vazia não deve lançar exceção")
	void deveTolerarAbaVazia() throws Exception {
		try (Planilha planilha = Planilha.nova("Vazia")) {
			assertDoesNotThrow(() -> { planilha.colarComoValores(); });
		}
	}
}
