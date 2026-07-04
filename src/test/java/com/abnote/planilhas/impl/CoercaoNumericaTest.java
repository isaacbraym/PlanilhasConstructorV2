package com.abnote.planilhas.impl;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.interfaces.IPlanilha;

/**
 * Testes da coerção numérica ao inserir dados: números viram número, mas dados
 * sensíveis de quem não programa (CEP, CPF, códigos com zero à esquerda) são
 * preservados como texto.
 */
@DisplayName("Coerção numérica ao inserir dados")
class CoercaoNumericaTest {

	@Test
	@DisplayName("Deve preservar zeros à esquerda como texto")
	void devePreservarZerosAEsquerda() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("T");
			planilha.selecionar().celula("A1").inserirDados("007");
			planilha.selecionar().celula("A2").inserirDados("01310-100");

			Sheet sheet = planilha.obterWorkbook().getSheetAt(0);
			Cell a1 = sheet.getRow(0).getCell(0);
			assertEquals(CellType.STRING, a1.getCellType());
			assertEquals("007", a1.getStringCellValue());

			Cell a2 = sheet.getRow(1).getCell(0);
			assertEquals(CellType.STRING, a2.getCellType());
			assertEquals("01310-100", a2.getStringCellValue());
		}
	}

	@Test
	@DisplayName("Deve converter números simples em numérico")
	void deveConverterNumerosSimples() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("T");
			planilha.selecionar().celula("A1").inserirDados("10");
			planilha.selecionar().celula("A2").inserirDados("3.5");
			planilha.selecionar().celula("A3").inserirDados("-400");

			Sheet sheet = planilha.obterWorkbook().getSheetAt(0);
			assertEquals(CellType.NUMERIC, sheet.getRow(0).getCell(0).getCellType());
			assertEquals(10.0, sheet.getRow(0).getCell(0).getNumericCellValue(), 0.001);
			assertEquals(3.5, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001);
			assertEquals(-400.0, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("Deve manter como texto números não limpos e dígitos longos")
	void deveManterTextoNaoNumerico() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("T");
			planilha.selecionar().celula("A1").inserirDados("NaN");
			planilha.selecionar().celula("A2").inserirDados("Infinity");
			planilha.selecionar().celula("A3").inserirDados("12345678901234567"); // 17 dígitos
			planilha.selecionar().celula("A4").inserirDados("Produto X");

			Sheet sheet = planilha.obterWorkbook().getSheetAt(0);
			assertEquals(CellType.STRING, sheet.getRow(0).getCell(0).getCellType());
			assertEquals(CellType.STRING, sheet.getRow(1).getCell(0).getCellType());
			assertEquals(CellType.STRING, sheet.getRow(2).getCell(0).getCellType());
			assertEquals("12345678901234567", sheet.getRow(2).getCell(0).getStringCellValue());
			assertEquals(CellType.STRING, sheet.getRow(3).getCell(0).getCellType());
		}
	}
}
