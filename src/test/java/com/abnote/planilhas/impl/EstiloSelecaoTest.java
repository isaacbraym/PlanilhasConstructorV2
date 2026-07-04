package com.abnote.planilhas.impl;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.interfaces.IPlanilha;

/**
 * Testes do direcionamento de estilos: garante que aplicarEstilos() após
 * selecionar uma célula estiliza exatamente aquela célula (regressão do bug em
 * que estilizava uma linha antiga inserida).
 */
@DisplayName("Direcionamento de aplicarEstilos()")
class EstiloSelecaoTest {

	private boolean negrito(Sheet sheet, int rowIndex, int colIndex) {
		Cell cell = sheet.getRow(rowIndex).getCell(colIndex);
		return ((XSSFCellStyle) cell.getCellStyle()).getFont().getBold();
	}

	@Test
	@DisplayName("Deve estilizar a célula selecionada, não a última linha inserida")
	void deveEstilizarCelulaSelecionada() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("T");
			planilha.selecionar().celula("A1").inserirDados("Titulo");
			planilha.selecionar().celula("A5").inserirDados("Rodape");

			// Seleciona A1 de novo e aplica negrito: deve afetar A1, não a linha 5.
			planilha.selecionar().celula("A1").aplicarEstilos().aplicarNegrito();

			Sheet sheet = planilha.obterWorkbook().getSheetAt(0);
			assertTrue(negrito(sheet, 0, 0), "A1 deve estar em negrito");
			assertFalse(negrito(sheet, 4, 0), "A5 não deveria estar em negrito");
		}
	}

	@Test
	@DisplayName("Deve estilizar a linha inteira recém-inserida via lista")
	void deveEstilizarLinhaInserida() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("T");
			planilha.selecionar().celula("A1").inserirDados("Nome,Idade,Cidade", ",");
			planilha.aplicarEstilos().aplicarNegrito();

			Sheet sheet = planilha.obterWorkbook().getSheetAt(0);
			assertTrue(negrito(sheet, 0, 0), "A1 deve estar em negrito");
			assertTrue(negrito(sheet, 0, 1), "B1 deve estar em negrito");
			assertTrue(negrito(sheet, 0, 2), "C1 deve estar em negrito");
		}
	}

	@Test
	@DisplayName("todasAsBordasEmTudo deve contornar a área usada sem erro")
	void deveContornarAreaUsada() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("T");
			planilha.selecionar().celula("A1").inserirDados("X");
			planilha.selecionar().celula("B2").inserirDados("Y");

			planilha.todasAsBordasEmTudo();

			Sheet sheet = planilha.obterWorkbook().getSheetAt(0);
			assertEquals(BorderStyle.THICK, sheet.getRow(0).getCell(0).getCellStyle().getBorderTop());
			assertEquals(BorderStyle.THICK, sheet.getRow(0).getCell(0).getCellStyle().getBorderLeft());
		}
	}

	@Test
	@DisplayName("todasAsBordasEmTudo em planilha vazia não deve lançar exceção")
	void deveTolerarPlanilhaVazia() throws Exception {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("Vazia");
			assertDoesNotThrow(() -> planilha.todasAsBordasEmTudo());
		}
	}
}
