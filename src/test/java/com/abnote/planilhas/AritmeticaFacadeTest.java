package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.exceptions.FormulaException;

/**
 * Testes das fórmulas aritméticas da facade (multiplicar, subtrair, dividir,
 * formula livre e preencherColuna).
 */
@DisplayName("Planilha — fórmulas aritméticas")
class AritmeticaFacadeTest {

	private double avaliar(Planilha planilha, String celula) {
		int coluna = celula.charAt(0) - 'A';
		int linha = Integer.parseInt(celula.substring(1)) - 1;
		Sheet sheet = planilha.workbook().getSheetAt(0);
		Cell cell = sheet.getRow(linha).getCell(coluna);
		FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
		avaliador.evaluateFormulaCell(cell);
		return cell.getNumericCellValue();
	}

	@Test
	@DisplayName("multiplicar deve gerar A*B e avaliar corretamente")
	void deveMultiplicar() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("B1", 2).escrever("C1", 10).multiplicar("D1", "B1", "C1");
			assertEquals("B1*C1", planilha.workbook().getSheetAt(0).getRow(0).getCell(3).getCellFormula());
			assertEquals(20.0, avaliar(planilha, "D1"), 0.001);
		}
	}

	@Test
	@DisplayName("subtrair e dividir devem avaliar corretamente")
	void deveSubtrairEDividir() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("B1", 10).escrever("C1", 4).subtrair("D1", "B1", "C1").dividir("E1", "B1", "C1");
			assertEquals(6.0, avaliar(planilha, "D1"), 0.001);
			assertEquals(2.5, avaliar(planilha, "E1"), 0.001);
		}
	}

	@Test
	@DisplayName("formula livre deve aceitar expressão do Excel")
	void deveAceitarFormulaLivre() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", 3).escrever("A2", 7).formula("A3", "=A1+A2");
			assertEquals("A1+A2", planilha.workbook().getSheetAt(0).getRow(2).getCell(0).getCellFormula());
			assertEquals(10.0, avaliar(planilha, "A3"), 0.001);
		}
	}

	@Test
	@DisplayName("preencherColuna deve repetir a fórmula trocando o número da linha")
	void devePreencherColuna() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Preço", "Qtd", "Total")
					.adicionarLinha(2, 3)
					.adicionarLinha(4, 5)
					.preencherColuna("C", 2, 3, "A{}*B{}");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("A2*B2", sheet.getRow(1).getCell(2).getCellFormula());
			assertEquals("A3*B3", sheet.getRow(2).getCell(2).getCellFormula());
			assertEquals(6.0, avaliar(planilha, "C2"), 0.001);
			assertEquals(20.0, avaliar(planilha, "C3"), 0.001);
		}
	}

	@Test
	@DisplayName("fórmula inválida deve lançar FormulaException amigável")
	void deveRejeitarFormulaInvalida() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(FormulaException.class, () -> planilha.formula("A1", "B2*"));
		}
	}
}
