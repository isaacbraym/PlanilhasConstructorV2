package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes do PROCV amigável (procurarValor / procurarValorNaAba).
 */
@DisplayName("Planilha — PROCV (VLOOKUP)")
class ProcvFacadeTest {

	private double avaliar(Planilha planilha, String aba, int linha, int coluna) {
		Cell cell = planilha.workbook().getSheet(aba).getRow(linha).getCell(coluna);
		FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
		avaliador.evaluateFormulaCell(cell);
		return cell.getNumericCellValue();
	}

	@Test
	@DisplayName("procurarValor deve trazer o dado da coluna indicada na mesma aba")
	void deveProcurarNaMesmaAba() throws Exception {
		try (Planilha planilha = Planilha.nova("Tabela")) {
			planilha.escreverLinha("A1", "Codigo", "Preco")
					.adicionarLinha("P1", 2.5)
					.adicionarLinha("P2", 15.9);
			// Procura "P2" (escrito em D1) na tabela A2:B3, trazendo a 2ª coluna (preço).
			planilha.escrever("D1", "P2").procurarValor("E1", "D1", "A2:B3", 2);

			assertEquals("VLOOKUP(D1,A2:B3,2,FALSE)",
					planilha.workbook().getSheetAt(0).getRow(0).getCell(4).getCellFormula());
			assertEquals(15.9, avaliar(planilha, "Tabela", 0, 4), 0.01);
		}
	}

	@Test
	@DisplayName("procurarValorNaAba deve buscar em outra aba")
	void deveProcurarEmOutraAba() throws Exception {
		try (Planilha planilha = Planilha.nova("Produtos")) {
			planilha.escreverLinha("A1", "Codigo", "Nome", "Preco")
					.adicionarLinha("P1", "Caneta", 2.5)
					.adicionarLinha("P2", "Caderno", 15.9);

			planilha.novaAba("Pedidos");
			planilha.escrever("A2", "P2").procurarValorNaAba("B2", "A2", "Produtos", "A2:C3", 3);

			assertEquals("VLOOKUP(A2,'Produtos'!A2:C3,3,FALSE)",
					planilha.workbook().getSheet("Pedidos").getRow(1).getCell(1).getCellFormula());
			assertEquals(15.9, avaliar(planilha, "Pedidos", 1, 1), 0.01);
		}
	}
}
