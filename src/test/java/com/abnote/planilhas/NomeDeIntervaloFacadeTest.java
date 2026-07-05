package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import com.abnote.planilhas.exceptions.DadosInvalidosException;
import com.abnote.planilhas.utils.PosicaoConverter;

/**
 * Testes de nomes de intervalo (named ranges) na facade.
 */
@DisplayName("Planilha — definirNome (nomes de intervalo)")
class NomeDeIntervaloFacadeTest {

	@Test
	@DisplayName("definirNome deve registrar o nome apontando para o intervalo, qualificado com a aba")
	void deveDefinirNome() throws Exception {
		try (Planilha planilha = Planilha.nova("Vendas")) {
			planilha.definirNome("Precos", "B2:B10");

			Name nome = planilha.workbook().getName("Precos");
			assertNotNull(nome);
			assertEquals("'Vendas'!$B$2:$B$10", nome.getRefersToFormula());
		}
	}

	@Test
	@DisplayName("Nome definido deve funcionar dentro de formula() e avaliar corretamente")
	void deveUsarNomeDentroDeFormula() throws Exception {
		try (Planilha planilha = Planilha.nova("Vendas")) {
			planilha.escreverColuna("B2", 10, 20, 30).definirNome("Precos", "B2:B4")
					.formula("D1", "SUM(Precos)");

			Cell celula = planilha.workbook().getSheetAt(0).getRow(0).getCell(3);
			assertEquals("SUM(Precos)", celula.getCellFormula());

			FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			avaliador.evaluateFormulaCell(celula);
			assertEquals(60.0, celula.getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("Nome definido deve funcionar nas fórmulas prontas de agregação")
	void deveUsarNomeNasFormulasProntasDeAgregacao() throws Exception {
		try (Planilha planilha = Planilha.nova("Vendas")) {
			planilha.escreverColuna("B2", 10, 20, 30)
					.definirNome("Precos", "B2:B4")
					.somar("D1", "Precos")
					.media("D2", "Precos")
					.contar("D3", "Precos")
					.minimo("D4", "Precos")
					.maximo("D5", "Precos");

			final FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			assertFormulaEAvaliacao(planilha, avaliador, "D1", "SUM(Precos)", 60.0);
			assertFormulaEAvaliacao(planilha, avaliador, "D2", "AVERAGE(Precos)", 20.0);
			assertFormulaEAvaliacao(planilha, avaliador, "D3", "COUNT(Precos)", 3.0);
			assertFormulaEAvaliacao(planilha, avaliador, "D4", "MIN(Precos)", 10.0);
			assertFormulaEAvaliacao(planilha, avaliador, "D5", "MAX(Precos)", 30.0);
		}
	}

	@Test
	@DisplayName("Nome definido deve funcionar dentro de procurarValor (PROCV)")
	void deveUsarNomeDentroDeProcurarValor() throws Exception {
		try (Planilha planilha = Planilha.nova("Tabela")) {
			planilha.escreverLinha("A1", "Codigo", "Preco").adicionarLinha("P1", 2.5).adicionarLinha("P2", 15.9)
					.definirNome("Tabela", "A2:B3")
					.escrever("D1", "P2")
					.procurarValor("E1", "D1", "Tabela", 2);

			Cell celula = planilha.workbook().getSheetAt(0).getRow(0).getCell(4);
			assertEquals("VLOOKUP(D1,Tabela,2,FALSE)", celula.getCellFormula());

			FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			avaliador.evaluateFormulaCell(celula);
			assertEquals(15.9, celula.getNumericCellValue(), 0.01);
		}
	}

	@ParameterizedTest
	@ValueSource(strings = { "1Nome", "Nome Com Espaco", "A1", "$B$2", "" })
	@DisplayName("definirNome deve recusar nomes inválidos com DadosInvalidosException, não IllegalArgumentException crua do POI")
	void deveRecusarNomeInvalido(final String nomeInvalido) throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(DadosInvalidosException.class, () -> planilha.definirNome(nomeInvalido, "A1:A2"));
		}
	}

	private void assertFormulaEAvaliacao(final Planilha planilha, final FormulaEvaluator avaliador,
			final String celula, final String formulaEsperada, final double numeroEsperado) {
		final int[] indices = PosicaoConverter.converterPosicao(celula);
		final Cell celulaFormula = planilha.workbook().getSheetAt(0).getRow(indices[1]).getCell(indices[0]);

		assertEquals(formulaEsperada, celulaFormula.getCellFormula());
		avaliador.evaluateFormulaCell(celulaFormula);
		assertEquals(numeroEsperado, celulaFormula.getNumericCellValue(), 0.001);
	}
}
