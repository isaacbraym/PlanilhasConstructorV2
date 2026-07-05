package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de busca/filtro de linhas na facade.
 */
@DisplayName("Planilha — busca e filtro de linhas")
class BuscaFacadeTest {

	/** Cria um cadastro: cabeçalho + Ana/SP, Bia/RJ, Cid/SP. */
	private Planilha cadastro() {
		Planilha planilha = Planilha.nova("Cadastro");
		planilha.escreverLinha("A1", "Nome", "Estado")
				.adicionarLinha("Ana", "SP")
				.adicionarLinha("Bia", "RJ")
				.adicionarLinha("Cid", "SP");
		return planilha;
	}

	@Test
	@DisplayName("buscarLinhas deve retornar os números (1-based) das linhas correspondentes")
	void deveBuscarLinhas() throws Exception {
		try (Planilha planilha = cadastro()) {
			assertEquals(Arrays.asList(2, 4), planilha.buscarLinhas("B", "SP"));
			assertEquals(Arrays.asList(3), planilha.buscarLinhas("B", "RJ"));
			assertTrue(planilha.buscarLinhas("B", "MG").isEmpty());
		}
	}

	@Test
	@DisplayName("contarLinhasOnde deve contar as correspondências")
	void deveContar() throws Exception {
		try (Planilha planilha = cadastro()) {
			assertEquals(2, planilha.contarLinhasOnde("B", "SP"));
			assertEquals(0, planilha.contarLinhasOnde("B", "MG"));
		}
	}

	@Test
	@DisplayName("copiarLinhasParaAba deve copiar sem alterar a origem")
	void deveCopiarParaAba() throws Exception {
		try (Planilha planilha = cadastro()) {
			planilha.copiarLinhasParaAba("B", "SP", "SomenteSP");

			Sheet destino = planilha.workbook().getSheet("SomenteSP");
			assertNotNull(destino);
			assertEquals("Ana", destino.getRow(0).getCell(0).getStringCellValue());
			assertEquals("Cid", destino.getRow(1).getCell(0).getStringCellValue());
			// Origem intacta: Cid continua na linha 4
			assertEquals("Cid", planilha.workbook().getSheet("Cadastro").getRow(3).getCell(0).getStringCellValue());
		}
	}

	@Test
	@DisplayName("removerLinhasOnde deve remover e subir as linhas de baixo")
	void deveRemover() throws Exception {
		try (Planilha planilha = cadastro()) {
			planilha.removerLinhasOnde("B", "SP");

			Sheet sheet = planilha.workbook().getSheet("Cadastro");
			assertEquals("Nome", sheet.getRow(0).getCell(0).getStringCellValue());
			assertEquals("Bia", sheet.getRow(1).getCell(0).getStringCellValue());
			assertEquals(1, sheet.getLastRowNum(), "Devem sobrar cabeçalho + 1 linha");
			assertTrue(planilha.buscarLinhas("B", "SP").isEmpty());
		}
	}

	@Test
	@DisplayName("moverLinhasParaAba deve copiar e remover da origem")
	void deveMover() throws Exception {
		try (Planilha planilha = cadastro()) {
			planilha.moverLinhasParaAba("B", "SP", "Arquivo");

			assertEquals(2, planilha.workbook().getSheet("Arquivo").getPhysicalNumberOfRows());
			assertEquals(0, planilha.contarLinhasOnde("B", "SP"));
			assertEquals(1, planilha.contarLinhasOnde("B", "RJ"));
		}
	}

	@Test
	@DisplayName("busca deve casar números sem a parte decimal (10, não 10.0)")
	void deveCasarNumeros() throws Exception {
		try (Planilha planilha = Planilha.nova("Notas")) {
			planilha.escreverLinha("A1", "Aluno", "Nota")
					.adicionarLinha("Ana", 10)
					.adicionarLinha("Bia", 7);
			List<Integer> comDez = planilha.buscarLinhas("B", "10");
			assertEquals(Arrays.asList(2), comDez);
		}
	}

	@Test
	@DisplayName("busca deve comparar pelo resultado avaliado de fórmulas")
	void deveBuscarPeloResultadoDeFormula() throws Exception {
		try (Planilha planilha = Planilha.nova("Calculos")) {
			planilha.escreverLinha("A1", "Base", "Dobro")
					.adicionarLinha(10, "")
					.adicionarLinha(7, "")
					.formula("B2", "A2*2")
					.formula("B3", "A3*2");

			assertEquals(Arrays.asList(2), planilha.buscarLinhas("B", "20"));
			assertTrue(planilha.buscarLinhas("B", "A2*2").isEmpty());

			planilha.copiarLinhasParaAba("B", "20", "Achados");
			Sheet destino = planilha.workbook().getSheet("Achados");
			assertEquals(10D, destino.getRow(0).getCell(0).getNumericCellValue());
			assertEquals("A1*2", destino.getRow(0).getCell(1).getCellFormula());
			FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			assertEquals(20.0, avaliador.evaluate(destino.getRow(0).getCell(1)).getNumberValue(), 0.001);
		}
	}
}
