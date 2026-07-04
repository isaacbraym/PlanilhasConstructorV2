package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.time.LocalDate;
import java.util.List;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de validação de entrada (número/inteiro/data dentro de limites) na
 * facade.
 */
@DisplayName("Planilha — validação de número/data com limites")
class ValidacaoDeEntradaFacadeTest {

	private DataValidationConstraint restricaoUnica(Planilha planilha) {
		Sheet sheet = planilha.workbook().getSheetAt(0);
		List<? extends DataValidation> validacoes = sheet.getDataValidations();
		assertEquals(1, validacoes.size());
		return validacoes.get(0).getValidationConstraint();
	}

	@Test
	@DisplayName("validarNumeroEntre deve criar restrição decimal com os limites informados")
	void deveValidarNumeroEntre() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.validarNumeroEntre("B2:B50", 0.0, 100.5);

			DataValidationConstraint restricao = restricaoUnica(planilha);
			assertEquals(ValidationType.DECIMAL, restricao.getValidationType());
			assertEquals("0.0", restricao.getFormula1());
			assertEquals("100.5", restricao.getFormula2());
		}
	}

	@Test
	@DisplayName("validarInteiroEntre deve criar restrição inteira com os limites informados")
	void deveValidarInteiroEntre() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.validarInteiroEntre("C2:C50", 1, 10);

			DataValidationConstraint restricao = restricaoUnica(planilha);
			assertEquals(ValidationType.INTEGER, restricao.getValidationType());
			assertEquals("1", restricao.getFormula1());
			assertEquals("10", restricao.getFormula2());
		}
	}

	@Test
	@DisplayName("validarDataEntre deve criar restrição de data com os limites informados")
	void deveValidarDataEntre() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.validarDataEntre("D2:D50", LocalDate.of(2024, 1, 1), LocalDate.of(2024, 12, 31));

			DataValidationConstraint restricao = restricaoUnica(planilha);
			assertEquals(ValidationType.DATE, restricao.getValidationType());
		}
	}

	@Test
	@DisplayName("Deve rejeitar (via Excel) um valor fora do limite ao reabrir o arquivo")
	void deveManterRestricaoAposSalvarEReabrir(@org.junit.jupiter.api.io.TempDir java.nio.file.Path pasta)
			throws Exception {
		String caminho = pasta.resolve("form.xlsx").toString();
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.validarInteiroEntre("A1:A1", 1, 5).salvar(caminho);
		}
		try (Planilha reaberta = Planilha.abrir(caminho)) {
			DataValidationConstraint restricao = restricaoUnica(reaberta);
			assertEquals(ValidationType.INTEGER, restricao.getValidationType());
		}
	}
}
