package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.exceptions.DadosInvalidosException;

/**
 * Testes de listas suspensas (validação de dados) na facade.
 */
@DisplayName("Planilha — listas suspensas (dropdown)")
class ListaSuspensaFacadeTest {

	private DataValidation validacaoUnica(Planilha planilha) {
		Sheet sheet = planilha.workbook().getSheetAt(0);
		List<? extends DataValidation> validacoes = sheet.getDataValidations();
		assertEquals(1, validacoes.size());
		return validacoes.get(0);
	}

	@Test
	@DisplayName("listaSuspensa deve criar validação com as opções fixas e mostrar a seta")
	void deveCriarListaComOpcoesFixas() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.listaSuspensa("C2:C10", "Pendente", "Pago", "Atrasado");

			DataValidation validacao = validacaoUnica(planilha);
			assertArrayEquals(new String[] { "Pendente", "Pago", "Atrasado" },
					validacao.getValidationConstraint().getExplicitListValues());
			assertTrue(validacao.getSuppressDropDownArrow(), "Seta do menu deve aparecer");
			assertEquals("C2:C10", validacao.getRegions().getCellRangeAddress(0).formatAsString());
		}
	}

	@Test
	@DisplayName("listaSuspensaDoIntervalo deve referenciar o intervalo de opções em absoluto")
	void deveCriarListaDeIntervalo() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("F2", "Norte", "Sul", "Leste", "Oeste");
			planilha.listaSuspensaDoIntervalo("A2:A10", "F2:F5");

			DataValidation validacao = validacaoUnica(planilha);
			assertEquals("$F$2:$F$5", validacao.getValidationConstraint().getFormula1());
		}
	}

	@Test
	@DisplayName("listaSuspensaDoIntervalo deve aceitar intervalo de opções já absoluto")
	void deveAceitarIntervaloDeOpcoesJaAbsoluto() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("F2", "Norte", "Sul", "Leste", "Oeste");
			planilha.listaSuspensaDoIntervalo("A2:A10", "$F$2:$F$5");

			DataValidation validacao = validacaoUnica(planilha);
			assertEquals("$F$2:$F$5", validacao.getValidationConstraint().getFormula1());
		}
	}

	@Test
	@DisplayName("listaSuspensaDoIntervalo deve referenciar opções em outra aba sem fórmula manual")
	void deveCriarListaDeIntervaloEmOutraAba(@TempDir Path pasta) throws Exception {
		File arquivo = pasta.resolve("lista-outra-aba.xlsx").toFile();
		try (Planilha planilha = Planilha.nova("Formulario")) {
			planilha.novaAba("Opcoes Aux")
					.escreverColuna("F2", "Pendente", "Pago", "Atrasado")
					.irParaAba("Formulario")
					.listaSuspensaDoIntervalo("A2:A10", "Opcoes Aux", "F2:F4")
					.salvar(arquivo.getAbsolutePath());
		}

		try (Workbook workbook = new XSSFWorkbook(arquivo)) {
			List<? extends DataValidation> validacoes = workbook.getSheet("Formulario").getDataValidations();
			assertEquals(1, validacoes.size());
			DataValidation validacao = validacoes.get(0);
			assertEquals("'Opcoes Aux'!$F$2:$F$4", validacao.getValidationConstraint().getFormula1());
			assertEquals("A2:A10", validacao.getRegions().getCellRangeAddress(0).formatAsString());
		}
	}

	@Test
	@DisplayName("listaSuspensaDoIntervalo deve lançar erro amigável quando a aba de opções não existe")
	void deveLancarParaAbaDeOpcoesInexistente() throws Exception {
		try (Planilha planilha = Planilha.nova("Formulario")) {
			IllegalArgumentException erro = assertThrows(IllegalArgumentException.class,
					() -> planilha.listaSuspensaDoIntervalo("A2:A10", "Opcoes", "F2:F4"));
			assertTrue(erro.getMessage().contains("Opcoes"));
		}
	}

	@Test
	@DisplayName("listaSuspensaDoIntervalo deve preservar referência já qualificada com aba")
	void devePreservarReferenciaComAba() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.novaAba("Opcoes");
			planilha.irParaAba("T");
			planilha.listaSuspensaDoIntervalo("A2:A10", "Opcoes!$A$1:$A$3");

			DataValidation validacao = validacaoUnica(planilha);
			assertEquals("Opcoes!$A$1:$A$3", validacao.getValidationConstraint().getFormula1());
		}
	}

	@Test
	@DisplayName("Deve lançar DadosInvalidosException sem opções")
	void deveLancarSemOpcoes() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(DadosInvalidosException.class, () -> planilha.listaSuspensa("A1:A5"));
		}
	}

	@Test
	@DisplayName("Deve lançar DadosInvalidosException quando as opções somam mais de 255 caracteres")
	void deveLancarListaMuitoLonga() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			String[] opcoesLongas = new String[10];
			for (int i = 0; i < opcoesLongas.length; i++) {
				opcoesLongas[i] = "Opcao-Bem-Longa-Numero-" + i + "-Para-Estourar-O-Limite-De-Duzentos-E-Cinquenta-E-Cinco";
			}
			assertThrows(DadosInvalidosException.class, () -> planilha.listaSuspensa("A1:A5", opcoesLongas));
		}
	}
}
