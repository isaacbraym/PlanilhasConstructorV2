package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.exceptions.DadosInvalidosException;

/**
 * Testes de ocultar/exibir linha, coluna e aba na facade.
 */
@DisplayName("Planilha — ocultar/exibir linha, coluna e aba")
class OcultarExibirFacadeTest {

	@Test
	@DisplayName("ocultarLinha e exibirLinha devem alternar a visibilidade da linha")
	void deveOcultarEExibirLinha() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A2", "conteúdo").ocultarLinha(2);
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.getRow(1).getZeroHeight());

			planilha.exibirLinha(2);
			assertFalse(sheet.getRow(1).getZeroHeight());
		}
	}

	@Test
	@DisplayName("ocultarColuna e exibirColuna devem alternar a visibilidade da coluna")
	void deveOcultarEExibirColuna() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.ocultarColuna("C");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.isColumnHidden(2));

			planilha.exibirColuna("C");
			assertFalse(sheet.isColumnHidden(2));
		}
	}

	@Test
	@DisplayName("ocultarAba e exibirAba devem alternar a visibilidade da aba, quando há outra visível")
	void deveOcultarEExibirAba() throws Exception {
		try (Planilha planilha = Planilha.nova("Principal")) {
			planilha.novaAba("Auxiliar").ocultarAba("Auxiliar");
			assertTrue(planilha.workbook().isSheetHidden(planilha.workbook().getSheetIndex("Auxiliar")));

			planilha.exibirAba("Auxiliar");
			assertFalse(planilha.workbook().isSheetHidden(planilha.workbook().getSheetIndex("Auxiliar")));
		}
	}

	@Test
	@DisplayName("ocultarAba deve recusar ocultar a única aba visível do arquivo")
	void deveRecusarOcultarUnicaAbaVisivel() throws Exception {
		try (Planilha planilha = Planilha.nova("Sozinha")) {
			assertThrows(DadosInvalidosException.class, () -> planilha.ocultarAba("Sozinha"));
			assertFalse(planilha.workbook().isSheetHidden(0), "Não deve ter sido ocultada");
		}
	}

	@Test
	@DisplayName("ocultarAba/exibirAba com nome inexistente devem lançar IllegalArgumentException")
	void deveLancarParaAbaInexistente() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(IllegalArgumentException.class, () -> planilha.ocultarAba("NaoExiste"));
			assertThrows(IllegalArgumentException.class, () -> planilha.exibirAba("NaoExiste"));
		}
	}
}
