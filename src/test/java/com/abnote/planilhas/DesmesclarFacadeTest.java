package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@code desmesclar} na facade.
 */
@DisplayName("Planilha — desmesclar")
class DesmesclarFacadeTest {

	@Test
	@DisplayName("Deve desfazer a mesclagem de um intervalo previamente mesclado")
	void deveDesmesclar() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", "Título").mesclar("A1:C1");
			assertEquals(1, planilha.workbook().getSheetAt(0).getNumMergedRegions());

			planilha.desmesclar("A1:C1");
			assertEquals(0, planilha.workbook().getSheetAt(0).getNumMergedRegions());
		}
	}

	@Test
	@DisplayName("Desmesclar um intervalo não mesclado não deve lançar exceção nem afetar outras mesclagens")
	void deveTolerarIntervaloNaoMesclado() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.mesclar("A1:B1");
			assertDoesNotThrow(() -> planilha.desmesclar("D1:E1"));
			assertEquals(1, planilha.workbook().getSheetAt(0).getNumMergedRegions(), "A1:B1 deve continuar mesclado");
		}
	}
}
