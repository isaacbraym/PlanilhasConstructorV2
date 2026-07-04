package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@link PositionManager}: a máquina de estados de seleção.
 */
@DisplayName("PositionManager — estado de seleção")
class PositionManagerTest {

	private PositionManager positionManager;

	@BeforeEach
	void setUp() {
		positionManager = new PositionManager();
	}

	@Test
	@DisplayName("naCelula deve marcar posição definida e índices corretos")
	void deveDefinirCelula() {
		positionManager.naCelula("B3");

		assertTrue(positionManager.isPosicaoDefinida());
		assertFalse(positionManager.isIntervaloDefinida());
		assertEquals(1, positionManager.getPosicaoInicialColuna());
		assertEquals(2, positionManager.getPosicaoInicialLinha());
	}

	@Test
	@DisplayName("noIntervalo deve marcar intervalo definido com início e fim")
	void deveDefinirIntervalo() {
		positionManager.noIntervalo("B2", "E10");

		assertTrue(positionManager.isIntervaloDefinida());
		assertEquals(1, positionManager.getPosicaoInicialColuna());
		assertEquals(1, positionManager.getPosicaoInicialLinha());
		assertEquals(4, positionManager.getPosicaoFinalColuna());
		assertEquals(9, positionManager.getPosicaoFinalLinha());
	}

	@Test
	@DisplayName("emTodaAPlanilha deve marcar toda a planilha")
	void deveDefinirTodaPlanilha() {
		positionManager.emTodaAPlanilha();
		assertTrue(positionManager.isTodaPlanilhaDefinida());
	}

	@Test
	@DisplayName("resetarPosicao deve limpar todos os estados")
	void deveResetar() {
		positionManager.naCelula("A1");
		positionManager.noIntervalo("A1", "B2");
		positionManager.emTodaAPlanilha();

		positionManager.resetarPosicao();

		assertFalse(positionManager.isPosicaoDefinida());
		assertFalse(positionManager.isIntervaloDefinida());
		assertFalse(positionManager.isTodaPlanilhaDefinida());
	}
}
