package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.exceptions.PosicaoInvalidaException;

/**
 * Testes de {@link PosicaoConverter}: conversão de posições e validação dos
 * limites do Excel .xlsx.
 */
@DisplayName("PosicaoConverter — conversões e limites do Excel")
class PosicaoConverterTest {

	@Test
	@DisplayName("Deve converter letras de coluna em índices 0-based")
	void deveConverterColunas() {
		assertEquals(0, PosicaoConverter.converterColuna("A"));
		assertEquals(25, PosicaoConverter.converterColuna("Z"));
		assertEquals(26, PosicaoConverter.converterColuna("AA"));
		assertEquals(701, PosicaoConverter.converterColuna("ZZ"));
		assertEquals(16383, PosicaoConverter.converterColuna("XFD"));
	}

	@Test
	@DisplayName("Deve aceitar coluna em minúsculas")
	void deveAceitarMinusculas() {
		assertEquals(0, PosicaoConverter.converterColuna("a"));
		assertArrayEquals(new int[] { 1, 1 }, PosicaoConverter.converterPosicao("b2"));
	}

	@Test
	@DisplayName("converterIndice deve ser inverso de converterColuna")
	void deveSerInverso() {
		assertEquals("A", PosicaoConverter.converterIndice(0));
		assertEquals("Z", PosicaoConverter.converterIndice(25));
		assertEquals("AA", PosicaoConverter.converterIndice(26));
		assertEquals("XFD", PosicaoConverter.converterIndice(16383));
	}

	@Test
	@DisplayName("Deve converter posição A1 em [coluna, linha] 0-based")
	void deveConverterPosicao() {
		assertArrayEquals(new int[] { 0, 0 }, PosicaoConverter.converterPosicao("A1"));
		assertArrayEquals(new int[] { 1, 1 }, PosicaoConverter.converterPosicao("B2"));
		assertArrayEquals(new int[] { 26, 9 }, PosicaoConverter.converterPosicao("AA10"));
	}

	@Test
	@DisplayName("Deve aceitar a última célula válida XFD1048576")
	void deveAceitarUltimaCelula() {
		assertArrayEquals(new int[] { 16383, 1048575 }, PosicaoConverter.converterPosicao("XFD1048576"));
	}

	@Test
	@DisplayName("Deve lançar PosicaoInvalidaException para entradas inválidas")
	void deveLancarParaEntradasInvalidas() {
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao(null));
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao(""));
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao("1A"));
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao("ABC"));
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao("A0"));
	}

	@Test
	@DisplayName("Deve lançar quando ultrapassa os limites do Excel")
	void deveLancarForaDosLimites() {
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterPosicao("A1048577"));
		assertThrows(PosicaoInvalidaException.class, () -> PosicaoConverter.converterColuna("XFE"));
	}

	@Test
	@DisplayName("Exceção deve carregar a posição inválida")
	void deveCarregarPosicaoInvalida() {
		PosicaoInvalidaException ex = assertThrows(PosicaoInvalidaException.class,
				() -> PosicaoConverter.converterPosicao("A0"));
		assertEquals("A0", ex.getPosicaoInvalida());
	}
}
