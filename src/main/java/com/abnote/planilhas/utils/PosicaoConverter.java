package com.abnote.planilhas.utils;

import com.abnote.planilhas.exceptions.PosicaoInvalidaException;

/**
 * Classe utilitária para conversão entre a notação de posição (por exemplo,
 * "B2") e índices numéricos (base 0).
 *
 * <p>
 * Valida os limites de uma planilha {@code .xlsx} (Excel 2007+): no máximo
 * {@value #MAX_LINHAS} linhas e {@value #MAX_COLUNAS} colunas (até "XFD").
 * </p>
 */
public final class PosicaoConverter {

	/** Máximo de linhas em uma planilha {@code .xlsx}. */
	public static final int MAX_LINHAS = 1_048_576;

	/** Máximo de colunas em uma planilha {@code .xlsx} (coluna "XFD"). */
	public static final int MAX_COLUNAS = 16_384;

	private PosicaoConverter() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Converte uma posição no formato "B2" em um array de inteiros onde o primeiro
	 * elemento é o índice da coluna (0-based) e o segundo é o índice da linha
	 * (0-based).
	 *
	 * @param posicao A posição no formato alfanumérico.
	 * @return Array de inteiros: [coluna, linha].
	 * @throws PosicaoInvalidaException se a posição for nula, vazia, mal formada ou
	 *                                  ultrapassar os limites do Excel.
	 */
	public static int[] converterPosicao(final String posicao) {
		if (posicao == null || posicao.trim().isEmpty()) {
			throw new PosicaoInvalidaException("Posição não pode ser nula ou vazia. Exemplo válido: 'A1', 'B10'",
					posicao);
		}
		final String limpo = posicao.trim().toUpperCase();
		if (!limpo.matches("^[A-Z]+[0-9]+$")) {
			throw new PosicaoInvalidaException("Posição inválida: '" + posicao
					+ "'. Use a(s) letra(s) da coluna seguidas do número da linha, ex.: 'A1', 'AB12'.", posicao);
		}
		final String colunaParte = limpo.replaceAll("\\d", "");
		final String linhaParte = limpo.replaceAll("\\D", "");
		final int coluna = converterColuna(colunaParte);
		final int linha = validarLinha(linhaParte, posicao);
		return new int[] { coluna, linha };
	}

	/**
	 * Converte uma letra ou conjunto de letras representando a coluna em um índice
	 * numérico (0-based).
	 *
	 * @param coluna A(s) letra(s) que representam a coluna (por exemplo, "A", "B",
	 *               "AA").
	 * @return Índice numérico da coluna.
	 * @throws PosicaoInvalidaException se a coluna for nula, vazia, contiver
	 *                                  caracteres inválidos ou ultrapassar "XFD".
	 */
	public static int converterColuna(final String coluna) {
		if (coluna == null || coluna.trim().isEmpty()) {
			throw new PosicaoInvalidaException("Coluna não pode ser nula ou vazia. Exemplo válido: 'A', 'B', 'AA'",
					coluna);
		}
		final String limpo = coluna.trim().toUpperCase();
		if (!limpo.matches("^[A-Z]+$")) {
			throw new PosicaoInvalidaException("Coluna inválida: '" + coluna + "'. Use apenas letras, ex.: 'A', 'AB'.",
					coluna);
		}
		long numero = 0;
		for (int i = 0; i < limpo.length(); i++) {
			numero = numero * 26 + (limpo.charAt(i) - ('A' - 1));
		}
		final long indice = numero - 1; // Ajusta para índice 0-based
		if (indice < 0 || indice >= MAX_COLUNAS) {
			throw new PosicaoInvalidaException(
					"Coluna '" + limpo + "' ultrapassa o limite do Excel (máximo: 'XFD', " + MAX_COLUNAS + " colunas).",
					coluna);
		}
		return (int) indice;
	}

	/**
	 * Converte um índice numérico de coluna (0-based) em sua representação
	 * alfabética (por exemplo, 0 -> "A").
	 *
	 * @param index Índice numérico da coluna.
	 * @return Representação alfabética da coluna.
	 */
	public static String converterIndice(int index) {
		StringBuilder result = new StringBuilder();
		index += 1; // Ajusta para 1-based
		while (index > 0) {
			int remainder = (index - 1) % 26;
			result.insert(0, (char) (remainder + 'A'));
			index = (index - 1) / 26;
		}
		return result.toString();
	}

	private static int validarLinha(final String linhaParte, final String posicaoOriginal) {
		final long linha1Based;
		try {
			linha1Based = Long.parseLong(linhaParte);
		} catch (NumberFormatException e) {
			throw new PosicaoInvalidaException("Linha inválida em '" + posicaoOriginal + "'.", posicaoOriginal);
		}
		if (linha1Based < 1 || linha1Based > MAX_LINHAS) {
			throw new PosicaoInvalidaException(
					"Linha " + linha1Based + " fora do intervalo do Excel (1 a " + MAX_LINHAS + ").", posicaoOriginal);
		}
		return (int) (linha1Based - 1);
	}
}
