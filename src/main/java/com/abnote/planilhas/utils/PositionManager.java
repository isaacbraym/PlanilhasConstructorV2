package com.abnote.planilhas.utils;

/**
 * Gerencia as posições e intervalos de células para operações na planilha.
 */
public class PositionManager {
	private int posicaoInicialColuna = 0;
	private int posicaoInicialLinha = 0;
	private int posicaoFinalColuna = 0;
	private int posicaoFinalLinha = 0;
	private boolean intervaloDefinida = false;
	private boolean posicaoDefinida = false;
	private boolean todaPlanilhaDefinida = false;

	/**
	 * Define a célula inicial a partir da posição informada (ex: "B2").
	 *
	 * @param posicao Posição no formato alfanumérico.
	 */
	public void naCelula(String posicao) {
		int[] posicoes = PosicaoConverter.converterPosicao(posicao);
		this.posicaoInicialColuna = posicoes[0];
		this.posicaoInicialLinha = posicoes[1];
		this.posicaoDefinida = true;
	}

	/**
	 * Define um intervalo a partir das posições informadas.
	 *
	 * @param posicaoInicial Posição inicial (ex: "B2").
	 * @param posicaoFinal   Posição final (ex: "E10").
	 */
	public void noIntervalo(String posicaoInicial, String posicaoFinal) {
		int[] inicio = PosicaoConverter.converterPosicao(posicaoInicial);
		int[] fim = PosicaoConverter.converterPosicao(posicaoFinal);
		this.posicaoInicialColuna = inicio[0];
		this.posicaoInicialLinha = inicio[1];
		this.posicaoFinalColuna = fim[0];
		this.posicaoFinalLinha = fim[1];
		this.intervaloDefinida = true;
	}

	/**
	 * Indica que as operações serão aplicadas em toda a planilha.
	 */
	public void emTodaAPlanilha() {
		this.todaPlanilhaDefinida = true;
	}

	public boolean isTodaPlanilhaDefinida() {
		return todaPlanilhaDefinida;
	}

	/**
	 * Reseta todas as definições de posição e intervalo.
	 */
	public void resetarPosicao() {
		this.posicaoInicialColuna = 0;
		this.posicaoInicialLinha = 0;
		this.posicaoFinalColuna = 0;
		this.posicaoFinalLinha = 0;
		this.posicaoDefinida = false;
		this.intervaloDefinida = false;
		this.todaPlanilhaDefinida = false;
	}

	// Getters e setters
	public int getPosicaoInicialColuna() {
		return posicaoInicialColuna;
	}

	public void setPosicaoInicialColuna(int posicaoInicialColuna) {
		this.posicaoInicialColuna = posicaoInicialColuna;
	}

	public int getPosicaoInicialLinha() {
		return posicaoInicialLinha;
	}

	public void setPosicaoInicialLinha(int posicaoInicialLinha) {
		this.posicaoInicialLinha = posicaoInicialLinha;
	}

	public int getPosicaoFinalColuna() {
		return posicaoFinalColuna;
	}

	public int getPosicaoFinalLinha() {
		return posicaoFinalLinha;
	}

	public boolean isIntervaloDefinida() {
		return intervaloDefinida;
	}

	public boolean isPosicaoDefinida() {
		return posicaoDefinida;
	}
}
