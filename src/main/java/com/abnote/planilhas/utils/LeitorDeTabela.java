package com.abnote.planilhas.utils;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Lê os dados de uma tabela (sem o cabeçalho) como uma lista de linhas,
 * detectando a largura pelo cabeçalho e a altura pela primeira coluna — a
 * mesma detecção usada por {@link TotalizadorDeTabela}.
 */
public final class LeitorDeTabela {

	private LeitorDeTabela() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Lê as linhas de dados da tabela (não inclui a linha de cabeçalho).
	 *
	 * @param sheet          A folha onde a tabela está.
	 * @param linhaCabecalho Índice (0-based) da linha de cabeçalho da tabela.
	 * @param colunaInicial  Índice (0-based) da primeira coluna da tabela.
	 * @return Lista de linhas; cada linha é uma lista de valores (ver
	 *         {@link LeitorDeCelulas#ler}). Lista vazia se não houver dados.
	 */
	public static List<List<Object>> ler(final Sheet sheet, final int linhaCabecalho, final int colunaInicial) {
		final int colunaFinal = TotalizadorDeTabela.ultimaColunaDoCabecalho(sheet, linhaCabecalho, colunaInicial);
		final int primeiraLinhaDeDados = linhaCabecalho + 1;
		final int ultimaLinhaDeDados = TotalizadorDeTabela.ultimaLinhaComDados(sheet, primeiraLinhaDeDados,
				colunaInicial);

		final FormulaEvaluator avaliador = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		final List<List<Object>> linhas = new ArrayList<>();
		for (int indiceLinha = primeiraLinhaDeDados; indiceLinha <= ultimaLinhaDeDados; indiceLinha++) {
			linhas.add(lerLinha(sheet.getRow(indiceLinha), colunaInicial, colunaFinal, avaliador));
		}
		return linhas;
	}

	private static List<Object> lerLinha(final Row linha, final int colunaInicial, final int colunaFinal,
			final FormulaEvaluator avaliador) {
		final List<Object> valores = new ArrayList<>();
		for (int coluna = colunaInicial; coluna <= colunaFinal; coluna++) {
			final Cell celula = linha == null ? null : linha.getCell(coluna);
			valores.add(LeitorDeCelulas.ler(celula, avaliador));
		}
		return valores;
	}
}
