package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Detecta os limites de uma tabela a partir da célula de cabeçalho e insere
 * uma linha de totais: soma (fórmula {@code SUM}) em cada coluna numérica, e o
 * texto "Total" na primeira coluna não numérica encontrada.
 */
public final class TotalizadorDeTabela {

	private TotalizadorDeTabela() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Insere a linha de totais logo após a última linha de dados da tabela.
	 *
	 * @param sheet          A folha onde a tabela está.
	 * @param linhaCabecalho Índice (0-based) da linha de cabeçalho da tabela.
	 * @param colunaInicial  Índice (0-based) da primeira coluna da tabela.
	 */
	public static void adicionarTotais(final Sheet sheet, final int linhaCabecalho, final int colunaInicial) {
		final int colunaFinal = ultimaColunaDoCabecalho(sheet, linhaCabecalho, colunaInicial);
		final int primeiraLinhaDeDados = linhaCabecalho + 1;
		final int ultimaLinhaDeDados = ultimaLinhaComDados(sheet, primeiraLinhaDeDados, colunaInicial);
		if (ultimaLinhaDeDados < primeiraLinhaDeDados) {
			return; // Tabela sem nenhuma linha de dados: nada a totalizar.
		}

		final Row linhaTotais = obterOuCriarLinha(sheet, ultimaLinhaDeDados + 1);
		boolean rotuloTotalEscrito = false;
		for (int coluna = colunaInicial; coluna <= colunaFinal; coluna++) {
			if (colunaEhNumerica(sheet, primeiraLinhaDeDados, ultimaLinhaDeDados, coluna)) {
				escreverSoma(linhaTotais, coluna, primeiraLinhaDeDados, ultimaLinhaDeDados);
			} else if (!rotuloTotalEscrito) {
				linhaTotais.createCell(coluna).setCellValue("Total");
				rotuloTotalEscrito = true;
			}
		}
	}

	private static void escreverSoma(final Row linhaTotais, final int coluna, final int primeiraLinhaDeDados,
			final int ultimaLinhaDeDados) {
		final String colunaLetra = PosicaoConverter.converterIndice(coluna);
		final String formula = String.format("SUM(%s%d:%s%d)", colunaLetra, primeiraLinhaDeDados + 1, colunaLetra,
				ultimaLinhaDeDados + 1);
		linhaTotais.createCell(coluna).setCellFormula(formula);
	}

	private static int ultimaColunaDoCabecalho(final Sheet sheet, final int linhaCabecalho, final int colunaInicial) {
		final Row cabecalho = sheet.getRow(linhaCabecalho);
		if (cabecalho == null) {
			return colunaInicial;
		}
		int coluna = colunaInicial;
		while (celulaTemConteudo(cabecalho.getCell(coluna))) {
			coluna++;
		}
		return coluna - 1;
	}

	private static int ultimaLinhaComDados(final Sheet sheet, final int primeiraLinhaDeDados,
			final int colunaReferencia) {
		int ultimaLinha = primeiraLinhaDeDados - 1;
		for (int indiceLinha = primeiraLinhaDeDados; indiceLinha <= sheet.getLastRowNum(); indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			if (linha == null || !celulaTemConteudo(linha.getCell(colunaReferencia))) {
				break;
			}
			ultimaLinha = indiceLinha;
		}
		return ultimaLinha;
	}

	private static boolean colunaEhNumerica(final Sheet sheet, final int primeiraLinha, final int ultimaLinha,
			final int coluna) {
		boolean encontrouAlgumaCelula = false;
		for (int indiceLinha = primeiraLinha; indiceLinha <= ultimaLinha; indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			final Cell celula = linha == null ? null : linha.getCell(coluna);
			if (celula == null || celula.getCellType() == CellType.BLANK) {
				continue;
			}
			if (celula.getCellType() != CellType.NUMERIC) {
				return false;
			}
			encontrouAlgumaCelula = true;
		}
		return encontrouAlgumaCelula;
	}

	private static boolean celulaTemConteudo(final Cell celula) {
		return celula != null && celula.getCellType() != CellType.BLANK;
	}

	private static Row obterOuCriarLinha(final Sheet sheet, final int indiceLinha) {
		final Row linha = sheet.getRow(indiceLinha);
		return linha != null ? linha : sheet.createRow(indiceLinha);
	}
}
