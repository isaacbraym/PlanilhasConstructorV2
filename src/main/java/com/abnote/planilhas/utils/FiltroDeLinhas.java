package com.abnote.planilhas.utils;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Localiza linhas de uma planilha cujo valor em uma coluna corresponde a um
 * critério. Base das operações de busca/filtro da facade.
 */
public final class FiltroDeLinhas {

	private FiltroDeLinhas() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Retorna os índices (0-based) das linhas cuja célula na coluna informada é
	 * igual ao valor procurado.
	 *
	 * @param sheet  A folha a pesquisar.
	 * @param coluna Índice da coluna (0-based).
	 * @param valor  Valor procurado (comparação por texto; inteiros sem ".0").
	 * @return Lista de índices de linha, em ordem crescente.
	 */
	public static List<Integer> encontrar(final Sheet sheet, final int coluna, final String valor) {
		final List<Integer> linhasEncontradas = new ArrayList<>();
		if (valor == null) {
			return linhasEncontradas;
		}
		final String alvo = valor.trim();
		for (int indiceLinha = 0; indiceLinha <= sheet.getLastRowNum(); indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			if (linha == null) {
				continue;
			}
			final Cell celula = linha.getCell(coluna);
			if (celula != null && valorComoTexto(celula).equals(alvo)) {
				linhasEncontradas.add(indiceLinha);
			}
		}
		return linhasEncontradas;
	}

	/**
	 * Representação textual de uma célula, com inteiros sem a parte decimal
	 * ({@code 10} em vez de {@code 10.0}) para facilitar a comparação.
	 *
	 * @param celula A célula.
	 * @return O texto equivalente ao conteúdo da célula.
	 */
	public static String valorComoTexto(final Cell celula) {
		switch (celula.getCellType()) {
		case STRING:
			return celula.getStringCellValue().trim();
		case NUMERIC:
			final double numero = celula.getNumericCellValue();
			if (numero == Math.rint(numero) && !Double.isInfinite(numero)) {
				return Long.toString((long) numero);
			}
			return Double.toString(numero);
		case BOOLEAN:
			return Boolean.toString(celula.getBooleanCellValue());
		case FORMULA:
			return celula.getCellFormula();
		default:
			return "";
		}
	}
}
