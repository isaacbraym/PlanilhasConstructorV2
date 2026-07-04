package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 * Lê o valor de uma célula como um tipo Java natural (o "inverso" da coerção
 * feita por {@link InsersorDeDados} ao escrever): número, texto, verdadeiro/
 * falso, data/hora, ou {@code null} para célula vazia/inexistente.
 */
public final class LeitorDeCelulas {

	private LeitorDeCelulas() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Lê o valor de uma célula, avaliando a fórmula se for o caso.
	 *
	 * @param celula    A célula a ler (pode ser {@code null}).
	 * @param avaliador Avaliador de fórmulas do workbook da célula.
	 * @return {@link Double}, {@link String}, {@link Boolean},
	 *         {@link java.time.LocalDateTime}, ou {@code null}.
	 */
	public static Object ler(final Cell celula, final FormulaEvaluator avaliador) {
		if (celula == null) {
			return null;
		}
		switch (celula.getCellType()) {
		case STRING:
			return celula.getStringCellValue();
		case NUMERIC:
			return DateUtil.isCellDateFormatted(celula) ? celula.getLocalDateTimeCellValue()
					: celula.getNumericCellValue();
		case BOOLEAN:
			return celula.getBooleanCellValue();
		case FORMULA:
			return lerValorCalculado(celula, avaliador.evaluate(celula));
		default:
			return null; // BLANK, ERROR
		}
	}

	/**
	 * Lê o valor de uma célula formatado exatamente como aparece no Excel (ex.:
	 * uma célula de moeda vira {@code "R$ 1.234,56"}, uma data vira
	 * {@code "04/07/2026"}).
	 *
	 * @param celula    A célula a ler (pode ser {@code null}).
	 * @param avaliador Avaliador de fórmulas do workbook da célula.
	 * @return O texto formatado, ou {@code ""} se a célula não existir.
	 */
	public static String comoTexto(final Cell celula, final FormulaEvaluator avaliador) {
		if (celula == null) {
			return "";
		}
		return new DataFormatter().formatCellValue(celula, avaliador);
	}

	private static Object lerValorCalculado(final Cell celulaFormula, final CellValue valorCalculado) {
		switch (valorCalculado.getCellType()) {
		case NUMERIC:
			return DateUtil.isCellDateFormatted(celulaFormula)
					? DateUtil.getLocalDateTime(valorCalculado.getNumberValue())
					: valorCalculado.getNumberValue();
		case STRING:
			return valorCalculado.getStringValue();
		case BOOLEAN:
			return valorCalculado.getBooleanValue();
		default:
			return null; // BLANK, ERROR
		}
	}
}
