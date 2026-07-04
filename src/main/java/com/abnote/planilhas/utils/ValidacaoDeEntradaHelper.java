package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Restringe o que pode ser digitado em um intervalo de células a um número ou
 * data dentro de limites — diferente de {@link ListaSuspensaHelper}, aqui não
 * há um menu de opções, só um limite mínimo/máximo aceito.
 */
public final class ValidacaoDeEntradaHelper {

	private ValidacaoDeEntradaHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Restringe o intervalo a números decimais entre dois limites (inclusive).
	 *
	 * @param sheet  A folha onde a validação será aplicada.
	 * @param regiao Célula(s) restringidas.
	 * @param minimo Valor mínimo aceito.
	 * @param maximo Valor máximo aceito.
	 */
	public static void numeroEntre(final XSSFSheet sheet, final CellRangeAddressList regiao, final double minimo,
			final double maximo) {
		final XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
		final DataValidationConstraint restricao = helper.createDecimalConstraint(OperatorType.BETWEEN,
				String.valueOf(minimo), String.valueOf(maximo));
		aplicar(sheet, helper, restricao, regiao);
	}

	/**
	 * Restringe o intervalo a números inteiros entre dois limites (inclusive).
	 *
	 * @param sheet  A folha onde a validação será aplicada.
	 * @param regiao Célula(s) restringidas.
	 * @param minimo Valor mínimo aceito.
	 * @param maximo Valor máximo aceito.
	 */
	public static void inteiroEntre(final XSSFSheet sheet, final CellRangeAddressList regiao, final int minimo,
			final int maximo) {
		final XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
		final DataValidationConstraint restricao = helper.createIntegerConstraint(OperatorType.BETWEEN,
				String.valueOf(minimo), String.valueOf(maximo));
		aplicar(sheet, helper, restricao, regiao);
	}

	/**
	 * Restringe o intervalo a datas entre dois limites (inclusive).
	 *
	 * @param sheet         A folha onde a validação será aplicada.
	 * @param regiao        Célula(s) restringidas.
	 * @param dataMinimaISO Data mínima aceita, no formato {@code yyyy-MM-dd}.
	 * @param dataMaximaISO Data máxima aceita, no formato {@code yyyy-MM-dd}.
	 */
	public static void dataEntre(final XSSFSheet sheet, final CellRangeAddressList regiao, final String dataMinimaISO,
			final String dataMaximaISO) {
		final XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
		final String formatoData = "yyyy-MM-dd";
		final DataValidationConstraint restricao = helper.createDateConstraint(OperatorType.BETWEEN, dataMinimaISO,
				dataMaximaISO, formatoData);
		aplicar(sheet, helper, restricao, regiao);
	}

	private static void aplicar(final XSSFSheet sheet, final XSSFDataValidationHelper helper,
			final DataValidationConstraint restricao, final CellRangeAddressList regiao) {
		final DataValidation validacao = helper.createValidation(restricao, regiao);
		validacao.setShowErrorBox(true);
		validacao.createErrorBox("Valor fora do permitido", "Digite um valor dentro do limite definido.");
		sheet.addValidationData(validacao);
	}
}
