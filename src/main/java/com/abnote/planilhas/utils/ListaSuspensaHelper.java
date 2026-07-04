package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Cria listas suspensas (validação de dados do tipo "Lista") em uma planilha.
 */
public final class ListaSuspensaHelper {

	private ListaSuspensaHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Cria uma lista suspensa com opções fixas (digitadas diretamente).
	 *
	 * @param sheet   A folha onde a validação será aplicada.
	 * @param regiao  Célula(s) que receberão a lista suspensa.
	 * @param opcoes  As opções que aparecerão no menu.
	 */
	public static void comOpcoesFixas(final XSSFSheet sheet, final CellRangeAddressList regiao,
			final String[] opcoes) {
		final XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
		final DataValidationConstraint restricao = helper.createExplicitListConstraint(opcoes);
		aplicar(sheet, helper, restricao, regiao);
	}

	/**
	 * Cria uma lista suspensa cujas opções vêm de um intervalo de células (ex.:
	 * uma coluna auxiliar com a lista de opções).
	 *
	 * @param sheet            A folha onde a validação será aplicada.
	 * @param regiao           Célula(s) que receberão a lista suspensa.
	 * @param formulaIntervalo Referência do intervalo com as opções (ex.:
	 *                         {@code "$D$2:$D$5"}).
	 */
	public static void doIntervalo(final XSSFSheet sheet, final CellRangeAddressList regiao,
			final String formulaIntervalo) {
		final XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
		final DataValidationConstraint restricao = helper.createFormulaListConstraint(formulaIntervalo);
		aplicar(sheet, helper, restricao, regiao);
	}

	private static void aplicar(final XSSFSheet sheet, final XSSFDataValidationHelper helper,
			final DataValidationConstraint restricao, final CellRangeAddressList regiao) {
		final DataValidation validacao = helper.createValidation(restricao, regiao);
		// Contraintuitivo, mas em XSSF é isto que EXIBE a seta do menu suspenso
		// (comportamento documentado do próprio Apache POI, não um bug).
		validacao.setSuppressDropDownArrow(true);
		validacao.setShowErrorBox(true);
		validacao.createErrorBox("Valor inválido", "Escolha um valor da lista.");
		sheet.addValidationData(validacao);
	}
}
