package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting.IconSet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFPatternFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;

/**
 * Aplica regras de formatação condicional (realce de células) em uma planilha.
 *
 * <p>
 * Assume sempre {@link XSSFSheet} (planilhas {@code .xlsx}), já que esta
 * biblioteca só cria/edita workbooks XSSF — a mesma premissa já usada em
 * outras partes do projeto (ex.: cores via {@link XSSFColor}).
 * </p>
 */
public final class FormatacaoCondicionalHelper {

	private FormatacaoCondicionalHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Pinta de fundo as células de {@code regioes} cujo valor é maior que
	 * {@code valor}.
	 */
	public static void realcarSeMaiorQue(final XSSFSheet sheet, final CellRangeAddress[] regioes, final double valor,
			final int red, final int green, final int blue) {
		aplicarComparacao(sheet, regioes, ComparisonOperator.GT, String.valueOf(valor), red, green, blue);
	}

	/**
	 * Pinta de fundo as células de {@code regioes} cujo valor é menor que
	 * {@code valor}.
	 */
	public static void realcarSeMenorQue(final XSSFSheet sheet, final CellRangeAddress[] regioes, final double valor,
			final int red, final int green, final int blue) {
		aplicarComparacao(sheet, regioes, ComparisonOperator.LT, String.valueOf(valor), red, green, blue);
	}

	/**
	 * Pinta de fundo as células de {@code regioes} cujo valor está entre
	 * {@code minimo} e {@code maximo} (inclusive).
	 */
	public static void realcarSeEntre(final XSSFSheet sheet, final CellRangeAddress[] regioes, final double minimo,
			final double maximo, final int red, final int green, final int blue) {
		final XSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
		final XSSFConditionalFormattingRule regra = scf.createConditionalFormattingRule(ComparisonOperator.BETWEEN,
				String.valueOf(minimo), String.valueOf(maximo));
		colorirFundo(regra, red, green, blue);
		scf.addConditionalFormatting(regioes, regra);
	}

	/**
	 * Pinta de fundo as células de {@code regioes} cujo valor é igual a
	 * {@code valor} (número ou texto).
	 */
	public static void realcarSeIgual(final XSSFSheet sheet, final CellRangeAddress[] regioes, final Object valor,
			final int red, final int green, final int blue) {
		aplicarComparacao(sheet, regioes, ComparisonOperator.EQUAL, formatarValor(valor), red, green, blue);
	}

	/**
	 * Aplica uma escala de 3 cores (vermelho → amarelo → verde, o "semáforo"
	 * clássico do Excel) sobre {@code regioes}, com o valor mínimo em vermelho, a
	 * mediana em amarelo e o máximo em verde.
	 */
	public static void aplicarEscalaDeCores(final XSSFSheet sheet, final CellRangeAddress[] regioes) {
		final XSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
		final XSSFConditionalFormattingRule regra = scf.createConditionalFormattingColorScaleRule();
		final ColorScaleFormatting escala = regra.getColorScaleFormatting();

		escala.setNumControlPoints(3);
		escala.setColors(new Color[] { corRgb(248, 105, 107), corRgb(255, 235, 132), corRgb(99, 190, 123) });

		final ConditionalFormattingThreshold[] limiares = escala.getThresholds();
		limiares[0].setRangeType(RangeType.MIN);
		limiares[1].setRangeType(RangeType.PERCENTILE);
		limiares[1].setValue(50d);
		limiares[2].setRangeType(RangeType.MAX);
		escala.setThresholds(limiares);

		scf.addConditionalFormatting(regioes, regra);
	}

	/**
	 * Aplica barras de dados no intervalo, usando a cor informada para representar
	 * visualmente a proporção entre o menor e o maior valor.
	 */
	public static void aplicarBarrasDeDados(final XSSFSheet sheet, final CellRangeAddress[] regioes, final int red,
			final int green, final int blue) {
		final XSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
		final XSSFConditionalFormattingRule regra = scf.createConditionalFormattingRule(corRgb(red, green, blue));
		scf.addConditionalFormatting(regioes, regra);
	}

	/**
	 * Aplica o conjunto de ícones de semáforo (verde/amarelo/vermelho) com os
	 * limiares percentuais padrão do Excel.
	 */
	public static void aplicarIconesSemaforo(final XSSFSheet sheet, final CellRangeAddress[] regioes) {
		final XSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
		final XSSFConditionalFormattingRule regra = scf.createConditionalFormattingRule(IconSet.GYR_3_TRAFFIC_LIGHTS);
		scf.addConditionalFormatting(regioes, regra);
	}

	private static void aplicarComparacao(final XSSFSheet sheet, final CellRangeAddress[] regioes, final byte operador,
			final String formula, final int red, final int green, final int blue) {
		final XSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
		final XSSFConditionalFormattingRule regra = scf.createConditionalFormattingRule(operador, formula);
		colorirFundo(regra, red, green, blue);
		scf.addConditionalFormatting(regioes, regra);
	}

	private static void colorirFundo(final XSSFConditionalFormattingRule regra, final int red, final int green,
			final int blue) {
		final XSSFPatternFormatting fundo = (XSSFPatternFormatting) regra.createPatternFormatting();
		fundo.setFillForegroundColor(corRgb(red, green, blue));
		fundo.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
	}

	private static XSSFColor corRgb(final int red, final int green, final int blue) {
		return new XSSFColor(new byte[] { (byte) red, (byte) green, (byte) blue }, null);
	}

	/**
	 * Formata um valor para uso em fórmula de condição: números ficam como estão,
	 * texto é envolvido em aspas duplas (sintaxe de fórmula do Excel).
	 */
	private static String formatarValor(final Object valor) {
		if (valor instanceof Number) {
			return valor.toString();
		}
		return "\"" + valor + "\"";
	}
}
