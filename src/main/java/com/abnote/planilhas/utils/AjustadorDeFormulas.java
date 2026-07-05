package com.abnote.planilhas.utils;

import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Ajusta fórmulas quando uma célula é copiada para outra posição.
 */
public final class AjustadorDeFormulas {

	private AjustadorDeFormulas() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Ajusta referências relativas da fórmula de {@code origem} para a posição de
	 * {@code destino}, como o Excel faz ao copiar uma célula com fórmula.
	 *
	 * @param origem  Célula de origem, do tipo {@code FORMULA}.
	 * @param destino Célula de destino.
	 * @return Fórmula ajustada, ou a fórmula original se o POI não conseguir
	 *         reescrever com segurança.
	 */
	public static String ajustarParaCopia(final Cell origem, final Cell destino) {
		return ajustar(origem.getSheet(), origem.getCellFormula(), origem.getRowIndex(), origem.getColumnIndex(),
				destino.getRowIndex(), destino.getColumnIndex());
	}

	/**
	 * Ajusta apenas a linha de uma fórmula ao reescrever uma célula na mesma coluna.
	 *
	 * @param sheet           Aba da fórmula.
	 * @param formulaOriginal Fórmula original.
	 * @param linhaOrigem     Linha original (0-based).
	 * @param linhaDestino    Linha de destino (0-based).
	 * @return Fórmula ajustada, ou a original em caso de fallback.
	 */
	public static String ajustarParaNovaLinha(final Sheet sheet, final String formulaOriginal, final int linhaOrigem,
			final int linhaDestino) {
		return ajustar(sheet, formulaOriginal, linhaOrigem, 0, linhaDestino, 0);
	}

	private static String ajustar(final Sheet sheet, final String formulaOriginal, final int linhaOrigem,
			final int colunaOrigem, final int linhaDestino, final int colunaDestino) {
		if (formulaOriginal == null || (linhaOrigem == linhaDestino && colunaOrigem == colunaDestino)) {
			return formulaOriginal;
		}
		try {
			final Workbook workbook = sheet.getWorkbook();
			final EvaluationWorkbook avaliacao = workbook.createEvaluationWorkbook();
			final FormulaParsingWorkbook parser = (FormulaParsingWorkbook) avaliacao;
			final FormulaRenderingWorkbook renderizador = (FormulaRenderingWorkbook) avaliacao;
			final int indiceAba = workbook.getSheetIndex(sheet);
			final Ptg[] tokens = FormulaParser.parse(formulaOriginal, parser, FormulaType.CELL, indiceAba,
					linhaOrigem);
			ajustarLinhas(sheet, workbook, indiceAba, tokens, linhaOrigem, linhaDestino);
			ajustarColunas(sheet, workbook, indiceAba, tokens, colunaOrigem, colunaDestino);
			return FormulaRenderer.toFormulaString(renderizador, tokens);
		} catch (RuntimeException e) {
			return formulaOriginal;
		}
	}

	private static void ajustarLinhas(final Sheet sheet, final Workbook workbook, final int indiceAba,
			final Ptg[] tokens, final int linhaOrigem, final int linhaDestino) {
		if (linhaOrigem == linhaDestino) {
			return;
		}
		FormulaShifter.createForRowCopy(indiceAba, sheet.getSheetName(), linhaOrigem, linhaOrigem,
				linhaDestino - linhaOrigem, workbook.getSpreadsheetVersion()).adjustFormula(tokens, indiceAba);
	}

	private static void ajustarColunas(final Sheet sheet, final Workbook workbook, final int indiceAba,
			final Ptg[] tokens, final int colunaOrigem, final int colunaDestino) {
		if (colunaOrigem == colunaDestino) {
			return;
		}
		FormulaShifter.createForColumnCopy(indiceAba, sheet.getSheetName(), colunaOrigem, colunaOrigem,
				colunaDestino - colunaOrigem, workbook.getSpreadsheetVersion()).adjustFormula(tokens, indiceAba);
	}
}
