package com.abnote.planilhas.calculos;

import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.*;
import com.abnote.planilhas.utils.LoggerUtil;
import com.abnote.planilhas.utils.PosicaoConverter;

public class Conversores {

	private static final Logger logger = LoggerUtil.getLogger(Conversores.class);

	/** Formato contábil brasileiro: símbolo à esquerda, valor alinhado à direita. */
	private static final String FORMATO_CONTABIL = "_-\"R$\" * #,##0.00_-;-\"R$\" * #,##0.00_-;_-\"R$\" * \"-\"??_-;_-@_-";

	/** Formato de moeda brasileira: "R$" grudado ao número. */
	private static final String FORMATO_MOEDA = "\"R$\" #,##0.00";

	/** Formato de texto puro do Excel. */
	private static final String FORMATO_TEXTO = "@";

	private Conversores() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Converte os valores de uma coluna para números, se possível.
	 *
	 * @param sheet          A folha da planilha a ser processada.
	 * @param posicaoInicial A posição inicial da coluna (ex: "J3").
	 */
	public static void converterEmNumero(Sheet sheet, String posicaoInicial) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell != null && cell.getCellType() == CellType.STRING) {
				tentarConverterCelulaEmNumero(cell, i);
			}
		}
	}

	/**
	 * Converte os valores de uma coluna para o formato contábil (R$).
	 *
	 * @param sheet          A folha da planilha a ser processada.
	 * @param posicaoInicial A posição inicial da coluna (ex: "J3").
	 * @param workbook       O workbook da planilha para criar estilos.
	 */
	public static void converterEmContabil(Sheet sheet, String posicaoInicial, Workbook workbook) {
		aplicarFormatoNumericoNaColuna(sheet, posicaoInicial, workbook, FORMATO_CONTABIL);
	}

	/**
	 * Converte os valores de uma coluna para o formato de moeda (R$ grudado ao valor).
	 *
	 * @param sheet          A folha da planilha a ser processada.
	 * @param posicaoInicial A posição inicial da coluna (ex: "J3").
	 * @param workbook       O workbook da planilha para criar estilos.
	 */
	public static void converterEmMoeda(Sheet sheet, String posicaoInicial, Workbook workbook) {
		aplicarFormatoNumericoNaColuna(sheet, posicaoInicial, workbook, FORMATO_MOEDA);
	}

	/**
	 * Converte os valores de uma coluna para texto puro, preservando a
	 * representação original (útil para manter zeros à esquerda de CEP/CPF).
	 *
	 * @param sheet          A folha da planilha a ser processada.
	 * @param posicaoInicial A posição inicial da coluna (ex: "J3").
	 * @param workbook       O workbook da planilha para criar estilos.
	 */
	public static void converterEmTexto(Sheet sheet, String posicaoInicial, Workbook workbook) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		CellStyle estiloTexto = workbook.createCellStyle();
		estiloTexto.setDataFormat(workbook.createDataFormat().getFormat(FORMATO_TEXTO));

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell == null) {
				continue;
			}
			if (cell.getCellType() == CellType.NUMERIC) {
				cell.setCellValue(formatarNumeroComoTexto(cell.getNumericCellValue()));
			}
			cell.setCellStyle(estiloTexto);
		}
	}

	private static void aplicarFormatoNumericoNaColuna(Sheet sheet, String posicaoInicial, Workbook workbook,
			String formato) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		CellStyle estiloNumerico = workbook.createCellStyle();
		estiloNumerico.setDataFormat(workbook.createDataFormat().getFormat(formato));

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell == null) {
				continue;
			}
			if (cell.getCellType() == CellType.STRING) {
				tentarConverterCelulaEmNumero(cell, i);
			}
			if (cell.getCellType() == CellType.NUMERIC) {
				cell.setCellStyle(estiloNumerico);
			}
		}
	}

	private static void tentarConverterCelulaEmNumero(Cell cell, int indiceLinha) {
		try {
			double valorNumerico = Double.parseDouble(cell.getStringCellValue());
			cell.setCellValue(valorNumerico);
		} catch (NumberFormatException e) {
			logger.fine("Célula em " + (indiceLinha + 1) + " não é numérica e foi ignorada.");
		}
	}

	private static String formatarNumeroComoTexto(double valor) {
		if (valor == Math.rint(valor) && !Double.isInfinite(valor)) {
			return Long.toString((long) valor);
		}
		return Double.toString(valor);
	}
}
