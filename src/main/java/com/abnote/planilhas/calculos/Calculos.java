package com.abnote.planilhas.calculos;

import org.apache.poi.ss.usermodel.*;
import com.abnote.planilhas.utils.PosicaoConverter;

public class Calculos {

	/**
	 * Soma os valores numéricos de uma coluna específica e insere a soma sem texto.
	 * A célula da soma mantém a mesma formatação das células somadas.
	 *
	 * @param sheet          A folha da planilha onde a soma será realizada.
	 * @param posicaoInicial A posição inicial da coluna a ser somada (ex: "J3").
	 */
	public static void somarColuna(Sheet sheet, String posicaoInicial) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		double soma = 0.0;
		CellStyle estiloSoma = null;
		Workbook workbook = sheet.getWorkbook();
		int ultimaLinhaDados = -1;

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell != null && cell.getCellType() == CellType.NUMERIC) {
				soma += cell.getNumericCellValue();
				if (estiloSoma == null) {
					estiloSoma = cell.getCellStyle();
				}
				ultimaLinhaDados = i;
			}
		}

		int linhaSoma = ultimaLinhaDados + 1;
		Row rowSoma = sheet.getRow(linhaSoma);
		if (rowSoma == null) {
			rowSoma = sheet.createRow(linhaSoma);
		}
		Cell cellSoma = rowSoma.getCell(coluna);
		if (cellSoma == null) {
			cellSoma = rowSoma.createCell(coluna);
		}
		cellSoma.setCellValue(soma);

		if (estiloSoma != null) {
			CellStyle somaStyle = workbook.createCellStyle();
			somaStyle.cloneStyleFrom(estiloSoma);
			cellSoma.setCellStyle(somaStyle);
		} else {
			CellStyle defaultStyle = workbook.createCellStyle();
			defaultStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
			cellSoma.setCellStyle(defaultStyle);
		}
	}

	/**
	 * Soma os valores numéricos de uma coluna específica e insere a soma com um
	 * texto descritivo. A célula da soma mantém a mesma formatação das células
	 * somadas.
	 *
	 * @param sheet          A folha da planilha onde a soma será realizada.
	 * @param posicaoInicial A posição inicial da coluna a ser somada (ex: "J3").
	 * @param texto          O texto descritivo que será inserido ao lado da soma.
	 */
	public static void somarColunaComTexto(Sheet sheet, String posicaoInicial, String texto) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		double soma = 0.0;
		int ultimaLinha = linhaInicial;
		CellStyle estiloSoma = null;
		Workbook workbook = sheet.getWorkbook();

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell != null && cell.getCellType() == CellType.NUMERIC) {
				soma += cell.getNumericCellValue();
				if (estiloSoma == null) {
					estiloSoma = cell.getCellStyle();
				}
			}
			ultimaLinha = i;
		}

		Row linhaSoma = sheet.createRow(ultimaLinha + 1);
		Cell cellTexto = linhaSoma.createCell(coluna - 1);
		cellTexto.setCellValue(texto);
		Cell cellSoma = linhaSoma.createCell(coluna);
		cellSoma.setCellValue(soma);

		if (estiloSoma != null) {
			CellStyle somaStyle = workbook.createCellStyle();
			somaStyle.cloneStyleFrom(estiloSoma);
			cellSoma.setCellStyle(somaStyle);
		} else {
			CellStyle defaultStyle = workbook.createCellStyle();
			defaultStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
			cellSoma.setCellStyle(defaultStyle);
		}
	}

	/**
	 * Multiplica os valores de duas colunas específicas e insere o resultado com um
	 * texto descritivo e a soma total dos resultados.
	 *
	 * @param sheet         A folha da planilha onde a operação será realizada.
	 * @param coluna1       A primeira coluna (ex: "D").
	 * @param coluna2       A segunda coluna (ex: "I").
	 * @param linhaInicial  A linha inicial para a operação.
	 * @param texto         O texto descritivo a ser inserido na linha de soma.
	 * @param colunaDestino A coluna onde o resultado será inserido (ex: "J").
	 */
	public static void multiplicarColunasComTexto(Sheet sheet, String coluna1, String coluna2, int linhaInicial,
			String texto, String colunaDestino) {
		int colunaIndex1 = PosicaoConverter.converterColuna(coluna1);
		int colunaIndex2 = PosicaoConverter.converterColuna(coluna2);
		int colunaDestinoIndex = PosicaoConverter.converterColuna(colunaDestino);

		int rowIndexInicial = linhaInicial - 1; // Ajuste para índice 0-based
		int ultimaLinhaDados = rowIndexInicial - 1;
		CellStyle estiloMultiplicacao = null;
		Workbook workbook = sheet.getWorkbook();
		DataFormat dataFormat = workbook.createDataFormat();

		for (int i = rowIndexInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell1 = row.getCell(colunaIndex1);
			Cell cell2 = row.getCell(colunaIndex2);
			if (cell1 != null && cell2 != null) {
				Cell cellDestino = row.getCell(colunaDestinoIndex);
				if (cellDestino == null) {
					cellDestino = row.createCell(colunaDestinoIndex);
				}
				String formula = coluna1 + (i + 1) + "*" + coluna2 + (i + 1);
				cellDestino.setCellFormula(formula);

				short dataFormatIndex = -1;
				if (cell1.getCellStyle() != null && cell1.getCellStyle().getDataFormat() != 0) {
					dataFormatIndex = cell1.getCellStyle().getDataFormat();
				} else if (cell2.getCellStyle() != null && cell2.getCellStyle().getDataFormat() != 0) {
					dataFormatIndex = cell2.getCellStyle().getDataFormat();
				}
				if (dataFormatIndex != -1) {
					CellStyle novoEstilo = workbook.createCellStyle();
					novoEstilo.setDataFormat(dataFormatIndex);
					cellDestino.setCellStyle(novoEstilo);
				} else {
					CellStyle defaultStyle = workbook.createCellStyle();
					defaultStyle.setDataFormat(dataFormat.getFormat("0.00"));
					cellDestino.setCellStyle(defaultStyle);
				}

				if (estiloMultiplicacao == null) {
					estiloMultiplicacao = cellDestino.getCellStyle();
				}
				ultimaLinhaDados = i;
			}
		}

		if (ultimaLinhaDados >= rowIndexInicial) {
			int linhaSoma = ultimaLinhaDados + 1;
			Row rowSoma = sheet.createRow(linhaSoma);
			Cell cellTexto = rowSoma.createCell(colunaDestinoIndex - 1);
			cellTexto.setCellValue(texto);
			Cell cellSoma = rowSoma.createCell(colunaDestinoIndex);
			String formulaSoma = "SUM(" + colunaDestino + linhaInicial + ":" + colunaDestino + (ultimaLinhaDados + 1)
					+ ")";
			cellSoma.setCellFormula(formulaSoma);
			if (estiloMultiplicacao != null) {
				CellStyle somaStyle = workbook.createCellStyle();
				somaStyle.setDataFormat(estiloMultiplicacao.getDataFormat());
				cellSoma.setCellStyle(somaStyle);
			} else {
				CellStyle defaultStyle = workbook.createCellStyle();
				defaultStyle.setDataFormat(dataFormat.getFormat("0.00"));
				cellSoma.setCellStyle(defaultStyle);
			}
		}
	}
}
