package com.abnote.planilhas.estilos.estilos;

import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.utils.PosicaoConverter;



/**
 * Classe responsável por aplicar estilos de borda em células, intervalos ou
 * posições específicas de uma planilha.
 */
public class BorderStyleHelper {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	/**
	 * Construtor para inicializar o BorderStyleHelper com um Workbook, Sheet e
	 * cache de estilos.
	 *
	 * @param workbook   O Workbook que contém a planilha.
	 * @param sheet      A Sheet onde os estilos serão aplicados.
	 * @param styleCache Cache para armazenar estilos já criados.
	 */
	public BorderStyleHelper(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	/**
	 * Aplica todas as bordas finas em um intervalo específico ou em toda a
	 * planilha.
	 *
	 * @param indiceInicioLinha  Índice da primeira linha do intervalo.
	 * @param indiceInicioColuna Índice da primeira coluna do intervalo.
	 * @param indiceFimLinha     Índice da última linha do intervalo.
	 * @param indiceFimColuna    Índice da última coluna do intervalo.
	 * @param isRange            Se verdadeiro, aplica em um intervalo; caso
	 *                           contrário, em toda a planilha.
	 */
	public void aplicarTodasAsBordas(int indiceInicioLinha, int indiceInicioColuna, int indiceFimLinha,
			int indiceFimColuna, boolean isRange) {
		if (isRange) {
			aplicarTodasAsBordasEmIntervalo(indiceInicioLinha, indiceInicioColuna, indiceFimLinha, indiceFimColuna);
		} else {
			aplicarTodasAsBordasNaPlanilha();
		}
	}

	// Método privado para aplicar todas as bordas em um intervalo específico de
	// células
	private void aplicarTodasAsBordasEmIntervalo(int indiceInicioLinha, int indiceInicioColuna, int indiceFimLinha,
			int indiceFimColuna) {
		for (int rowIdx = indiceInicioLinha; rowIdx <= indiceFimLinha; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null)
				continue;

			for (int colIdx = indiceInicioColuna; colIdx <= indiceFimColuna; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null || cell.getCellType() == CellType.BLANK) {
					continue; // Pular células vazias
				}
				aplicarBordasNaCelula(cell);
			}
		}
	}

	// Método privado para aplicar todas as bordas em toda a planilha
	private void aplicarTodasAsBordasNaPlanilha() {
		for (Row row : sheet) {
			if (row == null)
				continue;
			for (Cell cell : row) {
				if (cell == null || cell.getCellType() == CellType.BLANK) {
					continue; // Pular células vazias
				}
				aplicarBordasNaCelula(cell);
			}
		}
	}

	// Método privado para aplicar bordas finas em uma célula específica
	private void aplicarBordasNaCelula(Cell cell) {
		CellStyle estiloAtual = cell.getCellStyle();

		boolean possuiBordaEspessa = verificarBordaEspessa(estiloAtual);
		if (possuiBordaEspessa) {
			return; // Ignorar células com bordas espessas
		}

		String chaveCache = "borders_" + estiloAtual.hashCode();
		CellStyle estiloComBordas = styleCache.get(chaveCache);
		if (estiloComBordas == null) {
			estiloComBordas = criarEstiloComBordas(estiloAtual);
			styleCache.put(chaveCache, estiloComBordas);
		}

		cell.setCellStyle(estiloComBordas);
	}

	// Método privado para verificar se o estilo atual possui bordas espessas
	private boolean verificarBordaEspessa(CellStyle estilo) {
		return estilo.getBorderTop() == BorderStyle.NONE && estilo.getBorderBottom() == BorderStyle.NONE &&
				estilo.getBorderLeft() == BorderStyle.NONE && estilo.getBorderRight() == BorderStyle.NONE;
	}

	// Método privado para criar um novo estilo com bordas finas
	private CellStyle criarEstiloComBordas(CellStyle estiloOriginal) {
		CellStyle novoEstilo = workbook.createCellStyle();
		novoEstilo.cloneStyleFrom(estiloOriginal);
		novoEstilo.setBorderTop(BorderStyle.THIN);
		novoEstilo.setBorderBottom(BorderStyle.THIN);
		novoEstilo.setBorderLeft(BorderStyle.THIN);
		novoEstilo.setBorderRight(BorderStyle.THIN);
		return novoEstilo;
	}

	/**
	 * Aplica bordas finas em uma célula específica baseada na posição (e.g., "A1").
	 *
	 * @param posicao A posição da célula (ex: "A1").
	 */
	public void aplicarBordasNaCelula(String posicao) {
		int[] indicesPosicao = PosicaoConverter.converterPosicao(posicao);
		int coluna = indicesPosicao[0];
		int linha = indicesPosicao[1];
		Row row = sheet.getRow(linha);
		if (row == null) {
			row = sheet.createRow(linha);
		}
		Cell cell = row.getCell(coluna);
		if (cell == null) {
			cell = row.createCell(coluna);
		}
		aplicarBordasNaCelula(cell);
	}

	/**
	 * Aplica bordas finas entre duas posições específicas (e.g., "A1" até "C3").
	 *
	 * @param posicaoInicial A posição inicial (ex: "A1").
	 * @param posicaoFinal   A posição final (ex: "C3").
	 */
	public void aplicarBordasEntre(String posicaoInicial, String posicaoFinal) {
		int[] indicesInicio = PosicaoConverter.converterPosicao(posicaoInicial);
		int[] indicesFim = PosicaoConverter.converterPosicao(posicaoFinal);

		for (int rowIdx = indicesInicio[1]; rowIdx <= indicesFim[1]; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null) {
				row = sheet.createRow(rowIdx);
			}
			for (int colIdx = indicesInicio[0]; colIdx <= indicesFim[0]; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null) {
					cell = row.createCell(colIdx);
				}
				aplicarBordasNaCelula(cell);
			}
		}
	}

	/**
	 * Aplica bordas espessas nas bordas externas de um intervalo específico (e.g.,
	 * "A1" até "C3").
	 *
	 * @param posicaoInicial A posição inicial (ex: "A1").
	 * @param posicaoFinal   A posição final (ex: "C3").
	 */
	public void aplicarBordasEspessas(String posicaoInicial, String posicaoFinal) {
		int[] indicesInicio = PosicaoConverter.converterPosicao(posicaoInicial);
		int[] indicesFim = PosicaoConverter.converterPosicao(posicaoFinal);

		for (int rowIdx = indicesInicio[1]; rowIdx <= indicesFim[1]; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null) {
				row = sheet.createRow(rowIdx);
			}
			for (int colIdx = indicesInicio[0]; colIdx <= indicesFim[0]; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null) {
					cell = row.createCell(colIdx);
				}

				CellStyle estiloAtual = cell.getCellStyle();
				String chaveCache = "thickBorders_" + estiloAtual.hashCode() + "_r" + rowIdx + "_c" + colIdx;
				CellStyle estiloComBordasEspessas = styleCache.get(chaveCache);
				if (estiloComBordasEspessas == null) {
					estiloComBordasEspessas = criarEstiloComBordasEspessas(estiloAtual, rowIdx, colIdx, indicesInicio,
							indicesFim);
					styleCache.put(chaveCache, estiloComBordasEspessas);
				}

				cell.setCellStyle(estiloComBordasEspessas);
			}
		}
	}

	/**
	 * Aplica bordas espessas nas bordas internas e externas de um intervalo
	 * específico (e.g., "A1" até "C3").
	 *
	 * @param posicaoInicial A posição inicial (ex: "A1").
	 * @param posicaoFinal   A posição final (ex: "C3").
	 */
	public void aplicarBordasEspessasComInternas(String posicaoInicial, String posicaoFinal) {
		int[] indicesInicio = PosicaoConverter.converterPosicao(posicaoInicial);
		int[] indicesFim = PosicaoConverter.converterPosicao(posicaoFinal);

		for (int rowIdx = indicesInicio[1]; rowIdx <= indicesFim[1]; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null) {
				row = sheet.createRow(rowIdx);
			}
			for (int colIdx = indicesInicio[0]; colIdx <= indicesFim[0]; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null) {
					cell = row.createCell(colIdx);
				}

				CellStyle estiloAtual = cell.getCellStyle();
				StringBuilder chaveCacheBuilder = new StringBuilder("thickInternalBorders_");
				chaveCacheBuilder.append(estiloAtual.hashCode()).append("_r").append(rowIdx).append("_c")
						.append(colIdx);
				String chaveCache = chaveCacheBuilder.toString();

				CellStyle novoEstilo = styleCache.get(chaveCache);
				if (novoEstilo == null) {
					novoEstilo = criarEstiloComBordasEspessasComInternas(estiloAtual, rowIdx, colIdx, indicesInicio,
							indicesFim);
					styleCache.put(chaveCache, novoEstilo);
				}

				cell.setCellStyle(novoEstilo);
			}
		}
	}

	// Método privado para criar um estilo com bordas espessas nas bordas externas
	private CellStyle criarEstiloComBordasEspessas(CellStyle estiloOriginal, int rowIdx, int colIdx,
			int[] indicesInicio, int[] indicesFim) {
		CellStyle novoEstilo = workbook.createCellStyle();
		novoEstilo.cloneStyleFrom(estiloOriginal);

		// Aplicar bordas espessas nas bordas externas
		if (rowIdx == indicesInicio[1]) { // Primeira linha do intervalo
			novoEstilo.setBorderTop(BorderStyle.THICK);
		}
		if (rowIdx == indicesFim[1]) { // Última linha do intervalo
			novoEstilo.setBorderBottom(BorderStyle.THICK);
		}
		if (colIdx == indicesInicio[0]) { // Primeira coluna do intervalo
			novoEstilo.setBorderLeft(BorderStyle.THICK);
		}
		if (colIdx == indicesFim[0]) { // Última coluna do intervalo
			novoEstilo.setBorderRight(BorderStyle.THICK);
		}

		return novoEstilo;
	}

	// Método privado para criar um estilo com bordas espessas nas bordas externas e
	// finas nas internas
	private CellStyle criarEstiloComBordasEspessasComInternas(CellStyle estiloOriginal, int rowIdx, int colIdx,
			int[] indicesInicio, int[] indicesFim) {
		CellStyle novoEstilo = workbook.createCellStyle();
		novoEstilo.cloneStyleFrom(estiloOriginal);

		// Aplicar bordas finas em todas as direções
		novoEstilo.setBorderTop(BorderStyle.THIN);
		novoEstilo.setBorderBottom(BorderStyle.THIN);
		novoEstilo.setBorderLeft(BorderStyle.THIN);
		novoEstilo.setBorderRight(BorderStyle.THIN);

		// Aplicar bordas espessas nas bordas externas
		if (rowIdx == indicesInicio[1]) { // Primeira linha do intervalo
			novoEstilo.setBorderTop(BorderStyle.THICK);
		}
		if (rowIdx == indicesFim[1]) { // Última linha do intervalo
			novoEstilo.setBorderBottom(BorderStyle.THICK);
		}
		if (colIdx == indicesInicio[0]) { // Primeira coluna do intervalo
			novoEstilo.setBorderLeft(BorderStyle.THICK);
		}
		if (colIdx == indicesFim[0]) { // Última coluna do intervalo
			novoEstilo.setBorderRight(BorderStyle.THICK);
		}

		return novoEstilo;
	}
}
