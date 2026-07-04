package com.abnote.planilhas.estilos.estilos;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class CenterStyle {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public CenterStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	/**
	 * Centraliza as células em um intervalo específico ou em toda a planilha.
	 *
	 * @param startRow    Índice da primeira linha
	 * @param startColumn Índice da primeira coluna
	 * @param endRow      Índice da última linha
	 * @param endColumn   Índice da última coluna
	 * @param isRange     Se verdadeiro, aplica ao intervalo; caso contrário, a toda
	 *                    a planilha
	 */
	public void centralizarTudo(int startRow, int startColumn, int endRow, int endColumn, boolean isRange) {
		if (isRange) {
			centralizarIntervalo(startRow, startColumn, endRow, endColumn);
		} else {
			centralizarPlanilha();
		}
	}

	public void centralizarERedimensionarTudo() {
		centralizarPlanilha();
		redimensionarColuna();
	}

	private void centralizarIntervalo(int startRow, int startColumn, int endRow, int endColumn) {
		for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null)
				continue;

			for (int colIdx = startColumn; colIdx <= endColumn; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null)
					continue;

				aplicarCentralizacao(cell);
			}
		}
	}

	// Método privado para centralizar todas as células da planilha
	private void centralizarPlanilha() {
		for (Row row : sheet) {
			if (row == null)
				continue;

			for (Cell cell : row) {
				if (cell == null)
					continue;

				aplicarCentralizacao(cell);
			}
		}
	}

	// Método privado para aplicar a centralização a uma célula específica
	private void aplicarCentralizacao(Cell cell) {
		CellStyle originalStyle = cell.getCellStyle();

		// Verificar o alinhamento atual da célula
		HorizontalAlignment currentAlignment = originalStyle.getAlignment();
		if (currentAlignment == HorizontalAlignment.GENERAL || currentAlignment == null) {
			// Apenas centralizar se o alinhamento for 'GERAL' ou nulo
			CellStyle novoEstilo = criarEstiloCentralizado(originalStyle);
			cell.setCellStyle(novoEstilo);
		}
		// Caso contrário, não faz nada para não sobrescrever alinhamentos já definidos
	}

	// Método privado para criar um novo estilo com centralização
	private CellStyle criarEstiloCentralizado(CellStyle originalStyle) {
		// Verificar se já existe um estilo centralizado no cache
		String styleKey = "centralizado_" + originalStyle.hashCode();
		CellStyle novoEstilo = styleCache.get(styleKey);

		if (novoEstilo == null) {
			novoEstilo = workbook.createCellStyle();
			novoEstilo.cloneStyleFrom(originalStyle);
			novoEstilo.setAlignment(HorizontalAlignment.CENTER);
			novoEstilo.setVerticalAlignment(VerticalAlignment.CENTER);
			styleCache.put(styleKey, novoEstilo);
		}

		return novoEstilo;
	}

	public void redimensionarColuna() {
		// Avalia todas as fórmulas na planilha para que o autoSizeColumn funcione
		// corretamente
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		for (Row row : sheet) {
			if (row == null)
				continue;
			for (Cell cell : row) {
				if (cell == null)
					continue;
				if (cell.getCellType() == CellType.FORMULA) {
					evaluator.evaluateFormulaCell(cell);
				}
			}
		}

		int maxColumns = obterMaximoNumeroDeColunas();
		for (int i = 0; i < maxColumns; i++) {
			sheet.autoSizeColumn(i);
		}
	}

	// Método privado para determinar o número máximo de colunas na planilha
//    private int obterMaximoNumeroDeColunas() {
//        int maxColumns = 0;
//        for (Row row : sheet) {
//            if (row == null)
//                continue;
//
//            int lastCellNum = row.getLastCellNum();
//            if (lastCellNum > maxColumns) {
//                maxColumns = lastCellNum;
//            }
//        }
//        return maxColumns;
//    }
	private int obterMaximoNumeroDeColunas() {
		int maxColumns = 0;
		for (Row row : sheet) {
			if (row == null)
				continue;

			int lastCellNum = row.getLastCellNum();
			if (lastCellNum > maxColumns) {
				maxColumns = lastCellNum;
			}
		}
		return maxColumns;
	}
}
