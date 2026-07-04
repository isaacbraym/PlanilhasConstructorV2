package com.abnote.planilhas.estilos.estilos;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.util.CellIteratorUtil;

public class AlinhamentoStyle {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public AlinhamentoStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	// Método genérico para aplicar alinhamento e quebra de texto
	private void aplicarAlinhamento(HorizontalAlignment alignment, boolean quebraTexto, int rowIndex, int columnIndex,
			int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		CellIteratorUtil.forEachCell(sheet, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange, (Cell cell) -> applyAlignmentToCell(cell, alignment, quebraTexto));
	}

	public void alinharAEsquerda(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		aplicarAlinhamento(HorizontalAlignment.LEFT, false, rowIndex, columnIndex, startRowIndex, startColumnIndex,
				endRowIndex, endColumnIndex, isRange);
	}

	public void alinharADireita(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		aplicarAlinhamento(HorizontalAlignment.RIGHT, false, rowIndex, columnIndex, startRowIndex, startColumnIndex,
				endRowIndex, endColumnIndex, isRange);
	}

	public void quebrarTexto(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		aplicarAlinhamento(null, true, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	// Método para aplicar alinhamento e quebra de texto a uma célula
	private void applyAlignmentToCell(Cell cell, HorizontalAlignment alignment, boolean quebraTexto) {
		CellStyle currentStyle = cell.getCellStyle();
		String key = "alignment_" + currentStyle.hashCode() + "_" + (alignment != null ? alignment.name() : "null")
				+ "_" + quebraTexto;
		CellStyle newStyle = styleCache.get(key);
		if (newStyle == null) {
			newStyle = workbook.createCellStyle();
			newStyle.cloneStyleFrom(currentStyle);
			if (alignment != null) {
				newStyle.setAlignment(alignment);
			}
			newStyle.setWrapText(quebraTexto);
			styleCache.put(key, newStyle);
		}
		cell.setCellStyle(newStyle);
	}
}
