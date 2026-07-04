package com.abnote.planilhas.estilos.estilos;

import java.awt.Color;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class BackGroundColor {

	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public BackGroundColor(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	public void aplicarCorDeFundo(CorEnum corEnum, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		aplicarCorDeFundo(corEnum.getRed(), corEnum.getGreen(), corEnum.getBlue(), rowIndex, columnIndex, startRowIndex,
				startColumnIndex, endRowIndex, endColumnIndex, isRange);
	}

	public void aplicarCorDeFundo(String hexColor, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		Color color = hexToColor(hexColor);
		aplicarCorDeFundo(color.getRed(), color.getGreen(), color.getBlue(), rowIndex, columnIndex, startRowIndex,
				startColumnIndex, endRowIndex, endColumnIndex, isRange);
	}

	private Color hexToColor(String hexColor) {
		if (hexColor == null || !hexColor.matches("^#([A-Fa-f0-9]{6})$")) {
			throw new IllegalArgumentException("Código hexadecimal de cor inválido: " + hexColor);
		}
		return Color.decode(hexColor);
	}

	public void aplicarCorDeFundo(int red, int green, int blue, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		if (isRange) {
			for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
				Row row = sheet.getRow(rowIdx);
				if (row == null)
					continue;
				for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
					Cell cell = row.getCell(colIdx);
					if (cell == null)
						continue;
					applyBackgroundColorToCell(cell, red, green, blue);
				}
			}
		} else if (rowIndex != -1) {
			if (columnIndex == -1) {
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					for (Cell cell : row) {
						if (cell != null) {
							applyBackgroundColorToCell(cell, red, green, blue);
						}
					}
				}
			} else {
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					Cell cell = row.getCell(columnIndex);
					if (cell != null) {
						applyBackgroundColorToCell(cell, red, green, blue);
					}
				}
			}
		}
	}

	private void applyBackgroundColorToCell(Cell cell, int red, int green, int blue) {
	    CellStyle currentStyle = cell.getCellStyle();
	    CellStyle newStyle = workbook.createCellStyle();
	    newStyle.cloneStyleFrom(currentStyle);

	    if (newStyle instanceof XSSFCellStyle) {
	        XSSFCellStyle xssfCellStyle = (XSSFCellStyle) newStyle;
	        byte[] rgb = new byte[]{(byte) red, (byte) green, (byte) blue};
	        XSSFColor xssfColor = new XSSFColor(rgb, null);
	        xssfCellStyle.setFillForegroundColor(xssfColor);
	        xssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    } else if (newStyle instanceof HSSFCellStyle) {
	        short colorIndex = getNearestColorIndex(red, green, blue);
	        newStyle.setFillForegroundColor(colorIndex);
	        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    }
	    cell.setCellStyle(newStyle);
	}

	private short getNearestColorIndex(int red, int green, int blue) {
		return IndexedColors.AUTOMATIC.getIndex();
	}
}
