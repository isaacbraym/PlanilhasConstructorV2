package com.abnote.planilhas.estilos.estilos;

import java.awt.Color;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import com.abnote.planilhas.estilos.util.UtilEstiloFonte;

/**
 * Classe responsável por aplicar atributos de fonte em células, linhas ou
 * intervalos.
 */
public class Fontes {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public Fontes(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	public void aplicarAtributosFonte(FontAttributes attributes, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		iterateCells(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange,
				(Cell cell) -> applyFontAttributesToCell(cell, attributes));
	}

	public void aplicarItalico(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setItalic(true);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarSublinhado(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setUnderline(Font.U_SINGLE);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarTachado(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setStrikeout(true);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	private void iterateCells(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange, CellAction action) {
		if (isRange) {
			for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
				Row row = sheet.getRow(rowIdx);
				if (row == null)
					continue;
				for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
					Cell cell = row.getCell(colIdx);
					if (cell != null) {
						action.apply(cell);
					}
				}
			}
		} else if (rowIndex != -1) {
			if (columnIndex == -1) {
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					for (Cell cell : row) {
						if (cell != null) {
							action.apply(cell);
						}
					}
				}
			} else {
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					Cell cell = row.getCell(columnIndex);
					if (cell != null) {
						action.apply(cell);
					}
				}
			}
		}
	}

	@FunctionalInterface
	private interface CellAction {
		void apply(Cell cell);
	}

	private void applyFontAttributesToCell(Cell cell, FontAttributes attributes) {
		CellStyle currentStyle = cell.getCellStyle();
		String keyPrefix = "font_" + currentStyle.hashCode() + "_" + attributes.hashCode();
		UtilEstiloFonte.aplicarNovaFonte(cell, workbook, styleCache, keyPrefix, (Font currentFont) -> {
			Font newFont = workbook.createFont();
			copyFontAttributes(newFont, currentFont, attributes);
			return newFont;
		});
	}

	private void copyFontAttributes(Font newFont, Font currentFont, FontAttributes attributes) {
	    newFont.setFontName(attributes.getFontName() != null ? attributes.getFontName() : currentFont.getFontName());
	    newFont.setFontHeightInPoints(
	            attributes.getFontSize() != null ? attributes.getFontSize() : currentFont.getFontHeightInPoints());
	    newFont.setBold(attributes.isBold() != null ? attributes.isBold() : currentFont.getBold());
	    newFont.setItalic(attributes.isItalic() != null ? attributes.isItalic() : currentFont.getItalic());
	    newFont.setUnderline(
	            attributes.getUnderline() != null ? attributes.getUnderline() : currentFont.getUnderline());
	    newFont.setStrikeout(attributes.isStrikeout() != null ? attributes.isStrikeout() : currentFont.getStrikeout());
	    newFont.setCharSet(currentFont.getCharSet());
	    newFont.setTypeOffset(currentFont.getTypeOffset());
	    if (attributes.getColorRGB() != null) {
	        if (newFont instanceof XSSFFont) {
	            XSSFFont xssfFont = (XSSFFont) newFont;
	            // POI 5.x: Converter Color para byte[] 
	            Color cor = attributes.getColorRGB();
	            byte[] rgb = new byte[]{(byte) cor.getRed(), (byte) cor.getGreen(), (byte) cor.getBlue()};
	            XSSFColor xssfColor = new XSSFColor(rgb, null);
	            xssfFont.setColor(xssfColor);
	        } else if (newFont instanceof HSSFFont) {
	            short colorIndex = getNearestColorIndex(attributes.getColorRGB());
	            newFont.setColor(colorIndex);
	        }
	    } else {
	        newFont.setColor(currentFont.getColor());
	        if (currentFont instanceof XSSFFont && newFont instanceof XSSFFont) {
	            XSSFColor color = ((XSSFFont) currentFont).getXSSFColor();
	            ((XSSFFont) newFont).setColor(color);
	        }
	    }
	}

	private short getNearestColorIndex(Color color) {
		return IndexedColors.BLACK.getIndex();
	}

	public void aplicarCorFonte(CorEnum corEnum, int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes()
				.setColorRGB(new Color(corEnum.getRed(), corEnum.getGreen(), corEnum.getBlue()));
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarCorFonte(String hexColor, int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		Color color = hexToColor(hexColor);
		FontAttributes attributes = new FontAttributes().setColorRGB(color);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarCorFonte(int red, int green, int blue, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		Color color = new Color(red, green, blue);
		FontAttributes attributes = new FontAttributes().setColorRGB(color);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarFonte(String fontName, int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setFontName(fontName);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarFonte(FonteEnum fonteEnum, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		aplicarFonte(fonteEnum.getFontName(), rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarTamanhoFonte(int fontSize, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setFontSize((short) fontSize);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	private Color hexToColor(String hexColor) {
		if (hexColor == null || !hexColor.matches("^#([A-Fa-f0-9]{6})$")) {
			throw new IllegalArgumentException("Código hexadecimal de cor inválido: " + hexColor);
		}
		return Color.decode(hexColor);
	}
}
