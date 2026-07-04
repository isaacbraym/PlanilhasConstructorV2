package com.abnote.planilhas.estilos.estilos;

import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import com.abnote.planilhas.estilos.util.UtilEstiloFonte;

/**
 * Classe responsável por aplicar estilos de negrito em células, linhas ou
 * intervalos de uma planilha.
 */
public class BoldStyle {
    private final Workbook workbook;
    private final Sheet sheet;
    private final Map<String, CellStyle> styleCache;

    public BoldStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.styleCache = styleCache;
    }

    public void aplicarNegrito(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
            int endColumnIndex, boolean isRange) {
        if (isRange) {
            aplicarNegritoEmIntervalo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        } else if (rowIndex != -1) {
            if (columnIndex == -1) {
                aplicarNegritoEmLinha(rowIndex);
            } else {
                aplicarNegritoEmCelulaEspecifica(rowIndex, columnIndex);
            }
        }
    }

    private void aplicarNegritoEmIntervalo(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex) {
        for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null)
                continue;
            for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell == null)
                    continue;
                aplicarNegritoNaCelula(cell);
            }
        }
    }

    private void aplicarNegritoEmLinha(int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            for (Cell cell : row) {
                if (cell != null) {
                    aplicarNegritoNaCelula(cell);
                }
            }
        }
    }

    private void aplicarNegritoEmCelulaEspecifica(int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                aplicarNegritoNaCelula(cell);
            }
        }
    }

    private void aplicarNegritoNaCelula(Cell cell) {
        // Utiliza a classe utilitária para aplicar nova fonte com o atributo bold
        UtilEstiloFonte.aplicarNovaFonte(cell, workbook, styleCache, "bold_", 
            (Font fonteAtual) -> {
                Font novaFonte = workbook.createFont();
                copiarAtributosFonte(fonteAtual, novaFonte);
                novaFonte.setBold(true);
                return novaFonte;
            }
        );
    }

    private void copiarAtributosFonte(Font fonteOrigem, Font fonteDestino) {
        fonteDestino.setFontName(fonteOrigem.getFontName());
        fonteDestino.setFontHeightInPoints(fonteOrigem.getFontHeightInPoints());
        fonteDestino.setItalic(fonteOrigem.getItalic());
        fonteDestino.setStrikeout(fonteOrigem.getStrikeout());
        fonteDestino.setTypeOffset(fonteOrigem.getTypeOffset());
        fonteDestino.setUnderline(fonteOrigem.getUnderline());
        fonteDestino.setCharSet(fonteOrigem.getCharSet());
        fonteDestino.setColor(fonteOrigem.getColor());
        if (fonteOrigem instanceof XSSFFont && fonteDestino instanceof XSSFFont) {
            XSSFColor cor = ((XSSFFont) fonteOrigem).getXSSFColor();
            ((XSSFFont) fonteDestino).setColor(cor);
        }
    }
}
