package com.abnote.planilhas.estilos.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.util.function.Consumer;

/**
 * Utilitário para iterar sobre células em uma planilha.
 */
public class CellIteratorUtil {

    /**
     * Itera sobre as células com base nos parâmetros especificados e aplica a ação.
     *
     * @param sheet            A planilha.
     * @param rowIndex         Índice da linha se for célula única ou linha inteira (se columnIndex == -1).
     * @param columnIndex      Índice da coluna se for célula específica.
     * @param startRowIndex    Índice inicial do intervalo.
     * @param startColumnIndex Índice inicial da coluna do intervalo.
     * @param endRowIndex      Índice final do intervalo.
     * @param endColumnIndex   Índice final do intervalo.
     * @param isRange          Se verdadeiro, itera no intervalo; se falso, itera apenas na célula ou linha.
     * @param action           A ação a ser executada para cada célula.
     */
    public static void forEachCell(Sheet sheet,
                                   int rowIndex,
                                   int columnIndex,
                                   int startRowIndex,
                                   int startColumnIndex,
                                   int endRowIndex,
                                   int endColumnIndex,
                                   boolean isRange,
                                   Consumer<Cell> action) {
        if (isRange) {
            for (int r = startRowIndex; r <= endRowIndex; r++) {
                Row row = sheet.getRow(r);
                if (row == null)
                    continue;
                for (int c = startColumnIndex; c <= endColumnIndex; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null) {
                        action.accept(cell);
                    }
                }
            }
        } else if (rowIndex != -1) {
            if (columnIndex == -1) { // Aplica à linha inteira
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    for (Cell cell : row) {
                        if (cell != null)
                            action.accept(cell);
                    }
                }
            } else { // Aplica à célula específica
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null)
                        action.accept(cell);
                }
            }
        }
    }
}
