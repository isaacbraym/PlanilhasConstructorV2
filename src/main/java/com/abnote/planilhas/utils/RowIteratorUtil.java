package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.function.Consumer;

/**
 * Utilitário para iterar sobre as linhas de uma planilha.
 */
public class RowIteratorUtil {

    /**
     * Itera sobre as linhas da planilha no intervalo especificado e aplica uma ação.
     *
     * @param sheet       A planilha.
     * @param startRow    Índice da linha inicial (0-based).
     * @param endRow      Índice da linha final (0-based).
     * @param action      A ação a ser executada para cada linha não nula.
     */
    public static void forEachRow(Sheet sheet, int startRow, int endRow, Consumer<Row> action) {
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                action.accept(row);
            }
        }
    }
}
