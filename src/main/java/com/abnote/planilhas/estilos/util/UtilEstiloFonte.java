package com.abnote.planilhas.estilos.util;

import java.util.Map;
import java.util.function.Function;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Classe utilitária para aplicar uma nova fonte a uma célula. Centraliza a
 * lógica de clonagem do estilo atual, aplicação do modificador e uso de cache.
 */
public class UtilEstiloFonte {

	/**
	 * Aplica uma nova fonte a uma célula utilizando um modificador de fonte.
	 *
	 * @param cell         A célula a ser estilizada.
	 * @param workbook     O workbook da planilha.
	 * @param cache        Cache para armazenar estilos já criados.
	 * @param keyPrefix    Prefixo da chave para o cache.
	 * @param fontModifier Função que recebe a fonte atual e retorna a nova fonte desejada.
	 */
	public static void aplicarNovaFonte(Cell cell, Workbook workbook, Map<String, CellStyle> cache, String keyPrefix,
			Function<Font, Font> fontModifier) {
		CellStyle currentStyle = cell.getCellStyle();
		String key = keyPrefix + currentStyle.hashCode();
		CellStyle newStyle = cache.get(key);
		if (newStyle == null) {
			newStyle = workbook.createCellStyle();
			newStyle.cloneStyleFrom(currentStyle);
			Font currentFont = workbook.getFontAt(currentStyle.getFontIndex());
			Font newFont = fontModifier.apply(currentFont);
			newStyle.setFont(newFont);
			cache.put(key, newStyle);
		}
		cell.setCellStyle(newStyle);
	}
}
