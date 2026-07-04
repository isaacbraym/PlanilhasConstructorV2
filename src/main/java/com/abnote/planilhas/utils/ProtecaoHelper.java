package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Desbloqueia células dentro de uma planilha protegida (por padrão, o Excel
 * trata toda célula como "travada" — só surte efeito quando a planilha é
 * protegida com {@link Sheet#protectSheet(String)}).
 */
public final class ProtecaoHelper {

	private ProtecaoHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Marca as células de um intervalo como destravadas (editáveis mesmo com a
	 * planilha protegida). Cria a célula se ainda não existir.
	 *
	 * <p>
	 * Clona o {@link CellStyle} de cada célula antes de destravar — nunca muta o
	 * estilo original, que costuma ser compartilhado por muitas células (mutar
	 * em vez de clonar destravaria células sem relação nenhuma com o intervalo).
	 * </p>
	 *
	 * @param sheet  A folha onde o intervalo está.
	 * @param regiao O intervalo a destravar.
	 */
	public static void desbloquearIntervalo(final Sheet sheet, final CellRangeAddress regiao) {
		final Workbook workbook = sheet.getWorkbook();
		for (int indiceLinha = regiao.getFirstRow(); indiceLinha <= regiao.getLastRow(); indiceLinha++) {
			Row linha = sheet.getRow(indiceLinha);
			if (linha == null) {
				linha = sheet.createRow(indiceLinha);
			}
			for (int indiceColuna = regiao.getFirstColumn(); indiceColuna <= regiao.getLastColumn(); indiceColuna++) {
				final Cell celula = linha.getCell(indiceColuna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				final CellStyle estiloDestravado = workbook.createCellStyle();
				estiloDestravado.cloneStyleFrom(celula.getCellStyle());
				estiloDestravado.setLocked(false);
				celula.setCellStyle(estiloDestravado);
			}
		}
	}
}
