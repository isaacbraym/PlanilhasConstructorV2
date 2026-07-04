package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Adiciona comentários (notas) a células — o balãozinho amarelo que aparece
 * ao passar o mouse sobre a célula no Excel.
 */
public final class ComentarioHelper {

	/** Tamanho padrão do balão de comentário, em colunas/linhas ocupadas. */
	private static final int LARGURA_PADRAO_COLUNAS = 3;
	private static final int ALTURA_PADRAO_LINHAS = 4;

	private ComentarioHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Adiciona (ou substitui) o comentário de uma célula.
	 *
	 * @param sheet  A folha onde a célula está.
	 * @param linha  Índice (0-based) da linha da célula.
	 * @param coluna Índice (0-based) da coluna da célula.
	 * @param texto  Texto do comentário.
	 */
	public static void adicionar(final XSSFSheet sheet, final int linha, final int coluna, final String texto) {
		final XSSFDrawing desenho = sheet.createDrawingPatriarch();
		final XSSFClientAnchor ancora = desenho.createAnchor(0, 0, 0, 0, coluna, linha, coluna + LARGURA_PADRAO_COLUNAS,
				linha + ALTURA_PADRAO_LINHAS);
		final Comment comentario = desenho.createCellComment(ancora);
		comentario.setString(sheet.getWorkbook().getCreationHelper().createRichTextString(texto));

		Row row = sheet.getRow(linha);
		if (row == null) {
			row = sheet.createRow(linha);
		}
		final Cell celula = row.getCell(coluna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		celula.setCellComment(comentario);
	}
}
