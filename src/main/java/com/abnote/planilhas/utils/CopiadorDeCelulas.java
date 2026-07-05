package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Copia o valor e o estilo de uma célula para outra, célula a célula.
 *
 * <p>
 * Mais simples que {@link ManipuladorPlanilhaHelper}, que também copia células
 * mas como parte de um "recorte" temporário durante deslocamento de colunas.
 * Esta classe serve para cópia direta (duplicar coluna, duplicar linha, copiar
 * linha para outra aba), sem passar por um estado intermediário.
 * </p>
 */
public final class CopiadorDeCelulas {

	private CopiadorDeCelulas() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Copia o valor e o estilo da célula de origem para a célula de destino.
	 *
	 * @param origem  Célula com os dados a copiar.
	 * @param destino Célula que receberá os dados (deve já existir).
	 */
	public static void copiar(final Cell origem, final Cell destino) {
		destino.setCellStyle(origem.getCellStyle());
		switch (origem.getCellType()) {
		case STRING:
			destino.setCellValue(origem.getStringCellValue());
			break;
		case NUMERIC:
			destino.setCellValue(origem.getNumericCellValue());
			break;
		case BOOLEAN:
			destino.setCellValue(origem.getBooleanCellValue());
			break;
		case FORMULA:
			destino.setCellFormula(AjustadorDeFormulas.ajustarParaCopia(origem, destino));
			break;
		case BLANK:
			destino.setBlank();
			break;
		default:
			break;
		}
	}
}
