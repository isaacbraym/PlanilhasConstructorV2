package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Substitui fórmulas por seus valores calculados ("colar como valores" do
 * Excel) — útil antes de compartilhar um arquivo, para congelar os números e
 * impedir que fórmulas quebrem se a planilha for editada fora de contexto.
 */
public final class ColarComoValoresHelper {

	private ColarComoValoresHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Calcula e grava o valor de cada célula com fórmula dentro do intervalo,
	 * removendo a fórmula (a célula passa a ter um valor fixo).
	 *
	 * @param sheet  A folha onde o intervalo está.
	 * @param regiao O intervalo a processar.
	 */
	public static void aplicar(final Sheet sheet, final CellRangeAddress regiao) {
		final Workbook workbook = sheet.getWorkbook();
		final FormulaEvaluator avaliador = workbook.getCreationHelper().createFormulaEvaluator();

		for (int indiceLinha = regiao.getFirstRow(); indiceLinha <= regiao.getLastRow(); indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			if (linha == null) {
				continue;
			}
			for (int indiceColuna = regiao.getFirstColumn(); indiceColuna <= regiao.getLastColumn(); indiceColuna++) {
				final Cell celula = linha.getCell(indiceColuna);
				if (celula == null || celula.getCellType() != CellType.FORMULA) {
					continue;
				}
				substituirPeloValor(celula, avaliador.evaluate(celula));
			}
		}
	}

	private static void substituirPeloValor(final Cell celula, final CellValue valorCalculado) {
		// removeFormula() é necessário: setCellValue() sozinho NÃO limpa a fórmula
		// de uma célula do tipo FORMULA (só atualiza o valor em cache, mantendo o
		// tipo FORMULA) — comportamento verificado empiricamente no POI 5.2.5.
		celula.removeFormula();
		switch (valorCalculado.getCellType()) {
		case NUMERIC:
			celula.setCellValue(valorCalculado.getNumberValue());
			break;
		case STRING:
			celula.setCellValue(valorCalculado.getStringValue());
			break;
		case BOOLEAN:
			celula.setCellValue(valorCalculado.getBooleanValue());
			break;
		case ERROR:
			celula.setCellErrorValue(valorCalculado.getErrorValue());
			break;
		default:
			celula.setBlank();
			break;
		}
	}
}
