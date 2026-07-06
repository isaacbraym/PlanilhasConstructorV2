package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Utilitarios para localizar linhas considerando conteudo real, sem confundir
 * linhas apenas formatadas com dados preenchidos.
 */
public final class LinhasDaPlanilha {

	private LinhasDaPlanilha() {
		// Classe utilitaria: nao deve ser instanciada.
	}

	/**
	 * Retorna o indice (0-based) da proxima linha apos a ultima linha com alguma
	 * celula preenchida. Linhas criadas so por altura/estilo nao contam.
	 *
	 * @param sheet A folha a inspecionar.
	 * @return Indice 0-based da proxima linha para append.
	 */
	public static int proximaLinhaAposUltimaPreenchida(final Sheet sheet) {
		for (int indiceLinha = sheet.getLastRowNum(); indiceLinha >= 0; indiceLinha--) {
			if (linhaTemConteudo(sheet.getRow(indiceLinha))) {
				return indiceLinha + 1;
			}
		}
		return 0;
	}

	private static boolean linhaTemConteudo(final Row linha) {
		if (linha == null || linha.getFirstCellNum() < 0) {
			return false;
		}
		for (int indiceColuna = linha.getFirstCellNum(); indiceColuna < linha.getLastCellNum(); indiceColuna++) {
			if (celulaTemConteudo(linha.getCell(indiceColuna))) {
				return true;
			}
		}
		return false;
	}

	private static boolean celulaTemConteudo(final Cell celula) {
		if (celula == null || celula.getCellType() == CellType.BLANK) {
			return false;
		}
		if (celula.getCellType() == CellType.STRING) {
			return !celula.getStringCellValue().isEmpty();
		}
		return true;
	}
}
