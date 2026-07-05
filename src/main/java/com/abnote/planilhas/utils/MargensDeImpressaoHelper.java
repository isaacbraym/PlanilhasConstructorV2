package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.Sheet;

import com.abnote.planilhas.exceptions.DadosInvalidosException;

/**
 * Aplica margens de impressão usando centímetros na API pública e polegadas no
 * Apache POI, que é a unidade nativa gravada no arquivo Excel.
 */
public final class MargensDeImpressaoHelper {

	private static final double CENTIMETROS_POR_POLEGADA = 2.54;

	private MargensDeImpressaoHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Define as quatro margens principais da aba.
	 *
	 * @param sheet      Aba que receberá as margens.
	 * @param superiorCm Margem superior em centímetros.
	 * @param inferiorCm Margem inferior em centímetros.
	 * @param esquerdaCm Margem esquerda em centímetros.
	 * @param direitaCm  Margem direita em centímetros.
	 */
	public static void aplicar(final Sheet sheet, final double superiorCm, final double inferiorCm,
			final double esquerdaCm, final double direitaCm) {
		validarMargem("superior", superiorCm);
		validarMargem("inferior", inferiorCm);
		validarMargem("esquerda", esquerdaCm);
		validarMargem("direita", direitaCm);

		sheet.setMargin(PageMargin.TOP, paraPolegadas(superiorCm));
		sheet.setMargin(PageMargin.BOTTOM, paraPolegadas(inferiorCm));
		sheet.setMargin(PageMargin.LEFT, paraPolegadas(esquerdaCm));
		sheet.setMargin(PageMargin.RIGHT, paraPolegadas(direitaCm));
	}

	private static void validarMargem(final String nomeDaMargem, final double margemCm) {
		if (Double.isNaN(margemCm) || Double.isInfinite(margemCm) || margemCm < 0) {
			throw new DadosInvalidosException(
					"Margem de impressão deve ser um número em centímetros maior ou igual a zero",
					nomeDaMargem + "=" + margemCm);
		}
	}

	private static double paraPolegadas(final double centimetros) {
		return centimetros / CENTIMETROS_POR_POLEGADA;
	}
}
