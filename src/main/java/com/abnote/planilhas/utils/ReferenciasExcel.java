package com.abnote.planilhas.utils;

import java.util.Locale;

/**
 * Monta referências de célula/intervalo no formato esperado por fórmulas do
 * Excel, sem espalhar detalhes de aspas e cifrões pela facade.
 */
public final class ReferenciasExcel {

	private ReferenciasExcel() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Converte uma referência simples em absoluta.
	 *
	 * <p>
	 * Exemplos: {@code A1 -> $A$1}, {@code A1:B5 -> $A$1:$B$5}. Se a referência
	 * já vier qualificada com aba (contém {@code !}), ela é preservada como foi
	 * informada para não mudar uma fórmula avançada digitada manualmente.
	 * </p>
	 *
	 * @param intervalo Célula ou intervalo.
	 * @return Referência absoluta, quando aplicável.
	 */
	public static String absoluta(final String intervalo) {
		if (intervalo.contains("!")) {
			return intervalo;
		}
		final String[] partes = intervalo.split(":", 2);
		final StringBuilder referencia = new StringBuilder();
		for (int i = 0; i < partes.length; i++) {
			if (i > 0) {
				referencia.append(':');
			}
			referencia.append(celulaAbsoluta(partes[i]));
		}
		return referencia.toString();
	}

	/**
	 * Qualifica um intervalo com o nome da aba e o converte para absoluto.
	 *
	 * @param nomeDaAba Nome da aba que contém o intervalo.
	 * @param intervalo Célula ou intervalo dentro da aba.
	 * @return Referência como {@code 'Minha Aba'!$A$1:$A$5}.
	 */
	public static String qualificada(final String nomeDaAba, final String intervalo) {
		return "'" + nomeDaAba.replace("'", "''") + "'!" + absoluta(intervalo);
	}

	private static String celulaAbsoluta(final String celula) {
		final String semCifro = celula.trim().replace("$", "");
		PosicaoConverter.converterPosicao(semCifro);
		final String coluna = semCifro.replaceAll("\\d", "").toUpperCase(Locale.ROOT);
		final String linha = semCifro.replaceAll("\\D", "");
		return "$" + coluna + "$" + linha;
	}
}
