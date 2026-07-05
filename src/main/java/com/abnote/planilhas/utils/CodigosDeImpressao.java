package com.abnote.planilhas.utils;

/**
 * Traduz marcadores amigáveis (ex.: {@code "{pagina}"}) para os códigos de
 * cabeçalho/rodapé de impressão do Excel (ex.: {@code "&P"}) — a sintaxe
 * nativa do Excel é pouco descobrível para quem não a conhece de cor.
 */
public final class CodigosDeImpressao {

	private CodigosDeImpressao() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Substitui os marcadores amigáveis pelos códigos do Excel:
	 * {@code {pagina}} → número da página, {@code {total}} → total de páginas,
	 * {@code {data}} → data, {@code {hora}} → hora, {@code {arquivo}} → nome do
	 * arquivo, {@code {aba}} → nome da aba.
	 *
	 * @param texto Texto com marcadores amigáveis (pode ser {@code null}).
	 * @return O texto com os marcadores substituídos, ou {@code null} se a
	 *         entrada for {@code null}.
	 */
	public static String traduzir(final String texto) {
		if (texto == null) {
			return null;
		}
		return texto.replace("{pagina}", "&P").replace("{total}", "&N").replace("{data}", "&D")
				.replace("{hora}", "&T").replace("{arquivo}", "&F").replace("{aba}", "&A");
	}
}
