package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Cria e reaproveita (cache) os {@link CellStyle} de formatos comuns — data,
 * data/hora e porcentagem — para um workbook.
 *
 * <p>
 * Um {@code CellStyle} novo criado a cada célula esgota rápido o limite de
 * estilos de um workbook `.xlsx`; por isso cada formato é criado uma única vez
 * e reutilizado em todas as células que o pedirem.
 * </p>
 */
public final class FormatosDeCelula {

	private final Workbook workbook;
	private CellStyle estiloData;
	private CellStyle estiloDataHora;
	private CellStyle estiloPorcentagem;

	/**
	 * @param workbook O workbook para o qual os formatos serão criados.
	 */
	public FormatosDeCelula(final Workbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * Estilo de data no formato brasileiro (dd/MM/aaaa).
	 *
	 * @return O estilo, criando-o na primeira chamada.
	 */
	public CellStyle data() {
		if (estiloData == null) {
			estiloData = criarEstiloDeFormato("dd/MM/yyyy");
		}
		return estiloData;
	}

	/**
	 * Estilo de data e hora no formato brasileiro (dd/MM/aaaa HH:mm).
	 *
	 * @return O estilo, criando-o na primeira chamada.
	 */
	public CellStyle dataHora() {
		if (estiloDataHora == null) {
			estiloDataHora = criarEstiloDeFormato("dd/MM/yyyy HH:mm");
		}
		return estiloDataHora;
	}

	/**
	 * Estilo de porcentagem com duas casas decimais (ex.: 15,00%).
	 *
	 * @return O estilo, criando-o na primeira chamada.
	 */
	public CellStyle porcentagem() {
		if (estiloPorcentagem == null) {
			estiloPorcentagem = criarEstiloDeFormato("0.00%");
		}
		return estiloPorcentagem;
	}

	private CellStyle criarEstiloDeFormato(final String formato) {
		final CellStyle estilo = workbook.createCellStyle();
		estilo.setDataFormat(workbook.createDataFormat().getFormat(formato));
		return estilo;
	}

	/**
	 * Aplica um estilo a todas as células já existentes de uma coluna, a partir de
	 * uma linha inicial.
	 *
	 * @param sheet         A folha onde a coluna está.
	 * @param coluna        Índice da coluna (0-based).
	 * @param primeiraLinha Índice da primeira linha a considerar (0-based).
	 * @param estilo        O estilo a aplicar.
	 */
	public static void aplicarNaColuna(final Sheet sheet, final int coluna, final int primeiraLinha,
			final CellStyle estilo) {
		for (int indiceLinha = primeiraLinha; indiceLinha <= sheet.getLastRowNum(); indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			if (linha == null) {
				continue;
			}
			final Cell celula = linha.getCell(coluna);
			if (celula != null) {
				celula.setCellStyle(estilo);
			}
		}
	}
}
