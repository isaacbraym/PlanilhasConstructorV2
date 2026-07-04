package com.abnote.planilhas.utils;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Ordena as linhas de uma planilha pelo valor de uma coluna, preservando o
 * conteúdo e a formatação de cada linha (as células "viajam" juntas).
 *
 * <p>
 * A ordenação coloca números antes de texto; dentro de cada grupo, ordem
 * numérica ou alfabética (ignorando maiúsculas/minúsculas). Células vazias vão
 * para o fim (na ordem crescente).
 * </p>
 */
public final class OrdenadorDeLinhas {

	private OrdenadorDeLinhas() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Ordena as linhas a partir de {@code primeiraLinha} (0-based) até a última.
	 *
	 * @param sheet         A folha a ordenar.
	 * @param coluna        Índice da coluna-chave (0-based).
	 * @param crescente     {@code true} para A→Z / menor→maior.
	 * @param primeiraLinha Índice da primeira linha de dados (0-based). Use 1 para
	 *                      preservar um cabeçalho na linha 0.
	 */
	public static void ordenar(final Sheet sheet, final int coluna, final boolean crescente, final int primeiraLinha) {
		final int ultimaLinha = sheet.getLastRowNum();
		if (ultimaLinha <= primeiraLinha) {
			return; // 0 ou 1 linha de dados: nada a ordenar.
		}
		final List<LinhaCapturada> capturadas = new ArrayList<>();
		for (int indiceLinha = primeiraLinha; indiceLinha <= ultimaLinha; indiceLinha++) {
			capturadas.add(LinhaCapturada.de(sheet.getRow(indiceLinha), coluna));
		}
		capturadas.sort(crescente ? OrdenadorDeLinhas::compararChaves
				: (a, b) -> compararChaves(b, a));
		reescrever(sheet, capturadas, primeiraLinha);
	}

	private static int compararChaves(final LinhaCapturada a, final LinhaCapturada b) {
		if (a.chaveNumerica != null && b.chaveNumerica != null) {
			return Double.compare(a.chaveNumerica, b.chaveNumerica);
		}
		if (a.chaveNumerica != null) {
			return -1; // números antes de texto
		}
		if (b.chaveNumerica != null) {
			return 1;
		}
		if (a.chaveTexto == null) {
			return b.chaveTexto == null ? 0 : 1; // vazios por último
		}
		if (b.chaveTexto == null) {
			return -1;
		}
		return a.chaveTexto.compareToIgnoreCase(b.chaveTexto);
	}

	private static void reescrever(final Sheet sheet, final List<LinhaCapturada> capturadas, final int primeiraLinha) {
		for (int posicao = 0; posicao < capturadas.size(); posicao++) {
			final int indiceLinha = primeiraLinha + posicao;
			Row linha = sheet.getRow(indiceLinha);
			if (linha == null) {
				linha = sheet.createRow(indiceLinha);
			} else {
				limparLinha(linha);
			}
			for (final CelulaCapturada celula : capturadas.get(posicao).celulas) {
				celula.aplicar(linha);
			}
		}
	}

	private static void limparLinha(final Row linha) {
		for (int coluna = linha.getLastCellNum() - 1; coluna >= 0; coluna--) {
			final Cell celula = linha.getCell(coluna);
			if (celula != null) {
				linha.removeCell(celula);
			}
		}
	}

	/** Foto de uma linha: suas células e a chave de ordenação. */
	private static final class LinhaCapturada {
		private Double chaveNumerica;
		private String chaveTexto;
		private final List<CelulaCapturada> celulas = new ArrayList<>();

		private static LinhaCapturada de(final Row linha, final int coluna) {
			final LinhaCapturada capturada = new LinhaCapturada();
			if (linha == null) {
				return capturada;
			}
			for (final Cell celula : linha) {
				capturada.celulas.add(CelulaCapturada.de(celula));
			}
			capturada.definirChave(linha.getCell(coluna));
			return capturada;
		}

		private void definirChave(final Cell chave) {
			if (chave == null) {
				return;
			}
			switch (chave.getCellType()) {
			case NUMERIC:
				chaveNumerica = chave.getNumericCellValue();
				break;
			case STRING:
				chaveTexto = chave.getStringCellValue();
				break;
			case BOOLEAN:
				chaveTexto = Boolean.toString(chave.getBooleanCellValue());
				break;
			case FORMULA:
				chaveTexto = chave.getCellFormula();
				break;
			default:
				break;
			}
		}
	}

	/** Foto de uma célula: valor, tipo e estilo. */
	private static final class CelulaCapturada {
		private int coluna;
		private CellType tipo;
		private CellStyle estilo;
		private String texto;
		private double numero;
		private boolean logico;
		private String formula;

		private static CelulaCapturada de(final Cell celula) {
			final CelulaCapturada capturada = new CelulaCapturada();
			capturada.coluna = celula.getColumnIndex();
			capturada.tipo = celula.getCellType();
			capturada.estilo = celula.getCellStyle();
			switch (capturada.tipo) {
			case STRING:
				capturada.texto = celula.getStringCellValue();
				break;
			case NUMERIC:
				capturada.numero = celula.getNumericCellValue();
				break;
			case BOOLEAN:
				capturada.logico = celula.getBooleanCellValue();
				break;
			case FORMULA:
				capturada.formula = celula.getCellFormula();
				break;
			default:
				break;
			}
			return capturada;
		}

		private void aplicar(final Row linha) {
			final Cell nova = linha.createCell(coluna);
			switch (tipo) {
			case STRING:
				nova.setCellValue(texto);
				break;
			case NUMERIC:
				nova.setCellValue(numero);
				break;
			case BOOLEAN:
				nova.setCellValue(logico);
				break;
			case FORMULA:
				nova.setCellFormula(formula);
				break;
			default:
				break;
			}
			if (estilo != null) {
				nova.setCellStyle(estilo);
			}
		}
	}
}
