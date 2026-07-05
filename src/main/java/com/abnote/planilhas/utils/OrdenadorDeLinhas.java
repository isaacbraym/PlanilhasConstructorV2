package com.abnote.planilhas.utils;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

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
		final FormulaEvaluator avaliador = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		final List<LinhaCapturada> capturadas = new ArrayList<>();
		for (int indiceLinha = primeiraLinha; indiceLinha <= ultimaLinha; indiceLinha++) {
			capturadas.add(LinhaCapturada.de(sheet.getRow(indiceLinha), coluna, avaliador));
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
			final LinhaCapturada capturada = capturadas.get(posicao);
			capturada.aplicarAtributos(linha);
			for (final CelulaCapturada celula : capturada.celulas) {
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
		private boolean linhaExistente;
		private short altura;
		private boolean oculta;
		private CellStyle estilo;
		private Double chaveNumerica;
		private String chaveTexto;
		private final List<CelulaCapturada> celulas = new ArrayList<>();

		private static LinhaCapturada de(final Row linha, final int coluna, final FormulaEvaluator avaliador) {
			final LinhaCapturada capturada = new LinhaCapturada();
			if (linha == null) {
				return capturada;
			}
			capturada.linhaExistente = true;
			capturada.altura = linha.getHeight();
			capturada.oculta = linha.getZeroHeight();
			capturada.estilo = linha.getRowStyle();
			for (final Cell celula : linha) {
				capturada.celulas.add(CelulaCapturada.de(celula));
			}
			capturada.definirChave(linha.getCell(coluna), avaliador);
			return capturada;
		}

		private void definirChave(final Cell chave, final FormulaEvaluator avaliador) {
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
				definirChaveDeFormula(chave, avaliador);
				break;
			default:
				break;
			}
		}

		private void definirChaveDeFormula(final Cell chave, final FormulaEvaluator avaliador) {
			try {
				final CellValue resultado = avaliador.evaluate(chave);
				if (resultado == null) {
					return;
				}
				switch (resultado.getCellType()) {
				case NUMERIC:
					chaveNumerica = resultado.getNumberValue();
					break;
				case STRING:
					chaveTexto = resultado.getStringValue();
					break;
				case BOOLEAN:
					chaveTexto = Boolean.toString(resultado.getBooleanValue());
					break;
				default:
					chaveTexto = chave.getCellFormula();
					break;
				}
			} catch (RuntimeException e) {
				chaveTexto = chave.getCellFormula();
			}
		}

		private void aplicarAtributos(final Row linha) {
			if (linhaExistente) {
				linha.setHeight(altura);
				linha.setZeroHeight(oculta);
				linha.setRowStyle(estilo);
			} else {
				linha.setHeight(linha.getSheet().getDefaultRowHeight());
				linha.setZeroHeight(false);
				linha.setRowStyle(null);
			}
		}
	}

	/** Foto de uma célula: valor, tipo e estilo. */
	private static final class CelulaCapturada {
		private int coluna;
		private int linhaOrigem;
		private CellType tipo;
		private CellStyle estilo;
		private String texto;
		private double numero;
		private boolean logico;
		private String formula;

		private static CelulaCapturada de(final Cell celula) {
			final CelulaCapturada capturada = new CelulaCapturada();
			capturada.coluna = celula.getColumnIndex();
			capturada.linhaOrigem = celula.getRowIndex();
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
				nova.setCellFormula(ajustarFormulaParaLinha(nova.getSheet(), formula, linhaOrigem, nova.getRowIndex()));
				break;
			default:
				break;
			}
			if (estilo != null) {
				nova.setCellStyle(estilo);
			}
		}

		private String ajustarFormulaParaLinha(final Sheet sheet, final String formulaOriginal, final int origem,
				final int destino) {
			if (formulaOriginal == null || origem == destino) {
				return formulaOriginal;
			}
			try {
				final Workbook workbook = sheet.getWorkbook();
				final EvaluationWorkbook avaliacao = workbook.createEvaluationWorkbook();
				final FormulaParsingWorkbook parser = (FormulaParsingWorkbook) avaliacao;
				final FormulaRenderingWorkbook renderizador = (FormulaRenderingWorkbook) avaliacao;
				final int indiceAba = workbook.getSheetIndex(sheet);
				final Ptg[] tokens = FormulaParser.parse(formulaOriginal, parser, FormulaType.CELL, indiceAba, origem);
				final FormulaShifter shifter = FormulaShifter.createForRowCopy(indiceAba, sheet.getSheetName(), origem,
						origem, destino - origem, workbook.getSpreadsheetVersion());
				shifter.adjustFormula(tokens, indiceAba);
				return FormulaRenderer.toFormulaString(renderizador, tokens);
			} catch (RuntimeException e) {
				return formulaOriginal;
			}
		}
	}
}
