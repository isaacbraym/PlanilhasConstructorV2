package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;

import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.DataBarFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting.IconSet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Testes de formatação condicional (realce de células) na facade.
 */
@DisplayName("Planilha — formatação condicional")
class FormatacaoCondicionalFacadeTest {

	@TempDir
	File tempDir;

	private ConditionalFormatting regraUnica(Planilha planilha) {
		Sheet sheet = planilha.workbook().getSheetAt(0);
		assertEquals(1, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
		return sheet.getSheetConditionalFormatting().getConditionalFormattingAt(0);
	}

	private void assertCorDeFundo(ConditionalFormattingRule regra, CorEnum esperado) {
		XSSFColor cor = (XSSFColor) regra.getPatternFormatting().getFillForegroundColorColor();
		byte[] rgb = cor.getRGB();
		assertEquals((byte) esperado.getRed(), rgb[0]);
		assertEquals((byte) esperado.getGreen(), rgb[1]);
		assertEquals((byte) esperado.getBlue(), rgb[2]);
	}

	@Test
	@DisplayName("realcarSeMaiorQue deve criar regra GT com a cor de fundo")
	void deveRealcarSeMaiorQue() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", 10, 200, 30).realcarSeMaiorQue("A1:A3", 100, CorEnum.VERDE);

			ConditionalFormatting cf = regraUnica(planilha);
			ConditionalFormattingRule regra = cf.getRule(0);
			assertEquals(ComparisonOperator.GT, regra.getComparisonOperation());
			assertCorDeFundo(regra, CorEnum.VERDE);
			assertEquals("A1:A3", cf.getFormattingRanges()[0].formatAsString());
		}
	}

	@Test
	@DisplayName("realcarSeMenorQue deve criar regra LT")
	void deveRealcarSeMenorQue() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.realcarSeMenorQue("B2:B10", 0, CorEnum.VERMELHO_ESCURO);
			ConditionalFormattingRule regra = regraUnica(planilha).getRule(0);
			assertEquals(ComparisonOperator.LT, regra.getComparisonOperation());
		}
	}

	@Test
	@DisplayName("realcarSeEntre deve criar regra BETWEEN com os dois limites")
	void deveRealcarSeEntre() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.realcarSeEntre("C1:C5", 10, 20, CorEnum.AMARELO);
			ConditionalFormattingRule regra = regraUnica(planilha).getRule(0);
			assertEquals(ComparisonOperator.BETWEEN, regra.getComparisonOperation());
			assertEquals("10.0", regra.getFormula1());
			assertEquals("20.0", regra.getFormula2());
		}
	}

	@Test
	@DisplayName("realcarSeIgual deve aceitar texto (entre aspas na fórmula) e número")
	void deveRealcarSeIgual() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.realcarSeIgual("D1:D5", "Atrasado", CorEnum.VERMELHO_ESCURO);
			ConditionalFormattingRule regra = regraUnica(planilha).getRule(0);
			assertEquals(ComparisonOperator.EQUAL, regra.getComparisonOperation());
			assertEquals("\"Atrasado\"", regra.getFormula1());
		}
	}

	@Test
	@DisplayName("escalaDeCores deve criar uma regra de escala com 3 cores e limiares min/percentil/max")
	void deveAplicarEscalaDeCores() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escalaDeCores("E1:E10");
			ConditionalFormattingRule regra = regraUnica(planilha).getRule(0);
			ColorScaleFormatting escala = regra.getColorScaleFormatting();

			assertEquals(3, escala.getNumControlPoints());
			Color[] cores = escala.getColors();
			assertEquals(3, cores.length);
			assertEquals(RangeType.MIN, escala.getThresholds()[0].getRangeType());
			assertEquals(RangeType.PERCENTILE, escala.getThresholds()[1].getRangeType());
			assertEquals(50.0, escala.getThresholds()[1].getValue(), 0.001);
			assertEquals(RangeType.MAX, escala.getThresholds()[2].getRangeType());
		}
	}

	@Test
	@DisplayName("barrasDeDados deve criar data bar colorida e persistir em OOXML")
	void deveAplicarBarrasDeDados() throws Exception {
		final File arquivo = new File(tempDir, "barras.xlsx");
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", 10, 20, 30)
					.barrasDeDados("A1:A3", CorEnum.AZUL)
					.salvar(arquivo.getAbsolutePath());
		}

		try (XSSFWorkbook workbook = new XSSFWorkbook(arquivo)) {
			final ConditionalFormattingRule regra = workbook.getSheetAt(0).getSheetConditionalFormatting()
					.getConditionalFormattingAt(0).getRule(0);
			final DataBarFormatting barras = regra.getDataBarFormatting();
			final byte[] rgb = ((XSSFColor) barras.getColor()).getRGB();
			assertEquals(RangeType.MIN, barras.getMinThreshold().getRangeType());
			assertEquals(RangeType.MAX, barras.getMaxThreshold().getRangeType());
			assertArrayEquals(new byte[] { 0, 0, (byte) 255 }, rgb);
		}
	}

	@Test
	@DisplayName("iconesSemaforo deve criar conjunto de ícones de semáforo e persistir em OOXML")
	void deveAplicarIconesSemaforo() throws Exception {
		final File arquivo = new File(tempDir, "icones.xlsx");
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", 10, 20, 30).iconesSemaforo("A1:A3").salvar(arquivo.getAbsolutePath());
		}

		try (XSSFWorkbook workbook = new XSSFWorkbook(arquivo)) {
			final ConditionalFormattingRule regra = workbook.getSheetAt(0).getSheetConditionalFormatting()
					.getConditionalFormattingAt(0).getRule(0);
			final IconMultiStateFormatting icones = regra.getMultiStateFormatting();
			assertEquals(IconSet.GYR_3_TRAFFIC_LIGHTS, icones.getIconSet());
			assertEquals(RangeType.PERCENT, icones.getThresholds()[1].getRangeType());
			assertEquals(33.0, icones.getThresholds()[1].getValue(), 0.001);
			assertEquals(66.0, icones.getThresholds()[2].getValue(), 0.001);
		}
	}
}
