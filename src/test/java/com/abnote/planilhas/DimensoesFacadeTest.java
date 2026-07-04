package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.PaneInformation;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de porcentagem, largura/altura e congelamento na facade.
 */
@DisplayName("Planilha — porcentagem, dimensões e congelar")
class DimensoesFacadeTest {

	@Test
	@DisplayName("formatarComoPorcentagem deve aplicar formato com %")
	void deveFormatarPorcentagem() throws Exception {
		try (Planilha planilha = Planilha.nova("Pct")) {
			planilha.escrever("A1", 0.15).formatarComoPorcentagem("A1");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString().contains("%"));
		}
	}

	@Test
	@DisplayName("larguraColuna deve definir a largura da coluna")
	void deveDefinirLargura() throws Exception {
		try (Planilha planilha = Planilha.nova("Larg")) {
			planilha.larguraColuna("A", 20);
			assertEquals(20 * 256, planilha.workbook().getSheetAt(0).getColumnWidth(0));
		}
	}

	@Test
	@DisplayName("alturaLinha deve definir a altura da linha")
	void deveDefinirAltura() throws Exception {
		try (Planilha planilha = Planilha.nova("Alt")) {
			planilha.alturaLinha(1, 30);
			assertEquals(30.0f, planilha.workbook().getSheetAt(0).getRow(0).getHeightInPoints(), 0.1f);
		}
	}

	@Test
	@DisplayName("congelar deve criar painéis congelados")
	void deveCongelar() throws Exception {
		try (Planilha planilha = Planilha.nova("Congela")) {
			planilha.escrever("A1", "x").congelar(1, 2);
			PaneInformation painel = planilha.workbook().getSheetAt(0).getPaneInformation();
			assertNotNull(painel);
			assertTrue(painel.isFreezePane());
			assertEquals(1, painel.getHorizontalSplitPosition());
			assertEquals(2, painel.getVerticalSplitPosition());
		}
	}
}
