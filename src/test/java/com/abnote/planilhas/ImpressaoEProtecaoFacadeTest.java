package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de configuração de impressão e proteção de planilha na facade.
 */
@DisplayName("Planilha — impressão e proteção")
class ImpressaoEProtecaoFacadeTest {

	@Test
	@DisplayName("orientacaoPaisagem e orientacaoRetrato devem alternar a orientação de impressão")
	void deveDefinirOrientacao() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.orientacaoPaisagem();
			PrintSetup setup = planilha.workbook().getSheetAt(0).getPrintSetup();
			assertTrue(setup.getLandscape());

			planilha.orientacaoRetrato();
			assertFalse(setup.getLandscape());
		}
	}

	@Test
	@DisplayName("areaDeImpressao deve registrar a área no workbook")
	void deveDefinirAreaDeImpressao() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.areaDeImpressao("A1:F30");
			String area = planilha.workbook().getPrintArea(0);
			assertNotNull(area);
			assertTrue(area.contains("A1") && area.contains("F30") || area.contains("$A$1") && area.contains("$F$30"),
					"Área registrada: " + area);
		}
	}

	@Test
	@DisplayName("ajustarImpressaoEmPaginas deve ativar fitToPage com a largura/altura informadas")
	void deveAjustarImpressaoEmPaginas() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.ajustarImpressaoEmPaginas(1, 2);
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.getFitToPage());
			assertEquals(1, sheet.getPrintSetup().getFitWidth());
			assertEquals(2, sheet.getPrintSetup().getFitHeight());
		}
	}

	@Test
	@DisplayName("protegerPlanilha deve ativar a proteção da aba")
	void deveProtegerPlanilha() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertFalse(planilha.workbook().getSheetAt(0).getProtect());
			planilha.protegerPlanilha("1234");
			assertTrue(planilha.workbook().getSheetAt(0).getProtect());
		}
	}

	@Test
	@DisplayName("desbloquearCelulas deve destravar só o intervalo informado, sem afetar outras células")
	void deveDesbloquearApenasIntervalo() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			// Duas células que, sem nenhum estilo próprio, compartilham o MESMO
			// CellStyle padrão do workbook — o teste real é garantir que destravar
			// uma não destrava a outra (prova de que o estilo é clonado, não mutado).
			planilha.escrever("A1", "entrada").escrever("B1", "não deve mudar");

			planilha.desbloquearCelulas("A1:A1");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertFalse(sheet.getRow(0).getCell(0).getCellStyle().getLocked(), "A1 deve estar destravada");
			assertTrue(sheet.getRow(0).getCell(1).getCellStyle().getLocked(), "B1 deve continuar travada");
		}
	}

	@Test
	@DisplayName("desbloquearCelulas deve criar células vazias dentro do intervalo, se necessário")
	void deveCriarCelulaVaziaAoDesbloquear() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.desbloquearCelulas("C3:C3");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertNotNull(sheet.getRow(2));
			assertNotNull(sheet.getRow(2).getCell(2));
			assertFalse(sheet.getRow(2).getCell(2).getCellStyle().getLocked());
		}
	}

	@Test
	@DisplayName("Fluxo completo: desbloquear campo de entrada e depois proteger a planilha")
	void deveMontarFormularioProtegido() throws Exception {
		try (Planilha planilha = Planilha.nova("Formulario")) {
			planilha.escreverLinha("A1", "Nome:", "").desbloquearCelulas("B1:B1").protegerPlanilha("");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.getProtect());
			assertFalse(sheet.getRow(0).getCell(1).getCellStyle().getLocked(), "Campo de entrada editável");
			assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getLocked(), "Rótulo continua travado");
		}
	}
}
