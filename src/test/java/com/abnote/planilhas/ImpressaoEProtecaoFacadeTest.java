package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;

import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.exceptions.DadosInvalidosException;

/**
 * Testes de configuração de impressão e proteção de planilha na facade.
 */
@DisplayName("Planilha — impressão e proteção")
class ImpressaoEProtecaoFacadeTest {

	@TempDir
	File tempDir;

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
	@DisplayName("margensDeImpressao deve gravar margens em centímetros e persistir em OOXML")
	void deveDefinirMargensDeImpressao() throws Exception {
		final File arquivo = new File(tempDir, "margens.xlsx");
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.margensDeImpressao(2.54, 1.27, 0.635, 3.81).salvar(arquivo.getAbsolutePath());
		}

		try (XSSFWorkbook workbook = new XSSFWorkbook(arquivo)) {
			final Sheet sheet = workbook.getSheetAt(0);
			assertEquals(1.0, sheet.getMargin(PageMargin.TOP), 0.001);
			assertEquals(0.5, sheet.getMargin(PageMargin.BOTTOM), 0.001);
			assertEquals(0.25, sheet.getMargin(PageMargin.LEFT), 0.001);
			assertEquals(1.5, sheet.getMargin(PageMargin.RIGHT), 0.001);
		}
	}

	@Test
	@DisplayName("margensDeImpressao deve recusar margem negativa")
	void deveRecusarMargemNegativa() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(DadosInvalidosException.class,
					() -> planilha.margensDeImpressao(1.0, -0.1, 1.0, 1.0));
		}
	}

	@Test
	@DisplayName("cabecalhoDeImpressao(centro) deve definir só o texto central")
	void deveDefinirCabecalhoSoCentro() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.cabecalhoDeImpressao("Relatório Mensal");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("Relatório Mensal", sheet.getHeader().getCenter());
			assertEquals("", sheet.getHeader().getLeft());
			assertEquals("", sheet.getHeader().getRight());
		}
	}

	@Test
	@DisplayName("cabecalhoDeImpressao(esquerda, centro, direita) deve traduzir marcadores amigáveis para códigos do Excel")
	void deveDefinirCabecalhoComTresPartesEMarcadores() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.cabecalhoDeImpressao("{arquivo}", "Relatório", "{data}");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("&F", sheet.getHeader().getLeft());
			assertEquals("Relatório", sheet.getHeader().getCenter());
			assertEquals("&D", sheet.getHeader().getRight());
		}
	}

	@Test
	@DisplayName("rodapeDeImpressao(centro) deve definir só o texto central")
	void deveDefinirRodapeSoCentro() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.rodapeDeImpressao("Confidencial");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("Confidencial", sheet.getFooter().getCenter());
		}
	}

	@Test
	@DisplayName("rodapeDeImpressao deve traduzir {pagina} e {total} para os códigos &P e &N")
	void deveDefinirRodapeComPaginacao() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.rodapeDeImpressao("", "Página {pagina} de {total}", "");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("Página &P de &N", sheet.getFooter().getCenter());
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
