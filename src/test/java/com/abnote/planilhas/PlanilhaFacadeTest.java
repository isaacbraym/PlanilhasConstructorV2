package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Testes da facade amigável {@link Planilha}.
 */
@DisplayName("Planilha (facade amigável)")
class PlanilhaFacadeTest {

	@TempDir
	Path pasta;

	@Test
	@DisplayName("Deve escrever números, textos e preservar zeros à esquerda")
	void deveEscreverValores() throws Exception {
		try (Planilha planilha = Planilha.nova("Dados")) {
			planilha.escrever("A1", "Nome").escrever("B1", 10).escrever("C1", 3.5).escreverTexto("D1", "007");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals(CellType.STRING, sheet.getRow(0).getCell(0).getCellType());
			assertEquals("Nome", sheet.getRow(0).getCell(0).getStringCellValue());
			assertEquals(10.0, sheet.getRow(0).getCell(1).getNumericCellValue(), 0.001);
			assertEquals(3.5, sheet.getRow(0).getCell(2).getNumericCellValue(), 0.001);
			assertEquals(CellType.STRING, sheet.getRow(0).getCell(3).getCellType());
			assertEquals("007", sheet.getRow(0).getCell(3).getStringCellValue());
		}
	}

	@Test
	@DisplayName("escreverLinha, adicionarLinha e negrito devem montar uma lista")
	void deveMontarLista() throws Exception {
		try (Planilha planilha = Planilha.nova("Vendas")) {
			planilha.escreverLinha("A1", "Produto", "Preço", "Qtd").negrito("A1:C1").adicionarLinha("Caneta", 2.5, 100)
					.adicionarLinha("Caderno", 15.9, 30);

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("Produto", sheet.getRow(0).getCell(0).getStringCellValue());
			assertEquals("Caneta", sheet.getRow(1).getCell(0).getStringCellValue());
			assertEquals(100.0, sheet.getRow(1).getCell(2).getNumericCellValue(), 0.001);
			assertEquals("Caderno", sheet.getRow(2).getCell(0).getStringCellValue());
			assertTrue(((XSSFCellStyle) sheet.getRow(0).getCell(0).getCellStyle()).getFont().getBold());
			assertTrue(((XSSFCellStyle) sheet.getRow(0).getCell(2).getCellStyle()).getFont().getBold());
		}
	}

	@Test
	@DisplayName("escreverTabela deve escrever uma matriz de dados")
	void deveEscreverTabela() throws Exception {
		try (Planilha planilha = Planilha.nova("Tabela")) {
			planilha.escreverTabela("A1",
					Arrays.asList(Arrays.asList("A", "B"), Arrays.asList(1, 2), Arrays.asList(3, 4)));

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
			assertEquals(2.0, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.001);
			assertEquals(3.0, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("somar deve gravar fórmula que avalia corretamente")
	void deveSomar() throws Exception {
		try (Planilha planilha = Planilha.nova("Contas")) {
			planilha.escreverColuna("A1", 10, 20, 30).somar("A4", "A1:A3");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Cell soma = sheet.getRow(3).getCell(0);
			assertEquals("SUM(A1:A3)", soma.getCellFormula());
			FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			avaliador.evaluateFormulaCell(soma);
			assertEquals(60.0, soma.getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("formatarComoMoeda deve aplicar R$")
	void deveFormatarMoeda() throws Exception {
		try (Planilha planilha = Planilha.nova("Precos")) {
			planilha.escrever("A1", 1234.56).formatarComoMoeda("A1");
			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString().contains("R$"));
		}
	}

	@Test
	@DisplayName("duplicarAba deve copiar conteúdo em nova aba")
	void deveDuplicarAba() throws Exception {
		try (Planilha planilha = Planilha.nova("Original")) {
			planilha.escrever("A1", "Conteúdo").duplicarAba("Cópia");

			Workbook workbook = planilha.workbook();
			assertEquals(2, workbook.getNumberOfSheets());
			assertEquals("Conteúdo", workbook.getSheet("Cópia").getRow(0).getCell(0).getStringCellValue());
		}
	}

	@Test
	@DisplayName("duplicarAba deve recusar nome existente sem criar aba parcial")
	void deveRecusarDuplicarAbaComNomeExistenteSemMutacaoParcial() throws Exception {
		try (Planilha planilha = Planilha.nova("Original")) {
			planilha.novaAba("Existente").irParaAba("Original");

			assertThrows(IllegalArgumentException.class, () -> planilha.duplicarAba("Existente"));

			Workbook workbook = planilha.workbook();
			assertEquals(2, workbook.getNumberOfSheets());
			assertEquals("Original", workbook.getSheetName(0));
			assertEquals("Existente", workbook.getSheetName(1));
			assertNull(workbook.getSheet("Original (2)"));
		}
	}

	@Test
	@DisplayName("duplicarAba deve recusar nome inválido sem criar aba parcial")
	void deveRecusarDuplicarAbaComNomeInvalidoSemMutacaoParcial() throws Exception {
		try (Planilha planilha = Planilha.nova("Original")) {
			assertThrows(IllegalArgumentException.class, () -> planilha.duplicarAba("Nome/Invalido"));

			Workbook workbook = planilha.workbook();
			assertEquals(1, workbook.getNumberOfSheets());
			assertEquals("Original", workbook.getSheetName(0));
		}
	}

	@Test
	@DisplayName("duplicarColuna e moverColuna devem manipular colunas")
	void deveManipularColunas() throws Exception {
		try (Planilha planilha = Planilha.nova("Cols")) {
			planilha.escreverLinha("A1", "X", "Y").duplicarColuna("A", "C");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("X", sheet.getRow(0).getCell(2).getStringCellValue());

			planilha.moverColuna("A", "B");
			assertEquals("Y", sheet.getRow(0).getCell(0).getStringCellValue());
		}
	}

	@Test
	@DisplayName("mesclar deve criar região mesclada")
	void deveMesclar() throws Exception {
		try (Planilha planilha = Planilha.nova("Merge")) {
			planilha.escrever("A1", "Título").mesclar("A1:C1");
			assertEquals(1, planilha.workbook().getSheetAt(0).getNumMergedRegions());
		}
	}

	@Test
	@DisplayName("Deve salvar arquivo .xlsx que pode ser reaberto")
	void deveSalvar() throws Exception {
		String caminho = pasta.resolve("saida.xlsx").toString();
		try (Planilha planilha = Planilha.nova("Saida")) {
			planilha.escrever("A1", "Olá").salvar(caminho);
		}
		File arquivo = new File(caminho);
		assertTrue(arquivo.exists());
		try (Workbook wb = new XSSFWorkbook(arquivo)) {
			assertEquals("Olá", wb.getSheet("Saida").getRow(0).getCell(0).getStringCellValue());
		}
	}

	@Test
	@DisplayName("estilo() deve permitir encadear vários estilos de uma vez")
	void deveEncadearEstilos() throws Exception {
		try (Planilha planilha = Planilha.nova("Estilo")) {
			planilha.escrever("A1", "Destaque");
			planilha.estilo("A1").aplicarNegrito().corDeFundo(CorEnum.AMARELO).corFonte(CorEnum.AZUL);

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(((XSSFCellStyle) sheet.getRow(0).getCell(0).getCellStyle()).getFont().getBold());
		}
	}
}
