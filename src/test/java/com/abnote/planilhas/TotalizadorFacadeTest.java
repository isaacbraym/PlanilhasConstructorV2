package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Testes de {@code adicionarTotais} na facade.
 */
@DisplayName("Planilha — adicionarTotais")
class TotalizadorFacadeTest {

	@TempDir
	Path pasta;

	@Test
	@DisplayName("Deve somar as colunas numéricas e rotular a coluna de texto com 'Total'")
	void deveSomarColunasNumericas() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Preco", "Qtd")
					.adicionarLinha("Caneta", 2.5, 100)
					.adicionarLinha("Caderno", 15.9, 30)
					.adicionarTotais("A1");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Row linhaTotais = sheet.getRow(3);
			assertNotNull(linhaTotais, "Linha de totais deve existir na linha 4");

			assertEquals("Total", linhaTotais.getCell(0).getStringCellValue());

			Cell somaPreco = linhaTotais.getCell(1);
			assertEquals("SUM(B2:B3)", somaPreco.getCellFormula());
			Cell somaQtd = linhaTotais.getCell(2);
			assertEquals("SUM(C2:C3)", somaQtd.getCellFormula());

			FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			avaliador.evaluateFormulaCell(somaPreco);
			avaliador.evaluateFormulaCell(somaQtd);
			assertEquals(18.4, somaPreco.getNumericCellValue(), 0.001);
			assertEquals(130.0, somaQtd.getNumericCellValue(), 0.001);
		}
	}

	@Test
	@DisplayName("Deve parar de considerar colunas no primeiro cabeçalho vazio")
	void deveRespeitarLarguraDoCabecalho() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Qtd");
			// Coluna D tem dado numérico, mas está fora do cabeçalho (que termina em B).
			planilha.escrever("D2", 999);
			planilha.adicionarLinha("Caneta", 10).adicionarTotais("A1");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Row linhaTotais = sheet.getRow(2);
			assertNull(linhaTotais.getCell(3), "Coluna D está fora da tabela e não deve ser tocada");
		}
	}

	@Test
	@DisplayName("Tabela sem nenhuma linha de dados não deve gerar linha de totais")
	void deveIgnorarTabelaSemDados() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Qtd").adicionarTotais("A1");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertNull(sheet.getRow(1), "Não deve haver linha 2 (sem dados para totalizar)");
		}
	}

	@Test
	@DisplayName("Coluna com célula não numérica no meio dos dados não deve receber soma")
	void deveIgnorarColunaComTextoMisturado() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Observacao")
					.adicionarLinha("Caneta", "ok")
					.adicionarLinha("Caderno", "revisar")
					.adicionarTotais("A1");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Row linhaTotais = sheet.getRow(3);
			assertEquals("Total", linhaTotais.getCell(0).getStringCellValue());
			assertNull(linhaTotais.getCell(1), "Coluna só com texto não deve ganhar fórmula de soma");
		}
	}

	@Test
	@DisplayName("Deve somar coluna com formulas numericas")
	void deveSomarColunaComFormulasNumericas() throws Exception {
		String caminho = pasta.resolve("totais-formulas.xlsx").toString();
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Preco", "Qtd", "Total")
					.escreverLinha("A2", "Caneta", 2.5, 10, "")
					.formula("D2", "B2*C2")
					.escreverLinha("A3", "Caderno", 15, 2, "")
					.formula("D3", "B3*C3")
					.adicionarTotais("A1")
					.salvar(caminho);

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Cell somaTotal = sheet.getRow(3).getCell(3);
			assertEquals("SUM(D2:D3)", somaTotal.getCellFormula());
			FormulaEvaluator avaliador = planilha.workbook().getCreationHelper().createFormulaEvaluator();
			assertEquals(55.0, avaliador.evaluate(somaTotal).getNumberValue(), 0.001);
		}

		try (Workbook reaberto = new XSSFWorkbook(new File(caminho))) {
			Sheet sheet = reaberto.getSheetAt(0);
			Cell somaTotal = sheet.getRow(3).getCell(3);
			assertEquals("SUM(D2:D3)", somaTotal.getCellFormula());
			FormulaEvaluator avaliador = reaberto.getCreationHelper().createFormulaEvaluator();
			assertEquals(55.0, avaliador.evaluate(somaTotal).getNumberValue(), 0.001);
		}
	}
}
