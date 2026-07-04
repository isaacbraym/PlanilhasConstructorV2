package com.abnote.planilhas.formulas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.exceptions.FormulaException;
import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;

/**
 * Testes unitários para as fórmulas Excel implementadas em Sprint 5.
 * 
 * <p>
 * Testa as 11 fórmulas: SUM, AVERAGE, COUNT, MIN, MAX, IF, COUNTIF, SUMIF,
 * CONCATENATE, TODAY, NOW
 * </p>
 */
@DisplayName("Testes de Fórmulas Excel")
class FormulasTest {

	@TempDir
	File tempDir;

	private IPlanilha planilha;
	private File arquivoTeste;

	@BeforeEach
	void setUp() throws IOException {
		planilha = new PlanilhaXlsx();
		planilha.criarPlanilha("Teste");
		arquivoTeste = new File(tempDir, "teste-formulas.xlsx");
	}

	@AfterEach
	void tearDown() throws Exception {
		if (planilha != null) {
			planilha.close();
		}
	}

	// ========== TESTES DE FÓRMULAS MATEMÁTICAS ==========

	@Test
	@DisplayName("Deve inserir e avaliar fórmula SUM corretamente")
	void deveSomarValoresComFormulaSUM() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("A1").inserirDados("10");
		planilha.selecionar().celula("A2").inserirDados("20");
		planilha.selecionar().celula("A3").inserirDados("30");

		// ACT
		planilha.selecionar().celula("A4").formula().soma("A1:A3").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(3);
			assertNotNull(row, "Linha 4 deve existir"); // ✅ ADICIONADO

			Cell cell = row.getCell(0);
			assertNotNull(cell, "Célula A4 deve existir"); // ✅ ADICIONADO

			assertEquals(CellType.FORMULA, cell.getCellType(), "Célula deve conter fórmula");
			assertEquals("SUM(A1:A3)", cell.getCellFormula(), "Fórmula deve ser SUM(A1:A3)");

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(60.0, cell.getNumericCellValue(), 0.01, "Soma deve ser 60");
		}
	}

	@Test
	@DisplayName("Deve inserir e avaliar fórmula AVERAGE corretamente")
	void deveCalcularMediaComFormulaAVERAGE() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("B1").inserirDados("10");
		planilha.selecionar().celula("B2").inserirDados("20");
		planilha.selecionar().celula("B3").inserirDados("30");

		// ACT
		planilha.selecionar().celula("B4").formula().media("B1:B3").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(3);
			assertNotNull(row, "Linha 4 deve existir");

			Cell cell = row.getCell(1);
			assertNotNull(cell, "Célula B4 deve existir");

			assertEquals("AVERAGE(B1:B3)", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(20.0, cell.getNumericCellValue(), 0.01, "Média deve ser 20");
		}
	}

	@Test
	@DisplayName("Deve inserir e avaliar fórmula COUNT corretamente")
	void deveContarValoresComFormulaCOUNT() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("C1").inserirDados("10");
		planilha.selecionar().celula("C2").inserirDados("20");
		planilha.selecionar().celula("C3").inserirDados("30");
		planilha.selecionar().celula("C4").inserirDados("Texto");

		// ACT
		planilha.selecionar().celula("C5").formula().contar("C1:C4").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(4);
			assertNotNull(row, "Linha 5 deve existir");

			Cell cell = row.getCell(2);
			assertNotNull(cell, "Célula C5 deve existir");

			assertEquals("COUNT(C1:C4)", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(3.0, cell.getNumericCellValue(), 0.01, "Contagem deve ser 3 (ignora texto)");
		}
	}

	@Test
	@DisplayName("Deve inserir e avaliar fórmula MIN corretamente")
	void deveEncontrarMinimoComFormulaMIN() throws IOException, InvalidFormatException, ArquivoException {
		// ARRANGE
		planilha.selecionar().celula("D1").inserirDados("50");
		planilha.selecionar().celula("D2").inserirDados("10");
		planilha.selecionar().celula("D3").inserirDados("30");

		// ACT
		planilha.selecionar().celula("D4").formula().minimo("D1:D3").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(3);
			assertNotNull(row, "Linha 4 deve existir");

			Cell cell = row.getCell(3);
			assertNotNull(cell, "Célula D4 deve existir");

			assertEquals("MIN(D1:D3)", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(10.0, cell.getNumericCellValue(), 0.01, "Mínimo deve ser 10");
		}
	}

	@Test
	@DisplayName("Deve inserir e avaliar fórmula MAX corretamente")
	void deveEncontrarMaximoComFormulaMAX() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("E1").inserirDados("50");
		planilha.selecionar().celula("E2").inserirDados("100");
		planilha.selecionar().celula("E3").inserirDados("30");

		// ACT
		planilha.selecionar().celula("E4").formula().maximo("E1:E3").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(3);
			assertNotNull(row, "Linha 4 deve existir");

			Cell cell = row.getCell(4);
			assertNotNull(cell, "Célula E4 deve existir");

			assertEquals("MAX(E1:E3)", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(100.0, cell.getNumericCellValue(), 0.01, "Máximo deve ser 100");
		}
	}

	// ========== TESTES DE FÓRMULAS CONDICIONAIS ==========

	@Test
	@DisplayName("Deve inserir e avaliar fórmula IF com números corretamente")
	void deveAvaliarCondicaoIFComNumeros() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("F1").inserirDados("150");

		// ACT
		planilha.selecionar().celula("F2").formula().seEntao("F1>100", "Alto", "Baixo").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(1);
			assertNotNull(row, "Linha 2 deve existir");

			Cell cell = row.getCell(5);
			assertNotNull(cell, "Célula F2 deve existir");

			assertEquals("IF(F1>100,\"Alto\",\"Baixo\")", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals("Alto", cell.getStringCellValue(), "Resultado deve ser 'Alto'");
		}
	}

	@Test
	@DisplayName("Deve inserir e avaliar fórmula COUNTIF corretamente")
	void deveContarComCondicaoCOUNTIF() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("G1").inserirDados("50");
		planilha.selecionar().celula("G2").inserirDados("150");
		planilha.selecionar().celula("G3").inserirDados("200");
		planilha.selecionar().celula("G4").inserirDados("80");

		// ACT
		planilha.selecionar().celula("G5").formula().contarSe("G1:G4", ">100").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(4);
			assertNotNull(row, "Linha 5 deve existir");

			Cell cell = row.getCell(6);
			assertNotNull(cell, "Célula G5 deve existir");

			assertEquals("COUNTIF(G1:G4,\">100\")", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(2.0, cell.getNumericCellValue(), 0.01, "Contagem deve ser 2 (valores >100)");
		}
	}

	@Test
	@DisplayName("Deve inserir e avaliar fórmula SUMIF corretamente")
	void deveSomarComCondicaoSUMIF() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("H1").inserirDados("50");
		planilha.selecionar().celula("H2").inserirDados("150");
		planilha.selecionar().celula("H3").inserirDados("200");
		planilha.selecionar().celula("H4").inserirDados("80");

		// ACT
		planilha.selecionar().celula("H5").formula().somarSe("H1:H4", ">100").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(4);
			assertNotNull(row, "Linha 5 deve existir");

			Cell cell = row.getCell(7);
			assertNotNull(cell, "Célula H5 deve existir");

			assertEquals("SUMIF(H1:H4,\">100\")", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals(350.0, cell.getNumericCellValue(), 0.01, "Soma deve ser 350 (150+200)");
		}
	}

	// ========== TESTES DE FÓRMULAS ADICIONAIS ==========

	@Test
	@DisplayName("Deve inserir e avaliar fórmula CONCATENATE corretamente")
	void deveConcatenarTextosComCONCATENATE() throws IOException, InvalidFormatException {
		// ARRANGE
		planilha.selecionar().celula("I1").inserirDados("Java");
		planilha.selecionar().celula("I2").inserirDados("POI");

		// ACT
		planilha.selecionar().celula("I3").formula().concatenar("I1", "I2").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(2);
			assertNotNull(row, "Linha 3 deve existir");

			Cell cell = row.getCell(8);
			assertNotNull(cell, "Célula I3 deve existir");

			assertEquals("CONCATENATE(I1,I2)", cell.getCellFormula());

			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cell);
			assertEquals("JavaPOI", cell.getStringCellValue());
		}
	}

	@Test
	@DisplayName("Deve inserir fórmula TODAY corretamente")
	void deveInserirFormulaTODAY() throws IOException, InvalidFormatException {
		// ACT
		planilha.selecionar().celula("J1") // ✅ JÁ TEM naCelula()
				.formula().hoje().aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(0);
			assertNotNull(row, "Linha 1 deve existir"); // ✅ ADICIONADO

			Cell cell = row.getCell(9);
			assertNotNull(cell, "Célula J1 deve existir"); // ✅ ADICIONADO

			assertEquals("TODAY()", cell.getCellFormula());
			assertEquals(CellType.FORMULA, cell.getCellType());
		}
	}

	@Test
	@DisplayName("Deve inserir fórmula NOW corretamente")
	void deveInserirFormulaNOW() throws IOException, InvalidFormatException {
		// ACT
		planilha.selecionar().celula("K1").formula().agora().aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(0);
			assertNotNull(row, "Linha 1 deve existir");

			Cell cell = row.getCell(10);
			assertNotNull(cell, "Célula K1 deve existir");

			assertEquals("NOW()", cell.getCellFormula());
			assertEquals(CellType.FORMULA, cell.getCellType());
		}
	}

	// ========== TESTES DE VALIDAÇÃO ==========

	@Test
	@DisplayName("Deve lançar exceção ao tentar usar formula() sem selecionar célula")
	void deveLancarExcecaoSemSelecionarCelula() {
		// ACT & ASSERT
		IllegalStateException exception = assertThrows(IllegalStateException.class, () -> planilha.formula());

		assertTrue(exception.getMessage().contains("Nenhuma célula foi selecionada"),
				"Mensagem deve indicar que célula não foi selecionada");
	}

	@Test
	@DisplayName("Deve lançar exceção para range inválido")
	void deveLancarExcecaoParaRangeInvalido() throws IOException, InvalidFormatException {
	    // ACT & ASSERT
	    FormulaException exception = assertThrows(FormulaException.class,  // ← MUDOU AQUI
	            () -> planilha.selecionar().celula("A1").formula().soma("INVALIDO").aplicar()
	    );

	    assertTrue(exception.getMessage().toLowerCase().contains("formato"), 
	               "Mensagem deve indicar formato inválido");
	}

	@Test
	@DisplayName("Deve lançar exceção para range vazio")
	void deveLancarExcecaoParaRangeVazio() {
	    // ACT & ASSERT
	    FormulaException exception = assertThrows(FormulaException.class,  // ← MUDOU AQUI
	            () -> planilha.selecionar().celula("A1").formula().soma("").aplicar()
	    );

	    assertTrue(exception.getMessage().toLowerCase().contains("vazio"),
	               "Mensagem deve indicar que range não pode ser vazio");
	}

	@Test
	@DisplayName("Deve lançar exceção para condição IF sem operador")
	void deveLancarExcecaoParaCondicaoInvalida() {
	    // ACT & ASSERT
	    FormulaException exception = assertThrows(FormulaException.class,  // ← MUDOU AQUI
	            () -> planilha.selecionar().celula("A1").formula()
	                    .seEntao("A1", "Sim", "Não").aplicar()
	    );

	    assertTrue(exception.getMessage().contains("operador"), 
	               "Mensagem deve indicar necessidade de operador");
	}

	@Test
	@DisplayName("Deve lançar exceção ao aplicar formula sem naCelula")
	void deveLancarExcecaoSemSelecionarCelulaAntes() {
		// ACT & ASSERT
		IllegalStateException exception = assertThrows(IllegalStateException.class, () -> planilha.formula().aplicar());

		assertTrue(exception.getMessage().contains("Nenhuma célula foi selecionada"),
				"Mensagem deve indicar que célula não foi selecionada");
	}
	// ========== TESTES DE INTEGRAÇÃO ==========

	@Test
	@DisplayName("Deve permitir encadeamento de múltiplas operações")
	void devePermitirEncadeamentoDeOperacoes() throws IOException, InvalidFormatException {
		// ACT
		planilha.selecionar().celula("L1").inserirDados("10");
		planilha.selecionar().celula("L2").inserirDados("20");
		planilha.selecionar().celula("L3").formula().soma("L1:L2").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(2);
			assertNotNull(row, "Linha 3 deve existir");

			Cell cell = row.getCell(11);
			assertNotNull(cell, "Célula L3 deve existir");

			assertEquals("SUM(L1:L2)", cell.getCellFormula());
		}
	}

	@Test
	@DisplayName("Deve permitir inserir múltiplas fórmulas na mesma planilha")
	void devePermitirMultiplasFormulas() throws IOException, InvalidFormatException {
		// ACT
		planilha.selecionar().celula("M1").inserirDados("10");
		planilha.selecionar().celula("M2").inserirDados("20");
		planilha.selecionar().celula("M3").formula().soma("M1:M2").aplicar();
		planilha.selecionar().celula("M4").formula().media("M1:M2").aplicar();
		planilha.selecionar().celula("M5").formula().maximo("M1:M2").aplicar();

		planilha.salvar(arquivoTeste.getAbsolutePath());

		// ASSERT
		try (Workbook wb = new XSSFWorkbook(arquivoTeste)) {
			Sheet sheet = wb.getSheetAt(0);

			Row row3 = sheet.getRow(2);
			assertNotNull(row3, "Linha 3 deve existir");
			Cell cell3 = row3.getCell(12);
			assertNotNull(cell3, "Célula M3 deve existir");
			assertEquals("SUM(M1:M2)", cell3.getCellFormula());

			Row row4 = sheet.getRow(3);
			assertNotNull(row4, "Linha 4 deve existir");
			Cell cell4 = row4.getCell(12);
			assertNotNull(cell4, "Célula M4 deve existir");
			assertEquals("AVERAGE(M1:M2)", cell4.getCellFormula());

			Row row5 = sheet.getRow(4);
			assertNotNull(row5, "Linha 5 deve existir");
			Cell cell5 = row5.getCell(12);
			assertNotNull(cell5, "Célula M5 deve existir");
			assertEquals("MAX(M1:M2)", cell5.getCellFormula());
		}
	}
}