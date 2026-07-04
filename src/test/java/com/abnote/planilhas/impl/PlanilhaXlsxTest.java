package com.abnote.planilhas.impl;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.interfaces.IPlanilha;

/**
 * Testes de integração para PlanilhaXlsx.
 * 
 * Testa fluxos completos de uso da biblioteca.
 */
class PlanilhaXlsxTest {

	@TempDir
	Path pastaTemporaria; // JUnit cria pasta temporária automaticamente

	/**
	 * Teste: Fluxo básico - criar, inserir, salvar.
	 */
	@Test
	@DisplayName("Deve criar planilha, inserir dados e salvar arquivo")
	void deveCriarInserirESalvar() throws Exception {
		// ARRANGE
		String nomeArquivo = pastaTemporaria.resolve("teste_basico.xlsx").toString();

		// ACT
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("MinhaPlanilha");
			planilha.selecionar().celula("A1").inserirDados("Nome");
			planilha.selecionar().celula("A2").inserirDados("João Silva");
			planilha.salvar(nomeArquivo);
		}

		// ASSERT
		File arquivo = new File(nomeArquivo);
		assertTrue(arquivo.exists(), "Arquivo deve ter sido criado");
		assertTrue(arquivo.length() > 0, "Arquivo não deve estar vazio");

		// Verifica conteúdo
		try (Workbook wb = new XSSFWorkbook(arquivo)) {
			Sheet sheet = wb.getSheetAt(0);
			assertEquals("MinhaPlanilha", sheet.getSheetName());

			Cell cellA1 = sheet.getRow(0).getCell(0);
			assertEquals("Nome", cellA1.getStringCellValue());

			Cell cellA2 = sheet.getRow(1).getCell(0);
			assertEquals("João Silva", cellA2.getStringCellValue());
		}
	}

	/**
	 * Teste: Múltiplas abas.
	 */
	@Test
	@DisplayName("Deve criar múltiplas abas e alternar entre elas")
	void deveCriarMultiplasAbas() throws Exception {
		String nomeArquivo = pastaTemporaria.resolve("teste_multiplas_abas.xlsx").toString();

		try (IPlanilha planilha = new PlanilhaXlsx()) {
			// Aba 1
			planilha.criarPlanilha("Vendas");
			planilha.selecionar().celula("A1").inserirDados("Vendas 2024");

			// Aba 2
			planilha.criarSheet("Despesas");
			planilha.selecionarSheet("Despesas");
			planilha.selecionar().celula("A1").inserirDados("Despesas 2024");

			planilha.salvar(nomeArquivo);
		}

		// Verifica
		try (Workbook wb = new XSSFWorkbook(new File(nomeArquivo))) {
			assertEquals(2, wb.getNumberOfSheets(), "Deve ter 2 abas");
			assertEquals("Vendas", wb.getSheetAt(0).getSheetName());
			assertEquals("Despesas", wb.getSheetAt(1).getSheetName());

			assertEquals("Vendas 2024", wb.getSheet("Vendas").getRow(0).getCell(0).getStringCellValue());
			assertEquals("Despesas 2024", wb.getSheet("Despesas").getRow(0).getCell(0).getStringCellValue());
		}
	}

	/**
	 * Teste: API fluente funcionando.
	 */
	@Test
	@DisplayName("Deve encadear operações de forma fluente")
	void deveEncadearOperacoesFluentes() throws Exception {
		String nomeArquivo = pastaTemporaria.resolve("teste_fluente.xlsx").toString();

		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("Teste");

			// Encadeamento: inserir + converter + somar
			planilha.selecionar().celula("A1").inserir(100.0);
			planilha.selecionar().celula("A2").inserir(200.0);
			planilha.selecionar().celula("A3").inserir(300.0);

			planilha.converter().emContabil("A1").somarColuna("A1");

			planilha.salvar(nomeArquivo);
		}

		// Verifica que arquivo foi criado
		assertTrue(new File(nomeArquivo).exists());
	}

	/**
	 * Teste: AutoCloseable fecha workbook automaticamente.
	 */
	@Test
	@DisplayName("Deve fechar workbook automaticamente com try-with-resources")
	void deveFecharWorkbookAutomaticamente() throws Exception {
		IPlanilha planilha = new PlanilhaXlsx();
		planilha.criarPlanilha("Teste");

		// Fecha manualmente
		assertDoesNotThrow(() -> planilha.close(), "close() não deve lançar exceção");

		// Fechar novamente não deve dar erro
		assertDoesNotThrow(() -> planilha.close(), "close() múltiplo não deve lançar exceção");
	}

	/**
	 * Teste: Inserir dados de lista.
	 */
	@Test
	@DisplayName("Deve inserir dados de lista com delimitador")
	void deveInserirDadosComDelimitador() throws Exception {
		String nomeArquivo = pastaTemporaria.resolve("teste_delimitador.xlsx").toString();

		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("Dados");

			String dados = "Nome,Idade,Cidade";
			planilha.selecionar().celula("A1").inserirDados(dados, ",");

			planilha.salvar(nomeArquivo);
		}

		// Verifica
		try (Workbook wb = new XSSFWorkbook(new File(nomeArquivo))) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(0);

			assertEquals("Nome", row.getCell(0).getStringCellValue());
			assertEquals("Idade", row.getCell(1).getStringCellValue());
			assertEquals("Cidade", row.getCell(2).getStringCellValue());
		}
	}

	/**
	 * Teste: Erro ao selecionar aba inexistente.
	 */
	@Test
	@DisplayName("Deve lançar exceção ao selecionar aba inexistente")
	void deveLancarExcecaoAbaInexistente() {
		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("Aba1");

			// Tentar selecionar aba que não existe
			assertThrows(IllegalArgumentException.class, () -> planilha.selecionarSheet("AbaInexistente"),
					"Deve lançar IllegalArgumentException");
		} catch (Exception e) {
			fail("Não deveria lançar exceção no try-with-resources");
		}
	}

	/**
	 * Teste: Salvar sem definir diretório usa caminho completo.
	 */
	@Test
	@DisplayName("Deve salvar com caminho completo sem definir diretório")
	void deveSalvarComCaminhoCompleto() throws Exception {
		String caminhoCompleto = pastaTemporaria.resolve("teste_caminho_completo.xlsx").toString();

		try (IPlanilha planilha = new PlanilhaXlsx()) {
			planilha.criarPlanilha("Teste");
			planilha.selecionar().celula("A1").inserirDados("Teste");

			// Não definir diretorioSaida - usar caminho completo direto
			planilha.salvar(caminhoCompleto);
		}

		assertTrue(new File(caminhoCompleto).exists());
	}
}