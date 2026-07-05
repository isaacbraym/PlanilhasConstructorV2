package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;
import java.time.Duration;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Testes de carga da facade: não são benchmarks rígidos, mas protegem contra
 * regressões óbvias ao lidar com milhares de linhas.
 */
@DisplayName("Planilha — carga com milhares de linhas")
class CargaFacadeTest {

	@Test
	@DisplayName("Deve escrever, totalizar, salvar e reabrir milhares de linhas")
	void deveSalvarEReabrirMilharesDeLinhas(@TempDir Path pasta) {
		assertTimeout(Duration.ofSeconds(30), () -> {
			File arquivo = pasta.resolve("carga.xlsx").toFile();
			try (Planilha planilha = Planilha.nova("Carga")) {
				planilha.escreverLinha("A1", "Item", "Quantidade", "Valor");
				for (int indice = 1; indice <= 3000; indice++) {
					planilha.adicionarLinha("Item " + indice, indice, indice * 2);
				}
				planilha.somar("B3002", "B2:B3001")
						.somar("C3002", "C2:C3001")
						.congelarPrimeiraLinha()
						.filtrosNoCabecalho()
						.salvar(arquivo.getAbsolutePath());
			}

			try (Workbook workbook = new XSSFWorkbook(arquivo)) {
				Sheet sheet = workbook.getSheet("Carga");
				assertEquals(3001, sheet.getLastRowNum());
				assertEquals("Item 3000", sheet.getRow(3000).getCell(0).getStringCellValue());
				assertTrue(sheet.getPaneInformation().isFreezePane());
				assertEquals("A1:C1", ((XSSFSheet) sheet).getCTWorksheet().getAutoFilter().getRef());

				FormulaEvaluator avaliador = workbook.getCreationHelper().createFormulaEvaluator();
				Cell totalQuantidade = sheet.getRow(3001).getCell(1);
				Cell totalValor = sheet.getRow(3001).getCell(2);
				assertEquals("SUM(B2:B3001)", totalQuantidade.getCellFormula());
				assertEquals(4501500.0, avaliador.evaluate(totalQuantidade).getNumberValue(), 0.001);
				assertEquals(9003000.0, avaliador.evaluate(totalValor).getNumberValue(), 0.001);
			}
		});
	}

	@Test
	@DisplayName("Deve ordenar, buscar e copiar milhares de linhas")
	void deveOrdenarBuscarECopiarMilharesDeLinhas() {
		assertTimeout(Duration.ofSeconds(30), () -> {
			try (Planilha planilha = Planilha.nova("Pedidos")) {
				planilha.escreverLinha("A1", "Codigo", "Status");
				for (int codigo = 2000; codigo >= 1; codigo--) {
					String status = codigo % 2 == 0 ? "PAR" : "IMPAR";
					planilha.adicionarLinha(codigo, status);
				}

				planilha.ordenarPorCrescente("A").copiarLinhasParaAba("B", "PAR", "Pares");
				List<Integer> linhasPares = planilha.buscarLinhas("B", "PAR");
				Sheet pedidos = planilha.workbook().getSheet("Pedidos");
				Sheet pares = planilha.workbook().getSheet("Pares");

				assertEquals(1000, linhasPares.size());
				assertEquals(1.0, pedidos.getRow(1).getCell(0).getNumericCellValue(), 0.001);
				assertEquals(2000.0, pedidos.getRow(2000).getCell(0).getNumericCellValue(), 0.001);
				assertEquals(1000, pares.getPhysicalNumberOfRows());
				assertEquals("PAR", pares.getRow(0).getCell(1).getStringCellValue());
			}
		});
	}
}
