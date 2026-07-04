package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Path;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.exceptions.ArquivoException;

/**
 * Testes de abertura e edição de planilhas existentes pela facade.
 */
@DisplayName("Planilha — abrir e editar arquivos existentes")
class AbrirFacadeTest {

	@TempDir
	Path pasta;

	@Test
	@DisplayName("Deve abrir um arquivo existente, ler, editar e salvar de novo")
	void deveAbrirEditarSalvar() throws Exception {
		String caminho = pasta.resolve("dados.xlsx").toString();

		// Cria o arquivo original.
		try (Planilha planilha = Planilha.nova("Dados")) {
			planilha.escrever("A1", "Original").salvar(caminho);
		}

		// Abre, confere o conteúdo e acrescenta uma célula.
		try (Planilha planilha = Planilha.abrir(caminho)) {
			assertEquals("Original",
					planilha.workbook().getSheet("Dados").getRow(0).getCell(0).getStringCellValue());
			planilha.escrever("A2", "Adicionado").salvar(caminho);
		}

		// Reabre e verifica que a edição persistiu.
		try (Planilha planilha = Planilha.abrir(caminho)) {
			assertEquals("Adicionado",
					planilha.workbook().getSheet("Dados").getRow(1).getCell(0).getStringCellValue());
		}
	}

	@Test
	@DisplayName("Abrir arquivo inexistente deve lançar ArquivoException")
	void deveLancarParaArquivoInexistente() {
		String inexistente = pasta.resolve("nao-existe.xlsx").toString();
		assertThrows(ArquivoException.class, () -> Planilha.abrir(inexistente));
	}

	@Test
	@DisplayName("Abrir arquivo .xls (formato antigo) deve lançar ArquivoException amigável, não ClassCastException")
	void deveLancarParaArquivoXlsAntigo() throws Exception {
		String caminhoXls = pasta.resolve("antigo.xls").toString();
		try (HSSFWorkbook workbookAntigo = new HSSFWorkbook()) {
			workbookAntigo.createSheet("Dados");
			try (FileOutputStream saida = new FileOutputStream(new File(caminhoXls))) {
				workbookAntigo.write(saida);
			}
		}

		ArquivoException excecao = assertThrows(ArquivoException.class, () -> Planilha.abrir(caminhoXls));
		assertTrue(excecao.getMessage().contains(".xlsx"), "Mensagem deve orientar a converter para .xlsx");
	}
}
