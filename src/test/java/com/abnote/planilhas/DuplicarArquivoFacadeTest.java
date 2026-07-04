package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.exceptions.ArquivoException;

/**
 * Testes de {@code Planilha.duplicarArquivo}.
 */
@DisplayName("Planilha — duplicarArquivo")
class DuplicarArquivoFacadeTest {

	@TempDir
	Path pasta;

	@Test
	@DisplayName("Deve criar uma cópia completa e independente do arquivo")
	void deveDuplicarArquivo() throws Exception {
		String original = pasta.resolve("original.xlsx").toString();
		String copia = pasta.resolve("subpasta/copia.xlsx").toString();
		new File(pasta.resolve("subpasta").toString()).mkdirs();

		try (Planilha planilha = Planilha.nova("Dados")) {
			planilha.escrever("A1", "Conteúdo original").negrito("A1").salvar(original);
		}

		Planilha.duplicarArquivo(original, copia);

		assertTrue(new File(copia).exists());
		try (XSSFWorkbook wb = new XSSFWorkbook(new File(copia))) {
			assertEquals("Conteúdo original", wb.getSheet("Dados").getRow(0).getCell(0).getStringCellValue());
		}

		// Editar a cópia não deve afetar o original.
		try (Planilha copiaAberta = Planilha.abrir(copia)) {
			copiaAberta.escrever("A2", "Só na cópia").salvar(copia);
		}
		try (XSSFWorkbook wbOriginal = new XSSFWorkbook(new File(original))) {
			assertNull(wbOriginal.getSheet("Dados").getRow(1), "Original não deve ter sido alterado");
		}
	}

	@Test
	@DisplayName("Deve lançar ArquivoException se a origem não existir")
	void deveLancarParaOrigemInexistente() {
		String inexistente = pasta.resolve("nao-existe.xlsx").toString();
		String destino = pasta.resolve("destino.xlsx").toString();
		assertThrows(ArquivoException.class, () -> Planilha.duplicarArquivo(inexistente, destino));
	}
}
