package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Testes de comentários (notas) em células na facade.
 */
@DisplayName("Planilha — comentário em célula")
class ComentarioFacadeTest {

	@Test
	@DisplayName("Deve anexar o comentário com o texto correto à célula")
	void deveAdicionarComentario() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("B2", 100).comentario("B2", "Meta mensal de vendas");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Cell celula = sheet.getRow(1).getCell(1);
			Comment comentario = celula.getCellComment();

			assertNotNull(comentario);
			assertEquals("Meta mensal de vendas", comentario.getString().getString());
		}
	}

	@Test
	@DisplayName("Comentário deve sobreviver a salvar em disco e reabrir")
	void deveSobreviverAoSalvarEReabrir(@TempDir Path pasta) throws Exception {
		String caminho = pasta.resolve("com-comentario.xlsx").toString();
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.comentario("A1", "Não editar esta célula").salvar(caminho);
		}
		try (XSSFWorkbook wb = new XSSFWorkbook(new java.io.File(caminho))) {
			Comment comentario = wb.getSheetAt(0).getRow(0).getCell(0).getCellComment();
			assertNotNull(comentario);
			assertEquals("Não editar esta célula", comentario.getString().getString());
		}
	}
}
