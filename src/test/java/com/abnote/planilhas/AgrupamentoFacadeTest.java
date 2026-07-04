package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Testes de agrupamento de linhas/colunas (outline) na facade.
 *
 * <p>
 * Não testamos "recolher automaticamente": empiricamente,
 * {@code Sheet.setRowGroupCollapsed} marcou TODAS as linhas da planilha como
 * ocultas (não só o grupo) neste ambiente — um risco real de esconder dados
 * do usuário. Por isso a facade só expõe o agrupamento em si (o "+"/"-" que o
 * próprio usuário clica no Excel), nunca o recolhimento automático. Ver
 * AGENTS.md para os detalhes da investigação.
 * </p>
 */
@DisplayName("Planilha — agrupamento de linhas/colunas")
class AgrupamentoFacadeTest {

	@TempDir
	Path pasta;

	@Test
	@DisplayName("agruparLinhas deve marcar o outlineLevel das linhas do grupo, sem esconder nenhuma")
	void deveAgruparLinhasSemEsconderNada() throws Exception {
		String caminho = pasta.resolve("agrupado.xlsx").toString();
		try (Planilha planilha = Planilha.nova("T")) {
			for (int linha = 1; linha <= 6; linha++) {
				planilha.escrever("A" + linha, "linha" + linha);
			}
			planilha.agruparLinhas(2, 4).salvar(caminho);
		}

		try (XSSFWorkbook wb = new XSSFWorkbook(new File(caminho))) {
			Sheet sheet = wb.getSheetAt(0);
			for (int indice = 0; indice < 6; indice++) {
				Row linha = sheet.getRow(indice);
				int esperado = (indice >= 1 && indice <= 3) ? 1 : 0;
				assertEquals(esperado, linha.getOutlineLevel(), "Linha " + (indice + 1));
				assertFalse(linha.getZeroHeight(), "Linha " + (indice + 1) + " não deve ficar escondida");
			}
		}
	}

	@Test
	@DisplayName("agruparColunas deve marcar o outlineLevel das colunas do grupo")
	void deveAgruparColunas() throws Exception {
		String caminho = pasta.resolve("agrupado-colunas.xlsx").toString();
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "a", "b", "c", "d").agruparColunas("B", "C").salvar(caminho);
		}

		try (XSSFWorkbook wb = new XSSFWorkbook(new File(caminho))) {
			Sheet sheet = wb.getSheetAt(0);
			assertEquals(1, sheet.getColumnOutlineLevel(1)); // coluna B
			assertEquals(1, sheet.getColumnOutlineLevel(2)); // coluna C
			assertEquals(0, sheet.getColumnOutlineLevel(0)); // coluna A, fora do grupo
			assertEquals(0, sheet.getColumnOutlineLevel(3)); // coluna D, fora do grupo
		}
	}
}
