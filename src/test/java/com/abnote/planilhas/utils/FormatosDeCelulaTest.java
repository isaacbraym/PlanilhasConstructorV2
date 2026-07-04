package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@link FormatosDeCelula}: cache de estilos e aplicação em coluna.
 */
@DisplayName("FormatosDeCelula — cache de estilos comuns")
class FormatosDeCelulaTest {

	private Workbook workbook;
	private Sheet sheet;
	private FormatosDeCelula formatos;

	@BeforeEach
	void setUp() {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("T");
		formatos = new FormatosDeCelula(workbook);
	}

	@AfterEach
	void tearDown() throws Exception {
		workbook.close();
	}

	@Test
	@DisplayName("data() deve retornar o mesmo CellStyle em chamadas repetidas (cache)")
	void deveCachearEstiloDeData() {
		CellStyle primeira = formatos.data();
		CellStyle segunda = formatos.data();
		assertSame(primeira, segunda, "Deve reutilizar o mesmo CellStyle");
		assertTrue(primeira.getDataFormatString().contains("dd"));
	}

	@Test
	@DisplayName("dataHora() e porcentagem() devem ter formatos distintos e cacheados")
	void deveCachearOutrosFormatos() {
		assertSame(formatos.dataHora(), formatos.dataHora());
		assertSame(formatos.porcentagem(), formatos.porcentagem());
		assertTrue(formatos.dataHora().getDataFormatString().contains("HH"));
		assertTrue(formatos.porcentagem().getDataFormatString().contains("%"));
		assertNotEquals(formatos.data(), formatos.dataHora());
	}

	@Test
	@DisplayName("aplicarNaColuna deve aplicar o estilo apenas às células existentes da coluna")
	void deveAplicarNaColuna() {
		Row linha0 = sheet.createRow(0);
		linha0.createCell(0).setCellValue(1);
		Row linha1 = sheet.createRow(1);
		linha1.createCell(0).setCellValue(2);
		// Linha 2 propositalmente não criada (deve ser ignorada sem erro).

		FormatosDeCelula.aplicarNaColuna(sheet, 0, 0, formatos.porcentagem());

		assertEquals(formatos.porcentagem(), sheet.getRow(0).getCell(0).getCellStyle());
		assertEquals(formatos.porcentagem(), sheet.getRow(1).getCell(0).getCellStyle());
	}

	@Test
	@DisplayName("aplicarNaColuna deve respeitar a primeira linha informada")
	void deveRespeitarPrimeiraLinha() {
		Cell antes = sheet.createRow(0).createCell(0);
		Cell depois = sheet.createRow(1).createCell(0);
		CellStyle estiloOriginal = antes.getCellStyle();

		FormatosDeCelula.aplicarNaColuna(sheet, 0, 1, formatos.data());

		assertEquals(estiloOriginal, antes.getCellStyle(), "Linha antes do início não deve mudar");
		assertEquals(formatos.data(), depois.getCellStyle());
	}
}
