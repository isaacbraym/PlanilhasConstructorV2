package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de {@code formatarComoPersonalizado} na facade.
 */
@DisplayName("Planilha — formatarComoPersonalizado")
class FormatoPersonalizadoFacadeTest {

	@Test
	@DisplayName("Deve aplicar o formato numérico personalizado informado")
	void deveAplicarFormatoPersonalizado() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", 12.5).formatarComoPersonalizado("A1", "0.00 \"kg\"");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals("0.00 \"kg\"", sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString());
		}
	}

	@Test
	@DisplayName("Deve cachear o mesmo CellStyle para o mesmo formato")
	void deveCachearEstiloPorFormato() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", 1).escrever("A2", 2);
			planilha.formatarComoPersonalizado("A1", "0.0%");
			planilha.formatarComoPersonalizado("A2", "0.0%");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertEquals(sheet.getRow(0).getCell(0).getCellStyle(), sheet.getRow(1).getCell(0).getCellStyle());
		}
	}
}
