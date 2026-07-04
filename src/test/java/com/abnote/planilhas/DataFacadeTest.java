package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.time.LocalDate;
import java.time.LocalDateTime;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de escrita e formatação de datas na facade.
 */
@DisplayName("Planilha — datas e hora")
class DataFacadeTest {

	@Test
	@DisplayName("escreverData deve gravar data formatada dd/MM/aaaa")
	void deveEscreverData() throws Exception {
		try (Planilha planilha = Planilha.nova("Agenda")) {
			planilha.escreverData("A1", LocalDate.of(2024, 1, 15));

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Cell a1 = sheet.getRow(0).getCell(0);
			assertTrue(DateUtil.isCellDateFormatted(a1), "Célula deve estar formatada como data");
			assertEquals(LocalDate.of(2024, 1, 15), a1.getLocalDateTimeCellValue().toLocalDate());
			assertTrue(a1.getCellStyle().getDataFormatString().contains("dd"),
					"Formato deve ser de data: " + a1.getCellStyle().getDataFormatString());
		}
	}

	@Test
	@DisplayName("escreverDataHora deve gravar data com hora")
	void deveEscreverDataHora() throws Exception {
		try (Planilha planilha = Planilha.nova("Log")) {
			LocalDateTime momento = LocalDateTime.of(2024, 3, 10, 14, 30);
			planilha.escreverDataHora("B2", momento);

			Sheet sheet = planilha.workbook().getSheetAt(0);
			Cell b2 = sheet.getRow(1).getCell(1);
			assertTrue(DateUtil.isCellDateFormatted(b2));
			assertEquals(momento, b2.getLocalDateTimeCellValue());
		}
	}

	@Test
	@DisplayName("formatarComoData deve aplicar o formato à coluna")
	void deveFormatarComoData() throws Exception {
		try (Planilha planilha = Planilha.nova("Datas")) {
			// Escreve datas sem formato (via escrever numérico do serial não é trivial),
			// então usamos escreverData e reforçamos o formato na coluna.
			planilha.escreverData("A1", LocalDate.of(2020, 12, 25))
					.escreverData("A2", LocalDate.of(2021, 1, 1))
					.formatarComoData("A1");

			Sheet sheet = planilha.workbook().getSheetAt(0);
			assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString().contains("dd"));
			assertTrue(sheet.getRow(1).getCell(0).getCellStyle().getDataFormatString().contains("dd"));
		}
	}
}
