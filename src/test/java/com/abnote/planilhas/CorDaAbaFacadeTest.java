package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Testes de {@code corDaAba} na facade.
 */
@DisplayName("Planilha — corDaAba")
class CorDaAbaFacadeTest {

	@Test
	@DisplayName("Deve aplicar a cor RGB correta na aba")
	void deveAplicarCorNaAba() throws Exception {
		try (Planilha planilha = Planilha.nova("Vendas")) {
			planilha.corDaAba(CorEnum.VERDE);

			XSSFSheet sheet = (XSSFSheet) planilha.workbook().getSheetAt(0);
			XSSFColor cor = sheet.getTabColor();
			assertNotNull(cor);
			byte[] rgb = cor.getRGB();
			assertEquals((byte) CorEnum.VERDE.getRed(), rgb[0]);
			assertEquals((byte) CorEnum.VERDE.getGreen(), rgb[1]);
			assertEquals((byte) CorEnum.VERDE.getBlue(), rgb[2]);
		}
	}
}
