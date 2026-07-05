package com.abnote.planilhas.estilos;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.estilos.estilos.FonteEnum;

/**
 * Testes diretos de {@link EstiloCelula}: cada estilo aplicado à célula certa,
 * ao intervalo certo, ou à planilha inteira — conforme o construtor usado.
 *
 * <p>
 * Antes desta suíte, {@code estilos/} só era exercitado indiretamente via
 * testes de integração da facade.
 * </p>
 */
@DisplayName("EstiloCelula — estilos diretos")
class EstiloCelulaTest {

	private Workbook workbook;
	private Sheet sheet;

	@BeforeEach
	void setUp() {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("T");
		for (int linha = 0; linha < 3; linha++) {
			Row row = sheet.createRow(linha);
			for (int coluna = 0; coluna < 3; coluna++) {
				row.createCell(coluna).setCellValue("v" + linha + coluna);
			}
		}
	}

	@AfterEach
	void tearDown() throws Exception {
		workbook.close();
	}

	private Cell celula(int linha, int coluna) {
		return sheet.getRow(linha).getCell(coluna);
	}

	private XSSFFont fonteDe(Cell cell) {
		return ((XSSFCellStyle) cell.getCellStyle()).getFont();
	}

	// ---------- alvo: célula única x intervalo x planilha inteira ----------

	@Test
	@DisplayName("Construtor de célula única deve estilizar apenas aquela célula")
	void deveEstilizarApenasCelulaUnica() {
		new EstiloCelula(workbook, sheet, 0, 0).aplicarNegrito();

		assertTrue(fonteDe(celula(0, 0)).getBold());
		assertFalse(fonteDe(celula(0, 1)).getBold());
		assertFalse(fonteDe(celula(1, 0)).getBold());
	}

	@Test
	@DisplayName("Construtor de intervalo deve estilizar só as células dentro dele")
	void deveEstilizarApenasIntervalo() {
		new EstiloCelula(workbook, sheet, 0, 0, 1, 1).aplicarNegrito();

		assertTrue(fonteDe(celula(0, 0)).getBold());
		assertTrue(fonteDe(celula(1, 1)).getBold());
		assertFalse(fonteDe(celula(2, 2)).getBold(), "Fora do intervalo não deve mudar");
	}

	@Test
	@DisplayName("Construtor de planilha inteira deve estilizar todas as células existentes")
	void deveEstilizarPlanilhaInteira() {
		new EstiloCelula(workbook, sheet).aplicarNegrito();

		for (int linha = 0; linha < 3; linha++) {
			for (int coluna = 0; coluna < 3; coluna++) {
				assertTrue(fonteDe(celula(linha, coluna)).getBold(), "Célula " + linha + "," + coluna);
			}
		}
	}

	// ---------- fonte: negrito, itálico, sublinhado, tachado ----------

	@Test
	@DisplayName("Deve combinar negrito e itálico na mesma célula sem um apagar o outro")
	void deveCombinarNegritoEItalico() {
		EstiloCelula estilo = new EstiloCelula(workbook, sheet, 0, 0);
		estilo.aplicarNegrito().aplicarItalico();

		XSSFFont fonte = fonteDe(celula(0, 0));
		assertTrue(fonte.getBold());
		assertTrue(fonte.getItalic());
	}

	@Test
	@DisplayName("aplicarSublinhado e aplicarTachado devem marcar os atributos de fonte")
	void deveAplicarSublinhadoETachado() {
		new EstiloCelula(workbook, sheet, 0, 0).aplicarSublinhado().aplicarTachado();

		XSSFFont fonte = fonteDe(celula(0, 0));
		assertNotEquals(org.apache.poi.ss.usermodel.Font.U_NONE, fonte.getUnderline());
		assertTrue(fonte.getStrikeout());
	}

	// ---------- alinhamento ----------

	@Test
	@DisplayName("alinharAEsquerda e alinharADireita devem definir o alinhamento horizontal")
	void deveAlinhar() {
		new EstiloCelula(workbook, sheet, 0, 0).alinharAEsquerda();
		assertEquals(HorizontalAlignment.LEFT, celula(0, 0).getCellStyle().getAlignment());

		new EstiloCelula(workbook, sheet, 0, 1).alinharADireita();
		assertEquals(HorizontalAlignment.RIGHT, celula(0, 1).getCellStyle().getAlignment());
	}

	@Test
	@DisplayName("quebrarTexto deve ativar o quebra-linha da célula")
	void deveQuebrarTexto() {
		new EstiloCelula(workbook, sheet, 0, 0).quebrarTexto();
		assertTrue(celula(0, 0).getCellStyle().getWrapText());
	}

	// ---------- fonte (nome/enum/tamanho) e cores ----------

	@Test
	@DisplayName("fonte(String) e fonte(FonteEnum) devem trocar o nome da fonte")
	void deveTrocarFonte() {
		new EstiloCelula(workbook, sheet, 0, 0).fonte("Comic Sans MS");
		assertEquals("Comic Sans MS", fonteDe(celula(0, 0)).getFontName());

		new EstiloCelula(workbook, sheet, 0, 1).fonte(FonteEnum.CONSOLAS);
		assertEquals("Consolas", fonteDe(celula(0, 1)).getFontName());
	}

	@Test
	@DisplayName("fonteTamanho deve trocar o tamanho da fonte")
	void deveTrocarTamanhoDeFonte() {
		new EstiloCelula(workbook, sheet, 0, 0).fonteTamanho(20);
		assertEquals(20, fonteDe(celula(0, 0)).getFontHeightInPoints());
	}

	@Test
	@DisplayName("corFonte deve aceitar CorEnum, RGB e hexadecimal com o mesmo resultado")
	void deveAplicarCorDeFonte() {
		new EstiloCelula(workbook, sheet, 0, 0).corFonte(CorEnum.AZUL);
		new EstiloCelula(workbook, sheet, 0, 1).corFonte(0, 0, 255);
		new EstiloCelula(workbook, sheet, 0, 2).corFonte("#0000FF");

		for (int coluna = 0; coluna < 3; coluna++) {
			XSSFColor cor = (XSSFColor) fonteDe(celula(0, coluna)).getXSSFColor();
			byte[] rgb = cor.getRGB();
			assertEquals((byte) 0, rgb[0]);
			assertEquals((byte) 0, rgb[1]);
			assertEquals((byte) 255, rgb[2]);
		}
	}

	@Test
	@DisplayName("corDeFundo deve aceitar CorEnum, RGB e hexadecimal com o mesmo resultado")
	void deveAplicarCorDeFundo() {
		new EstiloCelula(workbook, sheet, 0, 0).corDeFundo(CorEnum.VERDE);
		new EstiloCelula(workbook, sheet, 0, 1).corDeFundo(0, 128, 0);
		new EstiloCelula(workbook, sheet, 0, 2).corDeFundo("#008000");

		for (int coluna = 0; coluna < 3; coluna++) {
			XSSFCellStyle estilo = (XSSFCellStyle) celula(0, coluna).getCellStyle();
			assertEquals(FillPatternType.SOLID_FOREGROUND, estilo.getFillPattern());
			byte[] rgb = estilo.getFillForegroundXSSFColor().getRGB();
			assertEquals((byte) 0, rgb[0]);
			assertEquals((byte) 128, rgb[1]);
			assertEquals((byte) 0, rgb[2]);
		}
	}

	@Test
	@DisplayName("corFonte/corDeFundo devem recusar hexadecimal inválido")
	void deveRecusarHexadecimalInvalido() {
		assertThrows(IllegalArgumentException.class,
				() -> new EstiloCelula(workbook, sheet, 0, 0).corFonte("0000FF"));
		assertThrows(IllegalArgumentException.class,
				() -> new EstiloCelula(workbook, sheet, 0, 0).corDeFundo("#GG0000"));
	}

	@Test
	@DisplayName("Estilo com columnIndex -1 deve aplicar fonte e fundo na linha inteira")
	void deveAplicarEstiloNaLinhaInteira() {
		new EstiloCelula(workbook, sheet, 1, -1).aplicarItalico().corDeFundo(CorEnum.AMARELO);

		for (int coluna = 0; coluna < 3; coluna++) {
			assertTrue(fonteDe(celula(1, coluna)).getItalic());
			assertEquals(FillPatternType.SOLID_FOREGROUND, celula(1, coluna).getCellStyle().getFillPattern());
			assertFalse(fonteDe(celula(0, coluna)).getItalic());
		}
	}

	// ---------- bordas ----------

	@Test
	@DisplayName("aplicarTodasAsBordas deve colocar borda fina mesmo em células sem borda prévia")
	void deveAplicarBordasFinasEmCelulasSemBorda() {
		// Regressão: BorderStyleHelper tinha a checagem de "pular borda espessa"
		// invertida e nunca aplicava borda em células sem nenhuma borda prévia.
		new EstiloCelula(workbook, sheet, 0, 0, 1, 1).aplicarTodasAsBordas();

		assertEquals(BorderStyle.THIN, celula(0, 0).getCellStyle().getBorderTop());
		assertEquals(BorderStyle.THIN, celula(1, 1).getCellStyle().getBorderBottom());
	}

	@Test
	@DisplayName("aplicarTodasAsBordas não deve rebaixar uma borda espessa já existente")
	void deveIgnorarCelulaComBordaEspessaAoAplicarBordasFinas() {
		new EstiloCelula(workbook, sheet, 0, 0).aplicarBordasEspessas("A1", "A1");
		new EstiloCelula(workbook, sheet, 0, 0, 0, 0).aplicarTodasAsBordas();

		assertEquals(BorderStyle.THICK, celula(0, 0).getCellStyle().getBorderTop(),
				"Borda espessa não deve virar fina");
	}

	@Test
	@DisplayName("aplicarBordasNaCelula deve aplicar borda apenas na posição informada")
	void deveAplicarBordaEmCelulaEspecifica() {
		new EstiloCelula(workbook, sheet, 0, 0).aplicarBordasNaCelula("B1");

		assertEquals(BorderStyle.THIN, celula(0, 1).getCellStyle().getBorderTop());
		assertEquals(BorderStyle.NONE, celula(0, 0).getCellStyle().getBorderTop());
	}

	@Test
	@DisplayName("aplicarBordasEspessasComInternas deve colocar THICK nas bordas externas e THIN nas internas")
	void deveAplicarBordasEspessasComInternas() {
		new EstiloCelula(workbook, sheet, 0, 0).aplicarBordasEspessasComInternas("A1", "C3");

		// Canto superior esquerdo: THICK em cima e à esquerda (bordas externas).
		CellStyle canto = celula(0, 0).getCellStyle();
		assertEquals(BorderStyle.THICK, canto.getBorderTop());
		assertEquals(BorderStyle.THICK, canto.getBorderLeft());

		// Célula do meio: bordas finas (internas) em todos os lados.
		CellStyle meio = celula(1, 1).getCellStyle();
		assertEquals(BorderStyle.THIN, meio.getBorderTop());
		assertEquals(BorderStyle.THIN, meio.getBorderLeft());
	}

	// ---------- centralização e redimensionamento ----------

	@Test
	@DisplayName("centralizarTudo deve centralizar apenas células com alinhamento GERAL")
	void deveCentralizarSoQuandoAlinhamentoEhGeral() {
		celula(0, 1).getCellStyle(); // GERAL por padrão
		new EstiloCelula(workbook, sheet, 0, 0).alinharAEsquerda(); // define LEFT explicitamente

		new EstiloCelula(workbook, sheet, 0, 0, 2, 2).centralizarTudo();

		assertEquals(HorizontalAlignment.LEFT, celula(0, 0).getCellStyle().getAlignment(),
				"Alinhamento explícito não deve ser sobrescrito");
		assertEquals(HorizontalAlignment.CENTER, celula(0, 1).getCellStyle().getAlignment());
		assertEquals(VerticalAlignment.CENTER, celula(0, 1).getCellStyle().getVerticalAlignment());
	}

	@Test
	@DisplayName("centralizarERedimensionarTudo deve centralizar e redimensionar coluna com fórmula avaliada")
	void deveCentralizarERedimensionarAvaliandoFormula() {
		celula(0, 0).setCellValue("Texto bem longo para redimensionamento");
		celula(0, 1).setCellFormula("A1");
		int larguraAntes = sheet.getColumnWidth(1);

		new EstiloCelula(workbook, sheet).centralizarERedimensionarTudo();

		assertEquals(HorizontalAlignment.CENTER, celula(0, 1).getCellStyle().getAlignment());
		assertTrue(sheet.getColumnWidth(1) > larguraAntes, "Coluna da fórmula deve aumentar após autoSize");
	}

	@Test
	@DisplayName("removerLinhasDeGrade deve desativar as linhas de grade da aba")
	void deveRemoverLinhasDeGrade() {
		new EstiloCelula(workbook, sheet, 0, 0).removerLinhasDeGrade();
		assertFalse(sheet.isDisplayGridlines());
	}

	// ---------- mesclar/duplicar via contornarTudo (planilha inteira) ----------

	@Test
	@DisplayName("contornarTudo deve contornar exatamente a área usada da planilha")
	void deveContornarAreaUsada() {
		new EstiloCelula(workbook, sheet).contornarTudo();

		assertEquals(BorderStyle.THICK, celula(0, 0).getCellStyle().getBorderTop());
		assertEquals(BorderStyle.THICK, celula(2, 2).getCellStyle().getBorderRight());
	}

	@Test
	@DisplayName("contornarTudo em planilha vazia não deve lançar exceção")
	void deveTolerarPlanilhaVaziaAoContornar() {
		Sheet vazia = workbook.createSheet("Vazia");
		assertDoesNotThrow(() -> new EstiloCelula(workbook, vazia).contornarTudo());
	}
}
