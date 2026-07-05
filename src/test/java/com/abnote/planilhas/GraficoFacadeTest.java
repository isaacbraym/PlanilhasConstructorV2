package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes de criação de gráficos (barras/pizza/linha) na facade.
 */
@DisplayName("Planilha — gráficos")
class GraficoFacadeTest {

	private XSSFChart unicoGrafico(Planilha planilha) {
		return unicoGrafico(planilha, planilha.workbook().getSheetAt(0).getSheetName());
	}

	private XSSFChart unicoGrafico(Planilha planilha, String nomeDaAba) {
		Sheet sheet = planilha.workbook().getSheet(nomeDaAba);
		XSSFDrawing drawing = (XSSFDrawing) sheet.getDrawingPatriarch();
		assertNotNull(drawing, "Deve existir um desenho na aba");
		List<XSSFChart> graficos = drawing.getCharts();
		assertEquals(1, graficos.size());
		return graficos.get(0);
	}

	@Test
	@DisplayName("graficoDeBarras deve criar gráfico de barras verticais com a série correta")
	void deveCriarGraficoDeBarras() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A2", "Jan", "Fev", "Mar").escreverColuna("B2", 100, 200, 150)
					.graficoDeBarras("Vendas", "A2:A4", "B2:B4", "D2");

			XSSFChart chart = unicoGrafico(planilha);
			assertEquals("Vendas", chart.getTitleText().getString());

			List<XDDFChartData> series = chart.getChartSeries();
			assertEquals(1, series.size());
			XDDFChartData dados = series.get(0);
			assertInstanceOf(XDDFBarChartData.class, dados);
			assertEquals(BarDirection.COL, ((XDDFBarChartData) dados).getBarDirection());

			XDDFChartData.Series serie = dados.getSeries(0);
			assertEquals(3, serie.getCategoryData().getPointCount());
			assertEquals("Jan", serie.getCategoryData().getPointAt(0));
			assertEquals(100.0, (Double) serie.getValuesData().getPointAt(0), 0.001);
		}
	}

	@Test
	@DisplayName("graficoDePizza deve criar gráfico sem eixos e com cores variadas")
	void deveCriarGraficoDePizza() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A2", "Norte", "Sul").escreverColuna("B2", 60, 40)
					.graficoDePizza("Regiões", "A2:A3", "B2:B3", "D2");

			XSSFChart chart = unicoGrafico(planilha);
			assertEquals("Regiões", chart.getTitleText().getString());
			assertTrue(chart.getAxes().isEmpty(), "Pizza não deve criar eixos");

			XDDFChartData dados = chart.getChartSeries().get(0);
			assertEquals(2, dados.getSeries(0).getCategoryData().getPointCount());
		}
	}

	@Test
	@DisplayName("graficoDeLinha deve criar gráfico de linha com eixos")
	void deveCriarGraficoDeLinha() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A2", "Sem1", "Sem2").escreverColuna("B2", 10, 20)
					.graficoDeLinha("Progresso", "A2:A3", "B2:B3", "D2");

			XSSFChart chart = unicoGrafico(planilha);
			assertEquals("Progresso", chart.getTitleText().getString());
			assertEquals(2, chart.getAxes().size(), "Linha deve ter eixo de categoria e de valor");
		}
	}

	@Test
	@DisplayName("graficoDeBarras(abaCategorias, abaValores) deve buscar categorias e valores em abas diferentes da aba do gráfico")
	void deveCriarGraficoDeBarrasEntreAbas() throws Exception {
		try (Planilha planilha = Planilha.nova("Produtos")) {
			planilha.escreverColuna("A2", "Caneta", "Caderno", "Lapis");
			planilha.novaAba("Vendas").escreverColuna("B2", 10.0, 20.0, 30.0);
			planilha.novaAba("Dashboard")
					.graficoDeBarras("Vendas por produto", "Produtos", "A2:A4", "Vendas", "B2:B4", "D2");

			XSSFChart chart = unicoGrafico(planilha, "Dashboard");
			XDDFChartData.Series serie = chart.getChartSeries().get(0).getSeries(0);

			assertEquals("Produtos!$A$2:$A$4", serie.getCategoryData().getDataRangeReference());
			assertEquals("Vendas!$B$2:$B$4", serie.getValuesData().getDataRangeReference());
			assertEquals("Caneta", serie.getCategoryData().getPointAt(0));
			assertEquals(10.0, (Double) serie.getValuesData().getPointAt(0), 0.001);
		}
	}

	@Test
	@DisplayName("graficoDeBarras entre abas deve lançar IllegalArgumentException para aba inexistente")
	void deveRecusarAbaInexistenteNoGraficoEntreAbas() throws Exception {
		try (Planilha planilha = Planilha.nova("Produtos")) {
			planilha.escreverColuna("A2", "Caneta").novaAba("Vendas").escreverColuna("B2", 10.0);
			assertThrows(IllegalArgumentException.class,
					() -> planilha.graficoDeBarras("X", "NaoExiste", "A1:A1", "Vendas", "B2:B2", "D2"));
		}
	}
}
