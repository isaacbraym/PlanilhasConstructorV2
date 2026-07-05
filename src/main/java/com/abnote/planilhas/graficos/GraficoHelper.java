package com.abnote.planilhas.graficos;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Cria gráficos (barras, pizza, linha) em uma planilha, a partir de intervalos
 * de categorias e valores.
 */
public final class GraficoHelper {

	/** Largura padrão do gráfico, em número de colunas ocupadas. */
	private static final int LARGURA_PADRAO_COLUNAS = 9;

	/** Altura padrão do gráfico, em número de linhas ocupadas. */
	private static final int ALTURA_PADRAO_LINHAS = 15;

	private GraficoHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Cria um gráfico de barras verticais.
	 *
	 * @param sheet         A folha onde o gráfico será desenhado.
	 * @param titulo        Título do gráfico (também usado como nome da série).
	 * @param categorias    Intervalo com os nomes das categorias (eixo X).
	 * @param valores       Intervalo com os valores numéricos (eixo Y).
	 * @param colunaAncora  Coluna (0-based) do canto superior esquerdo do gráfico.
	 * @param linhaAncora   Linha (0-based) do canto superior esquerdo do gráfico.
	 */
	public static void criarGraficoDeBarras(final XSSFSheet sheet, final String titulo,
			final CellRangeAddress categorias, final CellRangeAddress valores, final int colunaAncora,
			final int linhaAncora) {
		criarGraficoDeBarras(sheet, titulo, sheet, categorias, sheet, valores, colunaAncora, linhaAncora);
	}

	/**
	 * Cria um gráfico de barras verticais com categorias e valores vindos de
	 * abas diferentes entre si (e/ou diferentes da aba onde o gráfico é
	 * desenhado) — útil para um "dashboard" que resume dados de outras abas.
	 *
	 * @param sheet          A folha onde o gráfico será desenhado (ancoragem).
	 * @param titulo         Título do gráfico (também usado como nome da série).
	 * @param folhaCategorias Folha onde está o intervalo de categorias.
	 * @param categorias     Intervalo com os nomes das categorias (eixo X).
	 * @param folhaValores   Folha onde está o intervalo de valores.
	 * @param valores        Intervalo com os valores numéricos (eixo Y).
	 * @param colunaAncora   Coluna (0-based) do canto superior esquerdo do gráfico.
	 * @param linhaAncora    Linha (0-based) do canto superior esquerdo do gráfico.
	 */
	public static void criarGraficoDeBarras(final XSSFSheet sheet, final String titulo,
			final XSSFSheet folhaCategorias, final CellRangeAddress categorias, final XSSFSheet folhaValores,
			final CellRangeAddress valores, final int colunaAncora, final int linhaAncora) {
		final XSSFChart chart = criarChartVazio(sheet, titulo, colunaAncora, linhaAncora);
		final XDDFCategoryAxis eixoCategorias = chart.createCategoryAxis(AxisPosition.BOTTOM);
		final XDDFValueAxis eixoValores = chart.createValueAxis(AxisPosition.LEFT);

		final XDDFChartData dados = chart.createData(ChartTypes.BAR, eixoCategorias, eixoValores);
		((XDDFBarChartData) dados).setBarDirection(BarDirection.COL);

		adicionarSerie(dados, folhaCategorias, categorias, folhaValores, valores, titulo);
		chart.plot(dados);
	}

	/**
	 * Cria um gráfico de pizza.
	 *
	 * @param sheet         A folha onde o gráfico será desenhado.
	 * @param titulo        Título do gráfico.
	 * @param categorias    Intervalo com os nomes das fatias.
	 * @param valores       Intervalo com os valores numéricos de cada fatia.
	 * @param colunaAncora  Coluna (0-based) do canto superior esquerdo do gráfico.
	 * @param linhaAncora   Linha (0-based) do canto superior esquerdo do gráfico.
	 */
	public static void criarGraficoDePizza(final XSSFSheet sheet, final String titulo,
			final CellRangeAddress categorias, final CellRangeAddress valores, final int colunaAncora,
			final int linhaAncora) {
		criarGraficoDePizza(sheet, titulo, sheet, categorias, sheet, valores, colunaAncora, linhaAncora);
	}

	/**
	 * Cria um gráfico de pizza com categorias e valores vindos de abas
	 * diferentes entre si (e/ou diferentes da aba onde o gráfico é desenhado).
	 *
	 * @param sheet          A folha onde o gráfico será desenhado (ancoragem).
	 * @param titulo         Título do gráfico.
	 * @param folhaCategorias Folha onde está o intervalo de categorias.
	 * @param categorias     Intervalo com os nomes das fatias.
	 * @param folhaValores   Folha onde está o intervalo de valores.
	 * @param valores        Intervalo com os valores numéricos de cada fatia.
	 * @param colunaAncora   Coluna (0-based) do canto superior esquerdo do gráfico.
	 * @param linhaAncora    Linha (0-based) do canto superior esquerdo do gráfico.
	 */
	public static void criarGraficoDePizza(final XSSFSheet sheet, final String titulo,
			final XSSFSheet folhaCategorias, final CellRangeAddress categorias, final XSSFSheet folhaValores,
			final CellRangeAddress valores, final int colunaAncora, final int linhaAncora) {
		final XSSFChart chart = criarChartVazio(sheet, titulo, colunaAncora, linhaAncora);
		// Pizza não usa eixos — POI espera null nos dois parâmetros de eixo.
		final XDDFChartData dados = chart.createData(ChartTypes.PIE, null, null);
		dados.setVaryColors(true); // cada fatia com uma cor diferente

		adicionarSerie(dados, folhaCategorias, categorias, folhaValores, valores, titulo);
		chart.plot(dados);
	}

	/**
	 * Cria um gráfico de linha.
	 *
	 * @param sheet         A folha onde o gráfico será desenhado.
	 * @param titulo        Título do gráfico (também usado como nome da série).
	 * @param categorias    Intervalo com os nomes das categorias (eixo X).
	 * @param valores       Intervalo com os valores numéricos (eixo Y).
	 * @param colunaAncora  Coluna (0-based) do canto superior esquerdo do gráfico.
	 * @param linhaAncora   Linha (0-based) do canto superior esquerdo do gráfico.
	 */
	public static void criarGraficoDeLinha(final XSSFSheet sheet, final String titulo,
			final CellRangeAddress categorias, final CellRangeAddress valores, final int colunaAncora,
			final int linhaAncora) {
		criarGraficoDeLinha(sheet, titulo, sheet, categorias, sheet, valores, colunaAncora, linhaAncora);
	}

	/**
	 * Cria um gráfico de linha com categorias e valores vindos de abas
	 * diferentes entre si (e/ou diferentes da aba onde o gráfico é desenhado).
	 *
	 * @param sheet          A folha onde o gráfico será desenhado (ancoragem).
	 * @param titulo         Título do gráfico (também usado como nome da série).
	 * @param folhaCategorias Folha onde está o intervalo de categorias.
	 * @param categorias     Intervalo com os nomes das categorias (eixo X).
	 * @param folhaValores   Folha onde está o intervalo de valores.
	 * @param valores        Intervalo com os valores numéricos (eixo Y).
	 * @param colunaAncora   Coluna (0-based) do canto superior esquerdo do gráfico.
	 * @param linhaAncora    Linha (0-based) do canto superior esquerdo do gráfico.
	 */
	public static void criarGraficoDeLinha(final XSSFSheet sheet, final String titulo,
			final XSSFSheet folhaCategorias, final CellRangeAddress categorias, final XSSFSheet folhaValores,
			final CellRangeAddress valores, final int colunaAncora, final int linhaAncora) {
		final XSSFChart chart = criarChartVazio(sheet, titulo, colunaAncora, linhaAncora);
		final XDDFCategoryAxis eixoCategorias = chart.createCategoryAxis(AxisPosition.BOTTOM);
		final XDDFValueAxis eixoValores = chart.createValueAxis(AxisPosition.LEFT);

		final XDDFChartData dados = chart.createData(ChartTypes.LINE, eixoCategorias, eixoValores);
		adicionarSerie(dados, folhaCategorias, categorias, folhaValores, valores, titulo);
		chart.plot(dados);
	}

	private static XSSFChart criarChartVazio(final XSSFSheet sheet, final String titulo, final int colunaAncora,
			final int linhaAncora) {
		final XSSFDrawing desenho = sheet.createDrawingPatriarch();
		final XSSFClientAnchor ancora = desenho.createAnchor(0, 0, 0, 0, colunaAncora, linhaAncora,
				colunaAncora + LARGURA_PADRAO_COLUNAS, linhaAncora + ALTURA_PADRAO_LINHAS);
		final XSSFChart chart = desenho.createChart(ancora);
		chart.setTitleText(titulo);
		chart.setTitleOverlay(false);
		chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);
		return chart;
	}

	private static void adicionarSerie(final XDDFChartData dados, final XSSFSheet folhaCategorias,
			final CellRangeAddress categorias, final XSSFSheet folhaValores, final CellRangeAddress valores,
			final String titulo) {
		final XDDFCategoryDataSource fonteCategorias = XDDFDataSourcesFactory.fromStringCellRange(folhaCategorias,
				categorias);
		final XDDFNumericalDataSource<Double> fonteValores = XDDFDataSourcesFactory.fromNumericCellRange(folhaValores,
				valores);
		final XDDFChartData.Series serie = dados.addSeries(fonteCategorias, fonteValores);
		serie.setTitle(titulo, null);
	}
}
