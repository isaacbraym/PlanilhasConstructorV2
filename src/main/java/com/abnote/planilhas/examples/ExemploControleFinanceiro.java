package com.abnote.planilhas.examples;

import java.io.File;

import com.abnote.planilhas.Planilha;
import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Receita: orçamento mensal com totais automáticos, destaque de categorias
 * estouradas e um gráfico comparando orçado x gasto.
 *
 * <p>
 * Comandos demonstrados: {@code adicionarTotais}, {@code formatarComoMoeda},
 * {@code realcarSeMenorQue}, {@code graficoDeBarras}, {@code preencherColuna}.
 * </p>
 */
public class ExemploControleFinanceiro {

	public static void main(String[] args) {
		final String caminho = new File(System.getProperty("java.io.tmpdir"), "controle-financeiro.xlsx")
				.getAbsolutePath();

		try (Planilha planilha = Planilha.nova("Orçamento")) {

			// 1) Cabeçalho e dados de cada categoria.
			planilha.escreverLinha("A1", "Categoria", "Orçado", "Gasto", "Diferença");
			planilha.adicionarLinha("Moradia", 1500.00, 1500.00);
			planilha.adicionarLinha("Alimentação", 800.00, 950.00); // estourou
			planilha.adicionarLinha("Transporte", 400.00, 320.00);
			planilha.adicionarLinha("Lazer", 300.00, 410.00); // estourou
			planilha.adicionarLinha("Saúde", 200.00, 150.00);

			// 2) Diferença = Orçado - Gasto, linha por linha.
			planilha.preencherColuna("D", 2, 6, "B{}-C{}");

			// 3) Totais automáticos (soma Orçado, Gasto e Diferença).
			planilha.adicionarTotais("A1");

			// 4) Formatos e destaque visual.
			planilha.formatarComoMoeda("B2");
			planilha.formatarComoMoeda("C2");
			planilha.formatarComoMoeda("D2");
			planilha.realcarSeMenorQue("D2:D6", 0, CorEnum.VERMELHO_ESCURO); // categorias estouradas
			planilha.negrito("A1:D1").corDeFundo("A1:D1", CorEnum.AZUL_MARINHO).corDoTexto("A1:D1", CorEnum.BRANCO);
			planilha.negrito("A7:D7"); // linha de totais

			// 5) Gráfico do gasto por categoria (graficoDeBarras plota uma série por vez;
			// para comparar orçado x gasto lado a lado, crie dois gráficos separados).
			planilha.graficoDeBarras("Gasto por categoria", "A2:A6", "C2:C6", "F2");

			// 6) Acabamento.
			planilha.congelarPrimeiraLinha().ajustarColunas().salvar(caminho);

			System.out.println("Controle financeiro criado em: " + caminho);
		}
	}
}
