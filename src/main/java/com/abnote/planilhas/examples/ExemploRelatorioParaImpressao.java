package com.abnote.planilhas.examples;

import java.io.File;

import com.abnote.planilhas.Planilha;
import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Receita: um relatório de vendas pronto para imprimir em uma única página, e
 * já com os valores calculados "congelados" (sem depender de fórmulas) antes
 * de compartilhar o arquivo.
 *
 * <p>
 * Comandos demonstrados: {@code congelarPrimeiraLinha}, {@code
 * filtrosNoCabecalho}, {@code orientacaoPaisagem}, {@code areaDeImpressao},
 * {@code ajustarImpressaoEmPaginas}, {@code colarComoValores}.
 * </p>
 */
public class ExemploRelatorioParaImpressao {

	public static void main(String[] args) {
		final String caminho = new File(System.getProperty("java.io.tmpdir"), "relatorio-vendas.xlsx")
				.getAbsolutePath();

		try (Planilha planilha = Planilha.nova("Vendas do Trimestre")) {

			// 1) Cabeçalho e dados de vendas por vendedor.
			planilha.escreverLinha("A1", "Vendedor", "Janeiro", "Fevereiro", "Março", "Total");
			planilha.adicionarLinha("Ana", 12000.0, 15000.0, 11000.0);
			planilha.adicionarLinha("Bruno", 9800.0, 10200.0, 13500.0);
			planilha.adicionarLinha("Carla", 15600.0, 14100.0, 16700.0);

			// 2) Total por vendedor (soma das 3 colunas de mês).
			planilha.preencherColuna("E", 2, 4, "SUM(B{}:D{})");

			// 3) Formatos e visual.
			planilha.formatarComoMoeda("B2");
			planilha.negrito("A1:E1").corDeFundo("A1:E1", CorEnum.AZUL_MARINHO).corDoTexto("A1:E1", CorEnum.BRANCO);
			planilha.escalaDeCores("E2:E4"); // destaca quem vendeu mais/menos no trimestre

			// 4) Navegação: cabeçalho fixo e filtros.
			planilha.congelarPrimeiraLinha().filtrosNoCabecalho();

			// 5) Configuração de impressão: paisagem, só a tabela, cabendo em 1 página.
			planilha.orientacaoPaisagem();
			planilha.areaDeImpressao("A1:E4");
			planilha.ajustarImpressaoEmPaginas(1, 1);

			// 6) Antes de compartilhar: congela os totais calculados em valores fixos.
			planilha.colarComoValores("E2:E4");

			planilha.ajustarColunas().salvar(caminho);

			System.out.println("Relatório para impressão criado em: " + caminho);
		}
	}
}
