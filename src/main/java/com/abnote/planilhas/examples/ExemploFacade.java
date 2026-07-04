package com.abnote.planilhas.examples;

import java.io.File;
import java.util.Arrays;

import com.abnote.planilhas.Planilha;
import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Exemplo completo usando a facade amigável {@link Planilha}.
 *
 * <p>
 * Monta uma planilha de "Controle de Vendas" do zero, mostrando os comandos
 * mais usados por quem não programa: escrever, montar tabela, somar, formatar
 * como moeda e estilizar. Rode o {@code main} e abra o arquivo gerado.
 * </p>
 */
public class ExemploFacade {

	public static void main(String[] args) {
		final String caminho = new File(System.getProperty("java.io.tmpdir"), "controle-de-vendas.xlsx").getAbsolutePath();

		try (Planilha planilha = Planilha.nova("Vendas")) {

			// 1) Cabeçalho da tabela.
			planilha.escreverLinha("A1", "Produto", "Preço unitário", "Quantidade");

			// 2) Dados (uma linha por item, acrescentadas em sequência).
			planilha.adicionarLinha("Caneta azul", 2.50, 100);
			planilha.adicionarLinha("Caderno 96 folhas", 15.90, 30);
			planilha.adicionarLinha("Borracha", 1.20, 200);

			// 3) Totalizador: soma das quantidades na linha 5.
			planilha.escrever("A5", "TOTAL");
			planilha.somar("C5", "C2:C4");

			// 4) Formatação de moeda na coluna de preços.
			planilha.formatarComoMoeda("B2");

			// 5) Visual: cabeçalho em destaque, colunas ajustadas, filtros e grade limpa.
			planilha.negrito("A1:C1")
					.corDeFundo("A1:C1", CorEnum.AZUL_MARINHO)
					.corDoTexto("A1:C1", CorEnum.BRANCO)
					.centralizar("A1:C1")
					.negrito("A5:C5")
					.congelarPrimeiraLinha()
					.filtrosNoCabecalho()
					.removerLinhasDeGrade()
					.contornarTudo()
					.ajustarColunas();

			// 6) Salvar.
			planilha.salvar(caminho);

			System.out.println("Planilha criada em: " + caminho);
		}
	}
}
