package com.abnote.planilhas.interfaces;

import com.abnote.planilhas.estilos.EstiloCelula;

/**
 * Interface que define os métodos para aplicação de estilos na planilha.
 */
public interface IEstilos {

	/**
	 * Centraliza todas as células definidas.
	 *
	 * @return A instância de EstiloCelula para encadeamento.
	 */
	EstiloCelula centralizarTudo();

	/**
	 * Redimensiona as colunas com base no conteúdo.
	 *
	 * @return A instância de EstiloCelula para encadeamento.
	 */
	EstiloCelula redimensionarColunas();

	/**
	 * Remove as linhas de grade da planilha.
	 *
	 * @return A instância de EstiloCelula para encadeamento.
	 */
	EstiloCelula removerLinhasDeGrade();

	/**
	 * Aplica estilos na célula atual.
	 *
	 * @return A instância de EstiloCelula para encadeamento.
	 */
	EstiloCelula aplicarEstilosEmCelula();

	/**
	 * Aplica todas as bordas na planilha.
	 *
	 * @return A instância de EstiloCelula para encadeamento.
	 */
	EstiloCelula todasAsBordasEmTudo();
}
