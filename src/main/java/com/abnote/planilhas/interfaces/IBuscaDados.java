package com.abnote.planilhas.interfaces;

import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;

/**
 * Interface para busca de dados na planilha.
 */
public interface IBuscaDados {

	/**
	 * Busca dados na planilha a partir de um arquivo de entrada.
	 *
	 * @param caminhoArquivo Caminho do arquivo (TXT ou CSV) contendo os dados.
	 * @param delimitador    Delimitador usado no arquivo.
	 * @param coluna         Coluna onde os dados serão buscados (ex.: "B").
	 * @return Lista de linhas onde os dados foram encontrados.
	 * @throws IOException Se ocorrer erro na leitura do arquivo.
	 */
	List<Row> buscarDadosDeEm(String caminhoArquivo, String delimitador, String coluna) throws IOException;

	/**
	 * Busca dados na planilha a partir de uma lista de valores.
	 *
	 * @param valores Lista de valores a serem buscados.
	 * @param coluna  Coluna onde os dados serão buscados (ex.: "B").
	 * @return Lista de linhas onde os dados foram encontrados.
	 */
	List<Row> buscarDadosDeEm(List<String> valores, String coluna);

	/**
	 * Salva as linhas encontradas em uma nova aba.
	 *
	 * @param linhas      Linhas a serem salvas.
	 * @param nomeNovaAba Nome da nova aba.
	 */
	void salvarLinhasEmNovaAba(List<Row> linhas, String nomeNovaAba);

	/**
	 * Move as linhas encontradas para outra aba.
	 *
	 * @param linhas      Linhas a serem movidas.
	 * @param nomeDestino Nome da aba de destino.
	 */
	void moverLinhasParaAba(List<Row> linhas, String nomeDestino);

	/**
	 * Copia as linhas encontradas para outra aba.
	 *
	 * @param linhas      Linhas a serem copiadas.
	 * @param nomeDestino Nome da aba de destino.
	 */
	void copiarLinhasParaAba(List<Row> linhas, String nomeDestino);

	/**
	 * Remove as linhas encontradas da planilha.
	 *
	 * @param linhas Linhas a serem removidas.
	 */
	void removerLinhas(List<Row> linhas);
}
