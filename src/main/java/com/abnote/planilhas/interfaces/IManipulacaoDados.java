package com.abnote.planilhas.interfaces;

import java.util.List;
import com.abnote.planilhas.estilos.EstiloCelula;

/**
 * Interface que define as operações de manipulação de dados em uma planilha.
 */
public interface IManipulacaoDados {

//    /**
//     * Define a posição de uma célula.
//     *
//     * @param posicao Posição no formato alfanumérico (ex.: "B2")
//     * @return A instância atual para encadeamento.
//     */
//    IManipulacaoDados naCelula(String posicao);
//
//    /**
//     * Define um intervalo de células.
//     *
//     * @param posicaoInicial Posição inicial do intervalo (ex.: "B2")
//     * @param posicaoFinal   Posição final do intervalo (ex.: "E5")
//     * @return A instância atual para encadeamento.
//     */
//    IManipulacaoDados noIntervalo(String posicaoInicial, String posicaoFinal);

    /**
     * Insere dados em uma célula ou intervalo usando um delimitador.
     *
     * @param dados       Os dados a serem inseridos (pode ser String, List ou File).
     * @param delimitador O delimitador para separar os dados.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserirDados(Object dados, String delimitador);

    /**
     * Insere dados (String) em uma célula.
     *
     * @param valor O valor a ser inserido.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserirDados(String valor);

    /**
     * Insere uma lista de dados em uma célula ou intervalo.
     *
     * @param dados A lista de dados a serem inseridos.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserirDados(List<String> dados);

    /**
     * Insere uma lista de dados utilizando um delimitador.
     *
     * @param dados       A lista de dados.
     * @param delimitador O delimitador.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserirDados(List<String> dados, String delimitador);

    /**
     * Insere dados a partir de um arquivo.
     *
     * @param caminhoArquivo O caminho do arquivo.
     * @param delimitador    O delimitador presente no arquivo.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserirDadosArquivo(String caminhoArquivo, String delimitador);

    /**
     * Soma os valores de uma coluna.
     *
     * @param posicaoInicial A posição inicial da coluna (ex.: "J3").
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados somarColuna(String posicaoInicial);

    /**
     * Soma os valores de uma coluna e insere um texto descritivo.
     *
     * @param posicaoInicial A posição inicial da coluna (ex.: "J3").
     * @param texto          O texto descritivo.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados somarColunaComTexto(String posicaoInicial, String texto);

    /**
     * Multiplica os valores de duas colunas, insere um texto e coloca a soma total.
     *
     * @param coluna1       Primeira coluna (ex.: "D").
     * @param coluna2       Segunda coluna (ex.: "I").
     * @param linhaInicial  Linha inicial para o cálculo.
     * @param texto         Texto descritivo para a soma.
     * @param colunaDestino Coluna de destino (ex.: "J").
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados multiplicarColunasComTexto(String coluna1, String coluna2, int linhaInicial, String texto,
            String colunaDestino);

    /**
     * Mescla as células definidas no intervalo atual.
     *
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados mesclarCelulas();

    /**
     * Método alternativo para inserir dados de String.
     *
     * @param valor O valor a ser inserido.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserir(String valor);

    /**
     * Método alternativo para inserir dados inteiros.
     *
     * @param valor O valor inteiro a ser inserido.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserir(int valor);

    /**
     * Método alternativo para inserir dados do tipo double.
     *
     * @param valor O valor double a ser inserido.
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados inserir(double valor);

    /**
     * Retorna uma instância para aplicar estilos após operações de dados.
     * <p>
     * Este método permite encadeamento fluente entre operações de dados e estilos:
     * <pre>
     * planilha.naCelula("B2")
     *     .inserirDados("Valor")
     *     .aplicarEstilos()
     *     .aplicarNegrito();
     * </pre>
     * 
     * @return A instância de EstiloCelula para aplicar estilos.
     */
    EstiloCelula aplicarEstilos();

    /**
     * Define a posição na última linha da coluna indicada.
     *
     * @param coluna A coluna (ex.: "B").
     * @return A instância atual para encadeamento.
     */
    IManipulacaoDados naUltimaLinha(String coluna);
}