package com.abnote.planilhas.interfaces;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.utils.ManipuladorPlanilha;

public interface IPlanilhaBasica extends AutoCloseable {

    EstiloCelula aplicarEstilos();

    void criarPlanilha(String nomeSheet);

    /**
     * Abre uma planilha {@code .xlsx} existente a partir do caminho informado,
     * selecionando a primeira aba.
     *
     * @param caminhoArquivo Caminho do arquivo a abrir.
     * @throws com.abnote.planilhas.exceptions.ArquivoException se o arquivo não
     *         existir, não puder ser lido, ou não for um {@code .xlsx} (o
     *         formato antigo {@code .xls} não é suportado).
     */
    void abrirPlanilha(String caminhoArquivo);

    void criarSheet(String nomeSheet);

    void selecionarSheet(String nomeSheet);

    void salvar(String nomeArquivo);
    void setDiretorioSaida(String diretorioSaida);

    String getDiretorioSaida();

    Workbook obterWorkbook();

    IPlanilhaBasica emTodaAPlanilha();

    IPlanilhaBasica ultimaLinha(String coluna);

    ManipuladorPlanilha manipularPlanilha();
    IFormulas formula();
    @Override
    void close() throws Exception;
}