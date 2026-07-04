package com.abnote.planilhas.interfaces;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.utils.ManipuladorPlanilha;

public interface IPlanilhaBasica extends AutoCloseable {

    EstiloCelula aplicarEstilos();

    void criarPlanilha(String nomeSheet);

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