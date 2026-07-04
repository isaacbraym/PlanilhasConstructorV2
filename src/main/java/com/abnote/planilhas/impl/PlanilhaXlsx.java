package com.abnote.planilhas.impl;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IManipulacaoDados;

public class PlanilhaXlsx extends PlanilhaBase {

    @Override
    protected void inicializarWorkbook() {
        workbook = new XSSFWorkbook(); // Inicializa o XSSFWorkbook para arquivos .xlsx
    }

    @Override
    public EstiloCelula todasAsBordasEmTudo() {
        return super.todasAsBordasEmTudo();
    }

    @Override
    public IManipulacaoDados inserir(String valor) {
        return inserirDados(valor);
    }

    @Override
    public IManipulacaoDados inserir(int valor) {
        return inserirDados(String.valueOf(valor));
    }

    @Override
    public IManipulacaoDados inserir(double valor) {
        return inserirDados(String.valueOf(valor));
    }
}
