package com.abnote.planilhas.test;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Logger;

import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.estilos.estilos.FonteEnum;
import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.utils.LoggerUtil;

public class TestePlanilha {
    public static void main(String[] args) {
        final Logger logger = LoggerUtil.getLogger(TestePlanilha.class);

        try (IPlanilha planilha = new PlanilhaXlsx()) {
            
            String sheet1 = "Dados Brasileiros";
            String sheet2 = "TesteAba2";
            String sheet3 = "TesteAba3";

            String diretorioSaida = "C:\\opt\\tmp\\testePlanilhaSaidas";
            planilha.setDiretorioSaida(diretorioSaida);

            String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
            String nomeArquivoPlanilha = "\\planilhaTeste_" + timestamp + ".xlsx";
            String caminhoArquivo = planilha.getDiretorioSaida() + nomeArquivoPlanilha;
            String listaDeArquivos = "C:\\opt\\lista_dados_brasileiros.txt";
            String listaDeArquivos2 = "C:\\opt\\listaDados2.csv";

            String header = "NOME,CPF/CNPJ,ENDERECO,NUMERO,COMPLEMENTO,CEP,CIDADE,ESTADO,VALOR";

            // PLANILHA SHEET1
            planilha.criarPlanilha(sheet1);
            planilha.selecionarSheet(sheet1);
            planilha.selecionar().celula("B2").inserirDados(header, ",");
            planilha.aplicarEstilos()
                    .fonte(FonteEnum.TIMES_NEW_ROMAN)
                    .corFonte(CorEnum.LARANJA)
                    .fonteTamanho(14)
                    .corDeFundo(CorEnum.VERMELHO_ESCURO)
                    .aplicarNegrito();

            planilha.selecionar().celula("B3").inserirDados(listaDeArquivos, ";");
            planilha.emTodaAPlanilha().aplicarEstilos().fonte("Segoe UI").fonteTamanho(14);
            planilha.aplicarEstilos().aplicarBordasEspessasComInternas("B2", "J2");
            planilha.selecionar().celula("N11").inserirDados("TESTE");
            planilha.aplicarEstilosEmCelula().aplicarNegrito().corDeFundo(CorEnum.VERDE).corFonte(CorEnum.TURQUESA);

            planilha.converter().emContabil("J3").somarColunaComTexto("J3", "VALOR TOTAL DA SOMA");
            planilha.aplicarEstilos().aplicarItalico().aplicarNegrito();
            planilha.aplicarEstilos().aplicarBordasEntre("L2", "L200");
            planilha.selecionar().intervalo("C4", "C17").aplicarEstilos().aplicarNegrito();
            planilha.selecionar().intervalo("C5", "G5").aplicarEstilos().fonte("Calibri").fonteTamanho(18).aplicarNegrito();
            planilha.selecionar().intervalo("C4", "F4").aplicarEstilos().fonte(FonteEnum.EBRIMA).fonteTamanho(21).aplicarNegrito();
            planilha.manipularPlanilha().moverColuna("C", "F").logAlteracoes();
            planilha.manipularPlanilha().removerColuna("I").logAlteracoes();
            planilha.selecionar().intervalo("G10", "G20").aplicarEstilos().corDeFundo(CorEnum.ROXO).corFonte(CorEnum.BRANCO);
            planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
            planilha.selecionar().intervalo("G15", "G20").aplicarEstilos().aplicarItalico().aplicarNegrito().alinharADireita();
            planilha.selecionar().intervalo("G18", "G22").aplicarEstilos().aplicarNegrito().aplicarSublinhado().alinharAEsquerda();
            planilha.selecionar().intervalo("F20", "H20").aplicarEstilos().aplicarTachado().aplicarNegrito();
            planilha.selecionar().intervalo("G4", "H4").mesclarCelulas();
            planilha.aplicarEstilos().corDeFundo(CorEnum.VERMELHO_ESCURO).corFonte(CorEnum.BRANCO).aplicarNegrito();
            planilha.selecionar().intervalo("C12", "C15").mesclarCelulas();
            planilha.aplicarEstilos().corDeFundo(CorEnum.AZUL_CELESTE).aplicarItalico().aplicarTachado();
            planilha.converter().emContabil("J3").multiplicarColunasComTexto("D", "I", 3, "Total multiplicação", "J")
                    .aplicarEstilos().redimensionarColuna();
            planilha.ultimaLinha("I").aplicarEstilos().fonteTamanho(14).aplicarNegrito();
            planilha.naUltimaLinha("E").inserir("TESTEE").aplicarEstilos().aplicarNegrito().fonteTamanho(14)
                    .corDeFundo(CorEnum.LARANJA);
            planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
            planilha.inserirFiltros();

            // PLANILHA SHEET2
            planilha.criarSheet(sheet2);
            planilha.selecionarSheet(sheet2);
            planilha.selecionar().celula("C3").inserirDados(listaDeArquivos, ";");
            planilha.ultimaLinha("J").aplicarEstilos().aplicarNegrito().fonte("Arial").fonteTamanho(14);
            planilha.selecionar().intervalo("C4", "C17").aplicarEstilos().fonte("Another Danger - Demo").fonteTamanho(12)
                    .corDeFundo(CorEnum.AMARELO).aplicarNegrito();
            planilha.selecionar().intervalo("C4", "F4").aplicarEstilos().fonte(FonteEnum.VERDANA).fonteTamanho(21)
                    .corFonte(CorEnum.BRANCO).corDeFundo("#9400d3").aplicarNegrito();
            planilha.selecionar().intervalo("D11", "G11").aplicarEstilos().corFonte(CorEnum.BEGE).corDeFundo(90, 50, 128)
                    .aplicarNegrito();
            planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
            planilha.inserirFiltros();

            // PLANILHA SHEET3
            planilha.criarSheet(sheet3);
            planilha.selecionarSheet(sheet3);
            planilha.inserirDados(listaDeArquivos2, "|");
            planilha.converter().emNumero("K2").somarColunaComTexto("K2", "Totais:");
            planilha.converter().emNumero("L2").somarColuna("L2").aplicarEstilos().aplicarNegrito();
            planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
            planilha.inserirFiltros();

            // Teste de logs (descomente se necessário)
            // logger.severe("Esta é uma mensagem SEVERE");
            // logger.warning("Esta é uma mensagem WARNING");
            // logger.info("Esta é uma mensagem INFO");
            // logger.config("Esta é uma mensagem CONFIG");
            // logger.fine("Esta é uma mensagem FINE");
            // logger.finer("Esta é uma mensagem FINER");
            // logger.finest("Esta é uma mensagem FINEST");
            
            planilha.salvar(caminhoArquivo);
            
            logger.info("Planilha criada e salva com sucesso em: " + caminhoArquivo);
            
        } catch (IOException e) {
            logger.severe("Erro ao salvar a planilha: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            logger.severe("Erro inesperado: " + e.getMessage());
            e.printStackTrace();
        }
        // Workbook é fechado automaticamente aqui pelo try-with-resources
        
    }
}