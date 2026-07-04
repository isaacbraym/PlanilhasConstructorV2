package com.abnote.planilhas.examples;

import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;

import java.io.IOException;
import java.util.logging.Logger;

/**
 * Exemplo prático de uso das fórmulas Excel.
 * 
 * <p>
 * Demonstra como criar uma planilha de controle financeiro usando fórmulas
 * matemáticas e condicionais.
 * </p>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 */
public class ExemploFormulas {

	private static final Logger logger = Logger.getLogger(ExemploFormulas.class.getName());

	public static void main(String[] args) {
		String caminhoArquivo = new java.io.File(System.getProperty("java.io.tmpdir"), "exemplo-formulas.xlsx")
				.getAbsolutePath();

		try (IPlanilha planilha = new PlanilhaXlsx()) {

			planilha.criarPlanilha("Controle Financeiro");

			// ========== CABEÇALHO ==========
			planilha.selecionar().celula("A1").inserirDados("Mês");
			planilha.selecionar().celula("B1").inserirDados("Receita");
			planilha.selecionar().celula("C1").inserirDados("Despesa");
			planilha.selecionar().celula("D1").inserirDados("Saldo");
			planilha.selecionar().celula("E1").inserirDados("Status");

			// ========== DADOS ==========
			planilha.selecionar().celula("A2").inserirDados("Janeiro"); 
			planilha.selecionar().celula("B2").inserirDados("5000");
			planilha.selecionar().celula("C2").inserirDados("3200");

			planilha.selecionar().celula("A3").inserirDados("Fevereiro");
			planilha.selecionar().celula("B3").inserirDados("5500");
			planilha.selecionar().celula("C3").inserirDados("4100");

			planilha.selecionar().celula("A4").inserirDados("Março");
			planilha.selecionar().celula("B4").inserirDados("6000");
			planilha.selecionar().celula("C4").inserirDados("3800");

			planilha.selecionar().celula("A5").inserirDados("Abril");
			planilha.selecionar().celula("B5").inserirDados("4800");
			planilha.selecionar().celula("C5").inserirDados("5200");

			// ========== FÓRMULAS DE SALDO (Receita - Despesa) ==========
			// Nota: Como não temos subtração direta, usamos fórmulas customizadas
			// ou implementamos via DataManipulator

			logger.info("Inserindo dados de saldo manualmente (fórmula de subtração será implementada em v2)");
			planilha.selecionar().celula("D2").inserirDados("1800");
			planilha.selecionar().celula("D3").inserirDados("1400");
			planilha.selecionar().celula("D4").inserirDados("2200");
			planilha.selecionar().celula("D5").inserirDados("-400");

			// ========== FÓRMULA IF: Status baseado em saldo ==========
			planilha.selecionar().celula("E2").formula().seEntao("D2>0", "Positivo", "Negativo").aplicar();

			planilha.selecionar().celula("E3").formula().seEntao("D3>0", "Positivo", "Negativo").aplicar();

			planilha.selecionar().celula("E4").formula().seEntao("D4>0", "Positivo", "Negativo").aplicar();

			planilha.selecionar().celula("E5").formula().seEntao("D5>0", "Positivo", "Negativo").aplicar();

			// ========== TOTALIZADORES ==========
			planilha.selecionar().celula("A7").inserirDados("TOTAIS");

			// Total de Receitas
			planilha.selecionar().celula("B7").formula().soma("B2:B5").aplicar();

			// Total de Despesas
			planilha.selecionar().celula("C7").formula().soma("C2:C5").aplicar();

			// Total de Saldo
			planilha.selecionar().celula("D7").formula().soma("D2:D5").aplicar();

			// ========== ESTATÍSTICAS ==========
			planilha.selecionar().celula("A9").inserirDados("ESTATÍSTICAS");

			planilha.selecionar().celula("A10").inserirDados("Média de Receitas:");
			planilha.selecionar().celula("B10").formula().media("B2:B5").aplicar();

			planilha.selecionar().celula("A11").inserirDados("Maior Despesa:");
			planilha.selecionar().celula("B11").formula().maximo("C2:C5").aplicar();

			planilha.selecionar().celula("A12").inserirDados("Menor Despesa:");
			planilha.selecionar().celula("B12").formula().minimo("C2:C5").aplicar();

			// ========== ANÁLISES CONDICIONAIS ==========
			planilha.selecionar().celula("A14").inserirDados("ANÁLISES");

			planilha.selecionar().celula("A15").inserirDados("Meses Positivos:");
			planilha.selecionar().celula("B15").formula().contarSe("D2:D5", ">0").aplicar();

			planilha.selecionar().celula("A16").inserirDados("Receita de Meses >5000:");
			planilha.selecionar().celula("B16").formula().somarSe("B2:B5", ">5000").aplicar();

			// ========== FÓRMULAS ADICIONAIS ==========
			planilha.selecionar().celula("A18").inserirDados("Data de Geração:");
			planilha.selecionar().celula("B18").formula().hoje().aplicar();

			planilha.selecionar().celula("A19").inserirDados("Timestamp:");
			planilha.selecionar().celula("B19").formula().agora().aplicar();

			// ========== CONCATENAÇÃO ==========
			planilha.selecionar().celula("A21").inserirDados("Primeiro");
			planilha.selecionar().celula("B21").inserirDados("Nome");
			planilha.selecionar().celula("C21").formula().concatenar("A21", "B21").aplicar();

			// ========== SALVAR ==========
			planilha.salvar(caminhoArquivo);

			logger.info("Planilha criada com sucesso em: " + caminhoArquivo);
			logger.info("Fórmulas aplicadas:");
			logger.info("  - 4x IF (Status positivo/negativo)");
			logger.info("  - 3x SUM (Totalizadores)");
			logger.info("  - 1x AVERAGE (Média de receitas)");
			logger.info("  - 1x MAX (Maior despesa)");
			logger.info("  - 1x MIN (Menor despesa)");
			logger.info("  - 1x COUNTIF (Meses positivos)");
			logger.info("  - 1x SUMIF (Receita condicional)");
			logger.info("  - 1x TODAY (Data atual)");
			logger.info("  - 1x NOW (Data/hora atual)");
			logger.info("  - 1x CONCATENATE (Juntar textos)");

		} catch (IOException e) {
			logger.severe("Erro ao salvar a planilha: " + e.getMessage());
			e.printStackTrace();
		} catch (Exception e) {
			logger.severe("Erro inesperado: " + e.getMessage());
			e.printStackTrace();
		}
	}
}
