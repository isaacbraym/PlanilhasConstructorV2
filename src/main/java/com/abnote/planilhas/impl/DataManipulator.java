package com.abnote.planilhas.impl;

import java.util.List;
import java.util.stream.IntStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.abnote.planilhas.calculos.Calculos;
import com.abnote.planilhas.calculos.Conversores;
import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IManipulacaoDados;
import com.abnote.planilhas.utils.InsersorDeDados;
import com.abnote.planilhas.utils.PositionManager;
import com.abnote.planilhas.utils.PosicaoConverter;

public class DataManipulator implements IManipulacaoDados {

	private final Sheet sheet;
	private final Workbook workbook;
	private final PositionManager positionManager;
	private final InsersorDeDados insersorDeDados;
	private int ultimoIndiceDeLinhaInserido = -1;
	private int ultimoIndiceDeColunaInserido = -1;
	private int linhaSelecionadaAtual = -1;
	private int colunaSelecionadaAtual = -1;

	public DataManipulator(Workbook workbook, Sheet sheet, PositionManager positionManager) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.positionManager = positionManager;
		this.insersorDeDados = new InsersorDeDados(sheet, positionManager);
	}

	public IManipulacaoDados naCelula(String posicao) {
	    positionManager.naCelula(posicao);

	    int[] indices = com.abnote.planilhas.utils.PosicaoConverter.converterPosicao(posicao);
	    this.colunaSelecionadaAtual = indices[0];
	    this.linhaSelecionadaAtual = indices[1];
	    // Uma nova seleção substitui qualquer "última inserção" pendente, para que
	    // aplicarEstilos() sobre a célula selecionada não estilize uma linha antiga.
	    this.ultimoIndiceDeLinhaInserido = -1;
	    this.ultimoIndiceDeColunaInserido = -1;

	    return this;
	}

	public IManipulacaoDados noIntervalo(String posicaoInicial, String posicaoFinal) {
	    positionManager.noIntervalo(posicaoInicial, posicaoFinal);
	    
	    // Atualiza para a posição inicial do intervalo
	    int[] indices = com.abnote.planilhas.utils.PosicaoConverter.converterPosicao(posicaoInicial);
	    this.colunaSelecionadaAtual = indices[0];
	    this.linhaSelecionadaAtual = indices[1];
	    
	    return this;
	}

	@Override
	public IManipulacaoDados inserirDados(Object dados, String delimitador) {
		insersorDeDados.inserirDados(dados, delimitador);
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados inserirDados(String valor) {
		insersorDeDados.inserirDados(valor);
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados inserirDados(List<String> dados) {
		insersorDeDados.inserirDados(dados);
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados inserirDados(List<String> dados, String delimitador) {
		insersorDeDados.inserirDados(dados, delimitador);
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados inserirDadosArquivo(String caminhoArquivo, String delimitador) {
		insersorDeDados.inserirDadosArquivo(caminhoArquivo, delimitador);
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados somarColuna(String posicaoInicial) {
		Calculos.somarColuna(sheet, posicaoInicial);
		String colunaLetra = posicaoInicial.replaceAll("[0-9]", "");
		ultimaLinha(colunaLetra);
		ultimoIndiceDeColunaInserido = -1;
		return this;
	}

	@Override
	public IManipulacaoDados somarColunaComTexto(String posicaoInicial, String texto) {
		Calculos.somarColunaComTexto(sheet, posicaoInicial, texto);
		String colunaLetra = posicaoInicial.replaceAll("[0-9]", "");
		ultimaLinha(colunaLetra);
		ultimoIndiceDeColunaInserido = -1;
		return this;
	}

	@Override
	public IManipulacaoDados multiplicarColunasComTexto(String coluna1, String coluna2, int linhaInicial, String texto,
			String colunaDestino) {
		Calculos.multiplicarColunasComTexto(sheet, coluna1, coluna2, linhaInicial, texto, colunaDestino);
		String colunaLetra = colunaDestino;
		ultimaLinha(colunaLetra);
		ultimoIndiceDeColunaInserido = -1;
		return this;
	}

	@Override
	public IManipulacaoDados mesclarCelulas() {
		if (!positionManager.isIntervaloDefinida()) {
			throw new IllegalStateException(
					"É necessário definir um intervalo usando noIntervalo() antes de mesclar células.");
		}

		int startRow = positionManager.getPosicaoInicialLinha();
		int endRow = positionManager.getPosicaoFinalLinha();
		int startCol = positionManager.getPosicaoInicialColuna();
		int endCol = positionManager.getPosicaoFinalColuna();

		// Cria a região a ser mesclada
		org.apache.poi.ss.util.CellRangeAddress cellRangeAddress = new org.apache.poi.ss.util.CellRangeAddress(startRow,
				endRow, startCol, endCol);

		// Adiciona a região mesclada à planilha
		sheet.addMergedRegion(cellRangeAddress);

		return this;
	}

	/**
	 * Atualiza o índice da última linha inserida na coluna informada.
	 */
	public void ultimaLinha(String coluna) {
		int colunaIndex = PosicaoConverter.converterColuna(coluna);
		ultimoIndiceDeLinhaInserido = IntStream.rangeClosed(0, sheet.getLastRowNum()).filter(i -> {
			Row row = sheet.getRow(i);
			return row != null && row.getCell(colunaIndex) != null
					&& row.getCell(colunaIndex).getCellType() != CellType.BLANK;
		}).max().orElse(sheet.getLastRowNum());
	}

	@Override
	public IManipulacaoDados naUltimaLinha(String coluna) {
		int colunaIndex = PosicaoConverter.converterColuna(coluna);
		int lastRowIndex = IntStream.rangeClosed(0, sheet.getLastRowNum()).filter(i -> {
			Row row = sheet.getRow(i);
			return row != null && row.getCell(colunaIndex) != null
					&& row.getCell(colunaIndex).getCellType() != CellType.BLANK;
		}).max().orElse(-1);
		int nextRowIndex = lastRowIndex + 1;
		String colunaLetra = PosicaoConverter.converterIndice(colunaIndex);
		String posicao = colunaLetra + (nextRowIndex + 1);
		positionManager.naCelula(posicao);

		return this;
	}

	@Override
	public IManipulacaoDados inserir(String valor) {
		insersorDeDados.inserirDados(valor);
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados inserir(int valor) {
		insersorDeDados.inserirDados(String.valueOf(valor));
		updateLastInsertedIndices();
		return this;
	}

	@Override
	public IManipulacaoDados inserir(double valor) {
		insersorDeDados.inserirDados(String.valueOf(valor));
		updateLastInsertedIndices();
		return this;
	}

	private void updateLastInsertedIndices() {
		ultimoIndiceDeLinhaInserido = insersorDeDados.getUltimoIndiceDeLinhaInserido();
		ultimoIndiceDeColunaInserido = insersorDeDados.getUltimoIndiceDeColunaInserido();
	}

	public int getUltimoIndiceDeLinhaInserido() {
		return ultimoIndiceDeLinhaInserido;
	}

	public int getUltimoIndiceDeColunaInserido() {
		return ultimoIndiceDeColunaInserido;
	}
	/**
	 * Retorna o índice da linha atualmente selecionada por naCelula().
	 * Diferente de getUltimoIndiceDeLinhaInserido() que retorna a última linha com dados inseridos.
	 */
	public int getLinhaSelecionadaAtual() {
	    return linhaSelecionadaAtual;
	}

	/**
	 * Retorna o índice da coluna atualmente selecionada por naCelula().
	 * Diferente de getUltimoIndiceDeColunaInserido() que retorna a última coluna com dados inseridos.
	 */
	public int getColunaSelecionadaAtual() {
	    return colunaSelecionadaAtual;
	}
	@Override
	public EstiloCelula aplicarEstilos() {
		int rowIndex = getUltimoIndiceDeLinhaInserido();
		int columnIndex = getUltimoIndiceDeColunaInserido();
		if (rowIndex >= 0 && columnIndex >= 0) {
			return new EstiloCelula(workbook, sheet, rowIndex, columnIndex);
		} else {
			throw new IllegalStateException("Nenhuma célula disponível para aplicar estilos.");
		}
	}
}
