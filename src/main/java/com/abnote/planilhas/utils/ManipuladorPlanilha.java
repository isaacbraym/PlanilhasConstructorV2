package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.Sheet;
import java.util.HashMap;
import java.util.Map;

public class ManipuladorPlanilha {
	private Sheet sheet;
	private final LogsDeModificadores logs;
	private Map<Integer, ManipuladorPlanilhaHelper.CellData> colunaTemporaria = new HashMap<>();
	private int columnOffset;
	private ManipuladorPlanilhaHelper helper;

	// Construtor que determina automaticamente o columnOffset
	public ManipuladorPlanilha(Sheet sheet) {
		this(sheet, ManipuladorPlanilhaHelper.determinarColunaInicial(sheet));
	}

	// Construtor que permite definir o columnOffset manualmente
	public ManipuladorPlanilha(Sheet sheet, int columnOffset) {
		this.sheet = sheet;
		this.columnOffset = columnOffset;
		this.logs = new LogsDeModificadores();
		this.helper = new ManipuladorPlanilhaHelper(sheet, columnOffset);
	}

	public ManipuladorPlanilha moverColuna(String moverAColuna, String paraAPosicao) {
		int colunaOrigem = PosicaoConverter.converterColuna(moverAColuna) - columnOffset;
		int colunaDestino = PosicaoConverter.converterColuna(paraAPosicao) - columnOffset;

		if (colunaOrigem == colunaDestino) {
			return this;
		}

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		String headerOrigem = headerMap.get(colunaOrigem);
		LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerOrigem,
				PosicaoConverter.converterIndice(colunaOrigem + columnOffset),
				PosicaoConverter.converterIndice(colunaDestino + columnOffset));
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Deslocamento de colunas",
				mainMovement);

		// Copia a coluna de origem para armazenamento temporário
		colunaTemporaria = helper.copiarColuna(colunaOrigem);

		// Desloca as colunas conforme a direção do movimento
		if (colunaOrigem < colunaDestino) {
			helper.deslocarColunasParaEsquerda(colunaOrigem + 1, colunaDestino);
			helper.registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
		} else {
			helper.deslocarColunasParaDireita(colunaDestino, colunaOrigem - 1);
			helper.registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
		}

		// Cola a coluna temporária na posição de destino
		helper.colarColunaTemporaria(colunaDestino, colunaTemporaria);
		colunaTemporaria.clear();
		logs.adicionarLog(actionLog);

		return this;
	}

	public ManipuladorPlanilha removerColuna(String coluna) {
		int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;
		int lastColumn = helper.obterNumeroUltimaColuna();

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		String headerName = headerMap.get(colIndex);
		LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerName,
				PosicaoConverter.converterIndice(colIndex + columnOffset), null);
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Remoção de coluna", mainMovement);

		helper.removerCelulasDaColuna(colIndex);

		if (colIndex < lastColumn) {
			helper.deslocarColunasParaEsquerda(colIndex + 1, lastColumn);
			helper.registrarColunasDeslocadasRemocao(colIndex, lastColumn, headerMap, actionLog);
		}

		logs.adicionarLog(actionLog);
		return this;
	}

	public ManipuladorPlanilha limparColuna(String coluna) {
		int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;
		helper.limparColuna(colIndex);

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		String headerName = headerMap.get(colIndex);
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Limpeza de coluna",
				new LogsDeModificadores.ColumnMovement(headerName,
						PosicaoConverter.converterIndice(colIndex + columnOffset), null));
		logs.adicionarLog(actionLog);

		return this;
	}

	public ManipuladorPlanilha inserirColunaVaziaEntre(String colunaEsquerda, String colunaDireita) {
		int colEsquerdaIndex = PosicaoConverter.converterColuna(colunaEsquerda) - columnOffset;
		int colDireitaIndex = PosicaoConverter.converterColuna(colunaDireita) - columnOffset;

		helper.validarAdjacencia(colEsquerdaIndex, colDireitaIndex, colunaEsquerda, colunaDireita);

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		int posicaoInsercao = colDireitaIndex;
		int lastColumn = helper.obterNumeroUltimaColuna();

		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Inserção de coluna vazia",
				new LogsDeModificadores.ColumnMovement(null, colunaEsquerda, colunaDireita));

		if (posicaoInsercao <= lastColumn) {
			helper.deslocarColunasParaDireita(posicaoInsercao, lastColumn);
			helper.registrarColunasDeslocadasInsercao(posicaoInsercao, lastColumn, headerMap, actionLog);
			logs.adicionarLog(actionLog);
		}

		helper.definirLarguraNovaColuna(posicaoInsercao);
		return this;
	}

	public void logAlteracoes() {
		logs.exibirLogs();
	}
}
