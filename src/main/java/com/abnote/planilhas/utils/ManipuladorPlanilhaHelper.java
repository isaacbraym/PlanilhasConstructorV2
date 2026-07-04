package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;

import java.util.HashMap;
import java.util.Map;

/**
 * Classe utilitária para manipulação de planilhas usando Apache POI.
 */
public class ManipuladorPlanilhaHelper {
	private final Sheet sheet;
	private final int columnOffset;

	/**
	 * Construtor para inicializar o manipulador com uma planilha e um deslocamento
	 * de coluna.
	 *
	 * @param sheet        A planilha a ser manipulada.
	 * @param columnOffset O deslocamento a ser aplicado nas operações de coluna.
	 */
	public ManipuladorPlanilhaHelper(Sheet sheet, int columnOffset) {
		this.sheet = sheet;
		this.columnOffset = columnOffset;
	}

	/**
	 * Determina dinamicamente o índice inicial da coluna com base na primeira linha
	 * não vazia.
	 *
	 * @param sheet A planilha para determinar o índice da coluna inicial.
	 * @return O índice da coluna inicial.
	 */
	public static int determinarColunaInicial(Sheet sheet) {
		Row primeiraLinha = sheet.getRow(0);
		if (primeiraLinha != null) {
			for (Cell cell : primeiraLinha) {
				String valorCelula = obterValorCelulaComoString(cell);
				if (valorCelula != null && !valorCelula.trim().isEmpty()) {
					return cell.getColumnIndex();
				}
			}
		}
		return 0; // Padrão para a primeira coluna se nenhum cabeçalho for encontrado
	}

	/**
	 * Limpa os dados de uma célula de uma coluna específica, sem remover a coluna.
	 *
	 * @param colIndex Índice da coluna a ser limpa.
	 */
	public void limparColuna(int colIndex) {
		int ultimaLinha = sheet.getLastRowNum();
		for (int rowNum = 0; rowNum <= ultimaLinha; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null)
				continue;
			Cell celula = row.getCell(colIndex + columnOffset);
			if (celula != null) {
				celula.setCellType(CellType.BLANK);
			}
		}
	}

	/**
	 * Retorna o valor de uma célula como String.
	 *
	 * @param cell A célula.
	 * @return O valor da célula ou null.
	 */
	public static String obterValorCelulaComoString(Cell cell) {
		if (cell == null)
			return null;
		CellType tipoCelula = cell.getCellType();
		switch (tipoCelula) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			return DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue().toString()
					: Double.toString(cell.getNumericCellValue());
		case BOOLEAN:
			return Boolean.toString(cell.getBooleanCellValue());
		case FORMULA:
			return cell.getCellFormula();
		case ERROR:
			return Byte.toString(cell.getErrorCellValue());
		default:
			return null;
		}
	}

	/**
	 * Obtém o número da última coluna na planilha considerando o deslocamento.
	 *
	 * @return O índice da última coluna.
	 */
	public int obterNumeroUltimaColuna() {
		int maxColunas = 0;
		for (Row row : sheet) {
			if (row != null && row.getLastCellNum() > maxColunas) {
				maxColunas = row.getLastCellNum();
			}
		}
		return maxColunas - columnOffset - 1;
	}

	/**
	 * Obtém um mapa que relaciona os índices das colunas com os nomes dos
	 * cabeçalhos.
	 *
	 * @return Mapa de índices para nomes.
	 */
	public Map<Integer, String> obterMapaDeCabecalhos() {
		Map<Integer, String> mapaCabecalhos = new HashMap<>();
		int ultimaLinha = sheet.getLastRowNum();
		int ultimaColuna = obterNumeroUltimaColuna();
		for (int colIndex = 0; colIndex <= ultimaColuna; colIndex++) {
			String nomeCabecalho = encontrarCabecalho(colIndex, ultimaLinha);
			if (nomeCabecalho != null) {
				mapaCabecalhos.put(colIndex, nomeCabecalho);
			}
		}
		return mapaCabecalhos;
	}

	/**
	 * Encontra o cabeçalho de uma coluna específica.
	 *
	 * @param colIndex    Índice da coluna.
	 * @param ultimaLinha Última linha da planilha.
	 * @return O nome do cabeçalho ou null.
	 */
	private String encontrarCabecalho(int colIndex, int ultimaLinha) {
		for (int rowNum = 0; rowNum <= ultimaLinha; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null)
				continue;
			Cell celula = row.getCell(colIndex + columnOffset);
			String valorCelula = obterValorCelulaComoString(celula);
			if (valorCelula != null && !valorCelula.trim().isEmpty()) {
				return valorCelula;
			}
		}
		return null;
	}

	/**
	 * Valida se duas colunas são adjacentes.
	 *
	 * @param indiceEsquerda Índice da coluna à esquerda.
	 * @param indiceDireita  Índice da coluna à direita.
	 * @param colunaEsquerda Nome da coluna à esquerda.
	 * @param colunaDireita  Nome da coluna à direita.
	 * @throws IllegalArgumentException se as colunas não forem adjacentes.
	 */
	public void validarAdjacencia(int indiceEsquerda, int indiceDireita, String colunaEsquerda, String colunaDireita) {
		if (indiceDireita - indiceEsquerda != 1) {
			throw new IllegalArgumentException(String.format(
					"As colunas especificadas não são adjacentes. Certifique-se de que '%s' está imediatamente à direita de '%s'.",
					colunaDireita, colunaEsquerda));
		}
	}

	/**
	 * Define a largura da nova coluna na posição de inserção.
	 *
	 * @param posicaoInsercao Índice de inserção.
	 */
	public void definirLarguraNovaColuna(int posicaoInsercao) {
		sheet.setColumnWidth(posicaoInsercao + columnOffset, sheet.getDefaultColumnWidth() * 256);
	}

	/**
	 * Registra as colunas deslocadas após uma movimentação.
	 *
	 * @param colunaOrigem   Índice da coluna de origem.
	 * @param colunaDestino  Índice da coluna de destino.
	 * @param mapaCabecalhos Mapa de cabeçalhos.
	 * @param logAcoes       Objeto de log para registrar a ação.
	 */
	public void registrarColunasDeslocadas(int colunaOrigem, int colunaDestino, Map<Integer, String> mapaCabecalhos,
			LogsDeModificadores.ActionLog logAcoes) {
		if (colunaOrigem < colunaDestino) {
			for (int col = colunaOrigem + 1; col <= colunaDestino; col++) {
				adicionarMovimentacaoColuna(col, col - 1, mapaCabecalhos, logAcoes);
			}
		} else {
			for (int col = colunaOrigem - 1; col >= colunaDestino; col--) {
				adicionarMovimentacaoColuna(col, col + 1, mapaCabecalhos, logAcoes);
			}
		}
	}

	/**
	 * Registra as colunas deslocadas após a remoção de uma coluna.
	 *
	 * @param indiceColuna   Índice da coluna removida.
	 * @param ultimaColuna   Índice da última coluna.
	 * @param mapaCabecalhos Mapa de cabeçalhos.
	 * @param logAcoes       Objeto de log para registrar a ação.
	 */
	public void registrarColunasDeslocadasRemocao(int indiceColuna, int ultimaColuna,
			Map<Integer, String> mapaCabecalhos, LogsDeModificadores.ActionLog logAcoes) {
		for (int col = indiceColuna + 1; col <= ultimaColuna; col++) {
			adicionarMovimentacaoColuna(col, col - 1, mapaCabecalhos, logAcoes);
		}
	}

	/**
	 * Registra as colunas deslocadas após a inserção de uma coluna.
	 *
	 * @param posicaoInsercao Índice de inserção.
	 * @param ultimaColuna    Índice da última coluna.
	 * @param mapaCabecalhos  Mapa de cabeçalhos.
	 * @param logAcoes        Objeto de log para registrar a ação.
	 */
	public void registrarColunasDeslocadasInsercao(int posicaoInsercao, int ultimaColuna,
			Map<Integer, String> mapaCabecalhos, LogsDeModificadores.ActionLog logAcoes) {
		for (int col = ultimaColuna; col >= posicaoInsercao; col--) {
			adicionarMovimentacaoColuna(col, col + 1, mapaCabecalhos, logAcoes);
		}
	}

	/**
	 * Adiciona uma movimentação de coluna ao log.
	 *
	 * @param colunaAtual    Índice atual da coluna.
	 * @param colunaAlvo     Novo índice da coluna.
	 * @param mapaCabecalhos Mapa de cabeçalhos.
	 * @param logAcoes       Objeto de log.
	 */
	private void adicionarMovimentacaoColuna(int colunaAtual, int colunaAlvo, Map<Integer, String> mapaCabecalhos,
			LogsDeModificadores.ActionLog logAcoes) {
		String nomeCabecalho = mapaCabecalhos.get(colunaAtual);
		if (nomeCabecalho != null && !nomeCabecalho.trim().isEmpty()) {
			String indiceAnterior = PosicaoConverter.converterIndice(colunaAtual + columnOffset);
			String indiceNovo = PosicaoConverter.converterIndice(colunaAlvo + columnOffset);
			logAcoes.getShiftedColumns()
					.add(new LogsDeModificadores.ColumnMovement(nomeCabecalho, indiceAnterior, indiceNovo));
		}
	}

	/**
	 * Remove todas as células de uma coluna específica.
	 *
	 * @param colIndex Índice da coluna.
	 */
	public void removerCelulasDaColuna(int colIndex) {
		int ultimaLinha = sheet.getLastRowNum();
		for (int rowNum = 0; rowNum <= ultimaLinha; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null)
				continue;
			Cell celula = row.getCell(colIndex + columnOffset);
			if (celula != null) {
				row.removeCell(celula);
			}
		}
	}

	/**
	 * Copia uma coluna para um mapa temporário.
	 *
	 * @param colunaOrigem Índice da coluna a ser copiada.
	 * @return Mapa com os dados da coluna.
	 */
	public Map<Integer, CellData> copiarColuna(int colunaOrigem) {
		Map<Integer, CellData> colunaTemp = new HashMap<>();
		int ultimaLinha = sheet.getLastRowNum();
		for (int i = 0; i <= ultimaLinha; i++) {
			Row row = sheet.getRow(i);
			if (row == null)
				continue;
			Cell celulaOrigem = row.getCell(colunaOrigem + columnOffset);
			if (celulaOrigem != null) {
				CellData dadosCelula = new CellData();
				copiarDadosCelula(celulaOrigem, dadosCelula);
				colunaTemp.put(i, dadosCelula);
				row.removeCell(celulaOrigem);
			}
		}
		return colunaTemp;
	}

	/**
	 * Cola os dados de uma coluna temporária em uma coluna de destino.
	 *
	 * @param colunaDestino    Índice da coluna de destino.
	 * @param colunaTemporaria Mapa com os dados temporários.
	 */
	public void colarColunaTemporaria(int colunaDestino, Map<Integer, CellData> colunaTemporaria) {
		for (Map.Entry<Integer, CellData> entrada : colunaTemporaria.entrySet()) {
			int rowNum = entrada.getKey();
			CellData dadosCelula = entrada.getValue();
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				row = sheet.createRow(rowNum);
			}
			Cell celulaDestino = row.createCell(colunaDestino + columnOffset);
			colarDadosCelula(celulaDestino, dadosCelula);
		}
	}

	/**
	 * Copia os dados de uma célula para um objeto CellData.
	 *
	 * @param celulaOrigem A célula de origem.
	 * @param dadosCelula  Objeto para armazenar os dados.
	 */
	private void copiarDadosCelula(Cell celulaOrigem, CellData dadosCelula) {
		dadosCelula.setCellType(celulaOrigem.getCellType());
		dadosCelula.setCellStyle(celulaOrigem.getCellStyle());
		switch (celulaOrigem.getCellType()) {
		case STRING:
			dadosCelula.setStringValue(celulaOrigem.getStringCellValue());
			break;
		case NUMERIC:
			dadosCelula.setNumericValue(celulaOrigem.getNumericCellValue());
			break;
		case BOOLEAN:
			dadosCelula.setBooleanValue(celulaOrigem.getBooleanCellValue());
			break;
		case FORMULA:
			dadosCelula.setFormulaValue(celulaOrigem.getCellFormula());
			break;
		case ERROR:
			dadosCelula.setErrorValue(celulaOrigem.getErrorCellValue());
			break;
		case BLANK:
			break;
		default:
			break;
		}
	}

	/**
	 * Cola os dados de um objeto CellData em uma célula.
	 *
	 * @param celulaDestino A célula de destino.
	 * @param dadosCelula   Objeto com os dados.
	 */
	private void colarDadosCelula(Cell celulaDestino, CellData dadosCelula) {
		celulaDestino.setCellType(dadosCelula.getCellType());
		celulaDestino.setCellStyle(dadosCelula.getCellStyle());
		switch (dadosCelula.getCellType()) {
		case STRING:
			celulaDestino.setCellValue(dadosCelula.getStringValue());
			break;
		case NUMERIC:
			celulaDestino.setCellValue(dadosCelula.getNumericValue());
			break;
		case BOOLEAN:
			celulaDestino.setCellValue(dadosCelula.isBooleanValue());
			break;
		case FORMULA:
			celulaDestino.setCellFormula(dadosCelula.getFormulaValue());
			break;
		case ERROR:
			celulaDestino.setCellErrorValue(dadosCelula.getErrorValue());
			break;
		case BLANK:
			break;
		default:
			break;
		}
	}

	/**
	 * Desloca as colunas para a esquerda em um intervalo especificado.
	 *
	 * @param inicioColuna Índice inicial.
	 * @param fimColuna    Índice final.
	 */
	public void deslocarColunasParaEsquerda(int inicioColuna, int fimColuna) {
		int ultimaLinha = sheet.getLastRowNum();
		for (int col = inicioColuna; col <= fimColuna; col++) {
			int colunaAlvo = col - 1;
			deslocarColuna(col, colunaAlvo, ultimaLinha);
		}
	}

	/**
	 * Desloca as colunas para a direita em um intervalo especificado.
	 *
	 * @param inicioColuna Índice inicial.
	 * @param fimColuna    Índice final.
	 */
	public void deslocarColunasParaDireita(int inicioColuna, int fimColuna) {
		int ultimaLinha = sheet.getLastRowNum();
		for (int col = fimColuna; col >= inicioColuna; col--) {
			int colunaAlvo = col + 1;
			deslocarColuna(col, colunaAlvo, ultimaLinha);
		}
	}

	/**
	 * Desloca uma coluna específica em todas as linhas.
	 *
	 * @param colunaOrigem Índice de origem.
	 * @param colunaAlvo   Novo índice.
	 * @param ultimaLinha  Última linha da planilha.
	 */
	private void deslocarColuna(int colunaOrigem, int colunaAlvo, int ultimaLinha) {
		for (int rowNum = 0; rowNum <= ultimaLinha; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null)
				continue;
			Cell celulaOrigem = row.getCell(colunaOrigem + columnOffset);
			if (celulaOrigem != null) {
				Cell celulaAlvo = row.createCell(colunaAlvo + columnOffset);
				copiarValorEntreCelulas(celulaOrigem, celulaAlvo);
				copiarEstiloEntreCelulas(celulaOrigem, celulaAlvo);
				row.removeCell(celulaOrigem);
			} else {
				removerCelula(row, colunaAlvo + columnOffset);
			}
		}
	}

	/**
	 * Remove uma célula específica.
	 *
	 * @param row          A linha.
	 * @param indiceCelula Índice da célula.
	 */
	private void removerCelula(Row row, int indiceCelula) {
		Cell celula = row.getCell(indiceCelula);
		if (celula != null) {
			row.removeCell(celula);
		}
	}

	/**
	 * Copia o valor de uma célula para outra.
	 *
	 * @param origem  A célula de origem.
	 * @param destino A célula de destino.
	 */
	private void copiarValorEntreCelulas(Cell origem, Cell destino) {
		destino.setCellType(origem.getCellType());
		switch (origem.getCellType()) {
		case STRING:
			destino.setCellValue(origem.getStringCellValue());
			break;
		case NUMERIC:
			destino.setCellValue(origem.getNumericCellValue());
			break;
		case BOOLEAN:
			destino.setCellValue(origem.getBooleanCellValue());
			break;
		case FORMULA:
			destino.setCellFormula(origem.getCellFormula());
			break;
		case ERROR:
			destino.setCellErrorValue(origem.getErrorCellValue());
			break;
		case BLANK:
			break;
		default:
			break;
		}
	}

	/**
	 * Copia o estilo de uma célula para outra.
	 *
	 * @param origem  A célula de origem.
	 * @param destino A célula de destino.
	 */
	private void copiarEstiloEntreCelulas(Cell origem, Cell destino) {
		destino.setCellStyle(origem.getCellStyle());
	}

	/**
	 * Classe auxiliar para armazenar os dados de uma célula.
	 */
	public static class CellData {
		private String stringValue;
		private String formulaValue;
		private double numericValue;
		private boolean booleanValue;
		private byte errorValue;
		private CellType cellType;
		private CellStyle cellStyle;

		public String getStringValue() {
			return stringValue;
		}

		public void setStringValue(String stringValue) {
			this.stringValue = stringValue;
		}

		public String getFormulaValue() {
			return formulaValue;
		}

		public void setFormulaValue(String formulaValue) {
			this.formulaValue = formulaValue;
		}

		public double getNumericValue() {
			return numericValue;
		}

		public void setNumericValue(double numericValue) {
			this.numericValue = numericValue;
		}

		public boolean isBooleanValue() {
			return booleanValue;
		}

		public void setBooleanValue(boolean booleanValue) {
			this.booleanValue = booleanValue;
		}

		public byte getErrorValue() {
			return errorValue;
		}

		public void setErrorValue(byte errorValue) {
			this.errorValue = errorValue;
		}

		public CellType getCellType() {
			return cellType;
		}

		public void setCellType(CellType cellType) {
			this.cellType = cellType;
		}

		public CellStyle getCellStyle() {
			return cellStyle;
		}

		public void setCellStyle(CellStyle cellStyle) {
			this.cellStyle = cellStyle;
		}
	}
}
