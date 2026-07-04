package com.abnote.planilhas.impl;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.interfaces.IManipulacaoDados;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.interfaces.ISelecao;
import com.abnote.planilhas.utils.LoggerUtil;
import com.abnote.planilhas.utils.ManipuladorPlanilha;
import com.abnote.planilhas.utils.PosicaoConverter;
import com.abnote.planilhas.utils.PositionManager;
import com.abnote.planilhas.formulas.FormulaBuilder;
import com.abnote.planilhas.interfaces.IConversao;
import com.abnote.planilhas.interfaces.IFormulas;

public abstract class PlanilhaBase implements IPlanilha {

	private static final Logger logger = LoggerUtil.getLogger(PlanilhaBase.class);

	protected Workbook workbook;
	protected Sheet sheet;
	protected final PositionManager positionManager = new PositionManager();
	protected DataManipulator dataManipulator;
	protected StyleManager styleManager;
	protected ConversaoManager conversaoManager;
	protected SelecaoManager selecaoManager;
	private String diretorioSaida = "C:\\opt\\tmp\\testePlanilhaSaidas";

	protected abstract void inicializarWorkbook();

	// Método privado para (re)inicializar os manipuladores de dados e estilos
	private void initManipulators() {
	    positionManager.resetarPosicao();
	    dataManipulator = new DataManipulator(workbook, sheet, positionManager);
	    styleManager = new StyleManager(workbook, sheet, positionManager, dataManipulator);
	    conversaoManager = new ConversaoManager(sheet, workbook, this);
	    selecaoManager = new SelecaoManager(this, dataManipulator, positionManager);
	}
	@Override
	public void criarPlanilha(String nomeSheet) {
		logger.info("Iniciando a criação da planilha: " + nomeSheet);
		try {
			inicializarWorkbook();
			sheet = workbook.createSheet(nomeSheet);
			initManipulators();
			logger.info("Planilha '" + nomeSheet + "' criada com sucesso.");
		} catch (Exception e) {
			logger.severe("Erro ao criar a planilha '" + nomeSheet + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public void abrirPlanilha(String caminhoArquivo) {
		if (caminhoArquivo == null || caminhoArquivo.trim().isEmpty()) {
			throw new ArquivoException("Caminho do arquivo não pode ser nulo ou vazio", caminhoArquivo);
		}
		try (FileInputStream arquivoEntrada = new FileInputStream(caminhoArquivo)) {
			workbook = WorkbookFactory.create(arquivoEntrada);
			sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Planilha1");
			initManipulators();
			logger.info("Planilha aberta com sucesso: " + caminhoArquivo);
		} catch (IOException e) {
			logger.severe("Erro ao abrir a planilha '" + caminhoArquivo + "': " + e.getMessage());
			throw new ArquivoException("Erro ao abrir planilha. Verifique se o caminho existe e é um arquivo válido",
					caminhoArquivo, e);
		}
	}

	@Override
	public void criarSheet(String nomeSheet) {
		try {
			if (workbook.getSheet(nomeSheet) != null) {
				String msg = "A aba '" + nomeSheet + "' já existe!";
				logger.warning(msg);
				throw new IllegalArgumentException(msg);
			}
			sheet = workbook.createSheet(nomeSheet);
			initManipulators();
		} catch (Exception e) {
			logger.severe("Erro ao criar a aba '" + nomeSheet + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public void selecionarSheet(String nomeSheet) {
		logger.fine("Atuando na Sheet: " + nomeSheet);
		try {
			if (workbook == null) {
				String msg = "Workbook ainda não foi inicializado!";
				logger.severe(msg);
				throw new IllegalStateException(msg);
			}
			sheet = workbook.getSheet(nomeSheet);
			if (sheet == null) {
				String msg = "A aba '" + nomeSheet + "' não foi encontrada.";
				logger.warning(msg);
				throw new IllegalArgumentException(msg);
			}
			initManipulators();
		} catch (Exception e) {
			logger.severe("Erro ao selecionar a aba '" + nomeSheet + "': " + e.getMessage());
			throw e;
		}
	}

    @Override
    public void salvar(String nomeArquivo) { 
        if (nomeArquivo == null || nomeArquivo.trim().isEmpty()) {
            throw new ArquivoException(
                "Nome do arquivo não pode ser nulo ou vazio",
                nomeArquivo
            );
        }
        
        try (FileOutputStream arquivoSaida = new FileOutputStream(nomeArquivo)) {
            workbook.write(arquivoSaida);
            logger.info("Planilha salva com sucesso em: " + nomeArquivo);
        } catch (IOException e) {
            logger.severe("Erro ao salvar a planilha em '" + nomeArquivo + "': " + e.getMessage());
            throw new ArquivoException(
                "Erro ao salvar planilha. Verifique permissões e espaço em disco",
                nomeArquivo,
                e
            );
        }
    }
	
	@Override
	public IFormulas formula() {
	    int linhaAtual = dataManipulator.getLinhaSelecionadaAtual();
	    int colunaAtual = dataManipulator.getColunaSelecionadaAtual();
	    
	    if (linhaAtual < 0 || colunaAtual < 0) {
	        throw new IllegalStateException(
	            "Nenhuma célula foi selecionada. Use naCelula() antes de formula()"
	        );
	    }
	    
	    return new FormulaBuilder(workbook, sheet, linhaAtual, colunaAtual, this);
	}

	@Override
	public void setDiretorioSaida(String diretorioSaida) {
		this.diretorioSaida = diretorioSaida;
	}

	@Override
	public String getDiretorioSaida() {
		return diretorioSaida;
	}

	@Override
	public Workbook obterWorkbook() {
		return workbook;
	}

	@Override
	public IPlanilha emTodaAPlanilha() {
		try {
			positionManager.emTodaAPlanilha();
			return this;
		} catch (Exception e) {
			logger.severe("Erro ao aplicar operações em toda a planilha: " + e.getMessage());
			throw e;
		}
	}

	@Override
	public int getNumeroDeLinhas(String coluna) {
		int colunaIndex = PosicaoConverter.converterColuna(coluna);
		int lastRowNum = sheet.getLastRowNum();
		int numRows = 0;
		for (int i = 0; i <= lastRowNum; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(colunaIndex);
				if (cell != null && cell.getCellType() != CellType.BLANK) {
					numRows++;
				}
			}
		}
		return numRows;
	}

	@Override
	public int getNumeroDeColunasNaLinha(int linha) {
		int linhaIndex = linha - 1;
		Row row = sheet.getRow(linhaIndex);
		if (row == null) {
			return 0;
		}
		int numCols = 0;
		short lastCellNum = row.getLastCellNum();
		for (int i = 0; i < lastCellNum; i++) {
			Cell cell = row.getCell(i);
			if (cell != null && cell.getCellType() != CellType.BLANK) {
				numCols++;
			}
		}
		return numCols;
	}

	@Override
	public IPlanilha ultimaLinha(String coluna) {
		try {
			dataManipulator.ultimaLinha(coluna);
			return this;
		} catch (Exception e) {
			logger.severe("Erro ao encontrar a última linha na coluna '" + coluna + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public IManipulacaoDados naUltimaLinha(String coluna) {
		try {
			dataManipulator.naUltimaLinha(coluna);
			return dataManipulator;
		} catch (Exception e) {
			logger.severe("Erro ao encontrar a última linha na coluna '" + coluna + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public ManipuladorPlanilha manipularPlanilha() {
		return new ManipuladorPlanilha(sheet);
	}

	// Delegação dos métodos de IManipulacaoDados para dataManipulator

	@Override
	public IPlanilha inserirDados(Object dados, String delimitador) {
		dataManipulator.inserirDados(dados, delimitador);
		return this;
	}

	@Override
	public IPlanilha inserirDados(String valor) {
		dataManipulator.inserirDados(valor);
		return this;
	}

	@Override
	public IPlanilha inserirDados(java.util.List<String> dados) {
		dataManipulator.inserirDados(dados);
		return this;
	}

	@Override
	public IPlanilha inserirDados(java.util.List<String> dados, String delimitador) {
		dataManipulator.inserirDados(dados, delimitador);
		return this;
	}

	@Override
	public IPlanilha inserirDadosArquivo(String caminhoArquivo, String delimitador) {
		try {
			dataManipulator.inserirDadosArquivo(caminhoArquivo, delimitador);
		} catch (Exception e) {
			logger.severe("Erro ao inserir dados do arquivo '" + caminhoArquivo + "': " + e.getMessage());
			throw e;
		}
		return this;
	}

	@Override
	public IPlanilha somarColuna(String posicaoInicial) {
		dataManipulator.somarColuna(posicaoInicial);
		return this;
	}

	@Override
	public IPlanilha somarColunaComTexto(String posicaoInicial, String texto) {
		dataManipulator.somarColunaComTexto(posicaoInicial, texto);
		return this;
	}

	@Override
	public IPlanilha multiplicarColunasComTexto(String coluna1, String coluna2, int linhaInicial, String texto,
			String colunaDestino) {
		dataManipulator.multiplicarColunasComTexto(coluna1, coluna2, linhaInicial, texto, colunaDestino);
		return this;
	}

	@Override
	public IPlanilha mesclarCelulas() {
		dataManipulator.mesclarCelulas();
		return this;
	}

	@Override
	public IPlanilha inserirFiltros() {
		try {
			if (sheet == null) {
				throw new IllegalStateException(
						"Sheet não foi inicializada. Crie ou selecione uma planilha antes de inserir filtros.");
			}
			Row headerRow = encontrarLinhaDeCabecalho(sheet);
			if (headerRow != null) {
			    int headerRowIndex = headerRow.getRowNum();
			    int firstColumn = -1;
			    int lastColumn = -1;
			    short lastCellNum = headerRow.getLastCellNum();
			    for (int c = 0; c < lastCellNum; c++) {
			        Cell cell = headerRow.getCell(c);
			        if (cell != null && cell.getCellType() != CellType.BLANK && !cell.toString().trim().isEmpty()) {
			            if (firstColumn == -1) {
			                firstColumn = c;
			            }
			            lastColumn = c;
			        }
				}
				if (firstColumn != -1 && lastColumn != -1 && lastColumn >= firstColumn) {
					CellRangeAddress range = new CellRangeAddress(headerRowIndex, headerRowIndex, firstColumn,
							lastColumn);
					sheet.setAutoFilter(range);
					logger.info("Filtros aplicados na linha de cabeçalho: " + (headerRowIndex + 1) + ", colunas: "
							+ (firstColumn + 1) + " até " + (lastColumn + 1));
				} else {
					logger.warning("Não foi possível identificar colunas para aplicar filtros.");
				}
			} else {
				logger.warning("Não foi encontrada uma linha de cabeçalho para aplicar filtros.");
			}
		} catch (Exception e) {
			logger.severe("Erro ao inserir filtros: " + e.getMessage());
			throw e;
		}
		return this;
	}

	// Método auxiliar para encontrar a primeira linha com conteúdo não vazio(cabeçalho)
	private Row encontrarLinhaDeCabecalho(Sheet sheet) {
	    for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
	        Row row = sheet.getRow(i);
	        if (row != null && row.getLastCellNum() > 0) {
	            boolean linhaTemConteudo = false;
	            for (int c = 0; c < row.getLastCellNum(); c++) {
	                Cell cell = row.getCell(c);
	                if (cell != null && cell.getCellType() != CellType.BLANK && !cell.toString().trim().isEmpty()) {
	                    linhaTemConteudo = true;
	                    break;
	                }
	            }
	            if (linhaTemConteudo) {
	                return row;
	            }
	        }
	    }
	    return null;
	}

	// Delegação dos métodos de IEstilos para o StyleManager

	@Override
	public EstiloCelula aplicarEstilos() {
		return styleManager.aplicarEstilos();
	}

	@Override
	public EstiloCelula centralizarTudo() {
		return styleManager.centralizarTudo();
	}

	@Override
	public EstiloCelula redimensionarColunas() {
		return styleManager.redimensionarColunas();
	}

	@Override
	public EstiloCelula removerLinhasDeGrade() {
		return styleManager.removerLinhasDeGrade();
	}

	@Override
	public EstiloCelula aplicarEstilosEmCelula() {
		if (dataManipulator.getUltimoIndiceDeLinhaInserido() == -1
				|| dataManipulator.getUltimoIndiceDeColunaInserido() == -1) {
			return new EstiloCelula(workbook, sheet, -1, -1);
		}
		return new EstiloCelula(workbook, sheet, dataManipulator.getUltimoIndiceDeLinhaInserido(),
				dataManipulator.getUltimoIndiceDeColunaInserido());
	}

	@Override
	public EstiloCelula todasAsBordasEmTudo() {
		return new EstiloCelula(workbook, sheet).contornarTudo();
	}
	
	@Override
	public IConversao converter() {
	    return conversaoManager;
	}

	@Override
	public ISelecao selecionar() {
	    return selecaoManager;
	}
	
	/**
	 * Fecha o workbook e libera todos os recursos associados.
	 * Este método deve ser chamado sempre que terminar de trabalhar com a planilha.
	 * Recomenda-se usar try-with-resources para garantir fechamento automático.
	 * 
	 * @throws Exception se houver erro ao fechar o workbook
	 */
	@Override
	public void close() throws Exception {
	    if (workbook != null) {
	        try {
	            workbook.close();
	            logger.info("Workbook fechado com sucesso.");
	        } catch (IOException e) {
	            logger.severe("Erro ao fechar workbook: " + e.getMessage());
	            throw e;
	        }
	    }
	}
}
