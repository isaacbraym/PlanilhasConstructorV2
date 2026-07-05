package com.abnote.planilhas.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Pattern;
import java.util.logging.Logger;

import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.exceptions.DadosInvalidosException;

import org.apache.poi.ss.usermodel.*;

public class InsersorDeDados {

	private static final Logger logger = LoggerUtil.getLogger(InsersorDeDados.class);

	/**
	 * Reconhece apenas números "limpos": inteiros ou decimais, com sinal opcional e
	 * notação científica opcional. Evita que valores como "NaN", "Infinity" ou "1d"
	 * (aceitos por {@link Double#parseDouble}) sejam gravados como número por engano.
	 */
	private static final Pattern NUMERO_ESTRITO = Pattern.compile("^[+-]?(\\d+(\\.\\d+)?|\\.\\d+)([eE][+-]?\\d+)?$");

	private final Sheet sheet;
	private final PositionManager positionManager;
	private int ultimoIndiceDeLinhaInserido = -1;
	private int ultimoIndiceDeColunaInserido = -1;

	public InsersorDeDados(Sheet sheet, PositionManager positionManager) {
		this.sheet = sheet;
		this.positionManager = positionManager;
	}

	public void inserirDados(String valor) {
		definirPosicaoPadraoSeNecessario();

		Row linha = obterOuCriarLinha(positionManager.getPosicaoInicialLinha());
		Cell celula = linha.createCell(positionManager.getPosicaoInicialColuna());

		// Tenta converter para número, se falhar insere como string
		definirValorCelula(celula, valor);

		atualizarIndicesInseridos(positionManager.getPosicaoInicialLinha(), positionManager.getPosicaoInicialColuna());
	}

	public void inserirDados(Object dados, String delimitador) {
		if (dados == null) {
			throw new DadosInvalidosException("Dados não podem ser nulos");
		}
		if (dados instanceof List) {
			@SuppressWarnings("unchecked")
			List<String> lista = (List<String>) dados;
			inserirDados(lista, delimitador);
		} else if (dados instanceof String) {
			String str = (String) dados;
			if (Files.exists(Paths.get(str))) {
				inserirDadosArquivo(str, delimitador);
			} else {
				List<String> lista = Arrays.asList(dividirPreservandoVaziosFinais(str, delimitador));
				inserirDados(lista);
			}
		} else if (dados instanceof File) {
			inserirDadosArquivo(((File) dados).getPath(), delimitador);
		} else {
			throw new IllegalArgumentException("Tipo de dados não suportado: " + dados.getClass());
		}
	}

	public void inserirDados(List<String> dados) {
		validarDadosNaoNulos(dados);
		definirPosicaoPadraoSeNecessario();

		if (positionManager.isIntervaloDefinida()) {
			inserirDadosEmIntervalo(dados);
		} else {
			inserirDadosEmLinha(dados);
		}

		positionManager.resetarPosicao();
	}

	public void inserirDados(List<String> dados, String delimitador) {
		validarDadosNaoNulos(dados);
		if (delimitador == null) {
			throw new DadosInvalidosException("Delimitador não pode ser nulo");
		}
		if (delimitador.isEmpty()) {
			inserirDados(dados);
			return;
		}

		definirPosicaoPadraoSeNecessario();

		int linhaAtual = positionManager.getPosicaoInicialLinha();
		for (String linhaTexto : dados) {
			if (positionManager.isIntervaloDefinida() && linhaAtual > positionManager.getPosicaoFinalLinha()) {
				break;
			}
			String textoSeguro = linhaTexto == null ? "" : linhaTexto;
			String[] valores = dividirPreservandoVaziosFinais(textoSeguro, delimitador);
			inserirValoresEmLinha(linhaAtual, valores);
			linhaAtual++;
		}

		atualizarIndicesInseridos(linhaAtual - 1, ultimoIndiceDeColunaInserido);
		positionManager.setPosicaoInicialLinha(linhaAtual);
		positionManager.resetarPosicao();
	}

	public void inserirDadosArquivo(String caminhoArquivo, String delimitador) {
        definirPosicaoPadraoSeNecessario();
        
        if (caminhoArquivo == null || caminhoArquivo.trim().isEmpty()) {
            throw new ArquivoException(
                "Caminho do arquivo não pode ser nulo ou vazio",
                caminhoArquivo
            );
        }
        
        // [REMOVED] Validação do delimitador removida (pode ser vazio em casos válidos)

        try (BufferedReader br = new BufferedReader(new FileReader(caminhoArquivo))) {
            String linhaTexto;
            int linhaAtual = positionManager.getPosicaoInicialLinha();

            while ((linhaTexto = br.readLine()) != null) {
                String[] valores = dividirPreservandoVaziosFinais(linhaTexto, delimitador);
                inserirValoresEmLinha(linhaAtual, valores);

                linhaAtual++;
                if (positionManager.isIntervaloDefinida() && linhaAtual > positionManager.getPosicaoFinalLinha()) {
                    break;
                }
            }

            atualizarIndicesInseridos(linhaAtual - 1, positionManager.getPosicaoInicialColuna());
            positionManager.setPosicaoInicialLinha(linhaAtual);

        } catch (IOException e) {
            logger.severe("Erro ao ler o arquivo: " + e.getMessage());
            throw new ArquivoException(
                "Erro ao ler arquivo. Verifique se o arquivo existe e está acessível",
                caminhoArquivo,
                e
            );
        }

        positionManager.resetarPosicao();
    }

	private static String[] dividirPreservandoVaziosFinais(final String texto, final String delimitador) {
		if (delimitador == null) {
			throw new DadosInvalidosException("Delimitador não pode ser nulo");
		}
		if (delimitador.isEmpty()) {
			return texto.split(Pattern.quote(delimitador));
		}
		return texto.split(Pattern.quote(delimitador), -1);
	}

	/**
	 * Define o valor da célula, tentando converter para número quando possível.
	 * 
	 * @param celula Célula a receber o valor
	 * @param valor  String com o valor a ser inserido
	 */
	private void definirValorCelula(Cell celula, String valor) {
		if (valor == null || valor.trim().isEmpty()) {
			celula.setCellValue("");
			return;
		}

		String valorTrimmed = valor.trim();

		if (deveManterComoTexto(valorTrimmed)) {
			celula.setCellValue(valorTrimmed);
			return;
		}

		// Tenta converter para número
		try {
			double numeroDouble = Double.parseDouble(valorTrimmed);
			celula.setCellValue(numeroDouble);
		} catch (NumberFormatException e) {
			// Não é número, insere como string
			celula.setCellValue(valorTrimmed);
		}
	}

	/**
	 * Decide se um valor deve ser preservado como texto em vez de virar número.
	 *
	 * <p>
	 * Protege dados comuns de quem não programa contra a coerção numérica agressiva:
	 * </p>
	 * <ul>
	 * <li>Textos que não são números limpos (ex.: "NaN", "Infinity", "1d");</li>
	 * <li>Inteiros com zero à esquerda (ex.: "007", "01234" — CEP, códigos);</li>
	 * <li>Sequências longas de dígitos que perderiam precisão como {@code double}
	 * (ex.: CPF, CNPJ, cartões — 16+ dígitos).</li>
	 * </ul>
	 *
	 * @param valor Valor já sem espaços nas pontas.
	 * @return {@code true} se deve ser gravado como texto.
	 */
	private boolean deveManterComoTexto(final String valor) {
		if (!NUMERO_ESTRITO.matcher(valor).matches()) {
			return true;
		}
		final String semSinal = (valor.charAt(0) == '+' || valor.charAt(0) == '-') ? valor.substring(1) : valor;
		final boolean temZeroAEsquerda = semSinal.length() > 1 && semSinal.charAt(0) == '0'
				&& Character.isDigit(semSinal.charAt(1));
		if (temZeroAEsquerda) {
			return true;
		}
		final boolean inteiroMuitoLongo = semSinal.indexOf('.') < 0 && semSinal.length() >= 16;
		return inteiroMuitoLongo;
	}

	// Métodos auxiliares privados

	private void definirPosicaoPadraoSeNecessario() {
		if (!positionManager.isPosicaoDefinida() && !positionManager.isIntervaloDefinida()) {
			positionManager.setPosicaoInicialColuna(0);
			positionManager.setPosicaoInicialLinha(0);
		}
	}

	private void validarDadosNaoNulos(List<String> dados) {
		if (dados == null) {
			throw new DadosInvalidosException("Dados não podem ser nulos");
		}
	}

	private Row obterOuCriarLinha(int indiceLinha) {
		Row linha = sheet.getRow(indiceLinha);
		if (linha == null) {
			linha = sheet.createRow(indiceLinha);
		}
		return linha;
	}

	private void inserirDadosEmLinha(List<String> dados) {
		Row linha = obterOuCriarLinha(positionManager.getPosicaoInicialLinha());

		for (int i = 0; i < dados.size(); i++) {
			Cell celula = linha.createCell(positionManager.getPosicaoInicialColuna() + i);
			definirValorCelula(celula, dados.get(i)); // ✅ MUDANÇA AQUI
			ultimoIndiceDeColunaInserido = positionManager.getPosicaoInicialColuna() + i;
		}

		atualizarIndicesInseridos(positionManager.getPosicaoInicialLinha(), ultimoIndiceDeColunaInserido);
		positionManager.setPosicaoInicialLinha(positionManager.getPosicaoInicialLinha() + 1);
	}

	private void inserirDadosEmIntervalo(List<String> dados) {
		int linhaAtual = positionManager.getPosicaoInicialLinha();

		for (String dado : dados) {
			if (linhaAtual > positionManager.getPosicaoFinalLinha()) {
				break;
			}

			Row linha = obterOuCriarLinha(linhaAtual);

			for (int coluna = positionManager.getPosicaoInicialColuna(); coluna <= positionManager
					.getPosicaoFinalColuna(); coluna++) {
				Cell celula = linha.createCell(coluna);
				definirValorCelula(celula, dado); // ✅ MUDANÇA AQUI
			}

			linhaAtual++;
		}

		atualizarIndicesInseridos(linhaAtual - 1, positionManager.getPosicaoFinalColuna());
	}

	private void inserirValoresEmLinha(int indiceLinha, String[] valores) {
		Row linha = obterOuCriarLinha(indiceLinha);

		for (int i = 0; i < valores.length; i++) {
			int colunaAtual = positionManager.getPosicaoInicialColuna() + i;

			if (positionManager.isIntervaloDefinida() && colunaAtual > positionManager.getPosicaoFinalColuna()) {
				break;
			}

			Cell celula = linha.createCell(colunaAtual);
			definirValorCelula(celula, valores[i].trim()); // ✅ MUDANÇA AQUI
			ultimoIndiceDeColunaInserido = colunaAtual;
		}
	}

	private void atualizarIndicesInseridos(int linha, int coluna) {
		ultimoIndiceDeLinhaInserido = linha;
		ultimoIndiceDeColunaInserido = coluna;
	}

	// Getters para obter os últimos índices inseridos

	public int getUltimoIndiceDeLinhaInserido() {
		return ultimoIndiceDeLinhaInserido;
	}

	public int getUltimoIndiceDeColunaInserido() {
		return ultimoIndiceDeColunaInserido;
	}
}
