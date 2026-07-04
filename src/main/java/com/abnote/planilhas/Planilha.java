package com.abnote.planilhas;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.exceptions.DadosInvalidosException;
import com.abnote.planilhas.graficos.GraficoHelper;
import com.abnote.planilhas.imagens.ImagemHelper;
import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.utils.ComentarioHelper;
import com.abnote.planilhas.utils.CopiadorDeCelulas;
import com.abnote.planilhas.utils.FiltroDeLinhas;
import com.abnote.planilhas.utils.FormatacaoCondicionalHelper;
import com.abnote.planilhas.utils.FormatosDeCelula;
import com.abnote.planilhas.utils.ListaSuspensaHelper;
import com.abnote.planilhas.utils.OrdenadorDeLinhas;
import com.abnote.planilhas.utils.PosicaoConverter;
import com.abnote.planilhas.utils.ProtecaoHelper;
import com.abnote.planilhas.utils.TotalizadorDeTabela;
import com.abnote.planilhas.utils.ValidacaoDeEntradaHelper;

/**
 * Ponto de entrada <strong>amigável</strong> para criar planilhas Excel sem
 * precisar entender Apache POI nem a API fluente interna.
 *
 * <p>
 * Toda a API é feita de verbos simples que devolvem a própria planilha, então
 * você pode "encaixar" um comando no outro:
 * </p>
 *
 * <pre>{@code
 * try (Planilha planilha = Planilha.nova("Vendas")) {
 *     planilha.escreverLinha("A1", "Produto", "Preço", "Quantidade")
 *             .negrito("A1:C1")
 *             .adicionarLinha("Caneta", 2.50, 100)
 *             .adicionarLinha("Caderno", 15.90, 30)
 *             .somar("C5", "C2:C3")
 *             .formatarComoMoeda("B2")
 *             .ajustarColunas()
 *             .salvar("C:/tmp/vendas.xlsx");
 * }
 * }</pre>
 *
 * <p>
 * Precisa de algo avançado que não tem aqui? Use {@link #avancado()} para
 * acessar a API completa ({@link IPlanilha}) ou {@link #workbook()} para o
 * objeto do Apache POI.
 * </p>
 */
public final class Planilha implements AutoCloseable {

	private final IPlanilha planilha;
	private final FormatosDeCelula formatos;
	private String abaAtual;

	private Planilha(final IPlanilha planilhaInterna, final String abaAtiva) {
		this.planilha = planilhaInterna;
		this.abaAtual = abaAtiva;
		this.formatos = new FormatosDeCelula(planilhaInterna.obterWorkbook());
	}

	/**
	 * Cria uma planilha nova, já com a primeira aba pronta para uso.
	 *
	 * @param nomeDaAba Nome da primeira aba (ex.: "Vendas").
	 * @return A planilha criada.
	 */
	public static Planilha nova(final String nomeDaAba) {
		final PlanilhaXlsx interna = new PlanilhaXlsx();
		interna.criarPlanilha(nomeDaAba);
		return new Planilha(interna, nomeDaAba);
	}

	/**
	 * Abre uma planilha {@code .xlsx} existente para continuar editando com os
	 * mesmos comandos, posicionando-se na primeira aba.
	 *
	 * @param caminhoArquivo Caminho do arquivo a abrir (ex.: "C:/tmp/dados.xlsx").
	 * @return A planilha carregada do arquivo.
	 * @throws com.abnote.planilhas.exceptions.ArquivoException se o arquivo não
	 *         existir, não puder ser lido, ou não for um {@code .xlsx} (o
	 *         formato antigo {@code .xls} não é suportado — abra no Excel e
	 *         salve novamente como {@code .xlsx}).
	 */
	public static Planilha abrir(final String caminhoArquivo) {
		final PlanilhaXlsx interna = new PlanilhaXlsx();
		interna.abrirPlanilha(caminhoArquivo);
		return new Planilha(interna, interna.obterWorkbook().getSheetAt(0).getSheetName());
	}

	// ==================== ABAS ====================

	/**
	 * Cria uma nova aba e passa a trabalhar nela.
	 *
	 * @param nome Nome da nova aba.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha novaAba(final String nome) {
		planilha.criarSheet(nome);
		return irParaAba(nome);
	}

	/**
	 * Passa a trabalhar em uma aba já existente.
	 *
	 * @param nome Nome da aba.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha irParaAba(final String nome) {
		planilha.selecionarSheet(nome);
		this.abaAtual = nome;
		return this;
	}

	/**
	 * Duplica a aba atual com todo o conteúdo e formatação, e passa a trabalhar na
	 * cópia.
	 *
	 * @param novoNome Nome da aba duplicada.
	 * @return Esta planilha, agora posicionada na cópia.
	 */
	public Planilha duplicarAba(final String novoNome) {
		final Workbook workbook = planilha.obterWorkbook();
		final Sheet copia = workbook.cloneSheet(workbook.getSheetIndex(abaAtual));
		workbook.setSheetName(workbook.getSheetIndex(copia), novoNome);
		return irParaAba(novoNome);
	}

	// ==================== ESCRITA ====================

	/**
	 * Escreve um valor em uma célula. Números viram número; textos viram texto
	 * (com detecção inteligente para não estragar CEP/CPF com zero à esquerda).
	 *
	 * @param celula Posição da célula (ex.: "A1").
	 * @param valor  Valor a escrever (texto, número, etc.). {@code null} vira vazio.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escrever(final String celula, final Object valor) {
		if (valor == null) {
			return escreverTexto(celula, "");
		}
		if (valor instanceof Number) {
			planilha.selecionar().celula(celula).inserir(((Number) valor).doubleValue());
		} else {
			planilha.selecionar().celula(celula).inserirDados(valor.toString());
		}
		return this;
	}

	/**
	 * Escreve um valor <strong>sempre</strong> como texto, preservando exatamente o
	 * que foi digitado (ideal para CEP, CPF, telefone e códigos com zero à
	 * esquerda).
	 *
	 * @param celula Posição da célula (ex.: "A1").
	 * @param texto  Texto a escrever.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escreverTexto(final String celula, final String texto) {
		final int[] indices = PosicaoConverter.converterPosicao(celula);
		obterOuCriarCelula(indices[1], indices[0]).setCellValue(texto == null ? "" : texto);
		return this;
	}

	/**
	 * Escreve vários valores em sequência, na horizontal, a partir de uma célula.
	 *
	 * @param celulaInicial Célula onde começa (ex.: "A1").
	 * @param valores       Valores a escrever, um por coluna.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escreverLinha(final String celulaInicial, final Object... valores) {
		final int[] inicio = PosicaoConverter.converterPosicao(celulaInicial);
		for (int deslocamento = 0; deslocamento < valores.length; deslocamento++) {
			final String posicao = PosicaoConverter.converterIndice(inicio[0] + deslocamento) + (inicio[1] + 1);
			escrever(posicao, valores[deslocamento]);
		}
		return this;
	}

	/**
	 * Escreve vários valores em sequência, na vertical, a partir de uma célula.
	 *
	 * @param celulaInicial Célula onde começa (ex.: "A1").
	 * @param valores       Valores a escrever, um por linha.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escreverColuna(final String celulaInicial, final Object... valores) {
		final int[] inicio = PosicaoConverter.converterPosicao(celulaInicial);
		for (int deslocamento = 0; deslocamento < valores.length; deslocamento++) {
			final String posicao = PosicaoConverter.converterIndice(inicio[0]) + (inicio[1] + 1 + deslocamento);
			escrever(posicao, valores[deslocamento]);
		}
		return this;
	}

	/**
	 * Acrescenta uma linha de dados logo após a última linha já preenchida — ótimo
	 * para ir montando uma lista de itens.
	 *
	 * @param valores Valores da nova linha, um por coluna (a partir da coluna A).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha adicionarLinha(final Object... valores) {
		final Sheet sheet = sheetAtual();
		final int proximaLinha = sheet.getPhysicalNumberOfRows() == 0 ? 0 : sheet.getLastRowNum() + 1;
		return escreverLinha("A" + (proximaLinha + 1), valores);
	}

	/**
	 * Escreve uma tabela inteira (lista de linhas) a partir de uma célula.
	 *
	 * @param celulaInicial Canto superior esquerdo (ex.: "A1").
	 * @param linhas        Lista de linhas; cada linha é uma lista de valores.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escreverTabela(final String celulaInicial, final List<? extends List<?>> linhas) {
		final int[] inicio = PosicaoConverter.converterPosicao(celulaInicial);
		for (int numeroLinha = 0; numeroLinha < linhas.size(); numeroLinha++) {
			final String posicao = PosicaoConverter.converterIndice(inicio[0]) + (inicio[1] + 1 + numeroLinha);
			escreverLinha(posicao, linhas.get(numeroLinha).toArray());
		}
		return this;
	}

	// ==================== FÓRMULAS ====================

	/**
	 * Soma um intervalo e coloca o resultado em uma célula.
	 *
	 * @param celulaDestino Onde a soma aparece (ex.: "C5").
	 * @param intervalo     Intervalo a somar (ex.: "C2:C4").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha somar(final String celulaDestino, final String intervalo) {
		planilha.selecionar().celula(celulaDestino).formula().soma(intervalo).aplicar();
		return this;
	}

	/**
	 * Calcula a média de um intervalo e coloca o resultado em uma célula.
	 *
	 * @param celulaDestino Onde a média aparece.
	 * @param intervalo     Intervalo (ex.: "C2:C4").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha media(final String celulaDestino, final String intervalo) {
		planilha.selecionar().celula(celulaDestino).formula().media(intervalo).aplicar();
		return this;
	}

	/**
	 * Conta os valores numéricos de um intervalo.
	 *
	 * @param celulaDestino Onde a contagem aparece.
	 * @param intervalo     Intervalo (ex.: "C2:C4").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha contar(final String celulaDestino, final String intervalo) {
		planilha.selecionar().celula(celulaDestino).formula().contar(intervalo).aplicar();
		return this;
	}

	/**
	 * Encontra o menor valor de um intervalo.
	 *
	 * @param celulaDestino Onde o mínimo aparece.
	 * @param intervalo     Intervalo (ex.: "C2:C4").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha minimo(final String celulaDestino, final String intervalo) {
		planilha.selecionar().celula(celulaDestino).formula().minimo(intervalo).aplicar();
		return this;
	}

	/**
	 * Encontra o maior valor de um intervalo.
	 *
	 * @param celulaDestino Onde o máximo aparece.
	 * @param intervalo     Intervalo (ex.: "C2:C4").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha maximo(final String celulaDestino, final String intervalo) {
		planilha.selecionar().celula(celulaDestino).formula().maximo(intervalo).aplicar();
		return this;
	}

	/**
	 * Escreve um "se/então": testa uma condição e mostra um valor quando verdadeira
	 * e outro quando falsa.
	 *
	 * @param celulaDestino   Onde o resultado aparece.
	 * @param condicao        Condição (ex.: "B2&gt;100").
	 * @param seVerdadeiro    Valor quando a condição é verdadeira.
	 * @param seFalso         Valor quando a condição é falsa.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha seEntao(final String celulaDestino, final String condicao, final Object seVerdadeiro,
			final Object seFalso) {
		planilha.selecionar().celula(celulaDestino).formula().seEntao(condicao, seVerdadeiro, seFalso).aplicar();
		return this;
	}

	/**
	 * Multiplica duas células e coloca o resultado em outra (ex.: total = preço ×
	 * quantidade).
	 *
	 * @param celulaDestino Onde o resultado aparece (ex.: "D2").
	 * @param celula1       Primeira célula (ex.: "B2").
	 * @param celula2       Segunda célula (ex.: "C2").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha multiplicar(final String celulaDestino, final String celula1, final String celula2) {
		return formula(celulaDestino, celula1 + "*" + celula2);
	}

	/**
	 * Subtrai a segunda célula da primeira e coloca o resultado em outra.
	 *
	 * @param celulaDestino Onde o resultado aparece.
	 * @param celula1       Célula do minuendo (ex.: "B2").
	 * @param celula2       Célula do subtraendo (ex.: "C2").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha subtrair(final String celulaDestino, final String celula1, final String celula2) {
		return formula(celulaDestino, celula1 + "-" + celula2);
	}

	/**
	 * Divide a primeira célula pela segunda e coloca o resultado em outra.
	 *
	 * @param celulaDestino Onde o resultado aparece.
	 * @param celula1       Célula do dividendo (ex.: "B2").
	 * @param celula2       Célula do divisor (ex.: "C2").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha dividir(final String celulaDestino, final String celula1, final String celula2) {
		return formula(celulaDestino, celula1 + "/" + celula2);
	}

	/**
	 * Escreve uma fórmula do Excel em uma célula (para quem já conhece fórmulas).
	 * O "=" inicial é opcional.
	 *
	 * @param celulaDestino Onde a fórmula fica (ex.: "D2").
	 * @param formulaExcel  A fórmula (ex.: "B2*C2", "=SUM(A1:A9)").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formula(final String celulaDestino, final String formulaExcel) {
		planilha.selecionar().celula(celulaDestino).formula().personalizada(formulaExcel).aplicar();
		return this;
	}

	/**
	 * Preenche uma coluna com a mesma fórmula, linha por linha, trocando
	 * automaticamente o número da linha onde estiver <code>{}</code>.
	 *
	 * <p>Exemplo — coluna "Total" = Preço (B) × Quantidade (C), das linhas 2 a 10:</p>
	 * <pre>{@code
	 * planilha.preencherColuna("D", 2, 10, "B{}*C{}");
	 * // D2 = B2*C2, D3 = B3*C3, ... D10 = B10*C10
	 * }</pre>
	 *
	 * @param coluna        Coluna de destino (ex.: "D").
	 * @param linhaInicial  Primeira linha (ex.: 2).
	 * @param linhaFinal    Última linha (ex.: 10).
	 * @param modeloFormula Fórmula com <code>{}</code> no lugar do número da linha.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha preencherColuna(final String coluna, final int linhaInicial, final int linhaFinal,
			final String modeloFormula) {
		for (int linha = linhaInicial; linha <= linhaFinal; linha++) {
			formula(coluna + linha, modeloFormula.replace("{}", String.valueOf(linha)));
		}
		return this;
	}

	/**
	 * Procura o valor de uma célula em uma tabela e traz o dado de outra coluna
	 * (equivale ao PROCV/VLOOKUP do Excel, com correspondência exata).
	 *
	 * <p>Exemplo — buscar o preço do produto de A2 na tabela A2:C100:</p>
	 * <pre>{@code
	 * planilha.procurarValor("D2", "A2", "A2:C100", 3); // traz a 3ª coluna (preço)
	 * }</pre>
	 *
	 * @param celulaDestino    Onde o resultado aparece (ex.: "D2").
	 * @param celulaProcurada  Célula com o valor a procurar (ex.: "A2").
	 * @param intervaloTabela  Intervalo da tabela (ex.: "A2:C100").
	 * @param colunaResultado  Número da coluna da tabela a retornar (1 = primeira).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha procurarValor(final String celulaDestino, final String celulaProcurada,
			final String intervaloTabela, final int colunaResultado) {
		final String vlookup = "VLOOKUP(" + celulaProcurada + "," + intervaloTabela + "," + colunaResultado + ",FALSE)";
		return formula(celulaDestino, vlookup);
	}

	/**
	 * Igual a {@link #procurarValor}, mas a tabela está em <strong>outra aba</strong>.
	 *
	 * @param celulaDestino    Onde o resultado aparece.
	 * @param celulaProcurada  Célula com o valor a procurar.
	 * @param abaTabela        Aba onde está a tabela (ex.: "Produtos").
	 * @param intervaloTabela  Intervalo da tabela na outra aba (ex.: "A2:C100").
	 * @param colunaResultado  Número da coluna da tabela a retornar (1 = primeira).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha procurarValorNaAba(final String celulaDestino, final String celulaProcurada,
			final String abaTabela, final String intervaloTabela, final int colunaResultado) {
		return procurarValor(celulaDestino, celulaProcurada, "'" + abaTabela + "'!" + intervaloTabela, colunaResultado);
	}

	/**
	 * Dá um nome a um intervalo, para usar em fórmulas mais legíveis (ex.:
	 * {@code formula("D2", "SUM(Precos)")} em vez de {@code "SUM(B2:B100)"}).
	 *
	 * <p>
	 * Funciona com {@link #formula}, {@link #procurarValor} e
	 * {@link #procurarValorNaAba}. As fórmulas prontas ({@link #somar},
	 * {@link #media} etc.) ainda exigem um intervalo de células (ex.:
	 * {@code "B2:B100"}), não um nome — use {@code formula(celula, "SUM(Nome)")}
	 * nesse caso.
	 * </p>
	 *
	 * @param nome      Nome a definir (ex.: "Precos"). Sem espaços, começando com
	 *                  letra — mesmas regras do Excel para nomes de intervalo.
	 * @param intervalo Intervalo que o nome representa (ex.: "B2:B100"), na aba
	 *                  atual.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha definirNome(final String nome, final String intervalo) {
		final Name nomeDefinido = planilha.obterWorkbook().createName();
		nomeDefinido.setNameName(nome);
		nomeDefinido.setRefersToFormula("'" + abaAtual + "'!" + paraReferenciaAbsoluta(intervalo));
		return this;
	}

	/**
	 * Soma automaticamente cada coluna numérica de uma tabela, inserindo uma
	 * linha de totais logo abaixo da última linha de dados (fórmula
	 * {@code SUM}, não um valor fixo — atualiza sozinha se os dados mudarem).
	 * A primeira coluna não numérica (normalmente o rótulo) recebe o texto
	 * "Total".
	 *
	 * <p>Exemplo — tabela com cabeçalho em A1 e dados até a última linha
	 * preenchida na coluna A:</p>
	 * <pre>{@code
	 * planilha.escreverLinha("A1", "Produto", "Preço", "Qtd")
	 *         .adicionarLinha("Caneta", 2.5, 100)
	 *         .adicionarLinha("Caderno", 15.9, 30)
	 *         .adicionarTotais("A1");
	 * // Insere na linha seguinte: "Total", =SUM(B2:B3), =SUM(C2:C3)
	 * }</pre>
	 *
	 * @param celulaCabecalho Célula do canto superior esquerdo do cabeçalho da
	 *                        tabela (ex.: "A1").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha adicionarTotais(final String celulaCabecalho) {
		final int[] indices = PosicaoConverter.converterPosicao(celulaCabecalho);
		TotalizadorDeTabela.adicionarTotais(sheetAtual(), indices[1], indices[0]);
		return this;
	}

	// ==================== DATAS ====================

	/**
	 * Escreve uma data em uma célula, já formatada como dd/MM/aaaa.
	 *
	 * @param celula Posição da célula (ex.: "A1").
	 * @param data   A data a escrever.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escreverData(final String celula, final LocalDate data) {
		final int[] indices = PosicaoConverter.converterPosicao(celula);
		final Cell alvo = obterOuCriarCelula(indices[1], indices[0]);
		alvo.setCellValue(data);
		alvo.setCellStyle(formatos.data());
		return this;
	}

	/**
	 * Escreve uma data com hora em uma célula, formatada como dd/MM/aaaa HH:mm.
	 *
	 * @param celula   Posição da célula (ex.: "A1").
	 * @param dataHora A data e hora a escrever.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escreverDataHora(final String celula, final LocalDateTime dataHora) {
		final int[] indices = PosicaoConverter.converterPosicao(celula);
		final Cell alvo = obterOuCriarCelula(indices[1], indices[0]);
		alvo.setCellValue(dataHora);
		alvo.setCellStyle(formatos.dataHora());
		return this;
	}

	/**
	 * Aplica o formato de data (dd/MM/aaaa) às células da coluna, a partir da
	 * célula informada (útil quando a coluna já tem datas em formato de número).
	 *
	 * @param celulaInicial Primeira célula da coluna (ex.: "A2").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formatarComoData(final String celulaInicial) {
		aplicarEstiloNaColuna(celulaInicial, formatos.data());
		return this;
	}

	private void aplicarEstiloNaColuna(final String celulaInicial, final CellStyle estilo) {
		final int[] indices = PosicaoConverter.converterPosicao(celulaInicial);
		FormatosDeCelula.aplicarNaColuna(sheetAtual(), indices[0], indices[1], estilo);
	}

	// ==================== FORMATOS ====================

	/**
	 * Formata a coluna (a partir da célula) como moeda brasileira (R$ 1.234,56).
	 *
	 * @param celulaInicial Primeira célula da coluna a formatar (ex.: "B2").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formatarComoMoeda(final String celulaInicial) {
		planilha.converter().emMoeda(celulaInicial);
		return this;
	}

	/**
	 * Formata a coluna (a partir da célula) no padrão contábil (R$ alinhado).
	 *
	 * @param celulaInicial Primeira célula da coluna a formatar.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formatarComoContabil(final String celulaInicial) {
		planilha.converter().emContabil(celulaInicial);
		return this;
	}

	/**
	 * Converte a coluna (a partir da célula) para número.
	 *
	 * @param celulaInicial Primeira célula da coluna a converter.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formatarComoNumero(final String celulaInicial) {
		planilha.converter().emNumero(celulaInicial);
		return this;
	}

	/**
	 * Converte a coluna (a partir da célula) para texto puro.
	 *
	 * @param celulaInicial Primeira célula da coluna a converter.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formatarComoTexto(final String celulaInicial) {
		planilha.converter().emTexto(celulaInicial);
		return this;
	}

	/**
	 * Formata a coluna (a partir da célula) como porcentagem (ex.: 0,15 vira 15%).
	 *
	 * @param celulaInicial Primeira célula da coluna a formatar.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha formatarComoPorcentagem(final String celulaInicial) {
		aplicarEstiloNaColuna(celulaInicial, formatos.porcentagem());
		return this;
	}

	// ==================== FORMATAÇÃO CONDICIONAL ====================

	/**
	 * Pinta de fundo as células de um intervalo cujo valor é maior que o
	 * informado (ex.: destacar vendas acima de uma meta).
	 *
	 * @param intervalo Intervalo a observar (ex.: "B2:B20").
	 * @param valor     Valor de referência.
	 * @param cor       Cor de fundo aplicada quando a condição é verdadeira.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha realcarSeMaiorQue(final String intervalo, final double valor, final CorEnum cor) {
		FormatacaoCondicionalHelper.realcarSeMaiorQue(xssf(), regioesDe(intervalo), valor, cor.getRed(), cor.getGreen(),
				cor.getBlue());
		return this;
	}

	/**
	 * Pinta de fundo as células de um intervalo cujo valor é menor que o
	 * informado.
	 *
	 * @param intervalo Intervalo a observar (ex.: "B2:B20").
	 * @param valor     Valor de referência.
	 * @param cor       Cor de fundo aplicada quando a condição é verdadeira.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha realcarSeMenorQue(final String intervalo, final double valor, final CorEnum cor) {
		FormatacaoCondicionalHelper.realcarSeMenorQue(xssf(), regioesDe(intervalo), valor, cor.getRed(), cor.getGreen(),
				cor.getBlue());
		return this;
	}

	/**
	 * Pinta de fundo as células de um intervalo cujo valor está entre dois
	 * limites (inclusive).
	 *
	 * @param intervalo Intervalo a observar (ex.: "B2:B20").
	 * @param minimo    Limite inferior.
	 * @param maximo    Limite superior.
	 * @param cor       Cor de fundo aplicada quando a condição é verdadeira.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha realcarSeEntre(final String intervalo, final double minimo, final double maximo,
			final CorEnum cor) {
		FormatacaoCondicionalHelper.realcarSeEntre(xssf(), regioesDe(intervalo), minimo, maximo, cor.getRed(),
				cor.getGreen(), cor.getBlue());
		return this;
	}

	/**
	 * Pinta de fundo as células de um intervalo cujo valor é igual ao informado
	 * (aceita número ou texto).
	 *
	 * @param intervalo Intervalo a observar (ex.: "B2:B20").
	 * @param valor     Valor de referência (número ou texto, ex.: "Atrasado").
	 * @param cor       Cor de fundo aplicada quando a condição é verdadeira.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha realcarSeIgual(final String intervalo, final Object valor, final CorEnum cor) {
		FormatacaoCondicionalHelper.realcarSeIgual(xssf(), regioesDe(intervalo), valor, cor.getRed(), cor.getGreen(),
				cor.getBlue());
		return this;
	}

	/**
	 * Aplica uma escala de 3 cores no estilo "semáforo" (vermelho para os
	 * menores valores, amarelo para os medianos, verde para os maiores).
	 *
	 * @param intervalo Intervalo a colorir (ex.: "B2:B20").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha escalaDeCores(final String intervalo) {
		FormatacaoCondicionalHelper.aplicarEscalaDeCores(xssf(), regioesDe(intervalo));
		return this;
	}

	// ==================== VALIDAÇÃO DE DADOS (LISTA SUSPENSA / LIMITES) ====================

	/**
	 * Cria uma lista suspensa (menu de opções) em um intervalo de células, com
	 * opções fixas. Útil para formulários (ex.: "Status": Pendente/Pago/Atrasado).
	 *
	 * @param intervalo Célula(s) que receberão a lista suspensa (ex.: "C2:C50").
	 * @param opcoes    As opções que aparecerão no menu (ex.: "Sim", "Não").
	 * @return Esta planilha, para encadear comandos.
	 * @throws DadosInvalidosException se não houver opções, ou se a soma dos
	 *                                  textos ultrapassar 255 caracteres (limite
	 *                                  do Excel para listas com opções fixas —
	 *                                  nesse caso, use
	 *                                  {@link #listaSuspensaDoIntervalo}).
	 */
	public Planilha listaSuspensa(final String intervalo, final String... opcoes) {
		validarOpcoesDaLista(opcoes);
		ListaSuspensaHelper.comOpcoesFixas(xssf(), regiaoDe(intervalo), opcoes);
		return this;
	}

	/**
	 * Cria uma lista suspensa cujas opções vêm de um intervalo de células (ex.:
	 * uma coluna auxiliar com a lista de opções). Sem limite de 255 caracteres.
	 *
	 * @param intervaloDestino  Célula(s) que receberão a lista suspensa.
	 * @param intervaloOpcoes   Intervalo com as opções (ex.: "F2:F5").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha listaSuspensaDoIntervalo(final String intervaloDestino, final String intervaloOpcoes) {
		ListaSuspensaHelper.doIntervalo(xssf(), regiaoDe(intervaloDestino), paraReferenciaAbsoluta(intervaloOpcoes));
		return this;
	}

	/**
	 * Restringe o intervalo a números entre dois limites (inclusive) — quem tentar
	 * digitar um valor fora do limite vê um aviso de erro.
	 *
	 * @param intervalo Célula(s) restringidas (ex.: "B2:B50").
	 * @param minimo    Valor mínimo aceito.
	 * @param maximo    Valor máximo aceito.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha validarNumeroEntre(final String intervalo, final double minimo, final double maximo) {
		ValidacaoDeEntradaHelper.numeroEntre(xssf(), regiaoDe(intervalo), minimo, maximo);
		return this;
	}

	/**
	 * Restringe o intervalo a números **inteiros** entre dois limites (inclusive).
	 *
	 * @param intervalo Célula(s) restringidas (ex.: "B2:B50").
	 * @param minimo    Valor mínimo aceito.
	 * @param maximo    Valor máximo aceito.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha validarInteiroEntre(final String intervalo, final int minimo, final int maximo) {
		ValidacaoDeEntradaHelper.inteiroEntre(xssf(), regiaoDe(intervalo), minimo, maximo);
		return this;
	}

	/**
	 * Restringe o intervalo a datas entre dois limites (inclusive).
	 *
	 * @param intervalo Célula(s) restringidas (ex.: "B2:B50").
	 * @param minimo    Data mínima aceita.
	 * @param maximo    Data máxima aceita.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha validarDataEntre(final String intervalo, final LocalDate minimo, final LocalDate maximo) {
		ValidacaoDeEntradaHelper.dataEntre(xssf(), regiaoDe(intervalo), minimo.toString(), maximo.toString());
		return this;
	}

	private void validarOpcoesDaLista(final String[] opcoes) {
		if (opcoes == null || opcoes.length == 0) {
			throw new DadosInvalidosException("A lista suspensa precisa de pelo menos uma opção");
		}
		final int tamanhoTotal = String.join(",", opcoes).length();
		if (tamanhoTotal > 255) {
			throw new DadosInvalidosException("Lista de opções muito longa para o Excel (limite de 255 "
					+ "caracteres somados); use listaSuspensaDoIntervalo(...) referenciando uma coluna com as opções",
					tamanhoTotal);
		}
	}

	private String paraReferenciaAbsoluta(final String intervalo) {
		if (intervalo.contains("!")) {
			return intervalo; // Já qualificado com aba — mantém como está.
		}
		final String[] partes = intervalo.split(":", 2);
		final StringBuilder referencia = new StringBuilder();
		for (int i = 0; i < partes.length; i++) {
			if (i > 0) {
				referencia.append(':');
			}
			referencia.append(absoluta(partes[i].trim()));
		}
		return referencia.toString();
	}

	private String absoluta(final String celula) {
		final String colunaParte = celula.replaceAll("\\d", "");
		final String linhaParte = celula.replaceAll("\\D", "");
		return "$" + colunaParte + "$" + linhaParte;
	}

	// ==================== GRÁFICOS ====================

	/**
	 * Cria um gráfico de barras verticais a partir de um intervalo de categorias
	 * e um de valores.
	 *
	 * <p>Exemplo:</p>
	 * <pre>{@code
	 * planilha.graficoDeBarras("Vendas por mês", "A2:A5", "B2:B5", "D2");
	 * }</pre>
	 *
	 * @param titulo               Título do gráfico.
	 * @param intervaloCategorias  Intervalo com os nomes das categorias (ex.: "A2:A5").
	 * @param intervaloValores     Intervalo com os valores numéricos (ex.: "B2:B5").
	 * @param celulaSuperiorEsquerda Célula onde o canto superior esquerdo do
	 *                               gráfico será posicionado (ex.: "D2").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha graficoDeBarras(final String titulo, final String intervaloCategorias,
			final String intervaloValores, final String celulaSuperiorEsquerda) {
		final int[] ancora = PosicaoConverter.converterPosicao(celulaSuperiorEsquerda);
		GraficoHelper.criarGraficoDeBarras(xssf(), titulo, regioesDe(intervaloCategorias)[0],
				regioesDe(intervaloValores)[0], ancora[0], ancora[1]);
		return this;
	}

	/**
	 * Cria um gráfico de pizza a partir de um intervalo de categorias (fatias) e
	 * um de valores.
	 *
	 * @param titulo                 Título do gráfico.
	 * @param intervaloCategorias    Intervalo com os nomes das fatias.
	 * @param intervaloValores       Intervalo com os valores numéricos de cada fatia.
	 * @param celulaSuperiorEsquerda Célula onde o canto superior esquerdo do
	 *                               gráfico será posicionado.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha graficoDePizza(final String titulo, final String intervaloCategorias,
			final String intervaloValores, final String celulaSuperiorEsquerda) {
		final int[] ancora = PosicaoConverter.converterPosicao(celulaSuperiorEsquerda);
		GraficoHelper.criarGraficoDePizza(xssf(), titulo, regioesDe(intervaloCategorias)[0],
				regioesDe(intervaloValores)[0], ancora[0], ancora[1]);
		return this;
	}

	/**
	 * Cria um gráfico de linha a partir de um intervalo de categorias e um de
	 * valores.
	 *
	 * @param titulo                 Título do gráfico.
	 * @param intervaloCategorias    Intervalo com os nomes das categorias.
	 * @param intervaloValores       Intervalo com os valores numéricos.
	 * @param celulaSuperiorEsquerda Célula onde o canto superior esquerdo do
	 *                               gráfico será posicionado.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha graficoDeLinha(final String titulo, final String intervaloCategorias,
			final String intervaloValores, final String celulaSuperiorEsquerda) {
		final int[] ancora = PosicaoConverter.converterPosicao(celulaSuperiorEsquerda);
		GraficoHelper.criarGraficoDeLinha(xssf(), titulo, regioesDe(intervaloCategorias)[0],
				regioesDe(intervaloValores)[0], ancora[0], ancora[1]);
		return this;
	}

	// ==================== IMAGENS ====================

	/**
	 * Insere uma imagem (logo, foto, ícone) no tamanho natural do arquivo.
	 *
	 * @param celula         Célula onde o canto superior esquerdo da imagem ficará.
	 * @param caminhoArquivo Caminho do arquivo de imagem ({@code .png} ou
	 *                       {@code .jpg}/{@code .jpeg}).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha inserirImagem(final String celula, final String caminhoArquivo) {
		final int[] ancora = PosicaoConverter.converterPosicao(celula);
		ImagemHelper.inserir(xssf(), caminhoArquivo, ancora[0], ancora[1]);
		return this;
	}

	/**
	 * Insere uma imagem redimensionada por uma escala.
	 *
	 * @param celula         Célula onde o canto superior esquerdo da imagem ficará.
	 * @param caminhoArquivo Caminho do arquivo de imagem ({@code .png} ou
	 *                       {@code .jpg}/{@code .jpeg}).
	 * @param escala         Fator de escala (1.0 = tamanho original, 0.5 = metade,
	 *                       2.0 = dobro).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha inserirImagem(final String celula, final String caminhoArquivo, final double escala) {
		final int[] ancora = PosicaoConverter.converterPosicao(celula);
		ImagemHelper.inserir(xssf(), caminhoArquivo, ancora[0], ancora[1], escala);
		return this;
	}

	// ==================== COLUNAS ====================

	/**
	 * Move uma coluna inteira para a posição de outra.
	 *
	 * @param de   Coluna de origem (ex.: "C").
	 * @param para Coluna de destino (ex.: "F").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha moverColuna(final String de, final String para) {
		planilha.manipularPlanilha().moverColuna(de, para);
		return this;
	}

	/**
	 * Remove uma coluna inteira, deslocando as seguintes para a esquerda.
	 *
	 * @param coluna Coluna a remover (ex.: "I").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha removerColuna(final String coluna) {
		planilha.manipularPlanilha().removerColuna(coluna);
		return this;
	}

	/**
	 * Esvazia uma coluna, mantendo-a no lugar.
	 *
	 * @param coluna Coluna a limpar (ex.: "B").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha limparColuna(final String coluna) {
		planilha.manipularPlanilha().limparColuna(coluna);
		return this;
	}

	/**
	 * Insere uma coluna vazia entre duas colunas adjacentes.
	 *
	 * @param esquerda Coluna à esquerda (ex.: "A").
	 * @param direita  Coluna à direita, imediatamente ao lado (ex.: "B").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha inserirColunaEntre(final String esquerda, final String direita) {
		planilha.manipularPlanilha().inserirColunaVaziaEntre(esquerda, direita);
		return this;
	}

	/**
	 * Copia todo o conteúdo (e a formatação) de uma coluna para outra.
	 *
	 * @param origem  Coluna de origem (ex.: "A").
	 * @param destino Coluna de destino (ex.: "D").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha duplicarColuna(final String origem, final String destino) {
		final int colunaOrigem = PosicaoConverter.converterColuna(origem);
		final int colunaDestino = PosicaoConverter.converterColuna(destino);
		final Sheet sheet = sheetAtual();
		for (int indiceLinha = 0; indiceLinha <= sheet.getLastRowNum(); indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			if (linha == null) {
				continue;
			}
			final Cell celulaOrigem = linha.getCell(colunaOrigem);
			if (celulaOrigem == null) {
				continue;
			}
			CopiadorDeCelulas.copiar(celulaOrigem, obterOuCriarCelula(indiceLinha, colunaDestino));
		}
		return this;
	}

	/**
	 * Copia uma linha inteira (conteúdo e formatação) para outra linha.
	 *
	 * @param numeroLinhaOrigem  Número da linha de origem, começando em 1.
	 * @param numeroLinhaDestino Número da linha de destino, começando em 1.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha duplicarLinha(final int numeroLinhaOrigem, final int numeroLinhaDestino) {
		final Sheet sheet = sheetAtual();
		final Row linhaOrigem = sheet.getRow(numeroLinhaOrigem - 1);
		if (linhaOrigem == null) {
			return this;
		}
		Row linhaDestino = sheet.getRow(numeroLinhaDestino - 1);
		if (linhaDestino == null) {
			linhaDestino = sheet.createRow(numeroLinhaDestino - 1);
		}
		for (final Cell celulaOrigem : linhaOrigem) {
			final Cell celulaDestino = linhaDestino.getCell(celulaOrigem.getColumnIndex(),
					Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			CopiadorDeCelulas.copiar(celulaOrigem, celulaDestino);
		}
		return this;
	}

	// ==================== ESTILOS (um comando cada) ====================

	/**
	 * Deixa em negrito uma célula ou intervalo (ex.: "A1" ou "A1:C1").
	 *
	 * @param posicao Célula ou intervalo.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha negrito(final String posicao) {
		estiloDe(posicao).aplicarNegrito();
		return this;
	}

	/**
	 * Deixa em itálico uma célula ou intervalo.
	 *
	 * @param posicao Célula ou intervalo.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha italico(final String posicao) {
		estiloDe(posicao).aplicarItalico();
		return this;
	}

	/**
	 * Muda a cor do texto de uma célula ou intervalo.
	 *
	 * @param posicao Célula ou intervalo.
	 * @param cor     Cor do texto.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha corDoTexto(final String posicao, final CorEnum cor) {
		estiloDe(posicao).corFonte(cor);
		return this;
	}

	/**
	 * Muda a cor de fundo de uma célula ou intervalo.
	 *
	 * @param posicao Célula ou intervalo.
	 * @param cor     Cor de fundo.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha corDeFundo(final String posicao, final CorEnum cor) {
		estiloDe(posicao).corDeFundo(cor);
		return this;
	}

	/**
	 * Centraliza o conteúdo de uma célula ou intervalo.
	 *
	 * @param posicao Célula ou intervalo.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha centralizar(final String posicao) {
		estiloComoIntervalo(posicao).centralizarTudo();
		return this;
	}

	/**
	 * Troca a fonte (tipo de letra) de uma célula ou intervalo.
	 *
	 * @param posicao   Célula ou intervalo.
	 * @param nomeFonte Nome da fonte (ex.: "Arial", "Calibri").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha fonte(final String posicao, final String nomeFonte) {
		estiloDe(posicao).fonte(nomeFonte);
		return this;
	}

	/**
	 * Muda o tamanho da fonte de uma célula ou intervalo.
	 *
	 * @param posicao  Célula ou intervalo.
	 * @param tamanho  Tamanho da fonte em pontos (ex.: 14).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha tamanhoDaFonte(final String posicao, final int tamanho) {
		estiloDe(posicao).fonteTamanho(tamanho);
		return this;
	}

	/**
	 * Desenha bordas finas em todas as células de um intervalo.
	 *
	 * @param intervalo Intervalo (ex.: "A1:C10").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha bordas(final String intervalo) {
		estiloComoIntervalo(intervalo).aplicarTodasAsBordas();
		return this;
	}

	/**
	 * Mescla (junta) as células de um intervalo em uma só.
	 *
	 * @param intervalo Intervalo (ex.: "A1:C1").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha mesclar(final String intervalo) {
		final String[] partes = separarIntervalo(intervalo);
		planilha.selecionar().intervalo(partes[0], partes[1]).mesclarCelulas();
		return this;
	}

	/**
	 * Contorna toda a área usada da planilha com bordas (externas espessas,
	 * internas finas).
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha contornarTudo() {
		planilha.todasAsBordasEmTudo();
		return this;
	}

	/**
	 * Remove as linhas de grade (o "quadriculado" cinza) da aba atual.
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha removerLinhasDeGrade() {
		sheetAtual().setDisplayGridlines(false);
		return this;
	}

	/**
	 * Ajusta a largura de todas as colunas usadas ao conteúdo.
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha ajustarColunas() {
		final Sheet sheet = sheetAtual();
		final int ultimaColuna = ultimaColunaUsada(sheet);
		for (int coluna = 0; coluna <= ultimaColuna; coluna++) {
			sheet.autoSizeColumn(coluna);
		}
		return this;
	}

	/**
	 * Congela a primeira linha (cabeçalho fica fixo ao rolar a planilha).
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha congelarPrimeiraLinha() {
		sheetAtual().createFreezePane(0, 1);
		return this;
	}

	/**
	 * Congela um número de linhas e colunas (ficam fixas ao rolar).
	 *
	 * @param linhas  Quantas linhas de cima manter fixas (ex.: 1).
	 * @param colunas Quantas colunas da esquerda manter fixas (ex.: 0).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha congelar(final int linhas, final int colunas) {
		sheetAtual().createFreezePane(colunas, linhas);
		return this;
	}

	/**
	 * Define a largura de uma coluna (em número aproximado de caracteres, 0 a 255).
	 *
	 * @param coluna  Coluna a ajustar (ex.: "A").
	 * @param largura Largura aproximada em caracteres.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha larguraColuna(final String coluna, final int largura) {
		final int larguraSegura = Math.max(0, Math.min(largura, 255));
		sheetAtual().setColumnWidth(PosicaoConverter.converterColuna(coluna), larguraSegura * 256);
		return this;
	}

	/**
	 * Define a altura de uma linha (em pontos).
	 *
	 * @param numeroLinha Número da linha, começando em 1.
	 * @param altura      Altura em pontos (ex.: 30).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha alturaLinha(final int numeroLinha, final int altura) {
		final Sheet sheet = sheetAtual();
		Row linha = sheet.getRow(numeroLinha - 1);
		if (linha == null) {
			linha = sheet.createRow(numeroLinha - 1);
		}
		linha.setHeightInPoints(altura);
		return this;
	}

	/**
	 * Ativa os filtros (setinhas) na linha de cabeçalho.
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha filtrosNoCabecalho() {
		planilha.inserirFiltros();
		return this;
	}

	// ==================== ORDENAR ====================

	/**
	 * Ordena as linhas em ordem crescente (A→Z, menor→maior) pelo valor de uma
	 * coluna, mantendo a primeira linha como cabeçalho.
	 *
	 * @param coluna Coluna usada como critério (ex.: "B").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha ordenarPorCrescente(final String coluna) {
		return ordenarPor(coluna, true, 2);
	}

	/**
	 * Ordena as linhas em ordem decrescente (Z→A, maior→menor) pelo valor de uma
	 * coluna, mantendo a primeira linha como cabeçalho.
	 *
	 * @param coluna Coluna usada como critério (ex.: "B").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha ordenarPorDecrescente(final String coluna) {
		return ordenarPor(coluna, false, 2);
	}

	/**
	 * Ordena as linhas por uma coluna, com controle da primeira linha de dados
	 * (use 1 se a planilha não tiver cabeçalho).
	 *
	 * @param coluna       Coluna usada como critério (ex.: "B").
	 * @param crescente    {@code true} para A→Z / menor→maior.
	 * @param linhaInicial Número (começando em 1) da primeira linha de dados.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha ordenarPor(final String coluna, final boolean crescente, final int linhaInicial) {
		OrdenadorDeLinhas.ordenar(sheetAtual(), PosicaoConverter.converterColuna(coluna), crescente, linhaInicial - 1);
		return this;
	}

	// ==================== BUSCA / FILTRO DE LINHAS ====================

	/**
	 * Procura as linhas em que uma coluna tem determinado valor.
	 *
	 * @param coluna Coluna a pesquisar (ex.: "B").
	 * @param valor  Valor procurado (ex.: "SP").
	 * @return Números das linhas encontradas (começando em 1).
	 */
	public List<Integer> buscarLinhas(final String coluna, final String valor) {
		final int indiceColuna = PosicaoConverter.converterColuna(coluna);
		final List<Integer> numerosDeLinha = new ArrayList<>();
		for (final int indice : FiltroDeLinhas.encontrar(sheetAtual(), indiceColuna, valor)) {
			numerosDeLinha.add(indice + 1);
		}
		return numerosDeLinha;
	}

	/**
	 * Conta quantas linhas têm determinado valor em uma coluna.
	 *
	 * @param coluna Coluna a pesquisar (ex.: "B").
	 * @param valor  Valor procurado.
	 * @return Quantidade de linhas encontradas.
	 */
	public int contarLinhasOnde(final String coluna, final String valor) {
		return buscarLinhas(coluna, valor).size();
	}

	/**
	 * Copia para outra aba todas as linhas em que a coluna tem o valor informado.
	 * A aba de destino é criada se não existir.
	 *
	 * @param coluna      Coluna a pesquisar (ex.: "B").
	 * @param valor       Valor procurado (ex.: "SP").
	 * @param abaDestino  Aba onde as linhas serão copiadas.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha copiarLinhasParaAba(final String coluna, final String valor, final String abaDestino) {
		copiarLinhasEncontradas(coluna, valor, abaDestino);
		return this;
	}

	/**
	 * Move para outra aba (copia e depois remove) todas as linhas em que a coluna
	 * tem o valor informado.
	 *
	 * @param coluna      Coluna a pesquisar.
	 * @param valor       Valor procurado.
	 * @param abaDestino  Aba de destino.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha moverLinhasParaAba(final String coluna, final String valor, final String abaDestino) {
		copiarLinhasEncontradas(coluna, valor, abaDestino);
		return removerLinhasOnde(coluna, valor);
	}

	/**
	 * Remove todas as linhas em que a coluna tem o valor informado, subindo as
	 * linhas de baixo.
	 *
	 * @param coluna Coluna a pesquisar.
	 * @param valor  Valor procurado.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha removerLinhasOnde(final String coluna, final String valor) {
		final int indiceColuna = PosicaoConverter.converterColuna(coluna);
		final Sheet sheet = sheetAtual();
		final List<Integer> indices = FiltroDeLinhas.encontrar(sheet, indiceColuna, valor);
		for (int posicao = indices.size() - 1; posicao >= 0; posicao--) {
			removerLinhaComDeslocamento(sheet, indices.get(posicao));
		}
		return this;
	}

	// ==================== COMENTÁRIOS ====================

	/**
	 * Adiciona um comentário (nota) a uma célula — o balãozinho que aparece ao
	 * passar o mouse em cima no Excel. Útil para explicar uma fórmula ou dar uma
	 * instrução de preenchimento.
	 *
	 * @param celula Célula que receberá o comentário (ex.: "B2").
	 * @param texto  Texto do comentário.
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha comentario(final String celula, final String texto) {
		final int[] indices = PosicaoConverter.converterPosicao(celula);
		ComentarioHelper.adicionar(xssf(), indices[1], indices[0], texto);
		return this;
	}

	// ==================== IMPRESSÃO ====================

	/**
	 * Define a orientação de impressão como paisagem (deitada).
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha orientacaoPaisagem() {
		sheetAtual().getPrintSetup().setLandscape(true);
		return this;
	}

	/**
	 * Define a orientação de impressão como retrato (em pé) — padrão do Excel.
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha orientacaoRetrato() {
		sheetAtual().getPrintSetup().setLandscape(false);
		return this;
	}

	/**
	 * Define a área que será impressa (o resto da planilha não sai na impressão).
	 *
	 * @param intervalo Intervalo a imprimir (ex.: "A1:F30").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha areaDeImpressao(final String intervalo) {
		final Workbook workbook = planilha.obterWorkbook();
		workbook.setPrintArea(workbook.getSheetIndex(abaAtual), paraReferenciaAbsoluta(intervalo));
		return this;
	}

	/**
	 * Ajusta a impressão para caber em um número máximo de páginas de largura e
	 * altura (reduz a escala automaticamente, como o "Ajustar planilha em 1
	 * página" do Excel).
	 *
	 * @param larguraPaginas Número máximo de páginas na largura (ex.: 1).
	 * @param alturaPaginas  Número máximo de páginas na altura (ex.: 1).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha ajustarImpressaoEmPaginas(final int larguraPaginas, final int alturaPaginas) {
		final Sheet sheet = sheetAtual();
		sheet.setFitToPage(true);
		final PrintSetup configuracao = sheet.getPrintSetup();
		configuracao.setFitWidth((short) larguraPaginas);
		configuracao.setFitHeight((short) alturaPaginas);
		return this;
	}

	// ==================== PROTEÇÃO ====================

	/**
	 * Protege a aba atual: impede editar células travadas (todas, por padrão).
	 * Use {@link #desbloquearCelulas(String)} <strong>antes</strong> desta
	 * chamada para manter algumas células editáveis (ex.: campos de um
	 * formulário).
	 *
	 * @param senha Senha para desproteger depois (pode ser {@code null} ou vazia
	 *              — a planilha fica protegida sem exigir senha para editar via
	 *              código, apenas a interface do Excel respeita a proteção).
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha protegerPlanilha(final String senha) {
		sheetAtual().protectSheet(senha);
		return this;
	}

	/**
	 * Destrava um intervalo de células para que continue editável mesmo depois
	 * de {@link #protegerPlanilha(String)}. Chame antes de proteger.
	 *
	 * @param intervalo Intervalo a destravar (ex.: "B2:B10").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha desbloquearCelulas(final String intervalo) {
		ProtecaoHelper.desbloquearIntervalo(sheetAtual(), regioesDe(intervalo)[0]);
		return this;
	}

	// ==================== SALVAR / FECHAR ====================

	/**
	 * Salva a planilha no caminho informado.
	 *
	 * @param caminhoArquivo Caminho completo do arquivo (ex.: "C:/tmp/vendas.xlsx").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha salvar(final String caminhoArquivo) {
		planilha.salvar(caminhoArquivo);
		return this;
	}

	/**
	 * Salva a planilha em uma pasta, com o nome de arquivo informado.
	 *
	 * @param pasta        Pasta de destino (ex.: "C:/tmp").
	 * @param nomeArquivo  Nome do arquivo (ex.: "vendas.xlsx").
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha salvarNaPasta(final String pasta, final String nomeArquivo) {
		final String separador = pasta.endsWith("/") || pasta.endsWith("\\") ? "" : "/";
		return salvar(pasta + separador + nomeArquivo);
	}

	/**
	 * Fecha a planilha e libera os recursos. Chamado automaticamente ao usar
	 * try-with-resources.
	 */
	public void fechar() {
		try {
			planilha.close();
		} catch (Exception e) {
			throw new IllegalStateException("Erro ao fechar a planilha: " + e.getMessage(), e);
		}
	}

	/**
	 * Fecha a planilha (via try-with-resources). Diferente de
	 * {@link AutoCloseable#close()}, <strong>não obriga</strong> tratar exceção
	 * verificada — para manter a API simples.
	 */
	@Override
	public void close() {
		fechar();
	}

	// ==================== AVANÇADO ====================

	/**
	 * Devolve a célula ou intervalo como um {@link EstiloCelula}, para quem quiser
	 * encadear vários estilos de uma vez.
	 *
	 * @param posicaoOuIntervalo Célula (ex.: "A1") ou intervalo (ex.: "A1:C1").
	 * @return O construtor de estilos posicionado.
	 */
	public EstiloCelula estilo(final String posicaoOuIntervalo) {
		return estiloDe(posicaoOuIntervalo);
	}

	/**
	 * Dá acesso à API fluente completa ({@link IPlanilha}) para recursos avançados.
	 *
	 * @return A planilha subjacente.
	 */
	public IPlanilha avancado() {
		return planilha;
	}

	/**
	 * Dá acesso ao {@link Workbook} do Apache POI para recursos que a facade não
	 * cobre.
	 *
	 * @return O workbook subjacente.
	 */
	public Workbook workbook() {
		return planilha.obterWorkbook();
	}

	// ==================== INTERNOS ====================

	private EstiloCelula estiloDe(final String posicaoOuIntervalo) {
		if (posicaoOuIntervalo.contains(":")) {
			final String[] partes = separarIntervalo(posicaoOuIntervalo);
			planilha.selecionar().intervalo(partes[0], partes[1]);
		} else {
			planilha.selecionar().celula(posicaoOuIntervalo);
		}
		return planilha.aplicarEstilos();
	}

	private EstiloCelula estiloComoIntervalo(final String posicaoOuIntervalo) {
		final String[] partes = separarIntervalo(posicaoOuIntervalo);
		planilha.selecionar().intervalo(partes[0], partes[1]);
		return planilha.aplicarEstilos();
	}

	private String[] separarIntervalo(final String posicaoOuIntervalo) {
		if (posicaoOuIntervalo.contains(":")) {
			final String[] partes = posicaoOuIntervalo.split(":", 2);
			return new String[] { partes[0].trim(), partes[1].trim() };
		}
		return new String[] { posicaoOuIntervalo.trim(), posicaoOuIntervalo.trim() };
	}

	private Sheet sheetAtual() {
		return planilha.obterWorkbook().getSheet(abaAtual);
	}

	private XSSFSheet xssf() {
		return (XSSFSheet) sheetAtual();
	}

	private CellRangeAddress[] regioesDe(final String intervalo) {
		final String[] partes = separarIntervalo(intervalo);
		final int[] inicio = PosicaoConverter.converterPosicao(partes[0]);
		final int[] fim = PosicaoConverter.converterPosicao(partes[1]);
		return new CellRangeAddress[] { new CellRangeAddress(inicio[1], fim[1], inicio[0], fim[0]) };
	}

	private CellRangeAddressList regiaoDe(final String intervalo) {
		final String[] partes = separarIntervalo(intervalo);
		final int[] inicio = PosicaoConverter.converterPosicao(partes[0]);
		final int[] fim = PosicaoConverter.converterPosicao(partes[1]);
		return new CellRangeAddressList(inicio[1], fim[1], inicio[0], fim[0]);
	}

	private void copiarLinhasEncontradas(final String coluna, final String valor, final String abaDestino) {
		final int indiceColuna = PosicaoConverter.converterColuna(coluna);
		final Sheet origem = sheetAtual();
		final List<Integer> indices = FiltroDeLinhas.encontrar(origem, indiceColuna, valor);
		final Sheet destino = obterOuCriarAba(abaDestino);
		for (final int indiceOrigem : indices) {
			final Row linhaOrigem = origem.getRow(indiceOrigem);
			final int proximaLinha = destino.getPhysicalNumberOfRows() == 0 ? 0 : destino.getLastRowNum() + 1;
			final Row linhaDestino = destino.createRow(proximaLinha);
			for (final Cell celulaOrigem : linhaOrigem) {
				CopiadorDeCelulas.copiar(celulaOrigem, linhaDestino.createCell(celulaOrigem.getColumnIndex()));
			}
		}
	}

	private Sheet obterOuCriarAba(final String nome) {
		final Workbook workbook = planilha.obterWorkbook();
		final Sheet existente = workbook.getSheet(nome);
		return existente != null ? existente : workbook.createSheet(nome);
	}

	private void removerLinhaComDeslocamento(final Sheet sheet, final int indiceLinha) {
		final Row linha = sheet.getRow(indiceLinha);
		if (linha != null) {
			sheet.removeRow(linha);
		}
		if (indiceLinha < sheet.getLastRowNum()) {
			sheet.shiftRows(indiceLinha + 1, sheet.getLastRowNum(), -1);
		}
	}

	private Cell obterOuCriarCelula(final int indiceLinha, final int indiceColuna) {
		final Sheet sheet = sheetAtual();
		Row linha = sheet.getRow(indiceLinha);
		if (linha == null) {
			linha = sheet.createRow(indiceLinha);
		}
		return linha.getCell(indiceColuna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	}

	private int ultimaColunaUsada(final Sheet sheet) {
		int ultimaColuna = 0;
		for (int indiceLinha = 0; indiceLinha <= sheet.getLastRowNum(); indiceLinha++) {
			final Row linha = sheet.getRow(indiceLinha);
			if (linha != null && linha.getLastCellNum() - 1 > ultimaColuna) {
				ultimaColuna = linha.getLastCellNum() - 1;
			}
		}
		return ultimaColuna;
	}
}
