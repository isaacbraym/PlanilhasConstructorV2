package com.abnote.planilhas;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.utils.PosicaoConverter;

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
	private String abaAtual;

	private Planilha(final String nomePrimeiraAba) {
		this.planilha = new PlanilhaXlsx();
		this.planilha.criarPlanilha(nomePrimeiraAba);
		this.abaAtual = nomePrimeiraAba;
	}

	/**
	 * Cria uma planilha nova, já com a primeira aba pronta para uso.
	 *
	 * @param nomeDaAba Nome da primeira aba (ex.: "Vendas").
	 * @return A planilha criada.
	 */
	public static Planilha nova(final String nomeDaAba) {
		return new Planilha(nomeDaAba);
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
			copiarCelula(celulaOrigem, obterOuCriarCelula(indiceLinha, colunaDestino));
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
			copiarCelula(celulaOrigem, celulaDestino);
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
	 * Ativa os filtros (setinhas) na linha de cabeçalho.
	 *
	 * @return Esta planilha, para encadear comandos.
	 */
	public Planilha filtrosNoCabecalho() {
		planilha.inserirFiltros();
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

	private void copiarCelula(final Cell origem, final Cell destino) {
		destino.setCellStyle(origem.getCellStyle());
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
		case BLANK:
			destino.setBlank();
			break;
		default:
			break;
		}
	}
}
