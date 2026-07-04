package com.abnote.planilhas.estilos;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.estilos.AlinhamentoStyle;
import com.abnote.planilhas.estilos.estilos.BackGroundColor;
import com.abnote.planilhas.estilos.estilos.BoldStyle;
import com.abnote.planilhas.estilos.estilos.BorderStyleHelper;
import com.abnote.planilhas.estilos.estilos.CenterStyle;
import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.estilos.estilos.FonteEnum;
import com.abnote.planilhas.estilos.estilos.Fontes;
import com.abnote.planilhas.utils.PosicaoConverter;

/**
 * Classe responsável por aplicar diversos estilos em células, linhas ou
 * intervalos de uma planilha.
 */
public class EstiloCelula {
	private final Workbook workbook;
	private final Sheet sheet;

	private final int rowIndex;
	private final int columnIndex;

	private final int startRowIndex;
	private final int startColumnIndex;
	private final int endRowIndex;
	private final int endColumnIndex;

	private final boolean isRange;

	private final Map<String, org.apache.poi.ss.usermodel.CellStyle> styleCache = new HashMap<>();

	// Instâncias das classes auxiliares
	private final BoldStyle boldStyle;
	private final BorderStyleHelper borderStyleHelper;
	private final CenterStyle centerStyle;
	private final Fontes fontes;
	private final BackGroundColor backGroundColor;
	private final AlinhamentoStyle alinhamentoStyle;

	/**
	 * Construtor para aplicar estilos na planilha inteira.
	 *
	 * @param workbook O Workbook que contém a planilha.
	 * @param sheet    A Sheet onde os estilos serão aplicados.
	 */
	public EstiloCelula(Workbook workbook, Sheet sheet) {
		this(workbook, sheet, -1, -1, 0, 0, sheet.getLastRowNum(), getMaxColumnIndex(sheet));
	}

	/**
	 * Construtor para aplicar estilos em uma célula específica.
	 *
	 * @param workbook    O Workbook que contém a planilha.
	 * @param sheet       A Sheet onde a célula está.
	 * @param rowIndex    Índice da linha (0-based).
	 * @param columnIndex Índice da coluna (0-based).
	 */
	public EstiloCelula(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex) {
		this(workbook, sheet, rowIndex, columnIndex, -1, -1, -1, -1);
	}

	/**
	 * Construtor para aplicar estilos em um intervalo.
	 *
	 * @param workbook         O Workbook que contém a planilha.
	 * @param sheet            A Sheet onde o intervalo está.
	 * @param startRowIndex    Índice inicial da linha (0-based).
	 * @param startColumnIndex Índice inicial da coluna (0-based).
	 * @param endRowIndex      Índice final da linha (0-based).
	 * @param endColumnIndex   Índice final da coluna (0-based).
	 */
	public EstiloCelula(Workbook workbook, Sheet sheet, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex) {
		this(workbook, sheet, startRowIndex, startColumnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex);
	}

	/**
	 * Construtor interno que inicializa todos os campos e as classes auxiliares.
	 */
	private EstiloCelula(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex) {
		this.workbook = workbook;
		this.sheet = sheet;

		this.rowIndex = rowIndex;
		this.columnIndex = columnIndex;

		this.startRowIndex = startRowIndex;
		this.startColumnIndex = startColumnIndex;

		this.endRowIndex = endRowIndex;
		this.endColumnIndex = endColumnIndex;

		this.isRange = (endRowIndex != -1 && endColumnIndex != -1);

		// Inicialização das classes auxiliares com o cache de estilos
		this.boldStyle = new BoldStyle(workbook, sheet, styleCache);
		this.borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		this.centerStyle = new CenterStyle(workbook, sheet, styleCache);
		this.fontes = new Fontes(workbook, sheet, styleCache);
		this.backGroundColor = new BackGroundColor(workbook, sheet, styleCache);
		this.alinhamentoStyle = new AlinhamentoStyle(workbook, sheet, styleCache);
	}

	// Métodos para aplicar estilos

	public EstiloCelula aplicarItalico() {
		fontes.aplicarItalico(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	public EstiloCelula aplicarSublinhado() {
		fontes.aplicarSublinhado(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	public EstiloCelula aplicarTachado() {
		fontes.aplicarTachado(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	public EstiloCelula alinharAEsquerda() {
		alinhamentoStyle.alinharAEsquerda(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula alinharADireita() {
		alinhamentoStyle.alinharADireita(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula quebrarTexto() {
		alinhamentoStyle.quebrarTexto(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	private static int getMaxColumnIndex(Sheet sheet) {
		int maxColIndex = -1;
		for (int rowIdx = 0; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row != null && row.getLastCellNum() > maxColIndex) {
				maxColIndex = row.getLastCellNum();
			}
		}
		return maxColIndex - 1; // Ajuste para índice 0-based
	}

	public EstiloCelula aplicarNegrito() {
		boldStyle.aplicarNegrito(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	public EstiloCelula aplicarTodasAsBordas() {
		borderStyleHelper.aplicarTodasAsBordas(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula aplicarBordasNaCelula(String posicao) {
		borderStyleHelper.aplicarBordasNaCelula(posicao);
		return this;
	}

	public EstiloCelula aplicarBordasEntre(String posicaoInicial, String posicaoFinal) {
		borderStyleHelper.aplicarBordasEntre(posicaoInicial, posicaoFinal);
		return this;
	}

	public EstiloCelula aplicarBordasEspessas(String posicaoInicial, String posicaoFinal) {
		borderStyleHelper.aplicarBordasEspessas(posicaoInicial, posicaoFinal);
		return this;
	}

	public EstiloCelula aplicarBordasEspessasComInternas(String posicaoInicial, String posicaoFinal) {
		borderStyleHelper.aplicarBordasEspessasComInternas(posicaoInicial, posicaoFinal);
		return this;
	}

	public EstiloCelula centralizarTudo() {
		centerStyle.centralizarTudo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula centralizarERedimensionarTudo() {
		centerStyle.centralizarERedimensionarTudo();
		return this;
	}

	public EstiloCelula redimensionarColuna() {
		centerStyle.redimensionarColuna();
		return this;
	}

	public EstiloCelula removerLinhasDeGrade() {
		sheet.setDisplayGridlines(false);
		return this;
	}

	/**
	 * Aplica bordas espessas nas bordas externas e finas nas internas cobrindo toda
	 * a área usada deste estilo (útil para "contornar" a planilha inteira sem
	 * precisar informar posições fixas).
	 *
	 * @return Esta instância para encadeamento.
	 */
	public EstiloCelula contornarTudo() {
		if (startRowIndex < 0 || startColumnIndex < 0 || endRowIndex < 0 || endColumnIndex < 0) {
			return this; // Planilha vazia: nada a contornar.
		}
		String posicaoInicial = PosicaoConverter.converterIndice(startColumnIndex) + (startRowIndex + 1);
		String posicaoFinal = PosicaoConverter.converterIndice(endColumnIndex) + (endRowIndex + 1);
		borderStyleHelper.aplicarBordasEspessasComInternas(posicaoInicial, posicaoFinal);
		return this;
	}

	// Métodos para aplicar estilos de fonte

	public EstiloCelula fonte(String fontName) {
		fontes.aplicarFonte(fontName, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula fonte(FonteEnum fonteEnum) {
		fontes.aplicarFonte(fonteEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula fonteTamanho(int fontSize) {
		fontes.aplicarTamanhoFonte(fontSize, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corFonte(CorEnum corEnum) {
		fontes.aplicarCorFonte(corEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corFonte(int red, int green, int blue) {
		fontes.aplicarCorFonte(red, green, blue, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corFonte(String hexColor) {
		fontes.aplicarCorFonte(hexColor, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corDeFundo(CorEnum corEnum) {
		backGroundColor.aplicarCorDeFundo(corEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corDeFundo(int red, int green, int blue) {
		backGroundColor.aplicarCorDeFundo(red, green, blue, rowIndex, columnIndex, startRowIndex, startColumnIndex,
				endRowIndex, endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corDeFundo(String hexColor) {
		backGroundColor.aplicarCorDeFundo(hexColor, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}
}
