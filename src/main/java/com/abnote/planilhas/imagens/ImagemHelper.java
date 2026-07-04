package com.abnote.planilhas.imagens;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.exceptions.DadosInvalidosException;

/**
 * Insere imagens (PNG ou JPEG) em uma planilha, ancoradas em uma célula.
 */
public final class ImagemHelper {

	private ImagemHelper() {
		// Classe utilitária: não deve ser instanciada.
	}

	/**
	 * Insere uma imagem no tamanho natural (original) do arquivo.
	 *
	 * @param sheet          A folha onde a imagem será desenhada.
	 * @param caminhoArquivo Caminho do arquivo de imagem ({@code .png} ou
	 *                       {@code .jpg}/{@code .jpeg}).
	 * @param colunaAncora   Coluna (0-based) onde o canto superior esquerdo ficará.
	 * @param linhaAncora    Linha (0-based) onde o canto superior esquerdo ficará.
	 */
	public static void inserir(final XSSFSheet sheet, final String caminhoArquivo, final int colunaAncora,
			final int linhaAncora) {
		criarImagem(sheet, caminhoArquivo, colunaAncora, linhaAncora).resize();
	}

	/**
	 * Insere uma imagem redimensionada por uma escala (1.0 = tamanho original,
	 * 0.5 = metade, 2.0 = dobro).
	 *
	 * @param sheet          A folha onde a imagem será desenhada.
	 * @param caminhoArquivo Caminho do arquivo de imagem ({@code .png} ou
	 *                       {@code .jpg}/{@code .jpeg}).
	 * @param colunaAncora   Coluna (0-based) onde o canto superior esquerdo ficará.
	 * @param linhaAncora    Linha (0-based) onde o canto superior esquerdo ficará.
	 * @param escala         Fator de escala em relação ao tamanho original.
	 */
	public static void inserir(final XSSFSheet sheet, final String caminhoArquivo, final int colunaAncora,
			final int linhaAncora, final double escala) {
		final XSSFPicture picture = criarImagem(sheet, caminhoArquivo, colunaAncora, linhaAncora);
		// resize(double) do POI escala o tamanho ATUAL da âncora, não o tamanho
		// natural da imagem — e uma âncora recém-criada tem tamanho zero. É preciso
		// primeiro fixar o tamanho natural com resize() para então escalar a partir
		// dele (comportamento verificado empiricamente, não documentado pelo POI).
		picture.resize();
		picture.resize(escala);
	}

	private static XSSFPicture criarImagem(final XSSFSheet sheet, final String caminhoArquivo,
			final int colunaAncora, final int linhaAncora) {
		final byte[] bytes = lerArquivo(caminhoArquivo);
		final int tipo = tipoDaImagem(caminhoArquivo);
		final int indiceImagem = sheet.getWorkbook().addPicture(bytes, tipo);

		final XSSFDrawing desenho = sheet.createDrawingPatriarch();
		final XSSFClientAnchor ancora = desenho.createAnchor(0, 0, 0, 0, colunaAncora, linhaAncora, colunaAncora,
				linhaAncora);
		return desenho.createPicture(ancora, indiceImagem);
	}

	private static byte[] lerArquivo(final String caminhoArquivo) {
		try {
			return Files.readAllBytes(Paths.get(caminhoArquivo));
		} catch (IOException e) {
			throw new ArquivoException("Erro ao ler arquivo de imagem. Verifique se o caminho existe e é acessível",
					caminhoArquivo, e);
		}
	}

	private static int tipoDaImagem(final String caminhoArquivo) {
		final String caminhoMinusculo = caminhoArquivo.toLowerCase();
		if (caminhoMinusculo.endsWith(".png")) {
			return Workbook.PICTURE_TYPE_PNG;
		}
		if (caminhoMinusculo.endsWith(".jpg") || caminhoMinusculo.endsWith(".jpeg")) {
			return Workbook.PICTURE_TYPE_JPEG;
		}
		throw new DadosInvalidosException("Formato de imagem não suportado. Use .png, .jpg ou .jpeg",
				caminhoArquivo);
	}
}
