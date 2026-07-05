package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.exceptions.DadosInvalidosException;

/**
 * Testes de inserção de imagens (logo) na facade.
 */
@DisplayName("Planilha — inserir imagem")
class ImagemFacadeTest {

	@TempDir
	Path pasta;

	private String criarImagemPng(String nomeArquivo, int largura, int altura) throws IOException {
		BufferedImage imagem = new BufferedImage(largura, altura, BufferedImage.TYPE_INT_RGB);
		Graphics2D g = imagem.createGraphics();
		g.setColor(Color.RED);
		g.fillRect(0, 0, largura, altura);
		g.dispose();
		File arquivo = pasta.resolve(nomeArquivo).toFile();
		ImageIO.write(imagem, "png", arquivo);
		return arquivo.getAbsolutePath();
	}

	private XSSFPicture unicaImagem(Planilha planilha) {
		Sheet sheet = planilha.workbook().getSheetAt(0);
		XSSFDrawing drawing = (XSSFDrawing) sheet.getDrawingPatriarch();
		assertNotNull(drawing);
		assertEquals(1, drawing.getShapes().size());
		return (XSSFPicture) drawing.getShapes().get(0);
	}

	@Test
	@DisplayName("Deve inserir PNG ancorado na célula informada, no tamanho natural")
	void deveInserirImagemNoTamanhoNatural() throws Exception {
		String caminho = criarImagemPng("logo.png", 120, 60);
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.inserirImagem("B2", caminho);

			List<? extends PictureData> imagens = planilha.workbook().getAllPictures();
			assertEquals(1, imagens.size());

			XSSFPicture picture = unicaImagem(planilha);
			XSSFClientAnchor ancora = picture.getClientAnchor();
			assertEquals(1, ancora.getCol1(), "Coluna B é índice 1");
			assertEquals(1, ancora.getRow1(), "Linha 2 é índice 1");
		}
	}

	@Test
	@DisplayName("Deve preservar imagem inserida ao salvar e reabrir o XLSX")
	void devePreservarImagemEmRoundTrip() throws Exception {
		String caminho = criarImagemPng("logo.png", 120, 60);
		File arquivo = pasta.resolve("saida.xlsx").toFile();

		try (Planilha planilha = Planilha.nova("T")) {
			planilha.inserirImagem("B2", caminho).salvar(arquivo.getAbsolutePath());
		}

		try (FileInputStream entrada = new FileInputStream(arquivo);
				XSSFWorkbook workbook = new XSSFWorkbook(entrada)) {
			assertEquals(1, workbook.getAllPictures().size());
			assertEquals(Workbook.PICTURE_TYPE_PNG, workbook.getAllPictures().get(0).getPictureType());

			XSSFDrawing drawing = (XSSFDrawing) workbook.getSheetAt(0).getDrawingPatriarch();
			assertNotNull(drawing);
			assertEquals(1, drawing.getShapes().size());
		}
	}

	@Test
	@DisplayName("Deve redimensionar a imagem quando uma escala é informada")
	void deveInserirImagemComEscala() throws Exception {
		String caminho = criarImagemPng("logo.png", 100, 100);
		try (Planilha planilhaNatural = Planilha.nova("T")) {
			planilhaNatural.inserirImagem("A1", caminho);
			int larguraNatural = larguraEmEmu(planilhaNatural, unicaImagem(planilhaNatural).getClientAnchor());

			try (Planilha planilhaGrande = Planilha.nova("T")) {
				planilhaGrande.inserirImagem("A1", caminho, 3.0);
				int larguraGrande = larguraEmEmu(planilhaGrande, unicaImagem(planilhaGrande).getClientAnchor());

				assertTrue(larguraGrande > larguraNatural, "Imagem com escala 3x deve ocupar mais espaço");
			}
		}
	}

	/** Largura total (em EMU) do início da âncora até seu fim, assumindo colunas de largura padrão. */
	private int larguraEmEmu(Planilha planilha, XSSFClientAnchor ancora) {
		Sheet sheet = planilha.workbook().getSheetAt(0);
		int larguraColunaEmu = Units.columnWidthToEMU(sheet.getColumnWidth(ancora.getCol1()));
		return (ancora.getCol2() - ancora.getCol1()) * larguraColunaEmu + ancora.getDx2();
	}

	@Test
	@DisplayName("Deve lançar DadosInvalidosException para formato não suportado")
	void deveLancarFormatoNaoSuportado() throws Exception {
		File arquivo = pasta.resolve("logo.bmp").toFile();
		BufferedImage imagem = new BufferedImage(10, 10, BufferedImage.TYPE_INT_RGB);
		ImageIO.write(imagem, "bmp", arquivo);

		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(DadosInvalidosException.class,
					() -> planilha.inserirImagem("A1", arquivo.getAbsolutePath()));
		}
	}

	@Test
	@DisplayName("Deve lançar ArquivoException para arquivo inexistente")
	void deveLancarArquivoInexistente() throws Exception {
		String inexistente = pasta.resolve("nao-existe.png").toString();
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(ArquivoException.class, () -> planilha.inserirImagem("A1", inexistente));
		}
	}

	@Test
	@DisplayName("Deve lançar ArquivoException para caminho obrigatório ou inválido sem criar imagem parcial")
	void deveLancarArquivoExceptionParaCaminhoObrigatorioOuInvalido() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(ArquivoException.class, () -> planilha.inserirImagem("A1", null));
			assertThrows(ArquivoException.class, () -> planilha.inserirImagem("A1", " "));
			assertThrows(ArquivoException.class, () -> planilha.inserirImagem("A1", "\0.png"));

			assertTrue(planilha.workbook().getAllPictures().isEmpty());
			assertNull(planilha.workbook().getSheetAt(0).getDrawingPatriarch());
		}
	}

	@Test
	@DisplayName("Deve recusar escala inválida sem criar imagem parcial")
	void deveRecusarEscalaInvalidaSemCriarImagemParcial() throws Exception {
		String caminho = criarImagemPng("logo.png", 100, 100);
		try (Planilha planilha = Planilha.nova("T")) {
			assertThrows(DadosInvalidosException.class, () -> planilha.inserirImagem("A1", caminho, 0.0));
			assertThrows(DadosInvalidosException.class, () -> planilha.inserirImagem("A1", caminho, -0.5));
			assertThrows(DadosInvalidosException.class, () -> planilha.inserirImagem("A1", caminho, Double.NaN));
			assertThrows(DadosInvalidosException.class,
					() -> planilha.inserirImagem("A1", caminho, Double.POSITIVE_INFINITY));

			assertTrue(planilha.workbook().getAllPictures().isEmpty());
			assertNull(planilha.workbook().getSheetAt(0).getDrawingPatriarch());
		}
	}
}
