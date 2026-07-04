package com.abnote.planilhas.impl;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IEstilos;
import com.abnote.planilhas.utils.PositionManager;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class StyleManager implements IEstilos {

	private final Workbook workbook;
	private final Sheet sheet;
	private final PositionManager positionManager;
	private final DataManipulator dataManipulator;

	public StyleManager(Workbook workbook, Sheet sheet, PositionManager positionManager,
			DataManipulator dataManipulator) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.positionManager = positionManager;
		this.dataManipulator = dataManipulator;
	}

	public EstiloCelula aplicarEstilos() {
		EstiloCelula estilo;
		if (positionManager.isTodaPlanilhaDefinida()) {
			estilo = new EstiloCelula(workbook, sheet);
		} else if (positionManager.isIntervaloDefinida()) {
			estilo = new EstiloCelula(workbook, sheet, positionManager.getPosicaoInicialLinha(),
					positionManager.getPosicaoInicialColuna(), positionManager.getPosicaoFinalLinha(),
					positionManager.getPosicaoFinalColuna());
		} else if (dataManipulator.getUltimoIndiceDeLinhaInserido() != -1) {
			// Estiliza a linha inteira recém-inserida (comportamento após inserir dados).
			estilo = new EstiloCelula(workbook, sheet, dataManipulator.getUltimoIndiceDeLinhaInserido(), -1);
		} else if (positionManager.isPosicaoDefinida()) {
			// Célula única selecionada por selecionar().celula(...) sem inserção prévia.
			estilo = new EstiloCelula(workbook, sheet, positionManager.getPosicaoInicialLinha(),
					positionManager.getPosicaoInicialColuna());
		} else {
			estilo = new EstiloCelula(workbook, sheet, -1, -1);
		}
		// Reset opcional do positionManager para evitar efeitos colaterais
		positionManager.resetarPosicao();
		return estilo;
	}

	@Override
	public EstiloCelula centralizarTudo() {
		return aplicarEstilos().centralizarTudo();
	}

	@Override
	public EstiloCelula redimensionarColunas() {
		return aplicarEstilos().redimensionarColuna();
	}

	@Override
	public EstiloCelula removerLinhasDeGrade() {
		return aplicarEstilos().removerLinhasDeGrade();
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
}
