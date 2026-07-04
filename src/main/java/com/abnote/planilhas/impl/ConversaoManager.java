package com.abnote.planilhas.impl;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import com.abnote.planilhas.calculos.Conversores;
import com.abnote.planilhas.exceptions.PosicaoInvalidaException;
import com.abnote.planilhas.interfaces.IConversao;
import com.abnote.planilhas.interfaces.IPlanilha;

/**
 * Implementação do gerenciador de conversões de formato de células.
 * 
 * <p>
 * Delega as operações reais para a classe {@link Conversores}, atuando como um
 * wrapper que mantém a API fluente.
 * </p>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class ConversaoManager implements IConversao {

	private final Sheet sheet;
	private final Workbook workbook;
	private final IPlanilha planilhaRetorno;

	/**
	 * Construtor do gerenciador de conversões.
	 * 
	 * @param sheet           A folha da planilha onde as conversões serão aplicadas
	 * @param workbook        O workbook para criação de estilos
	 * @param planilhaRetorno Instância de IPlanilha para retorno fluente
	 */
	public ConversaoManager(Sheet sheet, Workbook workbook, IPlanilha planilhaRetorno) {
		this.sheet = sheet;
		this.workbook = workbook;
		this.planilhaRetorno = planilhaRetorno;
	}

	@Override
    public IPlanilha emNumero(String posicaoInicial) {
        validarPosicao(posicaoInicial);
        Conversores.converterEmNumero(sheet, posicaoInicial);
        return planilhaRetorno;
    }
    
    @Override
    public IPlanilha emContabil(String posicaoInicial) {
        validarPosicao(posicaoInicial);
        Conversores.converterEmContabil(sheet, posicaoInicial, workbook);
        return planilhaRetorno;
    }

	@Override
	public IPlanilha emTexto(String posicaoInicial) {
		validarPosicao(posicaoInicial);
		Conversores.converterEmTexto(sheet, posicaoInicial, workbook);
		return planilhaRetorno;
	}

	@Override
	public IPlanilha emMoeda(String posicaoInicial) {
		validarPosicao(posicaoInicial);
		Conversores.converterEmMoeda(sheet, posicaoInicial, workbook);
		return planilhaRetorno;
	}

	/**
	 * Valida se a posição informada não é nula ou vazia.
	 */
    private void validarPosicao(String posicao) { // [MODIFIED] SEM throws
        if (posicao == null || posicao.trim().isEmpty()) {
            throw new PosicaoInvalidaException(
                "Posição não pode ser nula ou vazia. Exemplo válido: 'A1', 'B10'",
                posicao
            );
        }
    }
}