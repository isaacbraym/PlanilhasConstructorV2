package com.abnote.planilhas.impl;

import com.abnote.planilhas.exceptions.PosicaoInvalidaException;
import com.abnote.planilhas.interfaces.IManipulacaoDados;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.interfaces.ISelecao;
import com.abnote.planilhas.utils.PositionManager;

/**
 * Implementação do gerenciador de seleções de posicionamento.
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class SelecaoManager implements ISelecao {
    
    private final IPlanilha planilha;
    private final DataManipulator dataManipulator;
    private final PositionManager positionManager;
    
    /**
     * Construtor do gerenciador de seleções.
     * 
     * @param planilha Instância de IPlanilha para retorno fluente
     * @param dataManipulator Manipulador de dados para operações internas
     * @param positionManager Gerenciador de posições
     */
    public SelecaoManager(IPlanilha planilha, DataManipulator dataManipulator, 
                         PositionManager positionManager) {
        this.planilha = planilha;
        this.dataManipulator = dataManipulator;
        this.positionManager = positionManager;
    }
    
    @Override
    public IPlanilha celula(String posicao) {
        validarPosicao(posicao);
        // Chama diretamente o dataManipulator (método interno)
        dataManipulator.naCelula(posicao);
        return planilha;
    }
    
    @Override
    public IPlanilha intervalo(String posicaoInicial, String posicaoFinal) {
        validarPosicao(posicaoInicial);
        validarPosicao(posicaoFinal);
        // Chama diretamente o dataManipulator (método interno)
        dataManipulator.noIntervalo(posicaoInicial, posicaoFinal);
        return planilha;
    }
    
    @Override
    public IPlanilha todaPlanilha() {
        // Chama diretamente o positionManager
        positionManager.emTodaAPlanilha();
        return planilha;
    }
    
    @Override
    public IManipulacaoDados ultimaLinha(String coluna) {
        validarColuna(coluna);
        // Retorna o dataManipulator após posicionar
        return dataManipulator.naUltimaLinha(coluna);
    }
    
    /**
     * Valida se a posição informada não é nula ou vazia.
     */
    private void validarPosicao(String posicao) { // [MODIFIED] SEM throws
        if (posicao == null || posicao.trim().isEmpty()) {
            throw new PosicaoInvalidaException(
                "Posição não pode ser nula ou vazia. Exemplo válido: 'A1', 'B10', 'AA100'",
                posicao
            );
        }
    }
    
    private void validarColuna(String coluna) { // [MODIFIED] SEM throws
        if (coluna == null || coluna.trim().isEmpty()) {
            throw new PosicaoInvalidaException(
                "Coluna não pode ser nula ou vazia. Exemplo válido: 'A', 'B', 'AA'",
                coluna
            );
        }
    }
}