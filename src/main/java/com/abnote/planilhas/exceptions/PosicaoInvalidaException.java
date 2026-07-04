package com.abnote.planilhas.exceptions;

/**
 * Exceção lançada quando uma posição de célula é inválida.
 * 
 * <p><strong>Unchecked exception</strong> - não precisa declarar throws.</p>
 * 
 * <p>Exemplos de situações que geram esta exceção:</p>
 * <ul>
 *   <li>Posição fora dos limites do Excel (ex: "XFE1048577")</li>
 *   <li>Formato inválido (ex: "123", "ABC")</li>
 *   <li>Range com sintaxe incorreta (ex: "A1-B10" ao invés de "A1:B10")</li>
 *   <li>Posição nula ou vazia</li>
 * </ul>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class PosicaoInvalidaException extends PlanilhaException {
    
    private static final long serialVersionUID = 1L;
    
    private final String posicaoInvalida;
    
    public PosicaoInvalidaException(String mensagem) {
        super(mensagem);
        this.posicaoInvalida = null;
    }
    
    public PosicaoInvalidaException(String mensagem, String posicaoInvalida) {
        super(String.format("%s. Posição inválida: '%s'", mensagem, posicaoInvalida));
        this.posicaoInvalida = posicaoInvalida;
    }
    
    public String getPosicaoInvalida() {
        return posicaoInvalida;
    }
}