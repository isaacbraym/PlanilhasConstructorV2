package com.abnote.planilhas.exceptions;

/**
 * Exceção base para todas as exceções específicas do ProjetoPlanilha3.
 * 
 * <p>Esta é uma exceção <strong>unchecked</strong> (RuntimeException) que NÃO
 * precisa ser declarada em assinaturas de métodos.</p>
 * 
 * <p>O desenvolvedor pode capturá-la opcionalmente para tratamento específico,
 * mas não é obrigado a fazê-lo.</p>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class PlanilhaException extends RuntimeException { // [MODIFIED] Exception → RuntimeException
    
    private static final long serialVersionUID = 1L;
    
    public PlanilhaException(String mensagem) {
        super(mensagem);
    }
    
    public PlanilhaException(String mensagem, Throwable causa) {
        super(mensagem, causa);
    }
    
    public PlanilhaException(Throwable causa) {
        super(causa);
    }
}