package com.abnote.planilhas.exceptions;

/**
 * Exceção lançada quando há erro ao processar ou validar dados.
 * 
 * <p><strong>Unchecked exception</strong> - não precisa declarar throws.</p>
 * 
 * <p>Exemplos de situações que geram esta exceção:</p>
 * <ul>
 *   <li>Conversão de String para número falhou</li>
 *   <li>Dado ultrapassa limite de caracteres (32.767/célula)</li>
 *   <li>Formato de data inválido</li>
 *   <li>Arquivo CSV com delimitador incorreto</li>
 *   <li>Dados nulos quando não permitidos</li>
 * </ul>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class DadosInvalidosException extends PlanilhaException {
    
    private static final long serialVersionUID = 1L;
    
    private final Object dadoInvalido;
    
    public DadosInvalidosException(String mensagem) {
        super(mensagem);
        this.dadoInvalido = null;
    }
    
    public DadosInvalidosException(String mensagem, Object dadoInvalido) {
        super(String.format("%s. Dado inválido: '%s'", mensagem, dadoInvalido));
        this.dadoInvalido = dadoInvalido;
    }
    
    public DadosInvalidosException(String mensagem, Throwable causa) {
        super(mensagem, causa);
        this.dadoInvalido = null;
    }
    
    public Object getDadoInvalido() {
        return dadoInvalido;
    }
}