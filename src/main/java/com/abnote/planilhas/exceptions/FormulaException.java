package com.abnote.planilhas.exceptions;

/**
 * Exceção lançada quando há erro ao construir ou aplicar fórmulas Excel.
 * 
 * <p><strong>Unchecked exception</strong> - não precisa declarar throws.</p>
 * 
 * <p>Exemplos de situações que geram esta exceção:</p>
 * <ul>
 *   <li>Range inválido em fórmula (ex: "INVALIDO")</li>
 *   <li>Condição sem operador de comparação em IF</li>
 *   <li>Critério inválido em COUNTIF/SUMIF</li>
 *   <li>Tentativa de aplicar fórmula sem definir célula</li>
 * </ul>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class FormulaException extends PlanilhaException {
    
    private static final long serialVersionUID = 1L;
    
    private final String formulaTentada;
    
    public FormulaException(String mensagem) {
        super(mensagem);
        this.formulaTentada = null;
    }
    
    public FormulaException(String mensagem, String formulaTentada) {
        super(String.format("%s. Fórmula tentada: '%s'", mensagem, formulaTentada));
        this.formulaTentada = formulaTentada;
    }
    
    public FormulaException(String mensagem, Throwable causa) {
        super(mensagem, causa);
        this.formulaTentada = null;
    }
    
    public String getFormulaTentada() {
        return formulaTentada;
    }
}