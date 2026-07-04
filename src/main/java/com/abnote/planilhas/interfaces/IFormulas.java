package com.abnote.planilhas.interfaces;

/**
 * Interface que define métodos para criação de fórmulas Excel.
 * 
 * <p>Esta interface permite a construção fluente de fórmulas através do método
 * {@link IPlanilhaBasica#formula()}, facilitando a aplicação de cálculos
 * em células da planilha.</p>
 * 
 * <p>Exemplo de uso:</p>
 * <pre>{@code
 * planilha.naCelula("D10")
 *         .formula()
 *         .soma("A1:A9")
 *         .aplicar();
 * }</pre>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public interface IFormulas {
    
    // ========== FÓRMULAS MATEMÁTICAS ==========
    
    /**
     * Cria fórmula SUM (soma) para o range especificado.
     * 
     * @param range Range de células no formato "A1:A10" ou "A1"
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range for inválido
     */
    IFormulas soma(String range);
    
    /**
     * Cria fórmula AVERAGE (média) para o range especificado.
     * 
     * @param range Range de células no formato "A1:A10" ou "A1"
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range for inválido
     */
    IFormulas media(String range);
    
    /**
     * Cria fórmula COUNT (contagem) para o range especificado.
     * 
     * @param range Range de células no formato "A1:A10" ou "A1"
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range for inválido
     */
    IFormulas contar(String range);
    
    /**
     * Cria fórmula MIN (mínimo) para o range especificado.
     * 
     * @param range Range de células no formato "A1:A10" ou "A1"
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range for inválido
     */
    IFormulas minimo(String range);
    
    /**
     * Cria fórmula MAX (máximo) para o range especificado.
     * 
     * @param range Range de células no formato "A1:A10" ou "A1"
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range for inválido
     */
    IFormulas maximo(String range);
    
    // ========== FÓRMULAS CONDICIONAIS ==========
    
    /**
     * Cria fórmula IF (se/então) com condição lógica.
     * 
     * <p>Exemplo:</p>
     * <pre>{@code
     * planilha.naCelula("E2")
     *         .formula()
     *         .seEntao("A2>100", "Alto", "Baixo")
     *         .aplicar();
     * }</pre>
     * 
     * @param condicao Condição lógica (ex: "A2>100", "B5=10")
     * @param valorVerdadeiro Valor retornado se condição for verdadeira
     * @param valorFalso Valor retornado se condição for falsa
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se condição for inválida
     */
    IFormulas seEntao(String condicao, Object valorVerdadeiro, Object valorFalso);
    
    /**
     * Cria fórmula COUNTIF (contar se) com critério.
     * 
     * <p>Exemplo:</p>
     * <pre>{@code
     * planilha.naCelula("F1")
     *         .formula()
     *         .contarSe("A1:A10", ">50")
     *         .aplicar();
     * }</pre>
     * 
     * @param range Range de células no formato "A1:A10"
     * @param criterio Critério de contagem (ex: ">50", "=10", "texto")
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range ou critério forem inválidos
     */
    IFormulas contarSe(String range, String criterio);
    
    /**
     * Cria fórmula SUMIF (somar se) com critério.
     * 
     * <p>Exemplo:</p>
     * <pre>{@code
     * planilha.naCelula("G1")
     *         .formula()
     *         .somarSe("A1:A10", ">100")
     *         .aplicar();
     * }</pre>
     * 
     * @param range Range de células no formato "A1:A10"
     * @param criterio Critério de soma (ex: ">100", "=50")
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se range ou critério forem inválidos
     */
    IFormulas somarSe(String range, String criterio);
    
    // ========== FÓRMULAS ADICIONAIS ==========
    
    /**
     * Cria fórmula CONCATENATE para juntar textos de múltiplos ranges.
     * 
     * <p>Exemplo:</p>
     * <pre>{@code
     * planilha.naCelula("H1")
     *         .formula()
     *         .concatenar("A1", "B1", "C1")
     *         .aplicar();
     * }</pre>
     * 
     * @param ranges Um ou mais ranges/células a concatenar
     * @return Esta instância para encadeamento fluente
     * @throws IllegalArgumentException se algum range for inválido
     */
    IFormulas concatenar(String... ranges);
    
    /**
     * Cria fórmula TODAY (data atual).
     * 
     * <p>Retorna a data atual do sistema sem hora.</p>
     * 
     * @return Esta instância para encadeamento fluente
     */
    IFormulas hoje();
    
    /**
     * Cria fórmula NOW (data e hora atual).
     * 
     * <p>Retorna a data e hora atual do sistema.</p>
     * 
     * @return Esta instância para encadeamento fluente
     */
    IFormulas agora();
    
    // ========== APLICAÇÃO ==========
    
    /**
     * Aplica a fórmula construída na célula atual da planilha.
     * 
     * <p>Este método finaliza a construção da fórmula e a insere
     * na célula previamente definida por {@link IPlanilhaBasica#naCelula(String)}.</p>
     * 
     * @return Interface principal da planilha para continuar encadeamento
     * @throws IllegalStateException se nenhuma fórmula foi definida
     */
    IPlanilha aplicar();
}
