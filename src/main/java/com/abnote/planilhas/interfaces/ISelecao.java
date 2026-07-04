package com.abnote.planilhas.interfaces;

/**
 * Interface para operações de seleção de posicionamento na planilha.
 * 
 * <p>Agrupa todas as operações de seleção de células, intervalos e áreas,
 * tornando a API mais semântica: "selecionar célula", "selecionar intervalo".</p>
 * 
 * <p>Exemplo de uso:</p>
 * <pre>{@code
 * planilha.selecionar().celula("A1").inserirDados("Teste");
 * planilha.selecionar().intervalo("A1", "D10").aplicarEstilos().aplicarNegrito();
 * planilha.selecionar().todaPlanilha().centralizarTudo();
 * }</pre>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public interface ISelecao {
    
    /**
     * Seleciona uma célula específica para operações subsequentes.
     * 
     * @param posicao Posição no formato alfanumérico (ex: "B2", "AA10", "XFD1048576")
     * @return A instância de IPlanilha para encadeamento fluente
     * @throws IllegalArgumentException se posicao for nula, vazia ou inválida
     */
    IPlanilha celula(String posicao);
    
    /**
     * Seleciona um intervalo de células para operações subsequentes.
     * 
     * @param posicaoInicial Posição inicial do intervalo (ex: "B2")
     * @param posicaoFinal Posição final do intervalo (ex: "E10")
     * @return A instância de IPlanilha para encadeamento fluente
     * @throws IllegalArgumentException se alguma posição for inválida
     */
    IPlanilha intervalo(String posicaoInicial, String posicaoFinal);
    
    /**
     * Seleciona toda a planilha para operações em massa.
     * 
     * <p>Útil para aplicar estilos, bordas ou formatações em toda a área
     * de dados da planilha.</p>
     * 
     * @return A instância de IPlanilha para encadeamento fluente
     */
    IPlanilha todaPlanilha();
    
    /**
     * Seleciona a próxima linha disponível após a última linha com dados
     * em uma coluna específica.
     * 
     * <p>Útil para adicionar novos dados ao final de uma lista existente.</p>
     * 
     * @param coluna A coluna de referência (ex: "B", "AA")
     * @return A instância de IManipulacaoDados para inserção de dados
     * @throws IllegalArgumentException se coluna for nula ou vazia
     */
    IManipulacaoDados ultimaLinha(String coluna);
}