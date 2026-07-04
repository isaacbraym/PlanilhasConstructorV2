package com.abnote.planilhas.interfaces;

/**
 * Interface para operações de conversão de formato de células.
 * 
 * <p>Agrupa todas as operações de conversão, tornando a API mais organizada
 * e descobrível através de autocomplete.</p>
 * 
 * <p>Exemplo de uso:</p>
 * <pre>{@code
 * planilha.converter().emContabil("A1");
 * planilha.converter().emNumero("B2");
 * planilha.converter().emTexto("C3");
 * }</pre>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public interface IConversao {
    
    /**
     * Converte os valores de uma coluna para numérico.
     * 
     * <p>Tenta converter valores do tipo String para Double.
     * Valores que não puderem ser convertidos são ignorados e
     * uma mensagem é registrada no log.</p>
     *
     * @param posicaoInicial A posição inicial da coluna (ex: "J3")
     * @return A instância de IPlanilha para encadeamento fluente
     * @throws IllegalArgumentException se posicaoInicial for nula ou vazia
     */
    IPlanilha emNumero(String posicaoInicial);
    
    /**
     * Converte os valores de uma coluna para formato contábil (R$ #,##0.00).
     * 
     * <p>Aplica formatação contábil brasileira com separador de milhares
     * e duas casas decimais. Valores não numéricos são convertidos primeiro.</p>
     *
     * @param posicaoInicial A posição inicial da coluna (ex: "J3")
     * @return A instância de IPlanilha para encadeamento fluente
     * @throws IllegalArgumentException se posicaoInicial for nula ou vazia
     */
    IPlanilha emContabil(String posicaoInicial);
    
    /**
     * Converte os valores de uma coluna para texto.
     * 
     * <p>Converte valores numéricos, datas e outros tipos para
     * formato de texto (String). Útil para preservar zeros à esquerda
     * ou formatar como texto puro.</p>
     *
     * @param posicaoInicial A posição inicial da coluna (ex: "J3")
     * @return A instância de IPlanilha para encadeamento fluente
     * @throws com.abnote.planilhas.exceptions.PosicaoInvalidaException se posicaoInicial for nula ou vazia
     */
    IPlanilha emTexto(String posicaoInicial);
    
    /**
     * Converte os valores de uma coluna para formato de moeda (R$).
     * 
     * <p>Similar a emContabil(), mas com símbolo de moeda (R$) visível.</p>
     *
     * @param posicaoInicial A posição inicial da coluna (ex: "J3")
     * @return A instância de IPlanilha para encadeamento fluente
     * @throws com.abnote.planilhas.exceptions.PosicaoInvalidaException se posicaoInicial for nula ou vazia
     */
    IPlanilha emMoeda(String posicaoInicial);
}