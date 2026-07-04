package com.abnote.planilhas.exceptions;

/**
 * Exceção lançada quando há erro de I/O ao manipular arquivos.
 * 
 * <p><strong>Unchecked exception</strong> - não precisa declarar throws.</p>
 * 
 * <p>Exemplos de situações que geram esta exceção:</p>
 * <ul>
 *   <li>Erro ao salvar planilha (disco cheio, permissão negada)</li>
 *   <li>Arquivo não encontrado ao importar CSV</li>
 *   <li>Formato de arquivo inválido</li>
 *   <li>Erro ao fechar workbook</li>
 *   <li>Caminho inválido ou inacessível</li>
 * </ul>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class ArquivoException extends PlanilhaException {
    
    private static final long serialVersionUID = 1L;
    
    private final String caminhoArquivo;
    
    public ArquivoException(String mensagem) {
        super(mensagem);
        this.caminhoArquivo = null;
    }
    
    public ArquivoException(String mensagem, String caminhoArquivo) {
        super(String.format("%s. Arquivo: '%s'", mensagem, caminhoArquivo));
        this.caminhoArquivo = caminhoArquivo;
    }
    
    public ArquivoException(String mensagem, String caminhoArquivo, Throwable causa) {
        super(String.format("%s. Arquivo: '%s'", mensagem, caminhoArquivo), causa);
        this.caminhoArquivo = caminhoArquivo;
    }
    
    public String getCaminhoArquivo() {
        return caminhoArquivo;
    }
}