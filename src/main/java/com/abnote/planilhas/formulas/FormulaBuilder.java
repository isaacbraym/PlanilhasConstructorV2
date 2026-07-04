package com.abnote.planilhas.formulas;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.exceptions.FormulaException;
import com.abnote.planilhas.interfaces.IFormulas;
import com.abnote.planilhas.interfaces.IPlanilha;

import java.util.Objects;

/**
 * Implementação do builder de fórmulas Excel.
 * 
 * <p>Esta classe constrói fórmulas Excel de forma fluente e as aplica
 * em células específicas da planilha usando Apache POI.</p>
 * 
 * @author ProjetoPlanilha3
 * @version 2.0.0
 * @since 2.0.0
 */
public class FormulaBuilder implements IFormulas {
    
    private final Workbook workbook;
    private final Sheet sheet;
    private final int linhaAtual;
    private final int colunaAtual;
    private final IPlanilha planilhaRetorno;
    private String formulaAtual;
    
    /**
     * Constrói um novo FormulaBuilder.
     * 
     * @param workbook Workbook do Apache POI
     * @param sheet Sheet atual
     * @param linhaAtual Índice da linha atual (0-based)
     * @param colunaAtual Índice da coluna atual (0-based)
     * @param planilhaRetorno Instância da planilha para retorno fluente
     */
    public FormulaBuilder(Workbook workbook, Sheet sheet, int linhaAtual, 
                         int colunaAtual, IPlanilha planilhaRetorno) {
        this.workbook = Objects.requireNonNull(workbook, "Workbook não pode ser nulo");
        this.sheet = Objects.requireNonNull(sheet, "Sheet não pode ser nulo");
        this.linhaAtual = linhaAtual;
        this.colunaAtual = colunaAtual;
        this.planilhaRetorno = Objects.requireNonNull(planilhaRetorno, "Planilha de retorno não pode ser nula");
        this.formulaAtual = null;
    }
    
    // ========== FÓRMULAS MATEMÁTICAS ==========
    
    @Override
    public IFormulas soma(String range) {
        validarRange(range);
        this.formulaAtual = String.format("SUM(%s)", range);
        return this;
    }
    
    @Override
    public IFormulas media(String range) {
        validarRange(range);
        this.formulaAtual = String.format("AVERAGE(%s)", range);
        return this;
    }
    
    @Override
    public IFormulas contar(String range) {
        validarRange(range);
        this.formulaAtual = String.format("COUNT(%s)", range);
        return this;
    }
    
    @Override
    public IFormulas minimo(String range) {
        validarRange(range);
        this.formulaAtual = String.format("MIN(%s)", range);
        return this;
    }
    
    @Override
    public IFormulas maximo(String range) {
        validarRange(range);
        this.formulaAtual = String.format("MAX(%s)", range);
        return this;
    }
    
    // ========== FÓRMULAS CONDICIONAIS ==========
    
    @Override
    public IFormulas seEntao(String condicao, Object valorVerdadeiro, Object valorFalso) {
        validarCondicao(condicao);
        
        String valorVerdadeiroStr = formatarValorFormula(valorVerdadeiro);
        String valorFalsoStr = formatarValorFormula(valorFalso);
        
        this.formulaAtual = String.format("IF(%s,%s,%s)", 
            condicao, valorVerdadeiroStr, valorFalsoStr);
        return this;
    }
    
    @Override
    public IFormulas contarSe(String range, String criterio) {
        validarRange(range);
        validarCriterio(criterio);
        
        String criterioFormatado = formatarCriterio(criterio);
        this.formulaAtual = String.format("COUNTIF(%s,%s)", range, criterioFormatado);
        return this;
    }
    
    @Override
    public IFormulas somarSe(String range, String criterio) {
        validarRange(range);
        validarCriterio(criterio);
        
        String criterioFormatado = formatarCriterio(criterio);
        this.formulaAtual = String.format("SUMIF(%s,%s)", range, criterioFormatado);
        return this;
    }
    
    // ========== FÓRMULAS ADICIONAIS ==========
    
    @Override
    public IFormulas concatenar(String... ranges) throws FormulaException {
        if (ranges == null || ranges.length == 0) {
            throw new IllegalArgumentException("Pelo menos um range deve ser fornecido para concatenar");
        }
        
        // Valida todos os ranges
        for (String range : ranges) {
            validarRange(range);
        }
        
        // Constrói CONCATENATE(A1,B1,C1,...)
        String rangesJoin = String.join(",", ranges);
        this.formulaAtual = String.format("CONCATENATE(%s)", rangesJoin);
        return this;
    }
    
    @Override
    public IFormulas hoje() {
        this.formulaAtual = "TODAY()";
        return this;
    }
    
    @Override
    public IFormulas agora() {
        this.formulaAtual = "NOW()";
        return this;
    }
    
    // ========== FÓRMULA PERSONALIZADA ==========

    @Override
    public IFormulas personalizada(String formulaExcel) {
        if (formulaExcel == null || formulaExcel.trim().isEmpty()) {
            throw new FormulaException("Fórmula não pode ser nula ou vazia");
        }
        String formulaLimpa = formulaExcel.trim();
        if (formulaLimpa.startsWith("=")) {
            formulaLimpa = formulaLimpa.substring(1).trim();
        }
        this.formulaAtual = formulaLimpa;
        return this;
    }

    // ========== APLICAÇÃO ==========

    @Override
    public IPlanilha aplicar() {
        if (formulaAtual == null || formulaAtual.trim().isEmpty()) {
            throw new IllegalStateException("Nenhuma fórmula foi definida para aplicar");
        }

        Cell celula = obterOuCriarCelula();
        try {
            celula.setCellFormula(formulaAtual);
        } catch (org.apache.poi.ss.formula.FormulaParseException e) {
            throw new FormulaException("Fórmula inválida para o Excel", formulaAtual);
        }

        return planilhaRetorno;
    }
    
    // ========== MÉTODOS PRIVADOS DE VALIDAÇÃO ==========
    
    /**
     * Valida formato básico de range de células.
     * 
     * @param range Range a validar
     * @throws FormulaException se range for inválido
     */
    private void validarRange(String range) throws FormulaException {
        if (range == null || range.trim().isEmpty()) {
            throw new FormulaException("Range não pode ser nulo ou vazio");
        }
        
        String rangeTrimmed = range.trim();
        
        // Regex para validar: A1 ou A1:B10 ou AA100 ou AA100:ZZ999
        if (!rangeTrimmed.matches("^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$")) {
            throw new FormulaException(
                "Formato de range inválido. Use formato 'A1' ou 'A1:B10'",
                rangeTrimmed
            );
        }
    }
    
    /**
     * Valida condição para fórmula IF.
     * 
     * @param condicao Condição a validar
     * @throws FormulaException se condição for inválida
     */
    private void validarCondicao(String condicao) throws FormulaException {
        if (condicao == null || condicao.trim().isEmpty()) {
            throw new FormulaException("Condição não pode ser nula ou vazia");
        }
        
        String condicaoTrimmed = condicao.trim();
        if (!condicaoTrimmed.matches(".*[><=!]+.*")) {
            throw new FormulaException(
                "Condição deve conter operador de comparação (>, <, =, !=, etc)", 
                condicaoTrimmed
            );
        }
    }
    
    /**
     * Valida critério para fórmulas COUNTIF/SUMIF.
     * 
     * @param criterio Critério a validar
     * @throws FormulaException se critério for inválido
     */
    private void validarCriterio(String criterio) throws FormulaException {
        if (criterio == null || criterio.trim().isEmpty()) {
            throw new FormulaException("Critério não pode ser nulo ou vazio");
        }
    }
    
    /**
     * Formata valor para uso em fórmula Excel.
     * 
     * <p>Strings são envolvidas em aspas duplas, números mantidos como estão.</p>
     * 
     * @param valor Valor a formatar
     * @return Valor formatado para fórmula
     */
    private String formatarValorFormula(Object valor) {
        if (valor == null) {
            return "\"\"";
        }
        
        if (valor instanceof Number) {
            return valor.toString();
        }
        
        // String - envolve em aspas
        return String.format("\"%s\"", valor.toString());
    }
    
    /**
     * Formata critério para uso em fórmulas COUNTIF/SUMIF.
     * 
     * <p>Critérios com operadores são envolvidos em aspas.</p>
     * 
     * @param criterio Critério a formatar
     * @return Critério formatado
     */
    private String formatarCriterio(String criterio) {
        String criterioTrimmed = criterio.trim();
        
        // Se já tem aspas, retorna como está
        if (criterioTrimmed.startsWith("\"") && criterioTrimmed.endsWith("\"")) {
            return criterioTrimmed;
        }
        
        // Se começa com operador, envolve em aspas
        if (criterioTrimmed.matches("^[><=!].*")) {
            return String.format("\"%s\"", criterioTrimmed);
        }
        
        // Tenta parsear como número
        try {
            Double.parseDouble(criterioTrimmed);
            return criterioTrimmed; // É número, retorna sem aspas
        } catch (NumberFormatException e) {
            // É texto, envolve em aspas
            return String.format("\"%s\"", criterioTrimmed);
        }
    }
    
    /**
     * Obtém célula existente ou cria nova na posição atual.
     * 
     * @return Célula na posição atual
     */
    private Cell obterOuCriarCelula() {
        // Garante que a linha existe
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(linhaAtual);
        if (row == null) {
            row = sheet.createRow(linhaAtual);
        }
        
        // Garante que a célula existe
        Cell celula = row.getCell(colunaAtual);
        if (celula == null) {
            celula = row.createCell(colunaAtual);
        }
        
        return celula;
    }
}
