package com.abnote.planilhas.calculos;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;

/**
 * Testes unitários para a classe Calculos.
 * 
 * Exemplos didáticos de como testar operações de cálculo em planilhas.
 */
class CalculosTest {

    private Workbook workbook;
    private Sheet sheet;

    /**
     * Executado ANTES de cada teste.
     * Cria um workbook e sheet limpos para cada teste.
     */
    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Teste");
    }

    /**
     * Executado APÓS cada teste.
     * Fecha o workbook para liberar memória.
     */
    @AfterEach
    void tearDown() throws Exception {
        if (workbook != null) {
            workbook.close();
        }
    }

    /**
     * Teste: Verifica se somarColuna() calcula corretamente.
     * 
     * Cenário:
     * - Inserir 3 valores numéricos em uma coluna
     * - Executar somarColuna()
     * - Verificar se a soma está correta
     */
    @Test
    @DisplayName("Deve somar valores numéricos de uma coluna corretamente")
    void deveSomarColunaCorretamente() {
        // ARRANGE (Preparar dados)
        criarCelula(0, 0, 10.0);   // A1 = 10.0
        criarCelula(1, 0, 20.5);   // A2 = 20.5
        criarCelula(2, 0, 15.3);   // A3 = 15.3

        // ACT (Executar ação)
        Calculos.somarColuna(sheet, "A1");

        // ASSERT (Verificar resultado)
        Row rowSoma = sheet.getRow(3); // Linha 4 (índice 3)
        assertNotNull(rowSoma, "Linha da soma deve existir");

        Cell cellSoma = rowSoma.getCell(0);
        assertNotNull(cellSoma, "Célula da soma deve existir");
        assertEquals(CellType.NUMERIC, cellSoma.getCellType(), "Célula deve ser numérica");

        double somaEsperada = 45.8;
        assertEquals(somaEsperada, cellSoma.getNumericCellValue(), 0.01, 
                "Soma deve ser 10.0 + 20.5 + 15.3 = 45.8");
    }


    /**
     * Teste: Verifica se converterEmNumero() transforma String em número.
     */
    @Test
    @DisplayName("Deve converter String para número")
    void deveConverterStringParaNumero() {
        // ARRANGE
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("123.45");

        // ACT
        Conversores.converterEmNumero(sheet, "A1");

        // ASSERT
        Cell cellConvertida = sheet.getRow(0).getCell(0);
        assertEquals(CellType.NUMERIC, cellConvertida.getCellType());
        assertEquals(123.45, cellConvertida.getNumericCellValue(), 0.01);
    }

    /**
     * Teste: Verifica comportamento com coluna vazia.
     */
    @Test
    @DisplayName("Deve retornar zero ao somar coluna vazia")
    void deveSomarColunaVazia() {
        // ARRANGE (sem dados)

        // ACT
        Calculos.somarColuna(sheet, "A1");

        // ASSERT
        Row rowSoma = sheet.getRow(0);
        if (rowSoma != null) {
            Cell cellSoma = rowSoma.getCell(0);
            if (cellSoma != null) {
                assertEquals(0.0, cellSoma.getNumericCellValue(), 0.01);
            }
        }
        // Se não criou linha/célula, também é comportamento válido
    }

    @Test
    @DisplayName("somarColunaComTexto deve preservar linhas existentes antes do total")
    void deveSomarColunaComTextoDepoisDaUltimaLinhaExistente() {
        criarCelula(0, 1, 10.0);
        criarCelula(1, 1, 20.0);
        Row linhaExistente = sheet.createRow(2);
        linhaExistente.createCell(3).setCellValue("Preservar");

        Calculos.somarColunaComTexto(sheet, "B1", "Total");

        assertEquals("Preservar", sheet.getRow(2).getCell(3).getStringCellValue());
        Row linhaTotal = sheet.getRow(3);
        assertEquals("Total", linhaTotal.getCell(0).getStringCellValue());
        assertEquals(30.0, linhaTotal.getCell(1).getNumericCellValue(), 0.001);
    }

    @Test
    @DisplayName("multiplicarColunasComTexto não deve apagar células já existentes na linha de total")
    void deveMultiplicarColunasComTextoSemApagarCelulasExistentes() {
        criarCelula(0, 3, 2.0);
        criarCelula(0, 8, 3.0);
        criarCelula(1, 3, 4.0);
        criarCelula(1, 8, 5.0);
        Row linhaExistente = sheet.createRow(2);
        linhaExistente.createCell(11).setCellValue("Preservar");

        Calculos.multiplicarColunasComTexto(sheet, "D", "I", 1, "Total", "J");

        Row linhaTotal = sheet.getRow(2);
        assertEquals("D1*I1", sheet.getRow(0).getCell(9).getCellFormula());
        assertEquals("D2*I2", sheet.getRow(1).getCell(9).getCellFormula());
        assertEquals("Total", linhaTotal.getCell(8).getStringCellValue());
        assertEquals("SUM(J1:J2)", linhaTotal.getCell(9).getCellFormula());
        assertEquals("Preservar", linhaTotal.getCell(11).getStringCellValue());
    }

    // ========== MÉTODOS AUXILIARES ==========

    /**
     * Helper: Cria uma célula numérica.
     */
    private void criarCelula(int rowIndex, int colIndex, double valor) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.createCell(colIndex);
        cell.setCellValue(valor);
    }
}
