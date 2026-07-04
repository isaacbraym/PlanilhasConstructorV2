package com.abnote.planilhas.estilos.estilos;

public enum CorEnum {
    VERMELHO_ESCURO(139, 0, 0, "\u001B[31m"),
    AZUL(0, 0, 255, "\u001B[34m"),
    VERDE(0, 128, 0, "\u001B[32m"),
    PRETO(0, 0, 0, "\u001B[30m"),
    BRANCO(255, 255, 255, "\u001B[37m"),
    CINZA_CLARO(211, 211, 211, "\u001B[37m"),
    AMARELO(255, 255, 0, "\u001B[33m"),
    LARANJA(255, 165, 0, null), // Não há código ANSI direto para laranja
    ROSA(255, 192, 203, null),
    ROXO(128, 0, 128, "\u001B[35m"),
    VIOLETA(238, 130, 238, null),
    AZUL_CELESTE(135, 206, 235, null),
    VERDE_LIMAO(50, 205, 50, null),
    MARROM(165, 42, 42, null),
    DOURADO(255, 215, 0, null),
    PRATA(192, 192, 192, null),
    BEGE(245, 245, 220, null),
    SALMAO(250, 128, 114, null),
    TURQUESA(64, 224, 208, "\u001B[36m"),
    LAVANDA(230, 230, 250, null),
    AZUL_MARINHO(0, 0, 128, null),
    CINZA_ESCURO(169, 169, 169, null),
    OLIVA(128, 128, 0, null),
    CORAL(255, 127, 80, null),
    MENTA(189, 252, 201, null),
    VERDE_OLIVA(107, 142, 35, null),
    VERDE_AGUA(32, 178, 170, null);

    private final int red;
    private final int green;
    private final int blue;
    private final String ansiCode;

    CorEnum(int red, int green, int blue, String ansiCode) {
        this.red = red;
        this.green = green;
        this.blue = blue;
        this.ansiCode = ansiCode;
    }

    public int getRed() {
        return red;
    }

    public int getGreen() {
        return green;
    }

    public int getBlue() {
        return blue;
    }

    public String getAnsiCode() {
        return ansiCode;
    }
}
