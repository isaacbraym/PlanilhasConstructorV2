package com.abnote.planilhas.estilos.estilos;

public enum FonteEnum {
    ARIAL("Arial"),
    SEGOE_UI("Segoe UI"),
    TIMES_NEW_ROMAN("Times New Roman"),
    CALIBRI("Calibri"),
    COURIER_NEW("Courier New"),
    GEORGIA("Georgia"),
    VERDANA("Verdana"),
    TAHOMA("Tahoma"),
    IMPACT("Impact"),
    TREBUCHET_MS("Trebuchet MS"),
    COMIC_SANS_MS("Comic Sans MS"),
    LUCIDA_CONSOLE("Lucida Console"),
    LUCIDA_SANS_UNICODE("Lucida Sans Unicode"),
    CONSOLAS("Consolas"),
    CAMBRIA("Cambria"),
    CANDARA("Candara"),
    CORBEL("Corbel"),
    GARAMOND("Garamond"),
    PALATINO_LINOTYPE("Palatino Linotype"),
    BOOK_ANTIQUE("Book Antiqua"),
    FRANKLIN_GOTHIC_MEDIUM("Franklin Gothic Medium"),
    MS_SANS_SERIF("MS Sans Serif"),
    MS_SERIF("MS Serif"),
    CENTURY_GOTHIC("Century Gothic"),
    EBRIMA("Ebrima"),
    SITKA("Sitka");

    private final String fontName;

    FonteEnum(String fontName) {
        this.fontName = fontName;
    }

    public String getFontName() {
        return fontName;
    }
}
