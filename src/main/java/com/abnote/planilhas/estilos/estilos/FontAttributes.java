package com.abnote.planilhas.estilos.estilos;

import java.awt.Color;

public class FontAttributes {
	private String fontName;
	private Short fontSize;
	private Boolean bold;
	private Boolean italic;
	private Byte underline;
	private Boolean strikeout;
	private Short colorIndex;
	private Color colorRGB;

	// Getters e Setters

	public String getFontName() {
		return fontName;
	}

	public FontAttributes setFontName(String fontName) {
		this.fontName = fontName;
		return this;
	}

	public Short getFontSize() {
		return fontSize;
	}

	public FontAttributes setFontSize(Short fontSize) {
		this.fontSize = fontSize;
		return this;
	}

	public Boolean isBold() {
		return bold;
	}

	public FontAttributes setBold(Boolean bold) {
		this.bold = bold;
		return this;
	}

	public Boolean isItalic() {
		return italic;
	}

	public FontAttributes setItalic(Boolean italic) {
		this.italic = italic;
		return this;
	}

	public Byte getUnderline() {
		return underline;
	}

	public FontAttributes setUnderline(Byte underline) {
		this.underline = underline;
		return this;
	}

	public Short getColorIndex() {
		return colorIndex;
	}

	public FontAttributes setColorIndex(Short colorIndex) {
		this.colorIndex = colorIndex;
		return this;
	}

	public Color getColorRGB() {
		return colorRGB;
	}

	public FontAttributes setColorRGB(Color colorRGB) {
		this.colorRGB = colorRGB;
		return this;
	}

	public Boolean isStrikeout() {
		return strikeout;
	}

	public FontAttributes setStrikeout(Boolean strikeout) {
		this.strikeout = strikeout;
		return this;
	}
}
