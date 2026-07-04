package com.abnote.planilhas.utils;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Formatter;
import java.util.logging.Level;
import java.util.logging.LogRecord;

import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Formatter que adiciona cores ANSI e formata as mensagens de log.
 */
public class ColorFormatter extends Formatter {

	private static final String RESET = "\u001B[0m";
	private static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

	@Override
	public String format(LogRecord record) {
		CorEnum cor;

		// Mapear os níveis de log para as cores definidas em CorEnum
		if (record.getLevel() == Level.SEVERE) {
			cor = CorEnum.VERMELHO_ESCURO;
		} else if (record.getLevel() == Level.WARNING) {
			cor = CorEnum.AMARELO;
		} else if (record.getLevel() == Level.INFO) {
			cor = CorEnum.VERDE;
		} else if (record.getLevel() == Level.CONFIG) {
			cor = CorEnum.TURQUESA;
		} else if (record.getLevel() == Level.FINE) {
			cor = CorEnum.AZUL;
		} else if (record.getLevel() == Level.FINER) {
			cor = CorEnum.ROXO;
		} else if (record.getLevel() == Level.FINEST) {
			cor = CorEnum.PRETO;
		} else {
			cor = CorEnum.BRANCO; // Nível padrão
		}

		String ansiCode = (cor.getAnsiCode() != null) ? cor.getAnsiCode() : "";
		StringBuilder builder = new StringBuilder();

		builder.append(ansiCode);

		// Incluir timestamp
		String timestamp = new SimpleDateFormat(DATE_FORMAT).format(new Date(record.getMillis()));
		builder.append("[").append(timestamp).append("] ");

		// Incluir nível de log
		builder.append("[").append(record.getLevel().getLocalizedName()).append("] ");

		// Incluir nome do logger (geralmente a classe)
		builder.append("[").append(record.getLoggerName()).append("] ");

		// Incluir a mensagem formatada
		builder.append(formatMessage(record));

		builder.append(RESET).append(System.lineSeparator());
		return builder.toString();
	}
}
