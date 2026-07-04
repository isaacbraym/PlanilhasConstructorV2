package com.abnote.planilhas.utils;

import java.util.logging.ConsoleHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Classe utilitária para configuração do logging do projeto.
 */
public class LoggerUtil {

	static {
		// Configurar o logger global
		Logger rootLogger = Logger.getLogger("");

		// Remover os handlers padrão
		Handler[] handlers = rootLogger.getHandlers();
		if (handlers != null) {
			for (Handler handler : handlers) {
				rootLogger.removeHandler(handler);
			}
		}

		// Criar um ConsoleHandler personalizado que escreve em System.out
		ConsoleHandler consoleHandler = new ConsoleHandler() {
			{
				setOutputStream(System.out);
			}
		};

		// Definir o nível para capturar todos os logs
		consoleHandler.setLevel(Level.ALL);
		rootLogger.setLevel(Level.ALL);

		// Definir o ColorFormatter para o handler
		consoleHandler.setFormatter(new ColorFormatter());

		// Adicionar o handler ao logger raiz
		rootLogger.addHandler(consoleHandler);
	}

	/**
	 * Retorna um logger para a classe informada.
	 *
	 * @param clazz A classe para a qual se deseja o logger.
	 * @return Logger configurado para a classe.
	 */
	public static Logger getLogger(Class<?> clazz) {
		return Logger.getLogger(clazz.getName());
	}
}
