package com.abnote.planilhas.utils;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("LogsDeModificadores")
class LogsDeModificadoresTest {

	private PrintStream saidaOriginal;
	private ByteArrayOutputStream saidaCapturada;

	@BeforeEach
	void setUp() {
		saidaOriginal = System.out;
		saidaCapturada = new ByteArrayOutputStream();
		System.setOut(new PrintStream(saidaCapturada));
	}

	@AfterEach
	void tearDown() {
		System.setOut(saidaOriginal);
	}

	@Test
	@DisplayName("exibirLogs deve imprimir deslocamento e limpar a fila")
	void deveExibirDeslocamentoELimparFila() {
		LogsDeModificadores logs = new LogsDeModificadores();
		LogsDeModificadores.ActionLog acao = new LogsDeModificadores.ActionLog("Deslocamento de colunas",
				new LogsDeModificadores.ColumnMovement("Total", "A", "C"));
		acao.getShiftedColumns().add(new LogsDeModificadores.ColumnMovement("Preco", "B", "A"));
		logs.adicionarLog(acao);

		logs.exibirLogs();

		String primeiraSaida = saidaCapturada.toString();
		assertTrue(primeiraSaida.contains("Deslocamento de colunas"));
		assertTrue(primeiraSaida.contains("Total"));
		assertTrue(primeiraSaida.contains("Preco"));
		assertTrue(primeiraSaida.contains("A"));
		assertTrue(primeiraSaida.contains("C"));

		saidaCapturada.reset();
		logs.exibirLogs();

		assertEquals("", saidaCapturada.toString(), "A fila de logs deve ser limpa após exibir");
	}

	@Test
	@DisplayName("exibirLogs deve cobrir remoção, inserção e limpeza")
	void deveExibirRemocaoInsercaoELimpeza() {
		LogsDeModificadores logs = new LogsDeModificadores();
		LogsDeModificadores.ActionLog remocao = new LogsDeModificadores.ActionLog("Remoção de coluna",
				new LogsDeModificadores.ColumnMovement("Status", "D", null));
		remocao.getShiftedColumns().add(new LogsDeModificadores.ColumnMovement("Data", "E", "D"));
		logs.adicionarLog(remocao);
		logs.adicionarLog(new LogsDeModificadores.ActionLog("Inserção de coluna vazia",
				new LogsDeModificadores.ColumnMovement(null, "A", "B")));
		logs.adicionarLog(new LogsDeModificadores.ActionLog("Limpeza de coluna",
				new LogsDeModificadores.ColumnMovement("Observacao", "F", null)));

		logs.exibirLogs();

		String saida = saidaCapturada.toString();
		assertTrue(saida.contains("Status"));
		assertTrue(saida.contains("Data"));
		assertTrue(saida.contains("A"));
		assertTrue(saida.contains("B"));
		assertTrue(saida.contains("Observacao"));
		assertTrue(saida.contains("F"));
	}

	@Test
	@DisplayName("tipo desconhecido deve ser ignorado sem imprimir")
	void deveIgnorarTipoDesconhecido() {
		LogsDeModificadores logs = new LogsDeModificadores();
		logs.adicionarLog(new LogsDeModificadores.ActionLog("Tipo futuro",
				new LogsDeModificadores.ColumnMovement("X", "A", "B")));

		logs.exibirLogs();

		assertEquals("", saidaCapturada.toString());
	}
}
