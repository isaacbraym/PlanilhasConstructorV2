package com.abnote.planilhas;

import static org.junit.jupiter.api.Assertions.*;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Testes da API de leitura da facade: {@code ler}, {@code lerTexto},
 * {@code lerNumero}, {@code lerData}, {@code lerTabela} e
 * {@code contarLinhasPreenchidas}.
 *
 * <p>
 * Antes desta leva, não havia nenhuma forma de ler um valor de volta sem
 * chamar {@code planilha.workbook()} e navegar o POI diretamente.
 * </p>
 */
@DisplayName("Planilha — API de leitura")
class LeituraFacadeTest {

	@Test
	@DisplayName("ler deve devolver String para célula de texto")
	void deveLerTexto() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", "Texto");
			assertEquals("Texto", planilha.ler("A1"));
		}
	}

	@Test
	@DisplayName("ler deve devolver Double para célula numérica")
	void deveLerNumeroComoDouble() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", 42.5);
			assertEquals(42.5, (Double) planilha.ler("A1"), 0.001);
		}
	}

	@Test
	@DisplayName("ler deve devolver LocalDateTime para célula de data")
	void deveLerDataComoLocalDateTime() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverData("A1", LocalDate.of(2024, 1, 15));
			Object valor = planilha.ler("A1");
			assertInstanceOf(java.time.LocalDateTime.class, valor);
			assertEquals(LocalDate.of(2024, 1, 15), ((java.time.LocalDateTime) valor).toLocalDate());
		}
	}

	@Test
	@DisplayName("ler deve devolver null para célula vazia ou inexistente")
	void deveDevolverNullParaCelulaVazia() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			assertNull(planilha.ler("Z99"));
		}
	}

	@Test
	@DisplayName("ler deve avaliar fórmulas antes de devolver o valor")
	void deveLerValorDeFormula() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", 10, 20, 30).somar("A4", "A1:A3");
			assertEquals(60.0, (Double) planilha.ler("A4"), 0.001);
		}
	}

	@Test
	@DisplayName("lerNumero deve devolver null quando a célula não é número")
	void deveLerNumeroOuNull() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", 10).escrever("A2", "texto");
			assertEquals(10.0, planilha.lerNumero("A1"), 0.001);
			assertNull(planilha.lerNumero("A2"));
		}
	}

	@Test
	@DisplayName("lerData deve devolver a data escrita com escreverData")
	void deveLerData() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverData("A1", LocalDate.of(2024, 12, 25));
			assertEquals(LocalDate.of(2024, 12, 25), planilha.lerData("A1"));
			assertNull(planilha.lerData("A2"), "Célula vazia não é data");
		}
	}

	@Test
	@DisplayName("lerTexto deve devolver o valor formatado como aparece no Excel")
	void deveLerTextoFormatado() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escrever("A1", 1234.5).formatarComoMoeda("A1");
			String textoFormatado = planilha.lerTexto("A1");
			assertTrue(textoFormatado.contains("R$"), "Deve conter o símbolo de moeda: " + textoFormatado);

			assertEquals("", planilha.lerTexto("Z99"), "Célula inexistente vira texto vazio");
		}
	}

	@Test
	@DisplayName("lerTabela deve devolver as linhas de dados, sem o cabeçalho")
	void deveLerTabela() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Preco").adicionarLinha("Caneta", 2.5).adicionarLinha("Caderno",
					15.9);

			List<List<Object>> dados = planilha.lerTabela("A1");

			assertEquals(2, dados.size());
			assertEquals("Caneta", dados.get(0).get(0));
			assertEquals(2.5, (Double) dados.get(0).get(1), 0.001);
			assertEquals("Caderno", dados.get(1).get(0));
		}
	}

	@Test
	@DisplayName("lerTabela deve devolver lista vazia quando não há dados")
	void deveLerTabelaVazia() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverLinha("A1", "Produto", "Preco");
			assertTrue(planilha.lerTabela("A1").isEmpty());
		}
	}

	@Test
	@DisplayName("contarLinhasPreenchidas deve contar as linhas com dado na coluna")
	void deveContarLinhasPreenchidas() throws Exception {
		try (Planilha planilha = Planilha.nova("T")) {
			planilha.escreverColuna("A1", "a", "b", "c");
			assertEquals(3, planilha.contarLinhasPreenchidas("A"));
			assertEquals(0, planilha.contarLinhasPreenchidas("Z"));
		}
	}
}
