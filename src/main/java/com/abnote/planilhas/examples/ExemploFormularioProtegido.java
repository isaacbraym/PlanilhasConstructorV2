package com.abnote.planilhas.examples;

import java.io.File;
import java.time.LocalDate;

import com.abnote.planilhas.Planilha;
import com.abnote.planilhas.estilos.estilos.CorEnum;

/**
 * Receita: um formulário de cadastro onde só os campos de entrada são
 * editáveis — o restante da planilha fica travado. Combina lista suspensa,
 * validação de limites e proteção.
 *
 * <p>
 * Comandos demonstrados: {@code listaSuspensa}, {@code validarDataEntre},
 * {@code validarNumeroEntre}, {@code comentario}, {@code desbloquearCelulas},
 * {@code protegerPlanilha}.
 * </p>
 */
public class ExemploFormularioProtegido {

	public static void main(String[] args) {
		final String caminho = new File(System.getProperty("java.io.tmpdir"), "cadastro-protegido.xlsx")
				.getAbsolutePath();

		try (Planilha planilha = Planilha.nova("Cadastro")) {

			// 1) Cabeçalho (fica travado depois de proteger).
			planilha.escreverLinha("A1", "Nome", "Cargo", "Data de Admissão", "Salário");
			planilha.negrito("A1:D1").corDeFundo("A1:D1", CorEnum.CINZA_ESCURO).corDoTexto("A1:D1", CorEnum.BRANCO);
			planilha.comentario("A1", "Preencha uma linha por funcionário. Só as células brancas são editáveis.");

			// 2) Área de entrada: 20 linhas em branco prontas para preencher.
			final int primeiraLinhaDeDados = 2;
			final int ultimaLinhaDeDados = 21;

			// 3) Lista suspensa de cargos.
			planilha.listaSuspensa("B" + primeiraLinhaDeDados + ":B" + ultimaLinhaDeDados, "Desenvolvedor(a)",
					"Analista", "Gerente", "Estagiário(a)");

			// 4) Só aceita datas de admissão a partir de 2015 até hoje.
			planilha.validarDataEntre("C" + primeiraLinhaDeDados + ":C" + ultimaLinhaDeDados, LocalDate.of(2015, 1, 1),
					LocalDate.now());

			// 5) Só aceita salário positivo (até um teto razoável).
			planilha.validarNumeroEntre("D" + primeiraLinhaDeDados + ":D" + ultimaLinhaDeDados, 0.01, 100000.0);

			// 6) Destrava a área de entrada ANTES de proteger — senão tudo fica travado.
			planilha.desbloquearCelulas("A" + primeiraLinhaDeDados + ":D" + ultimaLinhaDeDados);
			planilha.protegerPlanilha(""); // sem senha: só impede edição acidental

			planilha.ajustarColunas().salvar(caminho);

			System.out.println("Formulário protegido criado em: " + caminho);
		}
	}
}
