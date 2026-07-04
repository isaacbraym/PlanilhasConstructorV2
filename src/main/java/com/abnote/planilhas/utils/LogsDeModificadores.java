package com.abnote.planilhas.utils;

import java.util.ArrayList;
import java.util.List;

public class LogsDeModificadores {
	private List<ActionLog> actionLogs = new ArrayList<>();
	private static final String SEPARA_BLOCO = "____________________________________________________\n";
	private static final String SEPARA_BLOCO2 = "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\n";

	// Método para adicionar um log de ação
	public void adicionarLog(ActionLog log) {
		actionLogs.add(log);
	}

	// Método para exibir os logs de alterações
	public void exibirLogs() {
		for (ActionLog actionLog : actionLogs) {
			switch (actionLog.getActionType()) {
			case "Deslocamento de colunas":
				exibirDeslocamentoDeColunas(actionLog);
				break;
			case "Remoção de coluna":
				exibirRemocaoDeColuna(actionLog);
				break;
			case "Inserção de coluna vazia":
				exibirInsercaoDeColunaVazia(actionLog);
				break;
			case "Limpeza de coluna": // Novo caso para limpeza de coluna
				exibirLimpezaDeColuna(actionLog);
				break;
			default:
				// Tipos de ações no futuro
				break;
			}
		}
		// Limpa os logs após a exibição
		actionLogs.clear();
	}

	private void exibirDeslocamentoDeColunas(ActionLog actionLog) {
		System.out.println(SEPARA_BLOCO + "Deslocamento de colunas:");
		System.out.println(SEPARA_BLOCO2 + "Coluna \"" + actionLog.getMainMovement().getColumnName()
				+ "\": Índice anterior: " + actionLog.getMainMovement().getPreviousIndex() + "; Novo Índice: "
				+ actionLog.getMainMovement().getNewIndex());
		if (!actionLog.getShiftedColumns().isEmpty()) {
			System.out.println("Colunas deslocadas nesta ação:");
			for (ColumnMovement cm : actionLog.getShiftedColumns()) {
				System.out.println("\"" + cm.getColumnName() + "\": Índice anterior: " + cm.getPreviousIndex()
						+ "; Novo Índice: " + cm.getNewIndex());
			}
			System.out.println(SEPARA_BLOCO2);
		}
	}

	private void exibirRemocaoDeColuna(ActionLog actionLog) {
		System.out.println(SEPARA_BLOCO + "Remoção da coluna \"" + actionLog.getMainMovement().getColumnName()
				+ "\" em índice " + actionLog.getMainMovement().getPreviousIndex());
		if (!actionLog.getShiftedColumns().isEmpty()) {
			System.out.println(SEPARA_BLOCO2 + "Colunas deslocadas nesta ação:");
			for (ColumnMovement cm : actionLog.getShiftedColumns()) {
				System.out.println("\"" + cm.getColumnName() + "\": Índice anterior: " + cm.getPreviousIndex()
						+ "; Novo Índice: " + cm.getNewIndex());
			}
			System.out.println(SEPARA_BLOCO2);
		}
	}

	private void exibirLimpezaDeColuna(ActionLog actionLog) {
		System.out.println(SEPARA_BLOCO + "Limpeza da coluna \"" + actionLog.getMainMovement().getColumnName()
				+ "\" em índice " + actionLog.getMainMovement().getPreviousIndex());
		System.out.println(SEPARA_BLOCO2);
	}

	private void exibirInsercaoDeColunaVazia(ActionLog actionLog) {
		System.out.println(
				SEPARA_BLOCO + "Inserção de coluna vazia entre \"" + actionLog.getMainMovement().getPreviousIndex()
						+ "\" e \"" + actionLog.getMainMovement().getNewIndex() + "\"");
		if (!actionLog.getShiftedColumns().isEmpty()) {
			System.out.println(SEPARA_BLOCO2 + "Colunas deslocadas nesta ação:");
			for (ColumnMovement cm : actionLog.getShiftedColumns()) {
				System.out.println("\"" + cm.getColumnName() + "\": Índice anterior: " + cm.getPreviousIndex()
						+ "; Novo Índice: " + cm.getNewIndex());
			}
			System.out.println(SEPARA_BLOCO2);
		}
	}

	// Classes auxiliares internas

	public static class ColumnMovement {
		private String columnName;
		private String previousIndex;
		private String newIndex;

		public ColumnMovement(String columnName, String previousIndex, String newIndex) {
			this.columnName = columnName;
			this.previousIndex = previousIndex;
			this.newIndex = newIndex;
		}

		public String getColumnName() {
			return columnName;
		}

		public String getPreviousIndex() {
			return previousIndex;
		}

		public String getNewIndex() {
			return newIndex;
		}
	}

	public static class ActionLog {
		private String actionType;
		private ColumnMovement mainMovement;
		private List<ColumnMovement> shiftedColumns = new ArrayList<>();

		public ActionLog(String actionType, ColumnMovement mainMovement) {
			this.actionType = actionType;
			this.mainMovement = mainMovement;
		}

		public String getActionType() {
			return actionType;
		}

		public ColumnMovement getMainMovement() {
			return mainMovement;
		}

		public List<ColumnMovement> getShiftedColumns() {
			return shiftedColumns;
		}
	}
}