@AGENTS.md

Instruções específicas para o Claude Code neste projeto estão no `AGENTS.md`
acima (mesma fonte usada por Codex e demais agentes).

Atalhos úteis:

- Skills prontas para este repositório: `.claude/skills/` (invoque com
  `/planilha-facade`, `/planilha-nova-feature`, `/planilha-testes`).
- Antes de finalizar qualquer tarefa: `mvn clean test` deve ficar verde.
- API que o usuário final usa: a facade `com.abnote.planilhas.Planilha`.
