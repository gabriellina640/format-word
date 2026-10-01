# Usabilidade Word Implementation Plan

> **For agentic workers:** executar tarefas delimitadas e usar requesting-code-review para revisão independente final.

**Goal:** Usuário reconhece termos do Word e configura um perfil por seções curtas.

**Architecture:** Esquema de configuração preservado. Rótulos no módulo compartilhado; componente de recuo especial converte apresentação para valor assinado. Navegação filtra cartões existentes e revela seção em erro. Atalhos de espaçamento e seletor de cor escrevem nos mesmos campos persistidos.

**Tech Stack:** CustomTkinter/ttk, Python e unittest.

## Tarefas atômicas

- [x] 1. `app/ui_fields.py`, `tests/test_ui_values.py`: substituir rótulos/ajudas e resumo por termos Word, manter IDs e enumeração serializada. Verificar roundtrip completo e resumo (`.venv/bin/python -m unittest discover -s tests -p test_ui_values.py -v`).
- [x] 2. Criar `app/word_controls.py` e testes GUI: controle de recuo especial com seletores Nenhum/Primeira linha/Deslocado e valor positivo, sincronização bidirecional para StringVar canônica; invalidade nunca vira zero silenciosamente. Testar positivos/negativos/zero, decimal vírgula, valores externos, disabled e destruição dos traces.
- [x] 3. `app/ui.py`, `tests/test_ui_integration.py`: navegação seções e Anterior/Próximo, orientação de três passos, manter cartões e valores, revelar campos inválidos. Verificar campos inacessíveis por grupo oculto abrem grupo automaticamente e ausência de perda ao navegar.
- [x] 4. `app/ui.py`, `app/review_dialogs.py`: instalar controle especial, botões Simples/1,5 linhas/Duplo e indicação de unidade em Em; seletor de cor; atualizar textos de ajuda e ações. Regredir seletores de primeira escolha, configurações por tipo, importação e salvar/reabrir.
- [x] 5. `app/profile_import_dialog.py`, `app/selftest.py`, docs: alinhar termos e orientações, verificar resumo e atualizar smoke/navegação. Revisão independente de conformidade e qualidade, corrigir findings.
- [x] 6. Suíte completa com `FORMATWORD_GUI_TESTS=1`, compileall e diff-check; registrar em `docs/validation/2026-09-30-usabilidade-word.md`. Recriar app macOS, executar self-test e atualizar ZIP/guia.

Continuar na branch de funcionalidade já usada, preservando a implementação
anterior aprovada e o workspace compartilhado. Não publicar nem fazer merge.

Concluído: evidências em `docs/validation/2026-09-30-usabilidade-word.md`.
