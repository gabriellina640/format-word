# Seletores e importação de perfil Implementation Plan

> **For agentic workers:** Use executing-plans para execução e requesting-code-review para revisão independente. Etapas abaixo são verificáveis separadamente.

**Goal:** Seletores funcionam na primeira escolha; um DOCX preenche um novo perfil após revisão explícita.

**Architecture:** Correção localizada no agendamento da UI. Leitor independente de Tk resolve propriedades efetivas, agrupa alternativas por categoria e seção e devolve relatório. Diálogo modal deixa escolher alternativas e revisar resumo; confirmação cria rascunho, sem salvar automaticamente.

**Tech Stack:** Python, python-docx/lxml, CustomTkinter/ttk e unittest.

Ambiente: continuar na branch de funcionalidade `codex/formatacao-fiel`, no workspace compartilhado existente. Não há frentes em branches diferentes; arquivos de implementação delegados têm dono exclusivo.

## 1. Corrigir seletores (5–15 min)

Arquivos: `app/ui.py`, `tests/test_ui_integration.py`.

- [x] Adicionar teste que chama `_dropdown_callback` para cada opção, alterna nos dois sentidos e verifica valor antes/depois do processamento Tk.
- [x] Executar com `FORMATWORD_GUI_TESTS=1 .venv/bin/python -m unittest discover -s tests -p test_ui_integration.py -v`; registrar falha esperada de valor vazio.
- [x] Em `_changed`, agendar uma única atualização com `_schedule(0, callback)` em vez de executar `_refresh_form` dentro do trace. Callback limpa seu identificador; destruição usa `_after_ids` existente.
- [x] Testar campos dependentes, bloqueio durante lote e fechamento pendente.

## 2. Ler perfil do DOCX (tarefas de 5–15 min)

Arquivos: criar `app/profile_import.py`, `tests/test_profile_import.py`.

- [x] Testes de página e de corpo com formatação direta; rodar e confirmar falha por recurso ausente.
- [x] Implementar `inspect_profile(path, max_input_mb=50)` retornando relatório de alternativas de página/texto, avisos e `build_settings(page_index=0, selections=None)`.
- [x] Resolver herança de estilos, padrões do documento, fontes e cores de tema. Propriedades não resolvidas geram aviso e usam valor identificado como padrão.
- [x] Agrupar configurações completas por categoria; ordenar por frequência; textos curtos exemplificam cada opção. Preservar categorias ausentes. Não misturar propriedades de alternativas distintas silenciosamente.
- [x] Expor alternativas por seção; avisar cabeçalhos/rodapés e conteúdo não representável. Reusar validação de entrada e limites existentes. Não modificar arquivo fonte.
- [x] Testar estilos herdados, temas, seções divergentes, formatação mista, arquivo inválido, fonte imutável e aplicação em outro documento.

## 3. Integrar revisão e criação de rascunho (tarefas de 5–15 min)

Arquivos: criar `app/profile_import_dialog.py`; modificar `app/ui.py`, `tests/test_ui_integration.py`.

- [x] Testes da ação de importar, cancelar, confirmar e salvar/reabrir; nome sem colisão; bloqueio em lote e erro de leitura sem alteração.
- [x] Botão “Importar de Word” nos perfis. Importar apenas lê e abre diálogo; confirmação resolve rascunho anterior via fluxo existente e carrega novo perfil com nome único.
- [x] Diálogo ttk modal redimensionável com seletores de página/categoria, resumo rolável, avisos, Cancelar/Criar rascunho. Escolhas alteram resumo antes de confirmar. Erro de validação impede confirmação sem perder estado.
- [x] Visual: cores/tipografia nativas do app; espaçamento 8/16 px; foco de teclado e Escape; texto sem IDs/XML. Layout adapta às dimensões desktop suportadas (mínimo 760×600 no app), com rolagem no diálogo.
- [x] Verificar a interface real e a transação de cancelamento/confirmação; rascunho marcado como não salvo.

## 4. Revisão, documentação e entrega (tarefas de 5–15 min)

Arquivos: `README.md`, `docs/guia-rapido.md`, `app/selftest.py`, `docs/validation/2026-09-30-importacao-perfil.md`.

- [x] Revisão independente de conformidade, seguida de qualidade; corrigir achados relevantes e repetir testes afetados.
- [x] Documentar importação e limites, incluindo que preservar cabeçalho mantém o documento de destino, sem copiar timbrado.
- [x] Rodar suíte completa com interface, compileall e `git diff --check`; registrar comandos e resultados.
- [x] Reconstruir pacote macOS com comando já usado no projeto, executar self-test do binário e atualizar zip/guia locais. Não declarar teste Windows sem runner Windows.
- [x] Verificar diff, plano e artefatos finais antes de informar conclusão.

Concluído em 30/09/2026. Evidências: `docs/validation/2026-09-30-importacao-perfil.md`.
