# Formatação fiel — plano de implementação

> Execução: tarefas incrementais com testes de regressão e revisão independente
> conforme as skills executing-plans, test-driven-development e requesting-code-review.

**Goal:** aplicar perfis verificáveis em DOCX preservados, individualmente ou em lote.

**Architecture:** configuração tipada e validada compartilhada pela interface e
motor; edição do pacote DOCX original; lote com eventos e cancelamento; UI sem
editor destrutivo ou prévia fictícia.

**Tech Stack:** Python 3.14, python-docx 1.2, Pillow, CustomTkinter e unittest.

Design aprovado pelo usuário em 28/09/2026. Branch: `codex/formatacao-fiel`.
Trabalho no checkout compartilhado para entregar as alterações diretamente
no projeto. Baseline já registrado: 7 sucessos, 1 falha e 8 erros.

## 1. Configuração e migração (app/config.py; tests/test_config.py)

- [x] Escrever testes para números finitos/decimais, validação e round-trip.
- [x] Executar `.venv/bin/python -m unittest discover -s tests -p test_config.py -v`;
  confirmar falhas dos novos contratos antes da implementação.
- [x] Implementar `SettingsError(ValueError)` com atributo `field` e
  `validate_settings(settings, check_assets=False) -> FormatSettings`.
- [x] Campos: fonte/tamanho; alinhamento; modo/valor entrelinhas; antes/depois;
  recuos esquerda/direita/primeira linha; papel/orientação/dimensões/margens;
  flags opcionais de texto/paginação; cor; modos/imagens/distâncias/alinhamentos
  de cabeçalho/rodapé; sufixo e limite de entrada.
- [x] Persistir com `AppConfig(settings, stacks, active_stack)`; expor avisos em
  `ConfigStore.warnings`; migrar `justify_text` e `include_header/footer`.
  Manter `template_path` como bloqueio de revisão do legado.
- [x] Recuperar arquivo inválido com cópia de segurança antes de qualquer nova
  gravação; preservar configurações válidas; rejeitar tipos errados.
- [x] Importar PNG/JPEG sem deformar/recortar; cópia independente por perfil.
- [x] Reexecutar testes. Exemplo de contrato:

```python
settings = FormatSettings(font_size=12.5, font_name="Fonte personalizada")
store.save(AppConfig(settings=settings, stacks={"Perfil": settings}))
assert store.load().stacks["Perfil"] == settings
```

## 2. Preservação e fidelidade (app/formatter.py; tests/test_formatter.py)

- [x] Testes RED: tabelas/imagens/links/listas/seções preservados e primeira
  formatação configurada sem perda do negrito original.
- [x] Substituir reconstrução de texto por `Document(input_path)` e edição de
  parágrafos/runs do corpo, inclusive tabelas aninhadas e hyperlinks.
- [x] Aplicar propriedades de texto diretamente; retirar overrides temáticos
  conflitantes apenas quando a propriedade for explicitamente configurada.
- [x] Aplicar propriedades de parágrafo e de todas as seções. Limpar overrides
  automáticos que impediriam os valores de espaçamento/recuo configurados.
- [x] Bloquear conteúdo que não permita a formatação prometida (revisões,
  caixas de texto/objetos incompatíveis); comunicar campos e notas preservados.
- [x] Verificar DOCX e recursos antes de publicar saída. Arquivo original
  imutável; gravação temporária e publicação exclusiva sem sobrescrita.
- [x] Executar `.venv/bin/python -m unittest discover -s tests -p test_formatter.py -v`.

```python
result = format_document(source, output_dir, FormatSettings(font_size=36))
output = Document(result.output_path)
assert output.paragraphs[0].runs[0].font.size.pt == 36
assert len(output.tables) == len(Document(source).tables)
```

## 3. Cabeçalho/rodapé (mesmos arquivos e testes)

- [x] Testes para preservar/remover/substituir variantes de todas as seções.
- [x] Preservar conteúdo e relações por padrão; aplicar imagem com proporção
  original, largura, distância e alinhamento configurados.
- [x] Validar espaço útil antes de gerar o arquivo. Não ajustar margens
  silenciosamente. Impedir template legado ou imagem ausente.
- [x] Conferir XML dos cabeçalhos e reabrir imagens embutidas nos testes.

## 4. Lote (app/batch.py; tests/test_batch.py)

- [x] RED: um DOCX inválido entre dois válidos; cancelamento; nomes repetidos.
- [x] Implementar `BatchItem(input_path, result, error, cancelled)` e
  `format_documents_batch(paths, output_dir, settings, cancel_event, on_result)`.
- [x] Capturar snapshot de settings, continuar por arquivo e emitir progresso;
  cancelar entre arquivos sem interromper a gravação em andamento.
- [x] Executar `.venv/bin/python -m unittest discover -s tests -p test_batch.py -v`.

## 5. Interface (app/ui.py; app/ui_fields.py; tests/test_ui_values.py)

- [x] RED: parse de vírgula/decimais, erro por campo, resumo e valores salvos.
- [x] Substituir prévia/editor por seleção múltipla, resumo, aplicar/cancelar,
  lista de resultados e abrir arquivo/pasta.
- [x] Formulário em grupos com unidades e seleção explícita de preservar.
  Usar a mesma instância validada para resumo, perfis e aplicação.
- [x] CRUD/duplicação de perfil; detectar alterações não salvas antes de trocar,
  restaurar ou fechar; impedir nomes reservados e sobrescrita involuntária.
- [x] Thread de processamento publica eventos em `queue.Queue`; somente a
  thread principal acessa widgets. Bloquear controles durante execução.
- [x] Direção visual: desktop claro, tokens existentes azul/branco/cinza escuro,
  fonte de sistema, espaços 8/16/24, botões com pelo menos 36 px de altura;
  alternativa de abas densas descartada em favor de grupos com rolagem.
- [x] Teclado, foco, rótulos, contraste e estados de erro/busy; testar janelas
  768/1024/1440 px e limite mínimo desktop. 375 px não é alvo deste app desktop.

## 6. Integração, documentação e CI

- [x] Retirar `tests/` do `.gitignore`; preservar cópia dos testes antigos para
  auditoria local e registrar expectativas retiradas por mudança aprovada.
- [x] Remover PDF e dependência pypdf; ajustar requirements e scripts de build.
- [x] Configurar testes antes do build no CI Windows e testes também em Linux.
  Execução remota pendente: não há runner Windows/Linux neste ambiente macOS.
- [x] Atualizar README com instalação, opções/unidades, migração, fluxo e limites.

## 7. Revisão e validação final

- [x] Revisão independente de conformidade, seguida de qualidade; corrigir
  achados importantes com regressões.
- [x] `.venv/bin/python -m unittest discover -s tests -v` sem falhas.
- [x] `.venv/bin/python -m compileall -q app main.py tests` e `git diff --check`.
- [x] Smoke test gráfico com configuração temporária: criar/reabrir perfil,
  lote, cancelamento, redimensionamento, encerramento seguro.
- [x] Abrir DOCX de amostra no Microsoft Word se a automação local permitir;
  registrar explicitamente qualquer limite de validação visual/Windows.
- [x] Registrar evidência em `docs/validation/2026-09-28-formatacao-fiel.md`.

## Encerramento em 29/09/2026

87 testes passaram, inclusive 11 testes gráficos; build local macOS passou.
Revisão e limites registrados em [validação](../../validation/2026-09-28-formatacao-fiel.md).
A inspeção visual por screenshot não foi possível; testes de geometria e
interação foram executados. Atalhos e foco de teclado foram acrescentados após
a revisão final. Campos numéricos ficam editáveis em modos inativos para não
prender valores inválidos. Alterações permanecem no checkout da branch local,
sem publicação remota.
