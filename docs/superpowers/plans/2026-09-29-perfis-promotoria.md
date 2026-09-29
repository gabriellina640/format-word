# Perfis por promotor — plano de execução

**Objetivo:** resolver as limitações de formatação uniforme e ajustes de imagens,
mantendo o uso simples e a preservação dos documentos.

**Arquitetura:** regras validadas por categoria + inspeção compartilhada + ajustes
por arquivo com hash + edição/verificação do DOCX + interface de revisão.

**Execução:** testes antes das alterações críticas, etapas independentes com
revisão de código, sem reiniciar nem descartar as correções anteriores.

1. [x] `app/config.py`, `app/ui_fields.py`, `tests/test_config.py`,
   `tests/test_ui_values.py`: modo uniforme/por categoria; regras
   preservar/herdar/personalizar; mapa de estilos; imagens; conteúdo complexo.
   Validar round-trip, cópias profundas e rejeição de regras inválidas.
2. [x] `app/document_model.py`, `tests/test_document_model.py`: inspecionar
   parágrafos/imagens, classificar estilos e trechos protegidos; overrides
   associados ao hash. Testar documentos mistos, tabelas, caixas e revisões.
3. [x] `app/formatter.py`, `tests/test_formatter.py`: aplicar regras por categoria
   e verificar cada trecho; preservar conteúdo complexo com ressalvas; imagens
   com proporção e posição. Testar fidelidade e imutabilidade dos originais.
4. [x] `app/batch.py`, `tests/test_batch.py`: capturar regras e overrides por
   arquivo sem referências mutáveis; rejeitar revisão obsoleta sem prejudicar
   arquivos restantes.
5. [x] `app/ui.py`, novos módulos de diálogo e testes GUI: regras do promotor,
   revisão de tipos e imagens; persistência, estado não salvo, lote e teclado.
6. [x] `.github/workflows/build-windows-exe.yml`, scripts e documentação:
   teste automatizável do executável, fluxo de entrega e homologação simples.
7. [x] Revisão independente; suíte completa/GUI; build local; registro fiel das
   evidências e das validações que dependem de documentos/Windows externos.

Comandos de referência:

```sh
.venv/bin/python -m unittest discover -s tests -v
FORMATWORD_GUI_TESTS=1 .venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m compileall -q app main.py tests
git diff --check
```

Baseline da ampliação: 87 testes descobertos, 76 passaram e 11 gráficos pulados
na execução sem ambiente gráfico; rodada gráfica anterior de 87 testes passou.

## Encerramento

Implementação, revisão independente e validação local concluídas. 121 testes
passaram, incluindo 16 gráficos; build e execução do pacote macOS passaram.
Registro: `docs/validation/2026-09-29-perfis-promotoria.md`.

A preparação do build Windows está concluída com smoke obrigatório. Sua execução
no Windows e homologação com documentos reais da promotoria permanecem etapas
externas não realizadas nesta máquina; não são apresentadas como aprovadas.
