# Leitura UTF-8 nos testes do build Windows

Objetivo: corrigir a falha observada no workflow 36888169891, mantendo as
verificações do relatório e a matriz Windows/Linux Python 3.12/3.14.

Design da correção: `app/selftest.py` já grava JSON em UTF-8. O consumidor
`tests/test_selftest.py` deve declarar a mesma codificação nas leituras, sem
depender da configuração regional do Windows. Não mudar o formato do relatório
nem remover a verificação de importação.

- [x] Ler jobs e logs: ambos os Windows falham no mesmo assert de `Importação`;
  Linux passa e build fica skipped. Reproduzir com `Path.open` usando CP1252
  somente quando a leitura omite uma codificação explícita.
- [x] Alterar as leituras em `tests/test_selftest.py` para
  `report.read_text(encoding='utf-8')`, reutilizando o JSON no teste de sucesso.
- [x] Reexecutar os dois testes sob CP1252 e a suíte completa
  `.venv/bin/python -m unittest discover -s tests -v`; revisar o diff e registrar
  a evidência em `docs/validation/2026-10-01-windows-build-encoding.md`.

Limite de verificação: o ambiente local é macOS. A geração nativa do EXE e seu
smoke test exigem uma nova execução do workflow após enviar a correção.
