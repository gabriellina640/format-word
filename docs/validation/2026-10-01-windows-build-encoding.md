# Falha do build Windows — diagnóstico e validação local

Execução analisada: [Build Windows EXE — 36888169891](https://github.com/gabriellina640/format-word/actions/runs/36888169891),
commit `ffbb757e332b33a2062427b8a63c07261aeee7fc`.

- Windows Python 3.12 (job 110456729723) e 3.14 (job 110456730182): falha em
  `test_reports_success_and_cleans_temporary_documents`, linha 15.
- Ambos executaram 163 testes: uma falha, 36 testes gráficos ignorados.
- Ambos os jobs Linux, incluindo testes gráficos, passaram.
- `Build FormatadorDocumentos.exe` foi ignorado devido à dependência dos testes;
  portanto os logs não indicam falha do PyInstaller nessa execução.

## Causa e correção

O autoteste grava o relatório JSON explicitamente em UTF-8 com caracteres
acentuados. O teste lia o arquivo sem informar encoding. Sob o padrão CP1252 do
Windows, o JSON ainda é válido, mas `Importação` é decodificado incorretamente e
a asserção falha. O código de saída do autoteste já era zero.

As duas leituras em `tests/test_selftest.py` agora declaram `encoding='utf-8'`.
Mantidas as verificações de sucesso, importação, limpeza, falha e mensagem de
erro. Não foram alterados aplicativo, formato do relatório ou workflow.

## Verificações executadas

Reprodução com os testes existentes e `Path.open` configurado para CP1252
somente quando a codificação é omitida, sem mudar a configuração do sistema:

```python
original_open = Path.open
def windows_open(self, mode='r', buffering=-1, encoding=None, errors=None, newline=None):
    if 'b' not in mode and encoding in (None, 'locale'):
        encoding = 'cp1252'
    return original_open(self, mode, buffering, encoding, errors, newline)

suite = unittest.defaultTestLoader.discover('tests', pattern='test_selftest.py')
with patch.object(Path, 'open', windows_open):
    result = unittest.TextTestRunner(verbosity=2).run(suite)
```

- Antes: 2 testes, exatamente a mesma falha no assert de importação.
- Depois: 2 testes passaram sob a mesma simulação CP1252.
- `.venv/bin/python -m unittest discover -s tests -v`: 163 descobertos,
  127 passaram, 36 gráficos ignorados, nenhuma falha (2,195 s).
- `git diff --check`: sem erros.

A validação local foi feita no macOS. A correção precisa ser enviada ao GitHub
e o workflow executado novamente para comprovar a geração e o autoteste nativos
do EXE; este registro não declara que o build remoto já passou.

Revisão independente por `/root/review_ui`: sem achados; confirmado que as
asserções permanecem intactas e os dois testes passam com padrão CP1252 simulado.
