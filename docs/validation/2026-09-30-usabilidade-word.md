# Validação — usabilidade e termos do Word

## Entrega

- Editor de perfil dividido em Fonte, Parágrafo, Página, Cabeçalho e rodapé,
  Imagens e Opções avançadas, com acesso direto e Anterior/Próximo.
- Orientação inicial para importar ou criar, conferir e salvar. Salvar perfil
  permanece disponível abaixo do formulário.
- Termos padronizados: Cor da fonte, Centralizado, Espaçamento entre linhas,
  Antes/Depois, Em, Exatamente, Pelo menos, Manter com o próximo e Controle
  de viúvas/órfãs. Resumos e importação usam a mesma linguagem.
- Atalhos Simples, 1,5 linhas e Duplo; recuo Especial (Nenhum, Primeira linha,
  Deslocado) e Por em centímetros positivos; seletor visual de cor.
- Controles compartilhados com a edição por tipo de texto. Valores internos,
  IDs e esquema dos perfis existentes preservados, sem migração destrutiva.
- Erros ao salvar revelam a seção e o campo. Controles que recebem foco rolam
  para dentro da área visível, inclusive botões de cor e espaçamento.

## Testes

Rodada final em macOS arm64/Python 3.14.6: **163 testes, sem falhas nem skips,
em 20,042 segundos**. Inclui 31 testes de integração da interface, 5 testes
do controle de recuo especial, 16 testes do formulário e 18 da importação.

```sh
FORMATWORD_GUI_TESTS=1 .venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m compileall -q app tests
git diff --check
```

[Registro completo](2026-09-30-usabilidade-tests.txt). Compilação Python e
verificação de whitespace passaram.

As regressões verificam precisão, seleção nos dois sentidos, valor externo,
zero/positivo/negativo, inválidos, preservação de rascunho, atualização do modo
de espaçamento, cancelamento do seletor de cor, persistência por categoria,
navegação ao erro e teclado na janela mínima 760×600. A suíte também verifica
768/1024/1440 px, fluxos de importação e proteção dos documentos originais.

## Revisão independente

Revisão de conformidade e qualidade executada por subagente conforme a skill
local requesting-code-review. Foi identificado um P2: na janela mínima,
botões de cor e espaçamento podiam receber foco parcialmente fora do viewport.
Regressão confirmou 113 px para um viewport de 97 px antes da correção.

A correção calcula a posição do controle efetivamente focado e rola apenas o
necessário para mostrá-lo. O teste passou após a correção. Controles compostos
e botões adicionais dos diálogos também receberam tratamento de foco.
Nova revisão independente no Tk real reconferiu 760×600 e concluiu:
**nenhum achado crítico ou importante pendente**.

A tentativa de screenshot da janela isolada falhou no screencapture deste
ambiente; a verificação de layout usa coordenadas e dimensões reais do Tk.

## Empacotamento

Build PyInstaller macOS com CustomTkinter e PIL.ImageTk incluídos. Self-test
ampliado verifica navegação até Parágrafo, controle Deslocado, seletores,
importação, aplicação em outro Word e salvamento com configurações temporárias.
Executado no binário final com código de saída **0** e **ok: true** no
[relatório do executável](2026-09-30-usabilidade-frozen-self-test.json).

```sh
.venv/bin/python -m PyInstaller --noconfirm --clean --windowed --onedir \
  --name FormatWord --hidden-import PIL.ImageTk --collect-data customtkinter \
  --workpath /private/tmp/formatword-build-usability \
  --specpath /private/tmp/formatword-spec-usability main.py
dist/FormatWord.app/Contents/MacOS/FormatWord --self-test \
  docs/validation/2026-09-30-usabilidade-frozen-self-test.json
```

Artefatos locais: `dist/FormatWord.app`, `dist/FormatWord-macOS.zip` e
`dist/Guia-de-uso.md`. Trabalho preservado na branch `codex/formatacao-fiel`,
sem publicar nem fazer merge.

## Limites e referências

Validação técnica e de navegação, sem teste observado com usuários finais.
O aplicativo usa termos familiares ao Word, sem reproduzir seu editor completo
de estilos ou sua paginação. Configuração de página personalizada mantém o
modelo existente de lados menor/maior e orientação. Windows não foi executado
neste macOS; seu workflow continua utilizando o self-test ampliado.

Referências oficiais para a nomenclatura:
[recuos e espaçamento](https://support.microsoft.com/pt-br/word/adjust-indents-and-spacing),
[recuo especial](https://support.microsoft.com/pt-br/word/indent-the-second-line-in-word),
[quebras de linha e página](https://support.microsoft.com/pt-br/word/line-and-page-breaks).
