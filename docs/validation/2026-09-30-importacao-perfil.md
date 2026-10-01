# Validação — seletores e importação de perfil Word

## Implementação

- Correção dos seletores CustomTkinter: atualização do estado dos controles
  agendada após a escrita da seleção. O callback real do menu não deixa mais
  o campo vazio na primeira escolha. Campos dependentes, bloqueio durante lote
  e destruição com atualização pendente estão cobertos.
- Novo botão **Importar de Word**, leitura em segundo plano e revisão modal
  com opções por seção e categoria, resumo, exemplos e avisos.
- Confirmação cria um rascunho com nome sem colisão. Salvar continua explícito.
  Cancelar, erro de leitura e cancelamento durante processamento preservam o
  perfil anterior. Alterações não salvas usam a confirmação já existente.
- Leitor resolve propriedades de texto/página suportadas, formatação direta,
  herança, padrões e temas, incluindo mapeamento de cores. Alternativas
  incompatíveis com o perfil são bloqueadas sem ajuste silencioso dos valores.

## Evidência executada

Rodada final em macOS arm64, Python 3.14.6: **149 testes, sem falhas nem skips,
em 30,327 segundos**, incluindo **26 testes de interface Tk real** e **18 testes
dedicados à importação**.

```sh
FORMATWORD_GUI_TESTS=1 .venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m compileall -q app tests
git diff --check
```

- [Log completo](2026-09-30-importacao-tests.txt).
- Compilação Python e verificação de whitespace passaram.
- Regressão dos seletores reproduzida antes da correção: 51 subcasos falharam
  por valor vazio ou estado dependente incorreto; depois passaram.
- Testes de importação cobrem estilos herdados, temas, ênfase mista, opções
  coerentes, seções divergentes, paisagem, validação, arquivos inválidos,
  macros, entrada modificada e aplicação em outro DOCX. Fonte permanece intacta.
- Testes reais da UI cobrem cancelamento, fechamento durante leitura,
  responsividade, salvar/reabrir, nomes repetidos, mudanças de alternativas e
  confirmação bloqueada para geometria inválida.

## Revisão independente

Revisão de conformidade seguida de qualidade por subagente, conforme a skill
local requesting-code-review. UI aprovada; verificação adicional de layout
em 560×500, 740×600 e 1000×760, teclado Escape e processamento fora da thread Tk.

O leitor teve dois achados importantes, reproduzidos antes da correção:

1. Mapeamento de cores do documento não era aplicado ao tema.
2. Cor RGB armazenada recebia tint novamente quando o tema não era resolvido.

Ambos corrigidos com regressões. Nova revisão independente confirmou os
18 testes e os slots de texto/fundo, com parecer final: **nenhum achado crítico
ou importante pendente**. Conferência adicional confirmou editar o corpo após
importar e aplicar página personalizada em paisagem.

## Pacote local

Build PyInstaller macOS concluído, com CustomTkinter e PIL.ImageTk incluídos.
O self-test do aplicativo verifica importação, aplicação em outro documento,
seletores na primeira escolha, revisão e salvamento usando dados temporários.
Executado no binário final com código de saída **0** e **ok: true** no
[relatório do executável](2026-09-30-importacao-frozen-self-test.json).

Artefatos atualizados: `dist/FormatWord.app`, `dist/FormatWord-macOS.zip` e
`dist/Guia-de-uso.md`. Alterações mantidas na branch `codex/formatacao-fiel`
do workspace compartilhado, sem publicar ou fazer merge.

```sh
.venv/bin/python -m PyInstaller --noconfirm --clean --windowed --onedir \
  --name FormatWord --hidden-import PIL.ImageTk --collect-data customtkinter \
  --workpath /private/tmp/formatword-build-20260930 \
  --specpath /private/tmp/formatword-spec-20260930 main.py
dist/FormatWord.app/Contents/MacOS/FormatWord --self-test \
  docs/validation/2026-09-30-importacao-frozen-self-test.json
```

## Limites

- Referências e destinos de teste são sintéticos; não foram fornecidos modelos
  reais para homologação visual no Word da equipe.
- Perfil representa uma página global e um padrão por categoria. Avisos
  identificam alternativas, padrões não resolvidos e propriedades não suportadas.
- Cabeçalhos, rodapés, timbres, imagens, numerações e estruturas de tabelas da
  referência não são copiados. Preservar mantém os elementos do destino.
- Categorias dependem dos estilos e do contexto; conteúdo jurídico não é
  classificado por inferência textual. Propriedades condicionais de tabelas e
  fontes distintas por sistema de escrita não são reproduzidas separadamente.
- Conversão de tons de tema pode diferir do Word por pequeno arredondamento.
- Windows não foi executado neste macOS. O self-test ampliado será utilizado
  também pelo script/workflow Windows já existente.
- Callback global antigo do CustomTkinter em destruição programática foi
  reproduzido pela revisão sem importação; não atribuível a estas mudanças.

Fontes técnicas consultadas para as cores:
[mapeamento Microsoft](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.colorschememapping?view=openxml-3.0.1),
[cor armazenada e tema python-docx](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/text/font-color.html),
[tons de tema no Word](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/8229a077-7fc8-4fba-96cc-c77b6a4fc768).
