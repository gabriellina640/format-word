# Validação — perfis por promotor

## Resultado

Ampliação implementada na branch `codex/formatacao-fiel`. A rodada completa
passou: **121 testes, sem falhas e sem skips, em 22,099 segundos**, incluindo
**16 testes com janelas Tk reais**. Compilação Python e `git diff --check` passaram.
O ajuste final do smoke para verificar também a altura foi validado novamente
pelos seus dois testes e pelo executável recompilado.

- [Registro completo dos testes](2026-09-29-promotoria-tests.txt)
- [Relatório do executável macOS](2026-09-29-frozen-self-test.json)
- [Guia de uso](../guia-rapido.md)
- [Validação da etapa anterior](2026-09-28-formatacao-fiel.md)

Ambiente: macOS 26.5 arm64, Python 3.14.6, python-docx 1.2.0,
Pillow 12.2.0, CustomTkinter 5.2.2 e PyInstaller 6.20.0.

## Funcionalidades verificadas

- Regras distintas para corpo, títulos, subtítulos, citações, assinaturas,
  tabelas e caixas de texto. Preservar, herdar e personalizar; propriedades
  normalizadas, validação de geometria e persistência independente dos mapas.
- Identificação por estilos conhecidos e herdados, vínculos personalizados por
  nome/ID e classificação manual por documento. Sem inferir conteúdo jurídico.
- Revisões em texto, linhas e células de tabelas preservadas. Equações e objetos
  mantidos; modo estrito recusa conteúdo complexo. Caixas de texto podem receber
  regras próprias sem a fonte externa sobrescrever seus parágrafos internos.
- Inventário considera a imagem real, sem confundir sua caixa contenedora.
  Redimensionamento proporcional, alinhamento e movimento de imagens em linha;
  ordem de várias imagens conservada e opção preservar mantém alinhamento.
  Parágrafo próprio com recuos zerados evita ultrapassar margem pelo recuo do texto.
- Ajustes por documento vinculados ao hash do original. Arquivo alterado após
  revisão é bloqueado, sem impedir os demais arquivos do lote.
- Cópias profundas isolam regras, mapas e ajustes das alterações de interface
  ou callbacks durante processamento.
- Diálogos transacionais: cancelar não altera perfil/arquivo; confirmar marca
  alterações; salvar/restaurar e aplicar lote conservam as escolhas.
- Revisão de imagem inicia com a regra global; escolha explícita de preservar
  prevalece sobre a largura global do perfil.
- Toda saída é reaberta antes de ser publicada. O corpo serializado e as
  propriedades efetivas dos parágrafos/seções são verificados. Originais e
  resultados anteriores não são sobrescritos.

## Revisão de código

Revisão independente por subagente conforme a skill local, com conformidade de
escopo seguida de análise de qualidade. Achados corrigidos e reconferidos:

1. Revisões de linha/célula não protegiam todo seu conteúdo.
2. Uma caixa com imagem interna aparecia como imagem extra editável.
3. Herança do corpo ignorava sua regra personalizada/preservada.
4. Inserção repetida após o mesmo destino invertia imagens.
5. Mover com alinhamento preservar forçava alinhamento esquerdo.
6. Modo estrito deixava passar revisões de célula.
7. Largura de imagem ignorava o recuo de primeira linha.
8. Validação da interface misturava recuos de categorias independentes.
9. Preservar uma imagem eliminava o override e reaplicava largura global.

Também foi corrigida a consulta XML em elementos genéricos de caixas/âncoras,
que não recebem o resolvedor de namespaces de python-docx.
A última revisão reportou **nenhum achado bloqueante pendente**. Sugestão menor
de conferir a altura da imagem no smoke foi incorporada.

## Empacotamento executado

Build macOS `--onedir --windowed`, com dados CustomTkinter e PIL.ImageTk incluídos.
O binário congelado foi executado com `--self-test`, usando documentos e perfis
em pasta temporária. Confirmou categoria, fonte, largura/altura de imagem,
integridade do original, abertura da interface e gravação de perfil. Código de
saída 0 e relatório `ok: true`.

Artefatos locais:

- `dist/FormatWord.app`
- `dist/FormatWord-macOS.zip`
- `dist/Guia-de-uso.md`

O script Windows e o workflow incluem teste do executável gerado com limite de
120 segundos, código de saída e relatório JSON obrigatórios. O artifact remoto
só é publicado quando esse teste passa.

## Limites da evidência e entrega externa

**Não foi executado build/teste Windows nesta máquina macOS.** O workflow e o
script estão preparados, mas não se declara um EXE Windows já validado.
O pacote local é macOS arm64, sem assinatura de distribuição/notarização Apple.

Não foram fornecidos documentos reais da promotoria; as amostras dos testes
são sintéticas. A conferência no Word da etapa anterior foi parcial, descrita no
registro correspondente. Esta etapa não declara homologação visual da paginação.

O escopo suportado está no README: imagens flutuantes e em caixas são
preservadas, alterações controladas não são aceitas/rejeitadas, notas/comentários
mantêm seu conteúdo, e macros/assinaturas digitais/senha são bloqueados. Revisões
por arquivo duram a sessão; perfis e vínculos de estilos persistem ao salvar.
Para entrega operacional no Windows, executar o build lá e conferir documentos
representativos com as fontes e o Word usados pela equipe.
