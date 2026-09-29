# Validação da formatação fiel

Implementação aprovada em 28/09/2026 e verificação final realizada em
29/09/2026 na branch `codex/formatacao-fiel`.

## Resultado final

**87 testes passaram, sem falhas e sem skips**, em 15,428 s, incluindo interação
com janelas Tk reais. `compileall` e `git diff --check` também passaram.
Empacotamento local com PyInstaller concluído para macOS.

Ambiente: macOS 26.5 arm64, Python 3.14.6, python-docx 1.2.0,
Pillow 12.2.0, CustomTkinter 5.2.2 e PyInstaller 6.20.0.

| Área | Testes | Evidência principal |
| --- | ---: | --- |
| Configuração e perfis | 25 | Validação estrita, migração, backup, round-trip, imagens |
| Motor DOCX | 34 | Preservação do corpo, propriedades reabertas, cabeçalhos, saída segura |
| Lote | 7 | Falha individual, cancelamento, snapshot e nomes únicos |
| Valores e resumo da interface | 10 | Vírgula/decimais, todos os campos, precisão e opções efetivas |
| Interface real | 11 | CRUD, persistência, erros, lote, fechamento, tamanho e teclado |
| **Total** | **87** | **OK** |

Registro completo: [2026-09-29-tests.txt](2026-09-29-tests.txt).

Comandos executados:

```sh
FORMATWORD_GUI_TESTS=1 .venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m compileall -q app main.py tests
git diff --check
```

Tk precisou ser executado fora do sandbox, com autorização da ferramenta.
Os testes usam pastas temporárias; não alteram os perfis reais do usuário.
A descoberta sem `FORMATWORD_GUI_TESTS=1` pula somente os 11 testes gráficos.

## O que foi verificado

- Fonte, tamanho de meio ponto, fontes temáticas e fontes de caracteres complexos;
  alinhamentos, três modos de entrelinhas, espaços antes/depois, recuos, ênfase,
  cor, flags de paginação, margens e papel em todas as seções.
- Preservação de tabelas aninhadas/mescladas, hyperlinks, imagens, negrito
  original, títulos, listas, ordem, espaços, parágrafos vazios e quebras.
- Marcas de parágrafo recebem a fonte configurada, inclusive em parágrafos vazios.
- Cabeçalhos/rodapés e variantes preservados, removidos ou substituídos;
  imagens sem deformação por DPI e com geometria validada.
- Arquivo original inalterado; exportações simultâneas usam nomes diferentes;
  falhas de gravação não deixam arquivo final incompleto. Sistemas sem hardlinks
  usam criação exclusiva e limpeza de cópia malsucedida.
- A saída temporária é reaberta e suas propriedades são verificadas antes de
  ser publicada. Uma gravação que não contém a formatação solicitada é rejeitada.
- Configuração corrompida, raiz JSON inválida, tipos incorretos, números não
  finitos, imagens ausentes e perfis legados têm tratamento explícito.
- Lote continua após DOCX inválido; cancelamento preserva resultados concluídos;
  fechamento aguarda o arquivo em processamento.
- Interface em 768×600, 768×760, 1024×760 e 1440×760: lista utilizável e ações
  essenciais dentro da janela. Formulário e resumo têm rolagem.
- Tab/Enter/Espaço e atalhos: foco visível; Ctrl+Enter não dispara também o botão
  focado; campos inválidos continuam acessíveis para correção.

## Revisão de código

Revisões independentes de configuração, motor/lote e interface foram realizadas
por subagentes conforme as skills locais. Os achados importantes foram corrigidos
com regressões e reconferidos:

1. Um perfil legado bloqueava a gravação de todos os outros: validação de
   persistência foi separada da validação de execução.
2. JPEG importado podia crescer como PNG e se tornar inutilizável: limite de
   entrada separado do limite do recurso normalizado, mantendo limite de pixels.
3. Cabeçalho por imagem herdava marca de parágrafo com fonte gigante: marca de
   1 pt e espaçamentos explícitos.
4. Caminho de imagem com `~` passava na validação e falhava no motor: expansão
   aplicada também na abertura da imagem.
5. Campos numéricos inválidos ficavam desativados ao trocar modo: continuam
   acessíveis, sem substituir os valores digitados.
6. Perfil legado chamado “Configuração atual” ficava inacessível: migração com
   nome único, mantendo dados e seleção ativa.
7. Ctrl/Command+Enter podia invocar o botão em foco e iniciar o lote: bindings
   específicos interrompem a propagação; regressão gráfica confirmou a correção.

A revisão local também corrigiu ordem dos elementos OOXML, formatação da marca
no corpo, arredondamento do formulário, conteúdo do resumo e recursos não
suportados que antes poderiam resultar em sucesso indevido.

## Empacotamento

Build local concluído com `--onedir --windowed --collect-data customtkinter`,
PyInstaller 6.20.0 e ícone incluído com caminho absoluto no comando de validação.
O primeiro comando de teste falhou porque `--specpath` temporário mudou a base
para resolver `icone.ico`; o comando foi corrigido e o build final passou.

Artefato de validação local:
`/tmp/formatword-validation-20260928/dist/FormatWordValidation.app` (~49 MB).
Esse bundle é uma prova de empacotamento no macOS; não é instalador Windows nem
uma distribuição assinada/notarizada para terceiros. A execução da interface
foi validada pelo Python do projeto, não pelo bundle congelado.

O script PowerShell agora interrompe o build se instalação, testes ou
PyInstaller falharem. O workflow testa Windows/Linux, Python 3.12/3.14 e interface
em Xvfb antes de gerar o executável Windows. **Esses jobs remotos e o executável
Windows não foram executados nesta máquina macOS.**

## Conferência no Word e limites

Microsoft Word 16.109 abriu uma amostra sintética. A leitura via AppleScript
confirmou Arial, 12,5 pt, espaço antes de 3 pt, uma tabela e duas seções.
Margens/dimensões retornaram `missing value` nessa API; essas propriedades foram
verificadas por reabertura do DOCX nos testes, não confirmadas visualmente no Word.

A exportação PDF e comandos de fechamento por referência de documento retornaram
`-1708` na automação do Word. A captura da região da janela de teste também
falhou (`could not create image from rect`). Portanto, **não se declara aprovação
visual por screenshots nem equivalência de renderização entre editores**.
O teste de geometria/interação da janela foi executado com sucesso.

O aplicativo não promete todos os recursos do Word. O README e os resultados
explicam os recursos preservados sem formatação, as incompatibilidades bloqueadas
e a dependência de fontes instaladas. A garantia testada é a aplicação das
propriedades suportadas e a preservação dos elementos cobertos pelos testes.

## Suíte antiga e mudanças de escopo

Baseline: 7 sucessos, 1 falha e 8 erros, incluindo dois módulos inteiros que não
carregavam. Os testes estavam ignorados pelo Git. A nova suíte é versionável.
Cópias dos arquivos antigos foram preservadas localmente em
`/tmp/formatword-legacy-tests-20260928` antes das alterações.

Expectativas antigas substituídas pelo design aprovado:

- Tabelas achatadas como texto → tabelas preservadas.
- Margens de template substituindo o perfil → propriedades explícitas do perfil.
- Números truncados/limitados → validação sem substituição silenciosa.
- Templates ausentes ignorados → perfil bloqueado para revisão.
- APIs de pacotes portáteis e importação de template nunca implementadas → fora
  do escopo aprovado; não há controles anunciando essas capacidades.

PDF, editor de texto destrutivo, prévia aproximada e aplicação de templates
externos foram retirados conforme o design aprovado. Os arquivos antigos de
perfil, imagem e template não foram excluídos.
