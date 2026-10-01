# Format Word

Aplicativo desktop para aplicar perfis de formatação a **um ou vários arquivos
Word `.docx`**, preservando o conteúdo e salvando cópias em uma pasta escolhida.

## Como usar

1. Clique em **Adicionar DOCX** e selecione os documentos.
2. Escolha um perfil no seletor superior. Em **Perfis e formatação**, ajuste os
   valores, informe um nome e clique em **Salvar**. **Duplicar** cria outro perfil
   sem alterar o anterior. Sem nome, você salva a **Configuração atual**.
3. Se necessário, abra **Tipos de texto…** para personalizar títulos, citações,
   assinaturas, tabelas e caixas de texto. Selecione um documento e clique em
   **Revisar documento** para classificar trechos ou ajustar imagens.
4. Escolha a pasta de destino e confira o resumo em **Arquivos e resultados**.
5. Clique em **Aplicar aos arquivos**. Cada arquivo recebe um resultado próprio;
   uma falha não interrompe os demais. **Cancelar lote** cancela os próximos
   arquivos e deixa terminar o que já está em andamento.
6. Selecione o resultado para ver detalhes e use **Abrir DOCX selecionado** ou
   **Abrir pasta de saída**. Confira documentos com ressalvas no Word.

Os originais e as saídas anteriores não são sobrescritos. Arquivos com nomes
iguais recebem um número adicional. A exportação reabre o DOCX gravado e verifica
as propriedades configuradas antes de publicar o resultado.

Guia para o dia a dia: [Perfis por promotor](docs/guia-rapido.md).

## Criar perfil a partir de um Word

Em **Perfis e formatação → Importar de Word**, selecione um `.docx` já formatado.
A leitura acontece em segundo plano. Na revisão, confira a página e o padrão
de cada tipo de texto; havendo diferenças, escolha entre as alternativas
detectadas. O resumo acompanha as escolhas. Leia também **Avisos e limites**.

**Usar estas configurações** preenche o formulário e as regras por tipo.
Revise o nome e clique em **Salvar** para reutilizar o perfil. Nomes repetidos
recebem um número. Cancelar a revisão mantém o perfil e as edições anteriores;
confirmar oferece salvar/descartar/cancelar caso existam alterações não salvas.
O arquivo de referência nunca é alterado.

A importação lê fonte, tamanho, ênfase, cor, alinhamento, espaçamentos, recuos,
controles de parágrafo, papel, orientação, margens e distâncias de cabeçalho e
rodapé. Considera formatação direta, herança de estilos, padrões e tema do DOCX.
Alternativas de texto são ordenadas por frequência; diferenças dentro de uma
categoria exigem escolher um padrão. Categorias ausentes preservam o destino.
Valores não resolvidos têm avisos e sugestões explícitas para revisão.

O perfil usa uma configuração de página para todas as seções do destino.
**Cabeçalhos, rodapés, timbres, imagens, numerações e estruturas de tabelas da
referência não são copiados.** Preservar original mantém esses elementos do
documento de destino. O importador não adivinha títulos ou citações digitados
como texto comum; use **Revisar documento** para classificá-los. A importação
de configurações não equivale à reprodução integral de um template Word.

## Configurações disponíveis

O editor organiza os campos em **Fonte**, **Parágrafo**, **Página**, **Cabeçalho
e rodapé**, **Imagens** e **Opções avançadas**. Abra a seção desejada ou use
**Anterior/Próximo**; **Salvar perfil** fica acessível no rodapé. Erros ao salvar
abrem a seção do campo correspondente.

Controles usam linguagem familiar ao Word: atalhos **Simples/1,5 linhas/Duplo**,
**Espaçamento entre linhas**, **Especial: Primeira linha/Deslocado** com medida
positiva em **Por**, e seletor visual de **Cor da fonte**. As mesmas opções
aparecem na personalização por tipo de texto. Perfis já salvos são compatíveis.

| Grupo | Opções |
| --- | --- |
| Texto | Fonte livre, tamanho de 1 a 400 pt em incrementos de 0,5, negrito, itálico, sublinhado e cor hexadecimal |
| Parágrafo | Esquerda, centro, direita ou justificado; espaçamento antes/depois em pt; entrelinhas múltiplo, exato ou mínimo; recuos esquerdo/direito e primeira linha/deslocado em cm |
| Paginação | Manter com próximo, manter linhas juntas e controle de viúvas/órfãs |
| Página | A4, Carta ou papel personalizado; retrato/paisagem; quatro margens; distâncias de cabeçalho/rodapé |
| Cabeçalho e rodapé | Preservar original, remover ou aplicar imagem PNG/JPEG do perfil; largura e alinhamento da imagem |
| Tipos de texto | Corpo, título, subtítulo, citação, assinatura, tabela e caixa de texto: seguir corpo, preservar ou personalizar |
| Imagens no corpo | Largura proporcional e alinhamento por perfil; ajustes individuais e movimento antes/depois de um parágrafo |
| Saída | Sufixo e limite de tamanho por arquivo |

- Aceita vírgula ou ponto decimal. Valores inválidos são destacados para
  correção; o aplicativo não limita nem substitui silenciosamente os valores.
- **Todo o texto** aplica a mesma fonte e parágrafo ao corpo e tabelas.
  **Por tipo de texto** usa regras separadas; novos perfis preservam os tipos especiais
  até você configurá-los. O tipo é identificado pelo estilo do Word, pelo contexto
  de tabela/caixa ou pela revisão manual. Texto em estilo Normal não é adivinhado.
  As opções de página valem para todas as seções sem revisões protegidas.
- Negrito, itálico, sublinhado, cor e controles de paginação podem ser preservados.
  As demais opções aplicam os valores mostrados no formulário.
- Dimensões personalizadas só são usadas com papel **Personalizado**. Largura e
  alinhamento de imagem só são usados no modo **Imagem do perfil**. Números
  continuam acessíveis para correção, mesmo quando o modo não os utiliza.
- No papel personalizado, informe os lados menor e maior; a orientação define
  qual é a largura. Para recuo deslocado, escolha **Deslocado** e uma medida
  positiva em **Por (cm)**.
- Cabeçalhos e rodapés originais preservam seu texto, imagens e variações de
  página; a distância à borda segue o perfil. No modo imagem, todas as variantes
  recebem a imagem, com proporção preservada. A altura proporcional, a distância
  e uma folga de 0,1 cm precisam caber na margem. Caso contrário, ajuste os valores.
- Imagens importadas: até 8 MB e 40 megapixels. São copiadas para o perfil e
  normalizadas sem recortar ou deformar; o arquivo interno pode ser maior que o
  original. Alterar a imagem de um perfil não substitui imagens de outros perfis.
- A fonte escolhida precisa estar instalada no computador que abre o documento.
  Word e outros editores podem substituir fontes ausentes.
- O Word armazena medidas de parágrafo/página em unidades de 1/20 pt e múltiplos
  de entrelinhas em 1/240. Os testes e a verificação toleram apenas o arredondamento
  dessas unidades; fonte é armazenada em meios pontos.

## Preservação e limites

Tabelas, células mescladas, imagens, hyperlinks, listas, parágrafos vazios,
quebras e seções permanecem no documento. A lista mantém sua numeração e seus
símbolos; os parágrafos recebem os recuos do perfil. Larguras de tabelas e
posições de imagens flutuantes são preservadas, então confira seu encaixe ao
reduzir a área útil da página.

Campos e sumários são preservados, com ressalva: atualizá-los no Word pode
recalcular texto/formatação. Notas e comentários também são preservados, mas seu
texto não recebe o perfil. Tais situações aparecem nos detalhes do resultado.

Por padrão, trechos com alterações controladas, equações e objetos incorporados
são preservados com ressalvas. O aplicativo não aceita/rejeita revisões nem
reconstrói objetos. Caixas de texto recebem somente a regra de texto escolhida,
conservando tamanho/posição: confira seu encaixe no Word. Seções com revisões
preservam também suas configurações de página; substituir cabeçalho/rodapé nesses
casos é bloqueado. A opção **Recusar documento** bloqueia conteúdo complexo.
Documentos com macros, assinaturas digitais ou protegidos por senha são recusados.

Ajustes de imagens aceitam imagens em linha fora de caixas de texto e de revisões.
Redimensionar, alinhar ou mover cria um parágrafo próprio, com recuos zerados, sem
alterar o texto adjacente. A largura deve caber na área útil e a altura proporcional
na página. Mover exige um parágrafo de referência do corpo, fora de tabelas/caixas.
Imagens flutuantes são preservadas; para ajustá-las, converta para **Em linha** no
Word. Ajustes individuais podem preservar uma imagem mesmo quando o perfil define
uma largura global. A revisão fica vinculada ao conteúdo do arquivo: modificações
posteriores exigem revisá-lo novamente. Esses ajustes por arquivo valem para a
sessão atual; regras do perfil e vínculos de estilos persistem ao salvar o perfil.

PDF e `.doc` antigo não fazem parte deste fluxo. Abra/converta para `.docx` no
Word. A antiga prévia aproximada e o editor de texto simples foram retirados:
eles descartavam a estrutura do arquivo. A conferência visual é feita no DOCX
real, usando o editor instalado.

Teclado: Tab percorre os controles; Enter ou Espaço aciona o botão em foco.
Ctrl+S salva o perfil, Ctrl+Enter aplica, Ctrl+1/Ctrl+2 troca as abas e Esc
cancela os arquivos pendentes. No macOS, os atalhos também aceitam Command.

## Perfis antigos e recuperação

As configurações ficam na pasta do usuário (`FormatWord` em Application Support
no macOS, AppData/Roaming no Windows ou XDG_CONFIG_HOME no Linux).

Fontes, valores decimais e perfis são mantidos entre sessões. As opções antigas
compatíveis são migradas. Templates externos e deslocamentos sem equivalência
ficam bloqueados até **Revisar perfil antigo**: a tela explica o que será
removido do perfil, preserva os arquivos externos e pede que você revise os
modos suportados antes de salvar. Cabeçalhos/rodapés já presentes nos arquivos
Word continuam suportados.

Configurações corrompidas são recuperadas com aviso e cópia
`settings.<identificador>.bak`. Perfis válidos não são descartados por causa de
outro perfil inválido. Um perfil antigo chamado “Configuração atual” é renomeado
sem sobrescrever outros perfis. Importação/exportação de pacotes portáteis e
aplicação de templates Word externos não estão disponíveis nesta versão.

## Rodar localmente

Python 3.12 ou posterior com Tk disponível.

macOS/Linux:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

Windows:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

Algumas distribuições Linux e instalações Homebrew oferecem Tk em um pacote
separado. O comando `python -m tkinter` permite confirmar se a janela de teste
abre no seu ambiente.

## Testes e executável

Testes do motor, configuração e formulário, sem precisar abrir uma janela:

```sh
python -m unittest discover -s tests -v
```

Testes de integração com uma sessão gráfica, usando arquivos e perfis temporários:

```sh
FORMATWORD_GUI_TESTS=1 python -m unittest discover -s tests -p test_ui_integration.py -v
```

No PowerShell, defina `$env:FORMATWORD_GUI_TESTS="1"` antes do comando Python.
Em Linux sem monitor, use `xvfb-run -a` antes de `python`.

No Windows, `build_windows.ps1` executa os testes e gera
`dist/FormatadorDocumentos.exe`, abre o executável com arquivos temporários e
exige um relatório de sucesso em `dist/self-test.json`. O workflow do GitHub testa Python 3.12/3.14 em
Windows/Linux e a interface em Xvfb antes de gerar o artifact
`FormatadorDocumentos-windows-exe`. O build Windows precisa ser executado no
Windows; validação local no macOS não comprova o executável Windows.

Os resultados locais, correções de revisão e limites de validação estão em
[docs/validation/2026-09-29-perfis-promotoria.md](docs/validation/2026-09-29-perfis-promotoria.md).
Para os seletores e a importação de perfis, veja
[docs/validation/2026-09-30-importacao-perfil.md](docs/validation/2026-09-30-importacao-perfil.md).
A reorganização e os termos familiares ao Word estão registrados em
[docs/validation/2026-09-30-usabilidade-word.md](docs/validation/2026-09-30-usabilidade-word.md).
