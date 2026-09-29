# Format Word: formatação fiel e aplicação em lote

Status: aprovado em 28/09/2026; implementado e validado localmente em 29/09/2026.
Evidências e limites: [relatório de validação](../../validation/2026-09-28-formatacao-fiel.md).

## Objetivo

Permitir criar, salvar e reutilizar perfis de formatação em um ou mais arquivos
`.docx`, com correspondência verificável entre os valores informados e as
propriedades gravadas no Word. Preservar o conteúdo dos arquivos de entrada.

O pedido do usuário autoriza corrigir bugs e retirar recursos que não funcionem.
O AGENTS.md exige aprovação do design antes da implementação. Este documento
define as decisões concretas para essa aprovação.

## Diagnóstico executado em 28/09/2026

Aplicação desktop Python/CustomTkinter, com persistência JSON e python-docx.
O fluxo atual transforma o arquivo de entrada em uma lista de textos e depois
constrói um documento novo. A interface também usa esse fluxo na exportação.

### Falhas reproduzidas

| Caso | Resultado observado | Causa |
| --- | --- | --- |
| Documento com uma tabela entre parágrafos | Saída sem tabelas; conteúdo da tabela deslocado para o final | `app/formatter.py`, `_read_docx_paragraphs` e `format_paragraphs` |
| Documento com imagem, título, negrito e duas seções | Saída sem imagem, título convertido para Normal, negrito perdido e apenas uma seção | Reconstrução do documento a partir de texto simples |
| Template com cabeçalho/rodapé e ambas as opções desativadas | Cabeçalho e rodapé continuam presentes | `_apply_header_footer` retorna antes de consultar as opções |
| Template inexistente, com imagem de cabeçalho válida | Exportação informa sucesso sem aplicar a imagem | Fallback em `_create_output_document` combinado com retorno em `_apply_header_footer` |
| Tamanho de fonte `36` | Valor aplicado: `32` | Limitação silenciosa no parser da interface |
| Tamanho de fonte `12,5` | Valor aplicado: `12` | Conversão de ponto flutuante para inteiro |
| Tamanho de fonte `abc` | Valor aplicado: `12` | Substituição silenciosa por padrão |
| Arquivo de configuração contendo `[]` | `AttributeError` ao carregar | Raiz JSON não validada |
| Configuração com `line_spacing: "invalid"` | Valor inválido carregado | Validação filtra nomes, mas não tipos/valores |
| Arquivo de texto renomeado para `.docx`, importado como template | Aceito | Importação verifica extensão e tamanho, sem validar o documento |

Os experimentos usaram documentos sintéticos em diretório temporário, reabertura
dos arquivos gerados com python-docx e inspeção das propriedades. O hash do
arquivo de entrada permaneceu igual no experimento de preservação.

### Testes existentes

Comando executado:

```sh
.venv/bin/python -m unittest discover -s tests -v
```

Resultado: 16 entradas executadas pelo runner, 7 sucessos, 1 falha e 8 erros.
Duas dessas entradas são erros de importação de módulos de teste inteiros;
portanto, esse número não representa a execução de todos os casos existentes.

- `format_documents_batch` e `unique_stack_name` são importados pelos testes,
  mas não existem na aplicação.
- Há testes para importação de configurações de template e pacotes de perfis
  cujas funções não estão implementadas.
- Alguns testes antigos esperam que as margens do template substituam o perfil.
  Essa expectativa conflita com a exigência atual de fidelidade à configuração.
- `tests/` está ignorado pelo Git; o workflow Windows não executa testes.

### Constatações por leitura, ainda sem validação visual

- A prévia desenha apenas dez linhas, com dimensões e quebras aproximadas;
  não representa margens, paginação, tabelas ou justificação do documento real.
- Não há seleção múltipla nem fluxo de processamento em lote.
- Fonte personalizada é substituída por Arial ao recarregar o formulário.
- Cabeçalho/rodapé têm estados duplicados em duas abas, podendo divergir.
- A exportação acontece na thread da interface, podendo bloquear a janela.
- A janela de prévia tem canvas fixo e pode ocultar o editor em telas menores.

Microsoft Word foi localizado em `/Applications/Microsoft Word.app`.
Abertura/renderização no Word e interação gráfica ainda não foram testadas.

## Alternativas

1. **Recomendação: preservar o DOCX e aplicar propriedades sobre ele.** Mantém
   tabelas, imagens, links, listas e seções; exige substituir o fluxo de texto
   simples e alinhar configuração, perfis e exportação.
2. Manter o conversor de texto com pequenos ajustes. Menor alteração, porém
   continua perdendo estrutura e não atende ao objetivo do usuário.
3. Automatizar diretamente o Microsoft Word como motor obrigatório. Permite
   usar mais recursos nativos, porém exige Word instalado e uma implementação
   específica por sistema. Não é a base recomendada para este aplicativo.

## Comportamento proposto

### Documento e escopo das opções

- Abrir o DOCX original e salvar uma cópia formatada, sem reconstruir o corpo
  como texto. Preservar ordem, tabelas inclusive aninhadas, imagens, hyperlinks,
  listas, quebras, seções e conteúdo não alterado pelo perfil.
- Aplicar as opções de texto em parágrafos do corpo e células das tabelas.
  Cabeçalhos/rodapés têm regras próprias. Recursos especiais que não puderem ser
  formatados, como determinados objetos ou caixas de texto, devem ser detectados
  e informados; não apresentar esses trechos como formatados sem verificação.
- Não apagar nem regenerar campos, sumários, comentários ou revisões para
  conseguir alterar a formatação. Casos incompatíveis devem produzir indicação
  explícita e não um sucesso sem ressalvas.
- Fonte, tamanho com precisão compatível com Word, alinhamento (esquerda,
  centro, direita, justificado), espaçamento antes/depois, entrelinhas
  (múltiplo, exato, mínimo), recuos esquerdo/direito e primeira linha/deslocado.
- Margens das quatro bordas, papel A4/Carta ou dimensões personalizadas,
  orientação e distâncias de cabeçalho/rodapé. Aplicar nas seções abrangidas
  pelo perfil, com escopo global explícito na interface.
- Negrito, itálico, sublinhado e cor: preservar por padrão; quando houver
  sobrescrita explícita no perfil, aplicar e verificar a opção selecionada.
- Controles de manter com próximo, manter linhas juntas e viúvas/órfãs com
  estado explícito de preservar/ativar/desativar.
- Cada campo deve informar a unidade. Valores não finitos, fora da faixa ou
  incompatíveis com a página devem bloquear a operação e explicar o motivo.
  Não truncar, limitar ou substituir silenciosamente a escolha do usuário.
- Documentar a precisão de armazenamento do Word; diferenças apenas de
  arredondamento da unidade nativa serão verificadas com tolerância definida.

### Perfis

- Criar, selecionar, editar, duplicar e excluir perfis nomeados.
- Uma única representação das configurações entre formulário, perfil salvo,
  resumo e execução. Diferenciar alterações não salvas do perfil persistido.
- Preservar fontes personalizadas e valores decimais ao salvar/reabrir.
- Validar os perfis na leitura e na gravação. Configuração danificada deve ser
  recuperável com aviso, sem apagar silenciosamente os dados existentes.
- Migrar os campos antigos que tenham equivalência. Opções antigas removidas
  ou ambíguas devem exigir revisão do perfil, sem serem ignoradas.
- Importação/exportação portátil de perfis fica fora desta correção inicial;
  os testes locais referentes a APIs nunca implementadas serão identificados
  e substituídos por testes dos requisitos aprovados, sem ocultar regressões.

### Cabeçalho, rodapé e templates

- Modos explícitos e independentes: preservar o original, remover, aplicar
  imagem do perfil. Validar existência da imagem, proporção, tamanho e posição.
- Evitar os atuais deslocamentos com limites silenciosos. Expor distâncias e
  dimensões reais com unidades; rejeitar combinações inválidas.
- Retirar inicialmente a aplicação de cabeçalho/rodapé de template Word
  arbitrário: ela é incompatível com a preservação do documento no motor atual
  e exige transplante seguro de partes, relacionamentos e variações de seção.
  Os arquivos de template já salvos não serão apagados. Perfis que os utilizem
  devem apresentar aviso e pedir a escolha de um modo suportado antes de aplicar.
- Cabeçalhos/rodapés que já pertencem ao arquivo de entrada permanecem
  preservados no modo padrão, incluindo variações por página e seções.

### Aplicação em um ou mais arquivos

- Selecionar vários DOCX, mostrar a lista e permitir remover itens.
- Escolher perfil e destino; apresentar resumo dos valores efetivamente usados.
- Processar em segundo plano usando uma cópia imutável da configuração da
  execução. A interface recebe progresso e resultados por uma fila segura.
- Resultado individual por arquivo, com caminho gerado ou motivo da falha.
  Um arquivo inválido não cancela os demais. Impedir execuções duplicadas.
- Permitir cancelar entre arquivos e reportar processados/pendentes.
- Nunca sobrescrever a entrada nem saídas existentes. Publicar a saída completa
  de forma segura; evitar arquivos finais parciais quando houver falha.
- Oferecer abrir documento gerado e abrir pasta de saída.

### Simplificação da interface

- Fluxo: selecionar arquivos → escolher/editar perfil → conferir resumo → aplicar.
- Retirar o editor de texto simples e o desenho que hoje é apresentado como
  prévia do Word. Conferência visual será feita abrindo o DOCX gerado no editor
  instalado, sem prometer uma renderização equivalente dentro de um canvas.
- Retirar importação PDF deste fluxo, pois extração de texto não preserva a
  estrutura de um Word. Aceitar somente DOCX e explicar extensões rejeitadas.
- Organizar configurações por Texto, Parágrafo, Página e Cabeçalho/Rodapé,
  com rolagem, navegação por teclado e mensagens junto ao campo inválido.
- Testar uso em janela reduzida e escala de tela; manter ações essenciais
  acessíveis e evitar confirmação modal para cada ação rotineira.

## Estrutura da implementação após aprovação

- `app/config.py`: modelo, validação e persistência/migração de perfis.
- `app/formatter.py`: aplicação em documento preservado e gravação segura.
- Módulo dedicado ao lote se necessário: cancelamento e resultados por arquivo.
- `app/ui.py`: seleção múltipla, editor de perfil, resumo e progresso.
- `tests/`: regressões reais, fidelidade de cada campo e testes de integração.
- `.gitignore` e CI: versionar e executar testes antes de gerar executável.
- `README.md`: escopo suportado, operação e limites confirmados.

Antes de editar esses arquivos será escrito o plano executável em tarefas
atômicas, cada uma com arquivos-alvo e critérios de validação.

## Critérios de aceitação e validação

1. Testes que reproduzam os bugs devem falhar antes da correção correspondente.
2. Cada opção exposta terá teste de aplicação com reabertura do DOCX gerado e
   inspeção das propriedades/XML, inclusive valores que antes eram truncados.
3. Documentos sintéticos devem cobrir tabelas aninhadas/mescladas, hyperlinks,
   imagens, títulos, listas, múltiplas seções e variações de cabeçalho/rodapé.
4. Comparar conteúdo/ordem/relacionamentos relevantes antes e depois. Verificar
   hash dos originais e ausência de sobrescrita de saídas anteriores.
5. Testar perfis após reinício, migração, números com vírgula, campos vazios,
   tipos inválidos, JSON corrompido e caminhos de recursos inexistentes.
6. Testar lote com arquivos válidos e inválidos, nomes repetidos, destino
   inválido, falha de gravação, cancelamento e tentativa de execução duplicada.
7. Exercitar a interface real: criar perfil, salvar, reabrir, selecionar lote,
   aplicar, conferir resultados, redimensionar e fechar durante processamento.
8. Abrir amostras no Microsoft Word instalado e registrar o que foi inspecionado;
   se a automação gráfica estiver indisponível, registrar essa limitação.
9. Fazer code review com achados por severidade, corrigir os impeditivos e
   executar a verificação final sobre o estado final do código.
10. Registrar comandos, resultados e limitações. O executável Windows exige
    validação no Windows/CI; testes locais no macOS não comprovam esse build.

O compromisso é fidelidade das opções suportadas e resultados verificados.
Não declarar compatibilidade com todos os recursos do Word nem perfeição
universal sem evidência. Renderização depende também das fontes instaladas e
da versão do editor; isso não permite substituir valores silenciosamente.
