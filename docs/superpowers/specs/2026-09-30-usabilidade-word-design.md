# Usabilidade e linguagem familiar ao Word

Direção aprovada na conversa: configuração guiada com opções básicas primeiro,
exemplos e opções avançadas recolhidas; usuário pediu expressamente pensar a
usabilidade e usar os termos do próprio Word para evitar confusão.

## Escopo

- Organizar o perfil em seções navegáveis: Fonte, Parágrafo, Página,
  Cabeçalho e rodapé, Imagens e Opções avançadas. Começar em Fonte, mostrar
  indicação de etapa e permitir avançar/voltar ou abrir qualquer seção.
- Evidenciar importar um Word já formatado como ponto de partida, instruir
  conferir e salvar o perfil. Manter criação manual, duplicação e recuperação.
- Padronizar rótulos com Word PT-BR: Centralizado, Cor da fonte, Espaçamento
  entre linhas, Antes/Depois, Exatamente/Pelo menos/Múltiplo, Recuo à esquerda
  e à direita, Especial/Por, Manter com o próximo e Controle de viúvas/órfãs.
- Dar atalhos Simples/1,5 linhas/Duplo; informar unidade de Em conforme o modo.
- Recuo especial será escolhido como Nenhum, Primeira linha ou Deslocado, com
  magnitude positiva em Por. A representação interna assinada permanece.
- Oferecer seletor de cor e opção de manter cor original. Código hexadecimal
  permanece disponível como alternativa, sem obrigar conhecimento técnico.
- Usar nomes claros para opções do aplicativo: mesma formatação para todo o
  texto ou por tipo de texto; distinguir perfil salvo de estilos nativos Word.
- Configurações de paginação menos frequentes e opções de saída ficam em
  Opções avançadas. Erros abrem automaticamente a seção correspondente.
- Resumos, revisão de importação e regras por tipo usam os mesmos termos.

## Decisões

Manter componentes e cores existentes: superfície branca, fundo #f3f6fb,
texto #101828, ajuda #475467, ação #2563eb, erro #b42318; espaços 8/16 px.
Navegação por teclado, foco visível, nomes completos e rolagem independente.
Validar desktop 760×600, 1000×760 e janela ampla; não prometer interface móvel.

Uma troca somente de rótulos deixaria o formulário extenso. Um assistente
obrigatório tornaria a edição recorrente lenta. A solução é navegação guiada
com acesso direto a cada seção. Exemplos são explicativos, sem prometer prévia
fiel de paginação do Word.

## Compatibilidade e validação

Perfis salvos e valores numéricos mantêm esquema e precisão; mudança apenas
de apresentação e navegação. Entrada inválida continua editável e não é
substituída silenciosamente. Testar conversão recuo especial, atualização de
valores externos, atalhos entrelinhas, navegação para erro, estados durante
lote, revisão de categorias, importação e persistência; revisão independente,
suíte real Tk e novo pacote com self-test.

Referências:
- https://support.microsoft.com/pt-br/word/adjust-indents-and-spacing
- https://support.microsoft.com/pt-br/word/indent-the-second-line-in-word
- https://support.microsoft.com/pt-br/word/line-and-page-breaks
