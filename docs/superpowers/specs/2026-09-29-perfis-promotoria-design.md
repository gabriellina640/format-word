# Perfis por promotor e revisão de documentos

Direção apresentada ao usuário: regras distintas para partes do documento,
reconhecimento dos estilos existentes e confirmação manual quando necessário.
Autorização de execução reafirmada em 29/09/2026: “capriche e me entregue pronto”.

## Decisões

- Manter o fluxo principal de três passos: promotor/perfil, documentos, aplicar.
- Adicionar modo de regras por tipo de texto, com corpo, título, subtítulo,
  citação, assinatura, tabela e caixa de texto. Cada tipo pode seguir o corpo,
  preservar o original ou usar propriedades próprias. O modo uniforme continua
  disponível para compatibilidade com perfis anteriores.
- Identificar tipos pelos estilos Word, com mapa personalizável por perfil.
  Não adivinhar citações ou assinaturas pelo conteúdo jurídico. Estilos não
  reconhecidos são apresentados na revisão, e trechos podem receber um tipo
  explicitamente naquele arquivo. A revisão é necessária para diferenciar
  trechos que usam todos o mesmo estilo Normal.
- Mostrar revisão por documento com trechos e tipos, além das imagens existentes.
  Permitir largura proporcional e alinhamento de imagens; quando necessário,
  posicionar uma imagem em parágrafo próprio antes/depois do trecho escolhido.
  Não prometer um editor visual de páginas equivalente ao Word.
- Guardar ajustes individuais associados ao hash do documento. Se a entrada
  mudar após a revisão, bloquear esses ajustes até nova revisão.
- Permitir processar documentos com conteúdo complexo preservando os trechos
  protegidos e informando quais não receberam formatação. Revisões não serão
  aceitas/rejeitadas automaticamente; equações não serão convertidas para texto.
  Caixas de texto comuns poderão receber regras próprias, preservando sua caixa.
  Continua disponível o modo estrito de recusar documentos complexos.
- Arquivos com senha, macros ou assinatura digital continuam fora do fluxo.
  Não há garantia possível de modificar um arquivo e preservar sua assinatura
  digital original; esse limite não será disfarçado como sucesso.
- Reabrir o resultado e verificar propriedades por categoria e alterações de
  imagem, mantendo todas as proteções contra perda/sobrescrita do original.

## Arquitetura

Configuração: regras por categoria e mapa de estilos como parte do perfil.
Inspeção: lista de parágrafos/imagens com identificadores e hash da entrada.
Ajustes por documento: categoria de parágrafos e transformações de imagens,
separados das regras reutilizáveis do perfil. Lote captura cópias independentes
para impedir que edições na tela mudem uma execução em andamento.

Interface: botão de regras no editor do perfil e botão de revisar na lista de
arquivos. Controles com rótulos claros, foco/teclado, rolagem e erros junto ao
campo; sem exposição de XML, IDs técnicos ou JSON ao usuário.

## Alternativas consideradas

1. Estilos existentes + revisão explícita (escolhida): previsível, funciona sem
   Word instalado e não depende de inferências sobre o texto jurídico.
2. Classificação automática por conteúdo: pode confundir títulos/citações e
   atribuir regras incorretas; não será usada sem confirmação.
3. Word obrigatório como motor: amplia recursos nativos, mas depende de licença,
   instalação e automação específica do sistema; não é requisito do aplicativo.

## Validação e entrega

- Perfis com categorias diferentes e documentos com estilos misturados, tabelas,
  caixas de texto, equações e revisões preservadas.
- Imagens inline/flutuantes: proporção, ordem, largura e ajustes individuais.
- Detecção de entrada modificada após revisão; lote com perfis/ajustes isolados.
- Testes reais da interface para regras e revisão; suíte anterior preservada
  exceto expectativas alteradas explicitamente pelo novo suporte.
- Revisão independente de código e registro dos resultados finais.
- Windows: preparar teste do executável e executar em runner Windows quando
  disponível. Esta máquina é macOS; não declarar validação Windows sem execução.
- Homologação institucional: não foram fornecidos documentos reais. Testar com
  exemplos sintéticos representativos e deixar checklist curto para comparação
  com documentos anonimizados da instituição. Não confundir isso com aprovação
  de modelos reais de cada promotor.
