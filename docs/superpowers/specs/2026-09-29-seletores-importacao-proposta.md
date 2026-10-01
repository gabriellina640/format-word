# Seletores e criação de perfil a partir de Word

Status: aprovado pelo usuário em 29/09/2026: “pode implementar e deixar pronto”.

## Diagnóstico confirmado

Em 29/09/2026, execução da interface Tk real no macOS, com configuração
temporária, chamou o mesmo callback usado pelo menu do CTkComboBox.
Em 15 campos habilitados, a primeira seleção deixou o valor vazio e a segunda
gravou a opção. Campos: formatting_mode, bold, italic, underline, alignment,
line_spacing_mode, keep_with_next, keep_together, widow_control, paper_size,
orientation, header_mode, footer_mode, body_image_alignment e
complex_content_mode. Não foi uma simulação de cliques físicos no menu.

Causa: CTkComboBox._dropdown_callback habilita a entrada, apaga o texto e
insere o novo valor. A exclusão dispara o trace de StringVar; app/ui.py
executa _changed e _refresh_form de forma síncrona, restaurando readonly
antes da inserção. A inserção é ignorada. No segundo acionamento, a entrada
já está vazia e a operação de exclusão não produz a mesma interferência.

## Correção proposta

Adiar a atualização dos estados dos controles até terminar o callback de
seleção, consolidando atualizações pendentes. Manter atualização imediata
explícita para carregamento de perfil e início/fim de lote. Cancelar callbacks
pendentes ao destruir a janela. Preservar a validação e o resumo existentes.

Alternativa: substituir todos os seletores. Não recomendada porque aumenta
o escopo visual e de teclado sem necessidade de resolver a causa encontrada.

Critérios: todas as opções habilitadas devem funcionar na primeira seleção,
nos dois sentidos e repetidamente; os campos dependentes de cabeçalho/rodapé
devem acompanhar o novo modo; controles continuam bloqueados durante lote;
fechar com atualização pendente não deve produzir erro Tk.

Testes existentes que apenas usam StringVar.set não cobrem a sequência
delete/insert do menu. A regressão deve acionar o callback real de seleção.
Conferir também os diálogos ttk e o seletor de perfis, que usam outro fluxo.

## Viabilidade da importação

É possível preencher um perfil usando um arquivo .docx de referência, sem
Word instalado. O modelo atual armazena fonte, tamanho, ênfase, cor,
alinhamento, espaçamento, recuos, papel, orientação, margens, distâncias de
cabeçalho/rodapé e regras por categoria. O código atual inspeciona documentos,
mas ainda não extrai essas configurações para criar um perfil.

Fluxo recomendado: “Importar de Word” em Perfis → escolher .docx → revisar
configurações detectadas e diferenças → criar rascunho → salvar perfil.
Cancelar mantém o perfil atual. O documento de referência permanece intacto.

A leitura precisa resolver formatação direta, estilos herdados, padrões do
documento e referências ao tema. Propriedades sem resolução confiável devem
aparecer como pendências, nunca como valores extraídos com certeza.

Para categorias reconhecidas, sugerir regras por tipo e permitir revisão.
Se houver fontes/recuos diferentes na mesma categoria, ou margens/papéis
diferentes entre seções, apresentar alternativas ao usuário: o modelo atual
não representa todos esses padrões simultaneamente. Não deduzir o papel
jurídico de trechos apenas pelo conteúdo.

Cabeçalhos e rodapés arbitrários, tabelas complexas, numerações e objetos
flutuantes não cabem integralmente no perfil atual. O modo “Preservar original”
preserva os elementos do documento de destino; não copia os da referência.
A primeira versão deve deixar isso explícito na revisão. Copiar timbrados
completos exigiria uma ampliação própria do modelo e do formatador.

Alternativas: importar somente a configuração do corpo simplifica o trabalho,
mas perde diferenças entre categorias; copiar um template inteiro amplia a
fidelidade de elementos estruturais, mas exige outro escopo de implementação.
Recomendação: importação dos campos suportados com revisão de ambiguidades.

## Validação prevista após aprovação

Correção: reprodução falhando antes, regressões dos seletores passando depois,
suíte existente, code review e registro final dos resultados.

Importação: documentos sintéticos com estilos herdados, formatação direta,
temas, categorias e seções divergentes; erros de leitura; cancelar sem alterar
perfil; salvar/reabrir perfil; aplicar em outro documento e verificar as
propriedades resultantes. A implementação da importação depende da aprovação
desse escopo, pois a solicitação inicial pergunta sobre sua viabilidade.

Referências técnicas:
- https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html
- https://python-docx.readthedocs.io/en/latest/user/sections.html
