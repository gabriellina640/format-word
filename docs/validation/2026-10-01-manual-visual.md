# Manual visual — evidência de entrega

Pedido: manual com telas reais e instruções para que o destinatário consiga usar
a automação sem depender de uma apresentação presencial.

## Artefato

- `docs/manual-visual/Manual-Visual-Format-Word.pdf`: 19 páginas horizontais.
- Cópias idênticas em `dist/Manual-Visual-Format-Word.pdf` e
  `dist/manual-de-uso-formatador.pdf` (substitui o manual antigo nesse caminho).
- 19 marcadores de navegação, 7 atalhos internos, texto pesquisável e fontes
  utilizadas incorporadas. Tamanho: 3.733.471 bytes.
- SHA-256: `54f75e1099cafadde4fdc6b64192b82e3eef2c6f44c690b226c1e3c0ceafcc0c`.

Conteúdo: primeiro uso, importação de configurações de Word, salvar/reutilizar
perfil, processamento em lote, revisão opcional de texto/imagens, leitura de
resultados, todas as seis seções do perfil, regras por tipo e dúvidas frequentes.
O manual diferencia referência, perfil, entrada, salvar e aplicar, e descreve
os limites da importação sem prometer cópia integral do modelo.

## Captura e validação executadas

```sh
PYTHONPATH=/private/tmp/formatword-manual-tools:. .venv/bin/python docs/manual-visual/capturar_telas.py
PYTHONPATH=/private/tmp/formatword-manual-tools .venv/bin/python docs/manual-visual/gerar_manual.py
PYTHONPATH=/private/tmp/formatword-manual-tools .venv/bin/python docs/manual-visual/validar_manual.py
git diff --check
```

- 23 capturas PNG íntegras da interface real. Perfil e DOCX fictícios em diretório
  temporário; configuração pessoal não acessada.
- Importação/revisão/execução reais. Lote demonstrativo confirmou, nessa ordem,
  `Concluído`, `Concluído com ressalvas` e `Erro`. A falha usa um DOCX sintético
  intencionalmente inválido, sem envolver documentos reais do usuário.
- Todas as páginas renderizadas e inspecionadas. Conferidos controles,
  indicadores numéricos, caminhos temporários, legibilidade e limites de página.
- Validação estrutural: 238 caixas de texto sem colisões entre si; links com
  destinos válidos; fontes usadas incorporadas; texto sem caracteres de
  substituição nem caminhos pessoais; cópias distribuídas com o mesmo hash.
- Scripts e dependências são exclusivos da documentação. Nenhum código do
  aplicativo foi modificado; não houve necessidade de reconstruir o executável.

## Revisão independente

Revisor `/root/review_ui`, conforme skill local requesting-code-review.
Inspeção das 19 páginas, texto extraído e correspondência com o código do app.

Achados resolvidos e reconferidos:

1. Captura com resíduos da aba anterior: recapturada após redesenho completo.
2. Sobreposição de título/corpo na capa: espaçamento corrigido.
3. Ordem do vínculo de estilo: marcar antes de `Classificar selecionados`.
4. `Seguir corpo`: explicado como herança da regra de Corpo, inclusive ajustes
   personalizados, em vez de sempre usar diretamente o formulário principal.

Parecer final: nenhum achado crítico ou importante pendente; apto para entrega.
Limite: a aparência das janelas foi capturada no macOS; o manual informa que
aparência e seleção de arquivos podem variar conforme o sistema.
