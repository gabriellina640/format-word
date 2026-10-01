# Manual visual do Format Word

Arquivo para encaminhar ao usuário: **[Manual-Visual-Format-Word.pdf](Manual-Visual-Format-Word.pdf)**.
O destinatário precisa apenas do PDF e do aplicativo; não precisa dos scripts
nem da pasta de capturas.

O manual tem 19 páginas horizontais, texto pesquisável, marcadores de navegação
e atalhos clicáveis na página 2. A capa indica o caminho para o primeiro uso e
para quem já tem um perfil. As páginas 10–17 são consulta de ajustes opcionais.

As imagens são capturas da interface real no macOS, com dados fictícios criados
em uma pasta temporária e configuração isolada. As janelas nativas seguem o
tema do sistema; as imagens não foram redesenhadas. Recortes e marcadores são
aplicados durante a diagramação, preservando os PNG originais em `telas/`.

## Atualizar o manual

Ferramentas exclusivas da documentação, sem alteração nas dependências do app:

```sh
.venv/bin/python -m pip install --target /private/tmp/formatword-manual-tools \
  reportlab pymupdf pyobjc-framework-Cocoa pyobjc-framework-Quartz
```

Para recapturar, execute na raiz do repositório, em uma sessão gráfica macOS:

```sh
PYTHONPATH=/private/tmp/formatword-manual-tools:. .venv/bin/python docs/manual-visual/capturar_telas.py
```

O script abre somente janelas do aplicativo, cria documentos sintéticos,
importa um perfil, revisa texto e imagem e executa um lote real. Exige como
resultado `Concluído`, `Concluído com ressalvas` e `Erro`. O arquivo inválido é
intencional. A configuração pessoal do usuário não é acessada. Fechar e reabrir
visualmente cada janela garante um redesenho completo antes da captura.

Para diagramar e validar (sem abrir a interface):

```sh
PYTHONPATH=/private/tmp/formatword-manual-tools .venv/bin/python docs/manual-visual/gerar_manual.py
PYTHONPATH=/private/tmp/formatword-manual-tools .venv/bin/python docs/manual-visual/validar_manual.py
```

O gerador usa as fontes Arial do macOS, incorpora as fontes ao PDF e atualiza
`dist/Manual-Visual-Format-Word.pdf` e o nome antigo
`dist/manual-de-uso-formatador.pdf`. Não é necessário recompilar o aplicativo
para atualizar esse documento. Outros sistemas precisam de caminhos de fontes
equivalentes para executar o gerador; a leitura do PDF é independente disso.

Depois de modificar o conteúdo ou recapturar telas, confira visualmente todas
as páginas renderizadas. A validação estrutural não substitui essa revisão.
