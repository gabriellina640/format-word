# Format Word

Aplicativo desktop para receber um arquivo `.docx` ou `.pdf` e gerar um `.docx` formatado conforme configurações salvas pelo usuário.

## Funcionalidades

- Aba para subir arquivo e escolher pasta de saída.
- Aba de configurações persistentes.
- Fonte, tamanho, margens, espaçamento, recuo, justificação e sufixo configuráveis.
- Lista de fontes padrão em dropdown.
- Imagem de cabeçalho e rodapé salvas localmente com confirmação visual.
- Imagens adaptadas automaticamente a uma área fixa profissional para evitar cabeçalho ou rodapé gigantes.
- Stacks nomeadas de configuração, permitindo salvar perfis como `Documentos para PGJ` com fonte, margens, espaçamentos e imagens próprias.
- Dropdown em Configurações para selecionar stacks já salvas antes de editar ou excluir.
- Seleção da stack em dropdown antes de pré-visualizar e exportar cada documento.
- Abas e modal de pré-visualização com layout adaptável e rolagem para evitar conteúdo cortado em janelas menores.
- Opção rápida para aplicar ou não cabeçalho/rodapé na formatação.
- Modal de pré-visualização antes da exportação, com edição de texto, fonte, espaçamento e posição das imagens.
- Extração de parágrafos e texto de tabelas em arquivos `.docx`.
- Validação de extensão, limite de tamanho e recusa de PDFs protegidos por senha.
- Validação básica de assinatura PNG/JPEG no upload de imagens.
- Sanitização do nome de saída e geração sem sobrescrever arquivos existentes.
- Interface com resumo da configuração ativa, barra de progresso e bloqueio de cliques duplicados durante a conversão.

## Rodar localmente

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

No Windows:

```powershell
py -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Gerar executável

No Windows, execute:

```powershell
.\build_windows.ps1
```

O `.exe` ficará em `dist/FormatWord.exe`.

> Observação: executáveis Windows devem ser gerados no Windows. Se gerar no macOS, o PyInstaller cria um app/binário para macOS, não um `.exe`.

## Build automático no GitHub Actions

Ao enviar o projeto para o GitHub, o workflow `.github/workflows/build-windows-exe.yml` gera o executável em `windows-latest` e publica o arquivo como artifact chamado `FormatWord-windows-exe`.

## Observações

- O app aceita `.docx` e `.pdf`. Arquivos `.doc` antigos não são aceitos por segurança.
- A importação de PDF depende do texto estar extraível. PDF escaneado como imagem precisará de OCR em uma etapa futura.
- Tabelas de `.docx` são convertidas para linhas de texto separadas por `|`.
- As configurações são salvas no perfil do usuário, fora da pasta do projeto.
