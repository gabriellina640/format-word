$ErrorActionPreference = "Stop"

if (!(Test-Path ".venv")) {
    py -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\pyinstaller.exe --noconsole --onefile --name FormatadorDocumentos --icon icone.ico --hidden-import pypdf --hidden-import PIL.ImageTk --collect-data customtkinter main.py

Write-Host "Executavel gerado em dist\FormatadorDocumentos.exe"
