$ErrorActionPreference = "Stop"

if (!(Test-Path ".venv")) {
    py -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Falha ao criar ambiente Python." }
}

.\.venv\Scripts\python.exe -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar dependencias." }
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) { throw "Testes falharam; build cancelado." }
.\.venv\Scripts\pyinstaller.exe --noconsole --onefile --name FormatadorDocumentos --icon icone.ico --add-data "icone.ico;." --hidden-import PIL.ImageTk --collect-data customtkinter main.py
if ($LASTEXITCODE -ne 0) { throw "Falha ao gerar executavel." }

$reportPath = Join-Path $PWD "dist/self-test.json"
$process = Start-Process -FilePath "dist/FormatadorDocumentos.exe" -ArgumentList @("--self-test", "`"$reportPath`"") -PassThru
if (!$process.WaitForExit(120000)) { $process.Kill(); throw "Teste do executavel excedeu 120 segundos." }
if ($process.ExitCode -ne 0) { throw "Falha no teste do executavel." }
$report = Get-Content $reportPath -Raw | ConvertFrom-Json
if (!$report.ok) { throw "Executavel relatou falha." }
Write-Host "Executavel gerado e testado em dist\FormatadorDocumentos.exe"
