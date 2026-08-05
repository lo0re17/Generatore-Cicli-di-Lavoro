# Build dell'eseguibile standalone GeneratoreCicli.exe
# Uso:  powershell -ExecutionPolicy Bypass -File strumenti\build_exe.ps1
# Risultato in dist\GeneratoreCicli\ (cartella da distribuire cosi' com'e').

$radice = Split-Path -Parent $PSScriptRoot
Set-Location $radice

pyinstaller --noconfirm --clean --windowed --onedir `
    --name "GeneratoreCicli" `
    --distpath dist `
    app.py
if (-not $?) { Write-Error "Build PyInstaller fallita"; exit 1 }

# dati accanto all'exe (modificabili senza ricompilare)
Copy-Item -Recurse -Force templates dist\GeneratoreCicli\
Copy-Item -Recurse -Force anagrafica dist\GeneratoreCicli\
New-Item -ItemType Directory -Force dist\GeneratoreCicli\output | Out-Null
Copy-Item -Force README.md dist\GeneratoreCicli\

Write-Host "`nBuild completata: dist\GeneratoreCicli\GeneratoreCicli.exe"
