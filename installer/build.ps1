# Genera installer/Output/TTSWord-Setup.exe. Requisitos: ver "Generar el instalador" en CLAUDE.md.
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$python = Join-Path $root "helper\.venv\Scripts\python.exe"
$iscc = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "No encuentro Inno Setup 6: winget install JRSoftware.InnoSetup" }
if (-not (Test-Path "$root\helper\voices\es_AR-daniela-high.onnx")) { throw "Falta la voz de Piper en helper\voices" }

Push-Location "$root\addin"
try {
    npm run build
    if ($LASTEXITCODE) { throw "Falló npm run build" }
} finally { Pop-Location }

& $python -m PyInstaller --noconfirm --clean `
    --distpath "$PSScriptRoot\build\dist" --workpath "$PSScriptRoot\build\work" `
    "$PSScriptRoot\tts-helper.spec"
if ($LASTEXITCODE) { throw "Falló PyInstaller" }

# La versión del instalador sale del manifest: subila ahí para que Word tome la actualización.
$version = ([xml](Get-Content "$root\addin\manifest.xml")).OfficeApp.Version
& $iscc "/DAppVersion=$version" "$PSScriptRoot\tts-word.iss"
if ($LASTEXITCODE) { throw "Falló Inno Setup" }
Write-Host "Listo: $PSScriptRoot\Output\TTSWord-Setup.exe"
