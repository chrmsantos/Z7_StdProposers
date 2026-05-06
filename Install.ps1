#Requires -Version 5.1
<#
.SYNOPSIS
    Instala os executáveis compilados do Z7_StdProposers (parte Python) para o perfil do usuário.

.DESCRIPTION
    Copia as pastas dos executáveis compilados com PyInstaller (correct_grammar, config_prompt, chat_ia)
    da pasta do projeto para %LOCALAPPDATA%\Z7\Apps\Z7_StdProposers\ai\.

    Execute este script a partir da raiz do repositório após ter feito o build com ai\build_exe.ps1.

.EXAMPLE
    .\Install.ps1
    .\Install.ps1 -Verbose
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

# --- Caminhos ---
$srcBase  = Join-Path $PSScriptRoot "ai"
$destBase = Join-Path $env:LOCALAPPDATA "Z7\Apps\Z7_StdProposers\ai"
$runtimeDir = Join-Path $env:LOCALAPPDATA "Z7\Tmp\StdProposers"

$apps = @("correct_grammar", "config_prompt", "chat_ia")

# --- Validação de origem ---
$missing = $apps | Where-Object { -not (Test-Path (Join-Path $srcBase $_)) }
if ($missing) {
    Write-Error (
        "As seguintes pastas de executável não foram encontradas em '$srcBase':`n" +
        ($missing -join "`n") +
        "`n`nExecute 'ai\build_exe.ps1' antes de instalar."
    )
    exit 1
}

# --- Instalação ---
Write-Host "Destino: $destBase"
Write-Host ""

foreach ($app in $apps) {
    $src  = Join-Path $srcBase  $app
    $dest = Join-Path $destBase $app

    Write-Host "Instalando $app..." -NoNewline

    if (Test-Path $dest) {
        Remove-Item $dest -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item -Path "$src\*" -Destination $dest -Recurse -Force

    Write-Host " OK"
    Write-Verbose "  $src -> $dest"
}

# --- Pasta de runtime (logs, chave, prompt) ---
if (-not (Test-Path $runtimeDir)) {
    New-Item -ItemType Directory -Force -Path $runtimeDir | Out-Null
    Write-Verbose "Pasta de runtime criada: $runtimeDir"
}

Write-Host ""
Write-Host "Instalação concluída com sucesso."
Write-Host ""
Write-Host "Próximos passos:"
Write-Host "  1. Importe 'ai\WordMacro.bas' no editor VBA do Word (ALT+F11 > Arquivo > Importar)."
Write-Host "  2. Configure a chave Gemini executando a macro pela primeira vez."
