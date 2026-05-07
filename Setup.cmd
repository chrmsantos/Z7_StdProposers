@echo off
setlocal
set "Z7_SELF=%~f0"
set "Z7_DIR=%~dp0"
if "%Z7_DIR:~-1%"=="\" set "Z7_DIR=%Z7_DIR:~0,-1%"

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$f=$env:Z7_SELF; $t=(Get-Content $f -Encoding UTF8 -Raw) -split '(?m)^#!PS\r?\n',2; Invoke-Expression $t[1]"
set PS_ERR=%ERRORLEVEL%

echo.
if %PS_ERR% neq 0 echo  ERRO: Instalacao falhou - verifique as mensagens acima.
pause
exit /b %PS_ERR%

#!PS
$ErrorActionPreference = "Stop"
$ScriptDir = $env:Z7_DIR

Write-Host ""
Write-Host " Z7_StdProposers - Instalador" -ForegroundColor Cyan
Write-Host " =============================" -ForegroundColor Cyan
Write-Host ""

# ── Pre-requisito: Fechar o Word ─────────────────────────────────────────────
$wordProc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProc) {
    Write-Host " ATENCAO: O Word esta aberto e precisa ser fechado para continuar." -ForegroundColor Yellow
    Write-Host " Pressione Enter para fechar o Word forcadamente, ou S+Enter para cancelar: " -NoNewline -ForegroundColor Yellow
    $resp = Read-Host
    if ($resp -match '^[Ss]$') {
        throw "Instalacao cancelada pelo usuario."
    }
    Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    $wordProc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    if ($wordProc) {
        throw "Nao foi possivel fechar o Word. Feche manualmente e tente novamente."
    }
    Write-Host " Word encerrado." -ForegroundColor Green
    Write-Host ""
}

# ── Passo 1: Instalar executaveis ────────────────────────────────────────────
Write-Host "[1/3] Instalando executaveis Python..." -ForegroundColor Cyan

$srcBase    = Join-Path $ScriptDir "ai"
$destBase   = Join-Path $env:LOCALAPPDATA "Z7\Apps\Z7_StdProposers\ai"
$runtimeDir = Join-Path $env:LOCALAPPDATA "Z7\Tmp\StdProposers"
$apps       = @("correct_grammar", "config_prompt", "chat_ia")

$missing = $apps | Where-Object { -not (Test-Path (Join-Path $srcBase $_)) }
if ($missing) {
    throw "Pastas de executavel nao encontradas em '$srcBase': $($missing -join ', '). Execute 'ai\build_exe.ps1' primeiro."
}

foreach ($app in $apps) {
    $src  = Join-Path $srcBase  $app
    $dest = Join-Path $destBase $app
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item -Path "$src\*" -Destination $dest -Recurse -Force
    Write-Host "  Instalado: $app" -ForegroundColor Green
}

if (-not (Test-Path $runtimeDir)) {
    New-Item -ItemType Directory -Force -Path $runtimeDir | Out-Null
}

# ── Passo 2: Importar interface do Word ──────────────────────────────────────
Write-Host ""
Write-Host "[2/3] Importando personalizacao de interface do Word..." -ForegroundColor Cyan

$uiSource  = Join-Path $ScriptDir "ui\word_person.exportedUI"
$officeDir = Join-Path $env:LOCALAPPDATA "Microsoft\Office"
$wordUI    = Join-Path $officeDir "Word.officeUI"

if (-not (Test-Path $uiSource)) {
    Write-Warning "Arquivo de UI nao encontrado: '$uiSource'. Etapa ignorada."
} else {
    if (-not (Test-Path $officeDir)) { New-Item -ItemType Directory -Path $officeDir | Out-Null }
    if (Test-Path $wordUI) {
        $ts     = Get-Date -Format "yyyyMMdd_HHmmss"
        $backup = Join-Path $officeDir "Word.officeUI.backup_$ts"
        Copy-Item $wordUI $backup -Force
        Write-Host "  Backup criado: $(Split-Path $backup -Leaf)" -ForegroundColor DarkGray
    }
    Copy-Item $uiSource $wordUI -Force
    Write-Host "  Interface importada com sucesso." -ForegroundColor Green
}

# ── Passo 3: Importar modulos VBA ────────────────────────────────────────────
Write-Host ""
Write-Host "[3/3] Importando modulos VBA para Normal.dotm..." -ForegroundColor Cyan

$basDir  = Join-Path $ScriptDir "source\main"
$modules = @("Mod1Infrastructure.bas", "Mod2Engine.bas", "Mod3Pipeline.bas", "Mod4Main.bas")

$missingVba = $modules | Where-Object { -not (Test-Path (Join-Path $basDir $_)) }
if ($missingVba) {
    throw "Arquivos VBA nao encontrados em '$basDir': $($missingVba -join ', ')"
}

try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    Write-Host "  Conectado ao Word." -ForegroundColor DarkGray
} catch {
    Write-Host "  Abrindo o Word..." -ForegroundColor DarkGray
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.Documents.Add() | Out-Null
}

$doc = $word.NormalTemplate
if ($null -eq $doc) {
    throw "Normal.dotm nao encontrado via word.NormalTemplate."
}

try {
    $vbp = $doc.VBProject
} catch {
    throw @"
Acesso ao modelo de objeto VBA bloqueado.
Habilite em: Word > Arquivo > Opcoes > Central de Confiabilidade
             > Configuracoes de Macro
             > Confiar no acesso ao modelo de objeto do projeto VBA
"@
}

$components = $vbp.VBComponents
foreach ($file in $modules) {
    $name = [IO.Path]::GetFileNameWithoutExtension($file)
    try { $components.Remove($components.Item($name)) } catch {}
}
foreach ($file in $modules) {
    $name = [IO.Path]::GetFileNameWithoutExtension($file)
    $components.Import((Join-Path $basDir $file)) | Out-Null
    Write-Host "  Importado: $name" -ForegroundColor Green
}

# ── Conclusao ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host " Instalacao concluida com sucesso!" -ForegroundColor Green
Write-Host ""
Write-Host " Proximo passo: configure a chave Gemini executando a macro pela primeira vez." -ForegroundColor White
Write-Host ""
