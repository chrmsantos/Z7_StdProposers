#Requires -Version 5.1
<#
.SYNOPSIS
    Importa os 4 modulos VBA do projeto Z7_StdProposers no documento Word ativo.

.DESCRIPTION
    Remove modulos existentes com o mesmo nome e reimporta os .bas de source\main\.
    Requer que a macro "Confiar no acesso ao modelo de objeto do projeto VBA" esteja
    habilitada em Word > Opcoes > Central de Confiabilidade > Configuracoes de Macro.

.PARAMETER TargetDocument
    Nome do documento Word de destino (ex: "Normal" para o Normal.dotm).
    Se omitido, usa o documento ativo.

.EXAMPLE
    .\import_vba.ps1
    .\import_vba.ps1 -TargetDocument "Normal"
#>
param(
    [string]$TargetDocument = "Normal"
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$BasDir = Join-Path $PSScriptRoot "source\main"
$Modules = @(
    "Mod1Infrastructure.bas",
    "Mod2Engine.bas",
    "Mod3Pipeline.bas",
    "Mod4Main.bas"
)

# Verifica que todos os arquivos existem antes de comecar
foreach ($file in $Modules) {
    $path = Join-Path $BasDir $file
    if (-not (Test-Path $path)) {
        Write-Error "Arquivo nao encontrado: $path"
        exit 1
    }
}

# Conecta ao Word via COM
Write-Host "Conectando ao Word..." -ForegroundColor Cyan
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
} catch {
    Write-Error "Word nao esta aberto. Abra o Word e tente novamente."
    exit 1
}

# Seleciona o Normal.dotm (padrao) ou documento informado
if ($TargetDocument -eq "Normal") {
    $doc = $word.NormalTemplate
    if ($null -eq $doc) {
        Write-Error "Normal.dotm nao encontrado via word.NormalTemplate."
        exit 1
    }
} elseif ($TargetDocument -ne "") {
    try {
        $doc = $word.Documents.Item($TargetDocument)
    } catch {
        Write-Error "Documento '$TargetDocument' nao encontrado. Documentos abertos: $(($word.Documents | ForEach-Object { $_.Name }) -join ', ')"
        exit 1
    }
} else {
    $doc = $word.ActiveDocument
    if ($null -eq $doc) {
        Write-Error "Nenhum documento ativo no Word."
        exit 1
    }
}

Write-Host "Documento alvo: $($doc.Name)" -ForegroundColor Cyan

# Verifica acesso ao VBProject
try {
    $vbp = $doc.VBProject
} catch {
    Write-Error @"
Acesso ao modelo de objeto VBA bloqueado.

Para habilitar:
  Word > Arquivo > Opcoes > Central de Confiabilidade
  > Configuracoes de Macro > marcar
  'Confiar no acesso ao modelo de objeto do projeto VBA'
"@
    exit 1
}

$components = $vbp.VBComponents

# Passagem 1: remove todos os modulos pre-existentes
Write-Host "Removendo modulos pre-existentes..." -ForegroundColor Cyan
foreach ($file in $Modules) {
    $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file)
    try {
        $existing = $components.Item($moduleName)
        $components.Remove($existing)
        Write-Host "  Removido: $moduleName" -ForegroundColor Yellow
    } catch {
        # Modulo nao existia, nada a remover
    }
}

# Passagem 2: importa todos os modulos
Write-Host "Importando modulos..." -ForegroundColor Cyan
foreach ($file in $Modules) {
    $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $fullPath   = Join-Path $BasDir $file
    $components.Import($fullPath) | Out-Null
    Write-Host "  Importado: $moduleName  <-  $fullPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "Importacao concluida. $($Modules.Count) modulos carregados em '$($doc.Name)'." -ForegroundColor Green
