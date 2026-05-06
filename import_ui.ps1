#Requires -Version 5.1
<#
.SYNOPSIS
    Importa as personalizacoes da barra de opcoes e atalhos do Word (.exportedUI).

.DESCRIPTION
    Este script substitui o arquivo de configuracao de interface de usuario do Word (Word.officeUI)
    pelo arquivo customizado fornecido neste repositorio.
    O Word deve estar completamente fechado para que a copia seja efetuada com sucesso.
    Um backup do layout anterior do usuario e criado automaticamente.
#>

$ErrorActionPreference = "Stop"

$UiSource = Join-Path $PSScriptRoot "ui\word_person.exportedUI"

if (-not (Test-Path $UiSource)) {
    Write-Error "Arquivo de personalizacao de interface nao encontrado: $UiSource"
    exit 1
}

# 1. Verifica se o Word esta em execucao
$wordProcess = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProcess) {
    Write-Warning "O Microsoft Word esta aberto! Feche o Word completamente antes de importar a Interface."
    Write-Host "Por favor, feche todas as janelas do Word e execute este script novamente." -ForegroundColor Yellow
    exit 1
}

# 2. Localiza a pasta de configuracao do Office
$LocalAppData = [Environment]::GetFolderPath("LocalApplicationData")
$OfficeUIDir = Join-Path $LocalAppData "Microsoft\Office"
$WordUIFile = Join-Path $OfficeUIDir "Word.officeUI"

if (-not (Test-Path $OfficeUIDir)) {
    New-Item -ItemType Directory -Path $OfficeUIDir | Out-Null
}

# 3. Cria backup da UI atual (se existir)
if (Test-Path $WordUIFile) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $BackupFile = Join-Path $OfficeUIDir "Word.officeUI.backup_$timestamp"
    Copy-Item -Path $WordUIFile -Destination $BackupFile -Force
    Write-Host "Backup da UI anterior salvo em:" -ForegroundColor Cyan
    Write-Host "  $BackupFile" -ForegroundColor DarkGray
}

# 4. Copia a nova UI
Write-Host "Importando nova personalizacao de interface..." -ForegroundColor Cyan
Copy-Item -Path $UiSource -Destination $WordUIFile -Force

Write-Host "`nImportacao concluida com sucesso! Pode abrir o Word para visualizar a nova Guia." -ForegroundColor Green
