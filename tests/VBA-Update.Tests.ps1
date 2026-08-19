#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'Z7_STDPROPOSERS - Update VBA' {
    BeforeAll {
        $repoRoot = Get-RepoRoot
        $mainPath = Join-Path $repoRoot 'source\main'
        $modules = Get-ChildItem -Path $mainPath -Filter '*.bas' -File -ErrorAction Stop

        $mod1Path = Join-Path $mainPath 'Mod1Infrastructure.bas'
        $mod3Path = Join-Path $mainPath 'Mod3Pipeline.bas'

        $mod1 = Get-Content $mod1Path -Raw -Encoding UTF8
        $mod3 = Get-Content $mod3Path -Raw -Encoding UTF8
    }

    It 'Declara as rotinas de atualizacao corretas no modulo de infraestrutura' {
        $mod1 | Should Match 'Public Function CheckForUpdates\(Optional forceCheck As Boolean = False\) As Boolean'
        $mod1 | Should Match 'Public Function GetLocalVersion\(\) As String'
        $mod1 | Should Match 'Public Function GetRemoteVersion\(\) As String'
        $mod1 | Should Match 'Public Function CompareVersions\(ByVal version1 As String, ByVal version2 As String\) As Integer'
        $mod1 | Should Match 'Public Sub PromptForUpdate\(\)'
    }

    It 'Declara a rotina de atualizacao manual no modulo de pipeline' {
        $mod3 | Should Match 'Public Sub ExecutarInstalador\(\)'
    }

    It 'Contem tratamento de erro estruturado (ErrorHandler) nas rotinas de atualizacao' {
        # CheckForUpdates deve conter tratamento de erro
        $mod1 | Should Match '(?s)Public Function CheckForUpdates.*?On Error GoTo ErrorHandler'
        # PromptForUpdate deve conter tratamento de erro
        $mod1 | Should Match '(?s)Public Sub PromptForUpdate.*?On Error GoTo ErrorHandler'
        # ExecutarInstalador deve conter tratamento de erro
        $mod3 | Should Match '(?s)Public Sub ExecutarInstalador.*?On Error GoTo ErrorHandler'
    }

    It 'ErrorHandler loga descricao do erro com LOG_LEVEL_ERROR' {
        # CheckForUpdates ErrorHandler
        $mod1 | Should Match '(?s)ErrorHandler:.*?LogMessage.*?Err\.Description.*?LOG_LEVEL_ERROR'
        # PromptForUpdate ErrorHandler
        $mod1 | Should Match '(?s)Public Sub PromptForUpdate.*?ErrorHandler:.*?LogMessage.*?Err\.Description.*?LOG_LEVEL_ERROR'
        # ExecutarInstalador ErrorHandler
        $mod3 | Should Match '(?s)Public Sub ExecutarInstalador.*?ErrorHandler:.*?LogMessage.*?Err\.Description.*?LOG_LEVEL_ERROR'
    }

    It 'Rotinas de atualizacao tem Exit Sub/Function antes do ErrorHandler' {
        # CheckForUpdates deve ter Exit Function antes do ErrorHandler
        $mod1 | Should Match '(?s)Public Function CheckForUpdates.*?Exit Function.*?ErrorHandler:'
        # PromptForUpdate deve ter Exit Sub antes do ErrorHandler
        $mod1 | Should Match '(?s)Public Sub PromptForUpdate.*?Exit Sub.*?ErrorHandler:'
        # ExecutarInstalador deve ter Exit Sub antes do ErrorHandler
        $mod3 | Should Match '(?s)Public Sub ExecutarInstalador.*?Exit Sub.*?ErrorHandler:'
    }

    It 'Implementa logs detalhados (LogMessage) nos pontos decisivos das rotinas de atualizacao' {
        # Deve logar inicio e resultados de comparacao
        $mod1 | Should Match 'LogMessage "Iniciando verificacao de atualizacao\.\.\."'
        $mod1 | Should Match 'LogMessage "Versao local detectada: v"'
        $mod1 | Should Match 'LogMessage "Versao remota detectada: v"'

        # Deve logar decisoes interativas
        $mod1 | Should Match 'LogMessage "PromptForUpdate acionado interativamente pelo usuario"'
        $mod1 | Should Match 'LogMessage "Usuario aceitou a atualizacao para v"'
        $mod1 | Should Match 'LogMessage "Usuario recusou a atualizacao para v"'
        $mod1 | Should Match 'LogMessage "Disparando shell cmd para iniciar atualizacao: "'

        # Deve logar na atualizacao manual
        $mod3 | Should Match 'LogMessage "ExecutarInstalador acionado manualmente pelo usuario"'
        $mod3 | Should Match 'LogMessage "ExecutarInstalador: usuario confirmou execucao do instalador"'
        $mod3 | Should Match 'LogMessage "Disparando shell cmd para iniciar atualizacao manual: "'
    }

    It 'LogMessages das rotinas de atualizacao usam nivel de log correto' {
        # Extrai apenas as rotinas de atualizacao e verifica seus LogMessages
        # CheckForUpdates
        $checkBlock = [regex]::Match($mod1, '(?s)(Public Function CheckForUpdates.*?End Function)').Value
        $logLines = $checkBlock -split "`n" | Where-Object { $_ -match '^\s*LogMessage\s+"' }
        foreach ($line in $logLines) {
            $line | Should Match 'LOG_LEVEL_(INFO|WARNING|ERROR)'
        }
        # PromptForUpdate
        $promptBlock = [regex]::Match($mod1, '(?s)(Public Sub PromptForUpdate.*?End Sub)').Value
        $logLines2 = $promptBlock -split "`n" | Where-Object { $_ -match '^\s*LogMessage\s+"' }
        foreach ($line in $logLines2) {
            $line | Should Match 'LOG_LEVEL_(INFO|WARNING|ERROR)'
        }
        # ExecutarInstalador
        $execBlock = [regex]::Match($mod3, '(?s)(Public Sub ExecutarInstalador.*?End Sub)').Value
        $logLines3 = $execBlock -split "`n" | Where-Object { $_ -match '^\s*LogMessage\s+"' }
        foreach ($line in $logLines3) {
            $line | Should Match 'LOG_LEVEL_(INFO|WARNING|ERROR)'
        }
    }

    It 'CheckForUpdates tem log de resultado (atualizado ou atualizacao disponivel)' {
        $mod1 | Should Match 'LogMessage "Nova atualizacao disponivel: v"'
        $mod1 | Should Match 'LogMessage "Nenhuma atualizacao necessaria'
    }

    It 'ExecutarInstalador salva documentos abertos antes de executar' {
        $mod3 | Should Match '(?s)Public Sub ExecutarInstalador.*?doc\.Save'
    }

    It 'PromptForUpdate tem undoGroupEnabled guard' {
        $mod1 | Should Match '(?s)Public Sub PromptForUpdate.*?undoGroupEnabled'
    }
}
