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
        $mod1 | Should Match 'Public Function CheckForUpdates\(\) As Boolean'
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
}
