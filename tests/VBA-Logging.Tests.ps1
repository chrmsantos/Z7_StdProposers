#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'Z7_STDPROPOSERS - Logging VBA' {
    BeforeAll {
        $repoRoot = Get-RepoRoot
        $mainPath = Join-Path $repoRoot 'source\main'
        $modules = Get-ChildItem -Path $mainPath -Filter '*.bas' -File -ErrorAction Stop

        $mod3Path = Join-Path $mainPath 'Mod5Logging.bas'
        $mod4Path = Join-Path $mainPath 'Mod4Main.bas'
        $mod1Path = Join-Path $mainPath 'Mod1Infrastructure.bas'

        $mod3 = Get-Content $mod3Path -Raw -Encoding UTF8
        $mod4 = Get-Content $mod4Path -Raw -Encoding UTF8
        $mod1 = Get-Content $mod1Path -Raw -Encoding UTF8

        $allContent = ($modules | ForEach-Object { Get-Content $_.FullName -Raw -Encoding UTF8 }) -join "`n"
    }

    It 'Declara estrutura central de logging no pipeline' {
        $mod3 | Should Match 'Public Function InitializeLogging\(doc As Document\) As Boolean'
        $mod3 | Should Match 'Public Sub LogMessage\(message As String'
        $mod3 | Should Match 'Public Sub SafeFinalizeLogging\(\)'
    }

    It 'Mantem identificador de sessao e operacao no log' {
        $mod1 | Should Match 'Public currentLogSessionId As String'
        $mod1 | Should Match 'Public currentOperationId As String'
        $mod3 | Should Match '\[op=" & operationId & "\]'
    }

    It 'Inclui snapshots de contexto em pontos criticos do fluxo principal' {
        $mod4 | Should Match 'LogContextSnapshot doc, "INICIO"'
        $mod4 | Should Match 'LogContextSnapshot doc, "FIM"'
        $mod4 | Should Match 'LogContextSnapshot doc, "ERRO_CRITICO"'
    }

    It 'Registra eventos de erro e aviso em quantidade relevante' {
        $warnCount = ([regex]::Matches($allContent, 'LOG_LEVEL_WARNING')).Count
        $errorCount = ([regex]::Matches($allContent, 'LOG_LEVEL_ERROR')).Count

        $warnCount | Should BeGreaterThan 20
        $errorCount | Should BeGreaterThan 20
    }
}
