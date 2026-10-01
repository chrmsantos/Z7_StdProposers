#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1


Describe 'Z7_STDPROPOSERS - Mod_13_EmentaCoherence' {

    BeforeAll {
        $repoRoot = Get-RepoRoot
        $mainPath = Join-Path $repoRoot 'source\main'

        $mod13Path = Join-Path $mainPath 'Mod_13_EmentaCoherence.bas'
        $mod4Path = Join-Path $mainPath 'Mod_04_Main.bas'

        if (-not (Test-Path $mod13Path)) {
            throw "Mod_13_EmentaCoherence.bas nao encontrado em $mainPath"
        }

        $script:mod13Content = Get-Content $mod13Path -Raw -Encoding UTF8
        $script:mod4Content = Get-Content $mod4Path -Raw -Encoding UTF8

        # Apenas linhas de codigo (exclui comentarios) para contratos de seguranca
        $script:mod13Code = (($script:mod13Content -split "`n") | Where-Object {
            $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
        }) -join "`n"
    }

    It 'Tem Attribute VB_Name correto' {
        $script:mod13Content | Should Match 'Attribute VB_Name = "Mod_13_EmentaCoherence"'
    }

    It 'Tem Option Explicit' {
        $script:mod13Content | Should Match '(?m)^Option Explicit'
    }

    It 'Declara CheckEmentaCoherence (verificacao automatica do pipeline)' {
        $script:mod13Content | Should Match '(?m)^Public Function CheckEmentaCoherence\(doc As Document\) As Boolean'
    }

    It 'Declara ShowEmentaCoherenceWarning (aviso unico ao final)' {
        $script:mod13Content | Should Match '(?m)^Public Sub ShowEmentaCoherenceWarning\(\)'
    }

    It 'Declara macro manual VerificarCoerenciaEmenta' {
        $script:mod13Content | Should Match '(?m)^Public Sub VerificarCoerenciaEmenta\(\)'
    }

    It 'Declara autoteste TestarCoerenciaEmenta' {
        $script:mod13Content | Should Match '(?m)^Public Sub TestarCoerenciaEmenta\(\)'
    }

    It 'ECH_CompareTexts e funcao pura, testavel sem documento' {
        $script:mod13Content | Should Match '(?m)^Public Function ECH_CompareTexts\('
    }

    It 'Mod_13 e somente leitura (nao altera o documento)' {
        $script:mod13Code | Should Not Match '\.Delete'
        $script:mod13Code | Should Not Match 'InsertParagraph'
        $script:mod13Code | Should Not Match 'InsertBefore'
        $script:mod13Code | Should Not Match 'InsertAfter'
        $script:mod13Code | Should Not Match 'doc\.UndoClear'
    }

    It 'Mod_13 usa logging do projeto' {
        $script:mod13Content | Should Match 'LogMessage'
        $script:mod13Content | Should Match 'LOG_LEVEL_INFO'
        $script:mod13Content | Should Match 'LOG_LEVEL_WARNING'
    }

    It 'Mod_13 reutiliza helpers de Mod_01 (NormalizeForComparison, LevenshteinDistance)' {
        $script:mod13Content | Should Match 'NormalizeForComparison'
        $script:mod13Content | Should Match 'LevenshteinDistance'
    }

    It 'Mod_13 reutiliza ranges estruturais de Mod_04 (GetEmentaRange, GetCorpoRange)' {
        $script:mod13Content | Should Match 'GetEmentaRange'
        $script:mod13Content | Should Match 'GetCorpoRange'
    }

    It 'CheckEmentaCoherence e fail-open (falha nunca interrompe a padronizacao)' {
        $fn = [regex]::Match($script:mod13Content, '(?s)Public Function CheckEmentaCoherence\(.*?End Function')
        $fn.Success | Should Be $true
        $fn.Value | Should Match 'On Error GoTo ErrorHandler'
        # Em qualquer saida (inclusive erro) o pipeline deve seguir: retorna True
        $fn.Value | Should Match 'CheckEmentaCoherence = True'
        $fn.Value | Should Match '(?m)^\s*CheckEmentaCoherence = True\s*$'
    }

    It 'ShowEmentaCoherenceWarning limpa a mensagem pendente (exibe uma unica vez)' {
        $fn = [regex]::Match($script:mod13Content, '(?s)Public Sub ShowEmentaCoherenceWarning\(.*?End Sub')
        $fn.Success | Should Be $true
        $fn.Value | Should Match 'echLastWarning = ""'
    }

    It 'Mod_04 chama CheckEmentaCoherence apos BuildParagraphCache (estrutura identificada)' {
        $script:mod4Content | Should Match 'BuildParagraphCache doc[\s\S]*?CheckEmentaCoherence doc'
    }

    It 'Mod_04 exibe o aviso com ShowEmentaCoherenceWarning ao final do pipeline' {
        $script:mod4Content | Should Match 'ShowEmentaCoherenceWarning'
    }

    It 'Autoteste cobre casos que devem avisar e casos que devem passar' {
        $fn = [regex]::Match($script:mod13Content, '(?s)Public Sub TestarCoerenciaEmenta\(.*?End Sub')
        $fn.Success | Should Be $true
        $fn.Value | Should Match 'deve avisar'
        $fn.Value | Should Match 'deve passar'
    }

}
