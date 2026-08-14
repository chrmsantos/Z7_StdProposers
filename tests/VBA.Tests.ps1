#requires -Version 5.1
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'script: scope variables are used in Pester It blocks')]
param()
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'Z7_STDPROPOSERS - VBA Modular Architecture' {
    BeforeAll {
        $repoRoot = Get-RepoRoot
        $mainPath = Join-Path $repoRoot 'source\main'
        $script:modules = Get-ChildItem -Path $mainPath -Filter '*.bas' -File -ErrorAction Stop | Sort-Object Name

        $script:moduleNames = $script:modules | Select-Object -ExpandProperty Name
        $script:moduleContent = @{}
        foreach ($m in $script:modules) {
            $script:moduleContent[$m.Name] = Get-Content $m.FullName -Raw -Encoding UTF8
        }

        $script:allContent = ($script:moduleContent.Values) -join "`n"
    }

    Context 'Estrutura de modulos' {
        It 'Possui os modulos esperados' {
            ($script:moduleNames -contains 'Mod1Infrastructure.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod2Engine.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod3Pipeline.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod5Logging.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod4Main.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod6WordMacro.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod7Formatting.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod8Ementa.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod9SpecialParagraphs.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod10Validation.bas') | Should Be $true
        }

        It 'Todos os modulos tem Option Explicit' {
            foreach ($name in $script:moduleNames) {
                $script:moduleContent[$name] | Should Match '(?m)^Option Explicit'
            }
        }

        It 'Todos os modulos sao nao vazios' {
            foreach ($name in $script:moduleNames) {
                $script:moduleContent[$name].Length | Should BeGreaterThan 100
            }
        }
    }

    Context 'Pontos de entrada e pipeline' {
        It 'Tem entrypoint principal PadronizarDocumentoMain' {
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Sub PadronizarDocumentoMain\('
        }

        It 'Mantem as funcoes publicas de ranges estruturais' {
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetTituloRange\(doc As Document\) As Range'
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetEmentaRange\(doc As Document\) As Range'
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetVocativoRange\(doc As Document\) As Range'
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetProposicaoRange\(doc As Document\) As Range'
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetJustificativaRange\(doc As Document\) As Range'
        }

        It 'Logging esta em Mod5Logging' {
            $script:moduleContent['Mod5Logging.bas'] | Should Match 'Public Function InitializeLogging\(doc As Document\) As Boolean'
            $script:moduleContent['Mod5Logging.bas'] | Should Match 'Public Sub SafeFinalizeLogging\(\)'
        }

        It 'Rotinas tikinho tk estao em Mod9SpecialParagraphs' {
            $script:moduleContent['Mod9SpecialParagraphs.bas'] | Should Match 'Private Function IsTikinhoTk\(ByVal text As String\) As Boolean'
            $script:moduleContent['Mod9SpecialParagraphs.bas'] | Should Match 'Public Sub ReplaceTikinhoTkParagraphs\(doc As Document\)'
        }

        It 'Substituicoes de Jd e numero estao em Mod9SpecialParagraphs' {
            $script:moduleContent['Mod9SpecialParagraphs.bas'] | Should Match 'ExecuteFindReplace\(doc, " Jd ", " Jd. ", True\)'
        }

        It 'Rotina de justificativa esta em Mod9SpecialParagraphs' {
            $script:moduleContent['Mod9SpecialParagraphs.bas'] | Should Match 'Public Sub RemoveJustificativaColon\(doc As Document\)'
        }

        It 'Paginacao de requerimentos esta em Mod2Engine' {
            $script:moduleContent['Mod2Engine.bas'] | Should Match 'Public Function IsRequerimentoPageLine\(text As String\) As Boolean'
        }

        It 'Espaco nao separavel esta em Mod7Formatting' {
            $script:moduleContent['Mod7Formatting.bas'] | Should Match 'Public Sub EnsureNonBreakingSpaceAfterNo\(doc As Document\)'
        }
    }

    Context 'Qualidade basica de implementacao' {
        It 'Contem tratamento de erro amigavel e recuperacao' {
            $script:allContent | Should Match 'ShowUserFriendlyError'
            $script:allContent | Should Match 'EmergencyRecovery'
        }

        It 'Contem niveis de log basicos' {
            $script:allContent | Should Match 'LOG_LEVEL_INFO'
            $script:allContent | Should Match 'LOG_LEVEL_WARNING'
            $script:allContent | Should Match 'LOG_LEVEL_ERROR'
        }

        It 'Mantem estrutura de funcoes/subs balanceada por modulo' {
            foreach ($name in $script:moduleNames) {
                $content = $script:moduleContent[$name]
                $functionStarts = ([regex]::Matches($content, '(?m)^(Public |Private )?Function\s+\w+')).Count
                $functionEnds = ([regex]::Matches($content, '(?m)^End Function')).Count
                $subStarts = ([regex]::Matches($content, '(?m)^(Public |Private )?Sub\s+\w+')).Count
                $subEnds = ([regex]::Matches($content, '(?m)^End Sub')).Count

                $functionStarts | Should Be $functionEnds
                $subStarts | Should Be $subEnds
            }
        }
    }
}
