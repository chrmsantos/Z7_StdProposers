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
            ($script:moduleNames -contains 'Mod_01_Infrastructure.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_02_Engine.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_03_Pipeline.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_05_Logging.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_04_Main.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_06_WordMacro.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_07_Formatting.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_08_Ementa.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_09_SpecialParagraphs.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_10_Validation.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_11_RevisionText.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod_12_AIStructure.bas') | Should Be $true
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
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Sub PadronizarDocumentoMain\('
        }

        It 'Mantem as funcoes publicas de ranges estruturais' {
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Function GetTituloRange\(doc As Document\) As Range'
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Function GetEmentaRange\(doc As Document\) As Range'
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Function GetVocativoRange\(doc As Document\) As Range'
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Function GetCorpoRange\(doc As Document\) As Range'
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Function GetJustificativaRange\(doc As Document\) As Range'
        }

        It 'Logging esta em Mod_05_Logging' {
            $script:moduleContent['Mod_05_Logging.bas'] | Should Match 'Public Function InitializeLogging\(doc As Document\) As Boolean'
            $script:moduleContent['Mod_05_Logging.bas'] | Should Match 'Public Sub SafeFinalizeLogging\(\)'
        }

        It 'Rotinas tikinho tk estao em Mod_09_SpecialParagraphs' {
            $script:moduleContent['Mod_09_SpecialParagraphs.bas'] | Should Match 'Private Function IsTikinhoTk\(ByVal text As String\) As Boolean'
            $script:moduleContent['Mod_09_SpecialParagraphs.bas'] | Should Match 'Public Sub ReplaceTikinhoTkParagraphs\(doc As Document\)'
        }

        It 'Substituicoes de Jd e numero estao em Mod_09_SpecialParagraphs' {
            $script:moduleContent['Mod_09_SpecialParagraphs.bas'] | Should Match 'ExecuteFindReplace\(doc, " Jd ", " Jd. ", True\)'
        }

        It 'Rotina de justificativa esta em Mod_09_SpecialParagraphs' {
            $script:moduleContent['Mod_09_SpecialParagraphs.bas'] | Should Match 'Public Sub RemoveJustificativaColon\(doc As Document\)'
        }

        It 'Paginacao de requerimentos esta em Mod_02_Engine' {
            $script:moduleContent['Mod_02_Engine.bas'] | Should Match 'Public Function IsRequerimentoPageLine\(text As String\) As Boolean'
        }

        It 'Espaco nao separavel esta em Mod_07_Formatting' {
            $script:moduleContent['Mod_07_Formatting.bas'] | Should Match 'Public Sub EnsureNonBreakingSpaceAfterNo\(doc As Document\)'
        }

        It 'Revisao IA de texto selecionado esta em Mod_11_RevisionText' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match '(?m)^Public Sub TestarRevisaoTextoSelecionado\('
        }

        It 'Revisao IA de documento inteiro esta em Mod_11_RevisionText' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match '(?m)^Public Sub CorrigirProposituraComIA\('
        }

        It 'Diagnostico OpenRouter esta em Mod_11_RevisionText' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match '(?m)^Public Sub DiagnosticarOpenRouter\('
        }

        It 'Mod_11_RevisionText usa caminhos centralizados de Mod_01_Infrastructure' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'GetZ7StdProposersDataPath'
        }

        It 'Mod_11_RevisionText usa logging do projeto' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'LogMessage'
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'LOG_LEVEL_INFO'
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'LOG_LEVEL_ERROR'
        }

        It 'Diagnostico de Estrutura IA esta em Mod_12_AIStructure' {
            $script:moduleContent['Mod_12_AIStructure.bas'] | Should Match '(?m)^Public Sub DiagnosticarEstruturaIA\(\)'
        }

        It 'Teste de Estrutura IA esta em Mod_12_AIStructure' {
            $script:moduleContent['Mod_12_AIStructure.bas'] | Should Match '(?m)^Public Sub TestarEstruturaIADocumentoAtual\(\)'
        }

        It 'Mod_12_AIStructure usa logging do projeto' {
            $script:moduleContent['Mod_12_AIStructure.bas'] | Should Match 'LogMessage'
            $script:moduleContent['Mod_12_AIStructure.bas'] | Should Match 'LOG_LEVEL_INFO'
            $script:moduleContent['Mod_12_AIStructure.bas'] | Should Match 'LOG_LEVEL_ERROR'
        }

        It 'Mod_12_AIStructure usa caminhos centralizados de Mod_01_Infrastructure' {
            $script:moduleContent['Mod_12_AIStructure.bas'] | Should Match 'GetZ7StdProposersDataPath'
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
