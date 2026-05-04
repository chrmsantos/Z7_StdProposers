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
        It 'Possui os 4 modulos principais esperados' {
            ($script:moduleNames -contains 'Mod1Infrastructure.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod2Engine.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod3Pipeline.bas') | Should Be $true
            ($script:moduleNames -contains 'Mod4Main.bas') | Should Be $true
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
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetProposicaoRange\(doc As Document\) As Range'
            $script:moduleContent['Mod4Main.bas'] | Should Match '(?m)^Public Function GetJustificativaRange\(doc As Document\) As Range'
        }

        It 'Pipeline contem inicializacao e finalizacao de logging' {
            $script:moduleContent['Mod3Pipeline.bas'] | Should Match 'Public Function InitializeLogging\(doc As Document\) As Boolean'
            $script:moduleContent['Mod3Pipeline.bas'] | Should Match 'Public Sub SafeFinalizeLogging\(\)'
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
