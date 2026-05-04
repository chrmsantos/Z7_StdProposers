#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'Z7_STDPROPOSERS - VBA Identifier Functions (Modular)' {
    BeforeAll {
        $repoRoot = Get-RepoRoot
        $mod4Path = Join-Path $repoRoot 'source\main\Mod4Main.bas'
        if (-not (Test-Path $mod4Path)) {
            throw "Arquivo nao encontrado: $mod4Path"
        }

        $script:mod4Content = Get-Content $mod4Path -Raw -Encoding UTF8
    }

    Context 'Declaracoes publicas de funcoes de range' {
        It 'Declara GetTituloRange' {
            $script:mod4Content | Should Match '(?m)^Public Function GetTituloRange\(doc As Document\) As Range'
        }

        It 'Declara GetEmentaRange' {
            $script:mod4Content | Should Match '(?m)^Public Function GetEmentaRange\(doc As Document\) As Range'
        }

        It 'Declara GetProposicaoRange' {
            $script:mod4Content | Should Match '(?m)^Public Function GetProposicaoRange\(doc As Document\) As Range'
        }

        It 'Declara GetJustificativaRange' {
            $script:mod4Content | Should Match '(?m)^Public Function GetJustificativaRange\(doc As Document\) As Range'
        }

        It 'Declara GetDataRange' {
            $script:mod4Content | Should Match '(?m)^Public Function GetDataRange\(doc As Document\) As Range'
        }

        It 'Declara GetAssinaturaRange' {
            $script:mod4Content | Should Match '(?m)^Public Function GetAssinaturaRange\(doc As Document\) As Range'
        }
    }

    Context 'Validacoes de seguranca por indice' {
        It 'GetTituloRange valida limites de indice' {
            $script:mod4Content | Should Match 'If tituloParaIndex <= 0 Or tituloParaIndex > doc\.Paragraphs\.count Then Exit Function'
        }

        It 'GetEmentaRange valida limites de indice' {
            $script:mod4Content | Should Match 'If ementaParaIndex <= 0 Or ementaParaIndex > doc\.Paragraphs\.count Then Exit Function'
        }

        It 'GetJustificativaRange valida limites dos indices' {
            $script:mod4Content | Should Match 'If justificativaStartIndex <= 0 Or justificativaEndIndex <= 0 Then Exit Function'
        }

        It 'GetAssinaturaRange valida limites dos indices' {
            $script:mod4Content | Should Match 'If assinaturaStartIndex <= 0 Or assinaturaEndIndex <= 0 Then Exit Function'
        }
    }

    Context 'Padrao de retorno seguro' {
        It 'Inicializa retorno como Nothing nas funcoes Get*' {
            $functions = @('GetTituloRange','GetEmentaRange','GetProposicaoRange','GetJustificativaRange','GetDataRange','GetAssinaturaRange')
            foreach ($f in $functions) {
                $script:mod4Content | Should Match "(?s)Public Function $f.*?Set $f = Nothing"
            }
        }

        It 'Funcoes Get* possuem tratamento de erro' {
            $functions = @('GetTituloRange','GetEmentaRange','GetProposicaoRange','GetJustificativaRange','GetDataRange','GetAssinaturaRange')
            foreach ($f in $functions) {
                $script:mod4Content | Should Match "(?s)Public Function $f.*?ErrorHandler:"
            }
        }
    }
}
