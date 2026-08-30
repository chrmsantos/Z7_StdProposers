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

        It 'CorrigirProposituraComIA esta em Mod_11_RevisionText' {
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

        It 'Mod_11_RevisionText tem SanitizarTextoIA para prevenir erros de encoding' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'Private Function SanitizarTextoIA'
        }

        It 'DesescaparJSON usa vbCr (nao vbCrLf) para quebras de linha JSON' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            # Extrai o bloco Case "n" dentro de DesescaparJSON
            $match = [regex]::Match($mod11, 'Case "n"\s*\r?\n\s*(resultado\s*=\s*resultado\s*&\s*\w+)')
            $match.Success | Should Be $true
            $match.Groups[1].Value | Should Match 'vbCr\b'
            $match.Groups[1].Value | Should Not Match 'vbCrLf'
        }

        It 'SubstituirTextoPreservandoFormatacao sanitiza texto antes de substituir' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'novoTexto = SanitizarTextoIA\(novoTexto\)'
        }

        It 'ProcessarTextoComIA sanitiza resposta antes de retornar' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'SanitizarTextoIA\(LimparRespostaIA'
        }

        It 'SanitizarTextoIA remove caracteres de controle preservando CR e Tab' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            # Verifica que preserva vbCr e vbTab
            $mod11 | Should Match 'Case &H9, &HD'
            # Verifica que remove controles 0x01-0x08, 0x0B-0x0C, 0x0E-0x1F
            $mod11 | Should Match '&H0 To &H8.*&HA To &HC.*&HE To &H1F'
        }

        It 'DesescaparJSON valida codepoints Unicode problematicos' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            # Verifica que rejeita NUL
            $mod11 | Should Match 'Case &H0.*NUL'
            # Verifica que rejeita BOM markers
            $mod11 | Should Match '&HFFFE.*&HFFFF'
            # Verifica que rejeita surrogates isolados
            $mod11 | Should Match '&HD800 To &HDFFF'
        }

        It 'BytesParaStringUTF8 decodifica UTF-8 manualmente (sem ADODB.Stream Charset)' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            # Extrai o corpo da funcao BytesParaStringUTF8
            $match = [regex]::Match($mod11, 'Private Function BytesParaStringUTF8[\s\S]*?End Function')
            # Nao deve criar ADODB.Stream (uso real, nao comentarios)
            $match.Value | Should Not Match 'CreateObject.*ADODB'
            # Deve ter logica de decodificacao manual (2-byte sequence detection)
            $match.Value | Should Match 'b And &HE0.*= &HC0'
        }

        It 'LerArquivoUTF8 existe em Mod_11_RevisionText' {
            $script:moduleContent['Mod_11_RevisionText.bas'] | Should Match 'Private Function LerArquivoUTF8'
        }

        It 'CarregarPromptRevisao usa LerArquivoUTF8 (nao Line Input)' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            # Extrai o corpo de CarregarPromptRevisao
            $match = [regex]::Match($mod11, 'Private Function CarregarPromptRevisao[\s\S]*?End Function')
            $match.Value | Should Match 'LerArquivoUTF8'
            $match.Value | Should Not Match 'Line Input'
        }

        It 'CarregarModeloIA usa LerArquivoUTF8 (nao Line Input)' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function CarregarModeloIA[\s\S]*?End Function')
            $match.Value | Should Match 'LerArquivoUTF8'
            $match.Value | Should Not Match 'Line Input'
        }

        It 'AI_BytesParaStringUTF8 em Mod_12 decodifica UTF-8 manualmente' {
            $mod12 = $script:moduleContent['Mod_12_AIStructure.bas']
            # Extrai o corpo da funcao AI_BytesParaStringUTF8
            $match = [regex]::Match($mod12, 'Private Function AI_BytesParaStringUTF8[\s\S]*?End Function')
            # Nao deve criar ADODB.Stream (uso real, nao comentarios)
            $match.Value | Should Not Match 'CreateObject.*ADODB'
            # Deve ter logica de decodificacao manual
            $match.Value | Should Match 'b And &HE0.*= &HC0'
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

        It 'AutoOpen esta em Mod_04_Main para registrar atalhos' {
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Sub AutoOpen\(\)'
        }

        It 'RegistrarAtalhosTeclado esta em Mod_04_Main' {
            $script:moduleContent['Mod_04_Main.bas'] | Should Match '(?m)^Public Sub RegistrarAtalhosTeclado\(\)'
        }

        It 'RegistrarAtalhosTeclado registra Alt+P e Alt+C' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            $mod04 | Should Match 'PadronizarDocumentoMain'
            $mod04 | Should Match 'CorrigirProposituraComIA'
            $mod04 | Should Match 'wdKeyAlt'
        }

        It 'CriarAtalhosTeclado.bas nao existe mais (integrado em Mod_04)' {
            $mainPath = Join-Path (Get-RepoRoot) 'source\main'
            Test-Path (Join-Path $mainPath 'CriarAtalhosTeclado.bas') | Should Be $false
        }

    }

    Context 'UndoRecord - Seguranca de pilha de desfazer' {
        It 'NAO usa doc.UndoClear em nenhum modulo (causa entradas fantasmas)' {
            foreach ($name in $script:moduleNames) {
                $content = $script:moduleContent[$name]
                # Permite doc.UndoClear APENAS em linhas de comentario
                $codeLines = ($content -split "`n") | Where-Object {
                    $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
                }
                $codeContent = $codeLines -join "`n"
                $codeContent | Should Not Match 'doc\.UndoClear'
            }
        }

        It 'PadronizarDocumentoMain USA StartCustomRecord (integracao undo)' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            $codeLines = ($mod04 -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            $codeContent | Should Match 'StartCustomRecord'
        }

        It 'PadronizarDocumentoMain USA EndCustomRecord (integracao undo)' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            $codeLines = ($mod04 -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            $codeContent | Should Match 'EndCustomRecord'
        }

        It 'PadronizarDocumentoMain seta undoRecordActive apos StartCustomRecord' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            $codeLines = ($mod04 -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            $codeContent | Should Match 'undoRecordActive\s*=\s*True'
        }

        It 'EmergencyRecovery fecha UndoRecord quando undoRecordActive' {
            $mod01 = $script:moduleContent['Mod_01_Infrastructure.bas']
            $codeLines = ($mod01 -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            # Deve ter bloco que verifica undoRecordActive no EmergencyRecovery
            $codeContent | Should Match 'If undoRecordActive Then'
        }

        It 'CleanUp fecha EndCustomRecord DEPOIS de ScreenRefresh' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            $codeLines = ($mod04 -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            # Deve ter ScreenRefresh no CleanUp
            $codeContent | Should Match 'Application\.ScreenRefresh'
            # Deve ter EndCustomRecord no CleanUp
            $codeContent | Should Match 'EndCustomRecord'
            # Deve ter undoRecordActive = False
            $codeContent | Should Match 'undoRecordActive\s*=\s*False'
            # Verifica ordem no CleanUp: ScreenRefresh DEVE vir ANTES de EndCustomRecord
            $screenRefreshLine = -1
            $endCustomLine = -1
            $undoResetLine = -1
            $cleanUpFound = $false
            for ($i = 0; $i -lt ($codeLines.Count); $i++) {
                $line = $codeLines[$i]
                if ($line -match '^\s*CleanUp:') { $cleanUpFound = $true }
                if ($cleanUpFound) {
                    if ($line -match 'Application\.ScreenRefresh' -and $screenRefreshLine -eq -1) {
                        $screenRefreshLine = $i
                    }
                    if ($line -match 'EndCustomRecord' -and $endCustomLine -eq -1) {
                        $endCustomLine = $i
                    }
                    if ($line -match 'undoRecordActive\s*=\s*False' -and $undoResetLine -eq -1) {
                        $undoResetLine = $i
                    }
                }
            }
            $screenRefreshLine | Should BeGreaterThan -1
            $endCustomLine | Should BeGreaterThan -1
            $undoResetLine | Should BeGreaterThan -1
            # Ordem obrigatoria: ScreenRefresh < EndCustomRecord < undoRecordActive=False
            $endCustomLine | Should BeGreaterThan $screenRefreshLine
            $undoResetLine | Should BeGreaterThan $endCustomLine
        }

        It 'Nenhuma operacao perigosa apos EndCustomRecord no CleanUp' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            $codeLines = ($mod04 -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            
            # Encontra o bloco CleanUp principal (dentro de PadronizarDocumentoMain)
            $cleanUpStart = -1
            $endCustomLine = -1
            for ($i = 0; $i -lt ($codeLines.Count); $i++) {
                if ($codeLines[$i] -match '^\s*CleanUp:' -and $cleanUpStart -eq -1) {
                    $cleanUpStart = $i
                }
                if ($cleanUpStart -gt -1 -and $codeLines[$i] -match 'EndCustomRecord') {
                    $endCustomLine = $i
                    break
                }
            }
            
            $endCustomLine | Should BeGreaterThan -1
            
            # Verifica APENAS as linhas entre EndCustomRecord e o proximo Exit Sub ou End Sub
            for ($i = $endCustomLine + 1; $i -lt ($codeLines.Count); $i++) {
                $line = $codeLines[$i]
                # Para no fim da sub
                if ($line -match '^\s*Exit Sub' -or $line -match '^\s*End Sub') {
                    break
                }
                if ($line -match '^\s*SafeFinalizeLogging') { break }
                # Nenhuma operacao que toque documento/tela/DoEvents deve ocorrer apos EndCustomRecord
                $line | Should Not Match 'Selection\.'
                $line | Should Not Match 'doc\.Save'
                $line | Should Not Match 'doc\.UndoClear'
                $line | Should Not Match 'Application\.ScreenRefresh'
                $line | Should Not Match 'DoEvents'
                $line | Should Not Match 'Application\.OnRepeat'
            }
        }

        It 'PadronizarDocumentoMain nao usa Selection em lugar algum' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            # PadronizarDocumentoMain nao deve usar Selection em lugar algum para evitar entradas parasitas
            # Excluir linhas de comentario (que comecam com ') da verificacao
            $codeLines = ($mod04 -split "`n") | Where-Object { 
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            $codeContent | Should Not Match 'Selection\.'
        }

        It 'Mod_04_Main nao contem Application.OnRepeat (proibido por causa de instabilidade da pilha de undo)' {
            $mod04 = $script:moduleContent['Mod_04_Main.bas']
            # Excluir linhas de comentario (que comecam com ') da verificacao
            $codeLines = ($mod04 -split "`n") | Where-Object { 
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            $codeContent | Should Not Match 'Application\.OnRepeat'
        }

        It 'CorrigirProposituraComIA tambem nao usa doc.UndoClear' {
            $content = $script:moduleContent['Mod_11_RevisionText.bas']
            $codeLines = ($content -split "`n") | Where-Object {
                $_ -notmatch '^\s*\x27' -and $_ -notmatch '^\s*Rem\s'
            }
            $codeContent = $codeLines -join "`n"
            $codeContent | Should Not Match 'doc\.UndoClear'
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
