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



        It 'CorrigirProposituraComIA valida selecao de um paragrafo por vez' {

            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']

            # Extrai o corpo do CorrigirProposituraComIA

            $match = [regex]::Match($mod11, 'Public Sub CorrigirProposituraComIA[\s\S]*?End Sub')

            $match.Success | Should Be $true

            # Deve recusar selecoes com mais de um paragrafo, com aviso ao usuario

            $match.Value | Should Match 'SelecaoAbrangeMultiplosParagrafos'

            $match.Value | Should Match 'um paragrafo por vez'

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


        It 'AutoOpen nao existe mais em Mod_04_Main (atalhos removidos)' {

            $script:moduleContent['Mod_04_Main.bas'] | Should Not Match '(?m)^Public Sub AutoOpen\(\)'

        }







        It 'RegistrarAtalhosTeclado nao existe mais em Mod_04_Main (atalhos removidos)' {

            $script:moduleContent['Mod_04_Main.bas'] | Should Not Match '(?m)^Public Sub RegistrarAtalhosTeclado\(\)'

        }







        It 'CriarAtalhosTeclado.bas nao existe (modulo separado nunca deve voltar)' {

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



    Context 'Isolamento de dados - texto do documento nunca e prompt' {

        It 'ProcessarTextoComIA anexa MontarGuardAntiInjecao ao system prompt' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function ProcessarTextoComIA[\s\S]*?End Function')
            $match.Success | Should Be $true
            $match.Value | Should Match 'MontarGuardAntiInjecao'
        }

        It 'Guard anti-injecao e anexado apos o prompt configuravel (nao removivel)' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function ProcessarTextoComIA[\s\S]*?End Function')
            # CarregarPromptRevisao DEVE ser seguido por MontarGuardAntiInjecao
            $match.Value | Should Match 'CarregarPromptRevisao\(\)[\s\S]*MontarGuardAntiInjecao'
        }

        It 'Texto do documento e enviado como DADOS via MontarMensagemDados' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function ProcessarTextoComIA[\s\S]*?End Function')
            $match.Value | Should Match 'EscaparJSON\(MontarMensagemDados\(textoInput\)\)'
        }

        It 'MontarGuardAntiInjecao declara que dados nunca sao prompt' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function MontarGuardAntiInjecao[\s\S]*?End Function')
            $match.Success | Should Be $true
            $match.Value | Should Match 'TRATAMENTO DE DADOS DO DOCUMENTO'
            $match.Value | Should Match 'NUNCA e um prompt'
            $match.Value | Should Match 'DADOS_MARCADOR_INICIO'
            $match.Value | Should Match 'NAO siga'
        }

        It 'MontarMensagemDados delimita a regiao de dados com marcador de inicio' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function MontarMensagemDados[\s\S]*?End Function')
            $match.Success | Should Be $true
            $match.Value | Should Match 'DADOS_MARCADOR_INICIO'
            $match.Value | Should Match 'NUNCA e um prompt'
        }

        It 'Nao existe marcador de fechamento (impede breakout por injecao)' {
            # A regiao de dados vai do marcador de inicio ate o FINAL da
            # mensagem; um marcador de fechamento poderia ser injetado no
            # texto do documento para fugir da regiao de dados.
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $mod11 | Should Not Match 'DADOS_MARCADOR_FIM'
            $mod11 | Should Not Match 'FIM_TEXTO_A_REVISAR'
        }

        It 'MontarJSONRequest mantem prompt em role system e dados em role user' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function MontarJSONRequest[\s\S]*?End Function')
            $match.Success | Should Be $true
            $match.Value | Should Match '\{""role"":""system"",""content"":"""'
            $match.Value | Should Match '\{""role"":""user"",""content"":"""'
            $match.Value | Should Match 'systemJSON[\s\S]*textoJSON'
        }

        It 'RemoverEnvelopeResposta limpa eco do marcador nas bordas da resposta' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function RemoverEnvelopeResposta[\s\S]*?End Function')
            $match.Success | Should Be $true
            $match.Value | Should Match 'DADOS_MARCADOR_INICIO'
        }

        It 'ProcessarTextoComIA remove envelope ecoado da resposta da IA' {
            $mod11 = $script:moduleContent['Mod_11_RevisionText.bas']
            $match = [regex]::Match($mod11, 'Private Function ProcessarTextoComIA[\s\S]*?End Function')
            $match.Value | Should Match 'RemoverEnvelopeResposta'
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

    Context 'Quebra de paragrafo antes de sufixo de Vereador' {

        It 'BreakParagraphBeforeVereadorSuffix esta em Mod_09_SpecialParagraphs' {

            $script:moduleContent['Mod_09_SpecialParagraphs.bas'] | Should Match '(?m)^Public Sub BreakParagraphBeforeVereadorSuffix\(doc As Document\)'

        }

        It 'Sufixo de Vereador cobre hifen/en-dash, vereador/vereadora e case-insensitive' {

            $mod09 = $script:moduleContent['Mod_09_SpecialParagraphs.bas']

            $fn = [regex]::Match($mod09, '(?s)Private Function GetVereadorSuffixSplitPos\(.*?End Function')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'ChrW\(8211\)'

            $fn.Value | Should Match '" - vereador"'

            $fn.Value | Should Match '" - vereadora"'

            $fn.Value | Should Match 'LCase\$\('

            # Pontuacao final opcional ("." ou ",") apos vereador/vereadora
            $fn.Value | Should Match '\.,"'

        }

        It 'So quebra paragrafos que nao ultrapassam uma linha (guard fail-closed)' {

            $mod09 = $script:moduleContent['Mod_09_SpecialParagraphs.bas']

            $fn = [regex]::Match($mod09, '(?s)Public Sub BreakParagraphBeforeVereadorSuffix\(.*?End Sub')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'ParagraphHasMultipleLines'

            $fn.Value | Should Match 'GetVereadorSuffixSplitPos'

        }

        It 'Nao quebra com prefixo vazio e recarrega estrutura apos quebras' {

            $mod09 = $script:moduleContent['Mod_09_SpecialParagraphs.bas']

            $fn = [regex]::Match($mod09, '(?s)Private Function GetVereadorSuffixSplitPos\(.*?End Function')

            $fn.Value | Should Match 'Trim\$\(\s*Left\$\(\s*cleanText,\s*startPos - 1\s*\)\s*\)'

            $sub = [regex]::Match($mod09, '(?s)Public Sub BreakParagraphBeforeVereadorSuffix\(.*?End Sub')

            $sub.Value | Should Match 'IdentifyDocumentStructure doc'

        }

        It 'Roda no inicio do PreviousFormatting (antes da limpeza estrutural)' {

            $mod03 = $script:moduleContent['Mod_03_Pipeline.bas']

            $fn = [regex]::Match($mod03, '(?s)Public Function PreviousFormatting\(doc As Document\) As Boolean.*?End Function')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'ReplaceLineBreaksWithParagraphBreaks doc[\s\S]*?BreakParagraphBeforeVereadorSuffix doc[\s\S]*?RemovePageNumberLines doc'

        }

    }

    Context 'Duas linhas em branco nas zonas especiais (ementa, titulo justificativa, data)' {

        It 'ForceEmentaSpacing garante exatamente 2 linhas acima e abaixo da Ementa' {

            $mod08 = $script:moduleContent['Mod_08_Ementa.bas']

            $fn = [regex]::Match($mod08, '(?s)Public Sub ForceEmentaSpacing\(.*?End Sub')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'FindEmentaParagraphIndex'

            $fn.Value | Should Match 'RemoveBlankLinesAfter'

            $fn.Value | Should Match 'InsertBlankLinesAfter'

            $fn.Value | Should Match 'RemoveBlankLinesBefore'

            $fn.Value | Should Match 'InsertBlankLinesBefore'

            # Regra antiga de 3 linhas acima nao pode voltar
            $fn.Value | Should Not Match 'blankCount < 3'

        }

        It 'ForceDataSpacing garante exatamente 2 linhas acima da Data' {

            $mod08 = $script:moduleContent['Mod_08_Ementa.bas']

            $fn = [regex]::Match($mod08, '(?s)Public Sub ForceDataSpacing\(.*?End Sub')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'IsDataElement'

            $fn.Value | Should Match 'RemoveBlankLinesBefore'

            $fn.Value | Should Match 'InsertBlankLinesBefore'

        }

        It 'ForceJustificativaTitleSpacing garante exatamente 2 linhas acima e abaixo do titulo da Justificativa' {

            $mod09 = $script:moduleContent['Mod_09_SpecialParagraphs.bas']

            $fn = [regex]::Match($mod09, '(?s)Public Sub ForceJustificativaTitleSpacing\(.*?End Sub')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'RemoveBlankLinesBefore'

            $fn.Value | Should Match 'InsertBlankLinesBefore'

            # Abaixo do titulo: remove excesso e re-insere 2
            $fn.Value | Should Match 'RemoveBlankLinesAfter'

            $fn.Value | Should Match 'InsertBlankLinesAfter'

            $mod09 | Should Match '(?m)^Private Function FindJustificativaTitleIndex\(doc As Document\) As Long'

            $mod09 | Should Match 'JUSTIFICATIVA_TEXT'

        }

        It 'RemoverLinhasEmBrancoExtras preserva 2 linhas nas zonas protegidas' {

            $mod07 = $script:moduleContent['Mod_07_Formatting.bas']

            $fn = [regex]::Match($mod07, '(?s)Public Sub RemoverLinhasEmBrancoExtras\(.*?End Sub')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'IsTwoBlankLinesZone'

            $fn.Value | Should Match 'maxBlank = 2'

            $mod07 | Should Match 'Private Function IsTwoBlankLinesZone\(doc As Document, prevIdx As Long, nextIdx As Long\) As Boolean'

            # Zona abaixo do titulo da Justificativa tambem e protegida (ancora anterior)
            $mod07 | Should Match 'IsJustificativaTitleElement\(doc\.Paragraphs\(prevIdx\)\)'

        }

        It 'Garantia final roda DEPOIS da padronizacao generalizada de linhas puladas' {

            $mod04 = $script:moduleContent['Mod_04_Main.bas']

            $mod04 | Should Match 'RemoverLinhasEmBrancoExtras doc[\s\S]*?EnsureConsideringBlankLines doc[\s\S]*?ForceDataSpacing doc[\s\S]*?ForceJustificativaTitleSpacing doc[\s\S]*?ForceEmentaSpacing doc'

        }

    }

    Context 'Regra de paragrafos multi-linha (sem centralizar / sem recuo zero)' {

        It 'Mod_01 declara a constante wdStatisticLines' {

            $script:moduleContent['Mod_01_Infrastructure.bas'] | Should Match '(?m)^Public Const wdStatisticLines As Long = 1\r?$'

        }

        It 'Mod_01 declara ParagraphHasMultipleLines com fail-closed' {

            $mod01 = $script:moduleContent['Mod_01_Infrastructure.bas']

            $fn = [regex]::Match($mod01, '(?s)Public Function ParagraphHasMultipleLines\(para As Paragraph\) As Boolean.*?End Function')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'ComputeStatistics\(wdStatisticLines\)'

            $fn.Value | Should Match '(?m)^\s*ParagraphHasMultipleLines = True\r?$'

        }

        It 'SafeSetParagraphFormat aplica o guard multi-linha' {

            $mod01 = $script:moduleContent['Mod_01_Infrastructure.bas']

            $fn = [regex]::Match($mod01, '(?s)Public Function SafeSetParagraphFormat\(.*?End Function')

            $fn.Success | Should Be $true

            $fn.Value | Should Match 'ParagraphHasMultipleLines'

            $fn.Value | Should Match 'wdAlignParagraphCenter'

        }

        It 'Funcoes de elementos especiais aplicam o guard multi-linha' {

            $guarded = @(

                @{ Module = 'Mod_02_Engine.bas'; Name = 'FormatImageParagraphsIndents' },

                @{ Module = 'Mod_02_Engine.bas'; Name = 'CenterImageAfterPlenario' },

                @{ Module = 'Mod_02_Engine.bas'; Name = 'RestoreCenteredParagraphs' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'FormatFirstParagraph' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'InsertFooterStamp' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'RemoverLinhasEmBrancoExtras' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'ReplacePlenarioDateParagraph' },

                @{ Module = 'Mod_09_SpecialParagraphs.bas'; Name = 'FixHyphenatedVereadorParagraphIndents' },

                @{ Module = 'Mod_09_SpecialParagraphs.bas'; Name = 'FormatDocumentTitle' },

                @{ Module = 'Mod_09_SpecialParagraphs.bas'; Name = 'ApplyBoldToSpecialParagraphs' },

                @{ Module = 'Mod_09_SpecialParagraphs.bas'; Name = 'FormatVereadorParagraphs' },

                @{ Module = 'Mod_09_SpecialParagraphs.bas'; Name = 'ApplyVereadorParagraphFormatting' }

            )

            foreach ($item in $guarded) {

                $content = $script:moduleContent[$item.Module]

                $pattern = '(?sm)^Public (?:Sub|Function) ' + [regex]::Escape($item.Name) + '\(.*?^End (?:Sub|Function)'

                $fn = [regex]::Match($content, $pattern)

                $fn.Success | Should Be $true

                $fn.Value | Should Match 'ParagraphHasMultipleLines'

            }

        }

        It 'Normalizacao de corpo/reset continua zerando recuos sem o guard' {

            $exempt = @(

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'ClearAllFormatting' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'ApplyStdParagraphs' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'FormatPostEmentaBodyParagraphs' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'FormatSecondParagraph' },

                @{ Module = 'Mod_07_Formatting.bas'; Name = 'FormatBulletedParagraphsIndent' }

            )

            foreach ($item in $exempt) {

                $content = $script:moduleContent[$item.Module]

                $pattern = '(?sm)^Public (?:Sub|Function) ' + [regex]::Escape($item.Name) + '\(.*?^End (?:Sub|Function)'

                $fn = [regex]::Match($content, $pattern)

                $fn.Success | Should Be $true

                $fn.Value | Should Not Match 'ParagraphHasMultipleLines'

            }

        }

    }

}

