Attribute VB_Name = "Mod8SpecialParagraphs"
Option Explicit

' Mod8SpecialParagraphs
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)

Public Sub FixHyphenatedVereadorParagraphIndents(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim paraText As String
    Dim counter As Long
    Dim fixedCount As Long
    Dim blankRemovedCount As Long
    Dim i As Long

    fixedCount = 0
    blankRemovedCount = 0

    For i = 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For
        If i Mod 30 = 0 Then DoEvents

        Set para = doc.Paragraphs(i)
        paraText = Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")

        ' Normaliza espacos/tabs e hifens/travessoes para detectar o conteudo desejado
        Dim normText As String
        normText = Replace(paraText, vbTab, " ")
        normText = Replace(normText, ChrW(8209), "-") ' non-breaking hyphen
        normText = Replace(normText, ChrW(8211), "-") ' en dash
        normText = Replace(normText, ChrW(8212), "-") ' em dash
        normText = Replace(normText, ChrW(8722), "-") ' minus sign
        normText = Trim$(normText)
        Do While InStr(normText, "  ") > 0
            normText = Replace(normText, "  ", " ")
        Loop

        If normText = "- Vereador -" Or normText = "- Vereadora -" Or IsVereadorPattern(paraText) Then
            ' Remove paragrafos em branco imediatamente acima do "Vereador"
            Do While i > 1
                Set prevPara = doc.Paragraphs(i - 1)
                Dim prevClean As String
                prevClean = Trim(Replace(Replace(prevPara.Range.text, vbCr, ""), vbLf, ""))
                If prevClean = "" And Not HasVisualContent(prevPara) Then
                    prevPara.Range.Delete
                    i = i - 1
                    blankRemovedCount = blankRemovedCount + 1
                Else
                    Exit Do
                End If
            Loop

            ' Reobtem referencia apos possiveis remocoes
            If i > doc.Paragraphs.count Then Exit For
            Set para = doc.Paragraphs(i)

            On Error Resume Next

            With para.Format
                .leftIndent = 0
                .firstLineIndent = 0
                .RightIndent = 0
            End With

            With para.Range.ParagraphFormat
                .leftIndent = 0
                .firstLineIndent = 0
                .RightIndent = 0
            End With

            On Error GoTo ErrorHandler
            fixedCount = fixedCount + 1
        End If

        ' GARANTIA: Remove paragrafos em branco acima de vocativos (Senhores Vereadores,
        ' Senhoras Vereadoras, Senhores(as) Vereadores(as), Senhora Vereadora) quando
        ' estao separados do paragrafo exclusivo "Senhor Presidente,"
        If IsVocativoVereadorPresidenteParagraph(paraText) And i > 1 Then
            Dim vocCheckIdx As Long
            Dim vocBlankCount As Long
            Dim vocCheckPara As Paragraph
            Dim vocCheckText As String
            vocCheckIdx = i - 1
            vocBlankCount = 0

            ' Conta paragrafos em branco imediatamente acima
            Do While vocCheckIdx >= 1
                Set vocCheckPara = doc.Paragraphs(vocCheckIdx)
                vocCheckText = Trim(Replace(Replace(vocCheckPara.Range.text, vbCr, ""), vbLf, ""))
                If vocCheckText = "" And Not HasVisualContent(vocCheckPara) Then
                    vocBlankCount = vocBlankCount + 1
                    vocCheckIdx = vocCheckIdx - 1
                Else
                    Exit Do
                End If
            Loop

            ' Se encontrou paragrafos em branco e o paragrafo acima e "Senhor Presidente,"
            If vocBlankCount > 0 And vocCheckIdx >= 1 Then
                If IsSenhorPresidenteParagraph(doc.Paragraphs(vocCheckIdx).Range.text) Then
                    ' Remove os paragrafos em branco entre "Senhor Presidente," e o vocativo
                    Do While i > vocCheckIdx + 1
                        Set prevPara = doc.Paragraphs(i - 1)
                        prevPara.Range.Delete
                        i = i - 1
                        blankRemovedCount = blankRemovedCount + 1
                    Loop

                    ' Reobtem referencia apos possiveis remocoes
                    If i > doc.Paragraphs.count Then Exit For
                    Set para = doc.Paragraphs(i)

                    LogMessage "Paragrafos em branco entre 'Senhor Presidente,' e vocativo removidos (final): " & vocBlankCount, LOG_LEVEL_INFO
                End If
            End If
        End If
    Next i

    If blankRemovedCount > 0 Then
        LogMessage "Paragrafos em branco acima de 'Vereador' removidos (final): " & blankRemovedCount, LOG_LEVEL_INFO
    End If

    If fixedCount > 0 Then
        LogMessage "Recuos ajustados para " & fixedCount & " paragrafo(s) de Vereador/Vereadora", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao ajustar recuos de Vereador/Vereadora: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORCE EMENTA SPACING - Garante 3 paragrafos em branco acima e 2 abaixo da Ementa
' Executada como ULTIMA etapa para nao ser desfeita por processamento posterior.
'================================================================================

Public Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    Dim testRange As Range

    ' Verifica se o documento esta protegido
    If doc.ProtectionType <> wdNoProtection Then
        LogMessage "Documento protegido - formatacao de titulo ignorada", LOG_LEVEL_INFO
        FormatDocumentTitle = True
        Exit Function
    End If

    ' Testa se e possivel editar o primeiro paragrafo
    On Error Resume Next
    Set testRange = doc.Paragraphs(1).Range
    If testRange Is Nothing Then
        Err.Clear
        On Error GoTo ErrorHandler
        LogMessage "Range invalido - formatacao de titulo ignorada", LOG_LEVEL_INFO
        FormatDocumentTitle = True
        Exit Function
    End If
    ' Tenta modificar uma propriedade para verificar acesso de escrita
    Dim originalBold As Boolean
    originalBold = testRange.Font.Bold
    testRange.Font.Bold = originalBold
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        LogMessage "Selecao protegida - formatacao de titulo ignorada", LOG_LEVEL_INFO
        FormatDocumentTitle = True
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Encontra o primeiro paragrafo com texto (apos exclusao de linhas em branco)
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i

    If paraText = "" Then
        LogMessage "Nenhum texto encontrado para formatacao do titulo", LOG_LEVEL_WARNING
        FormatDocumentTitle = True
        Exit Function
    End If

    ' Remove ponto final se existir
    If Right(paraText, 1) = "." Then
        paraText = Left(paraText, Len(paraText) - 1)
    End If

    ' Verifica se e uma proposicao (para aplicar substituicao $NUMERO$/$ANO$)
    Dim isProposition As Boolean
    Dim firstWord As String

    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
        ' Compara diretamente com as formas acentuadas e nao-acentuadas
        If firstWord = "indica" & Chr(231) & Chr(227) & "o" _
        Or firstWord = "indicacao" _
        Or firstWord = "requerimento" _
        Or firstWord = "mo" & Chr(231) & Chr(227) & "o" _
        Or firstWord = "mocao" Then
            isProposition = True
        End If
    End If

    ' Se for proposicao, substitui pelo padrao implementado (sempre com o final padronizado)
    If isProposition Then
        ' Isola a parte textual do titulo, ignorando formatacoes numericas antigas
        Dim baseText As String
        baseText = paraText
        
        ' Remover sufixos padronizados se existirem, para nao acumular
        baseText = Replace(baseText, "$NUMERO$/$ANO$", "")
        baseText = Trim(baseText)
        
        ' Remover ultimos "palavras" se forem numeros ou fracoes irrelevantes
        Do While True
            words = Split(baseText, " ")
            If UBound(words) <= 0 Then Exit Do
            Dim lastW As String
            lastW = words(UBound(words))
            
            ' Verifica se a ultima palavra e uma abreviacao/indicador de numero ou fracao
            Dim isNumIndicator As Boolean
            isNumIndicator = False
            
            Dim uLastW As String
            uLastW = UCase(lastW)
            
            If IsNumeric(lastW) Or InStr(lastW, "/") > 0 Then
                isNumIndicator = True
            ElseIf uLastW = "N" Or uLastW = "N." Or uLastW = "NO" Or uLastW = "N.O" Or uLastW = "NO." Then
                isNumIndicator = True
            ElseIf uLastW = "N" & Chr(186) Or uLastW = "N" & Chr(176) Then
                isNumIndicator = True
            ElseIf uLastW = "N." & Chr(186) Or uLastW = "N." & Chr(176) Then
                isNumIndicator = True
            End If
            
            If isNumIndicator Then
                baseText = Left(baseText, Len(baseText) - Len(lastW))
                baseText = Trim(baseText)
            Else
                Exit Do
            End If
        Loop
        
        newText = baseText & " N" & Chr(186) & " $NUMERO$/$ANO$"
    Else
        ' Se nao for proposicao, mantem o texto original
        newText = paraText
    End If

    ' SEMPRE aplica formatacao de titulo: caixa alta, negrito, sublinhado (preserva marca de paragrafo)
    Dim titleRng As Range
    Set titleRng = firstPara.Range
    If titleRng.End > titleRng.Start Then
        titleRng.End = titleRng.End - 1
    End If
    titleRng.text = UCase(newText)

    ' Formatacao completa do titulo (primeira linha)
    With firstPara.Range.Font
        .Bold = True
        .Underline = wdUnderlineSingle
        .AllCaps = True
    End With

    With firstPara.Format
        .alignment = wdAlignParagraphCenter
        .leftIndent = 0
        .firstLineIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 6  ' Pequeno espaco apos o titulo
    End With

    If isProposition Then
        LogMessage "Titulo de proposicao formatado: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    Else
        LogMessage "Primeira linha formatada como titulo: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    End If

    FormatDocumentTitle = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao do titulo: " & Err.Description, LOG_LEVEL_ERROR
    FormatDocumentTitle = False
End Function

'================================================================================
' FORMATACAO DE PARAGRAFOS "CONSIDERANDO" E "ANTE O EXPOSTO"
'================================================================================

Public Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim anteExpostoFormatted As Long
    Dim i As Long
    Dim nextChar As String

    ' Percorre todos os paragrafos procurando por "considerando" ou "ante o exposto" no inicio
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Verifica se o paragrafo comeca com "considerando" (ignorando maiusculas/minusculas)
        If Len(paraText) >= 12 And LCase(Left(paraText, 12)) = "considerando" Then
            ' Verifica se apos "considerando" vem espaco, virgula, ponto-e-virgula ou fim da linha
            If Len(paraText) > 12 Then
                nextChar = Mid(paraText, 13, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    ' E realmente "considerando" no inicio do paragrafo
                    Set rng = para.Range

                    ' CORRECAO: Usa Find/Replace para preservar espacamento
                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "considerando"
                        .Replacement.text = "CONSIDERANDO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False
                        .Forward = True
                        .Wrap = wdFindStop

                        ' Limita a busca ao inicio do paragrafo
                        rng.End = rng.Start + 15

                        If .Execute(Replace:=True) Then
                            totalFormatted = totalFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Paragrafo contem apenas "considerando"
                Set rng = para.Range
                rng.End = rng.Start + 12

                With rng
                    .text = "CONSIDERANDO"
                    .Font.Bold = True
                End With

                totalFormatted = totalFormatted + 1
            End If

        ' Verifica se o paragrafo comeca com "ante o exposto" (14 caracteres)
        ElseIf Len(paraText) >= 14 And LCase(Left(paraText, 14)) = "ante o exposto" Then
            ' Verifica se apos "ante o exposto" vem espaco, virgula, ponto-e-virgula ou fim
            If Len(paraText) > 14 Then
                nextChar = Mid(paraText, 15, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    Set rng = para.Range

                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "ante o exposto"
                        .Replacement.text = "ANTE O EXPOSTO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False
                        .Forward = True
                        .Wrap = wdFindStop

                        rng.End = rng.Start + 17

                        If .Execute(Replace:=True) Then
                            anteExpostoFormatted = anteExpostoFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Paragrafo contem apenas "ante o exposto"
                Set rng = para.Range
                rng.End = rng.Start + 14

                With rng
                    .text = "ANTE O EXPOSTO"
                    .Font.Bold = True
                End With

                anteExpostoFormatted = anteExpostoFormatted + 1
            End If
        End If
    Next i

    If totalFormatted > 0 Then
        LogMessage "Formatacao 'CONSIDERANDO' aplicada: " & totalFormatted & " ocorrencia(s)", LOG_LEVEL_INFO
    End If
    If anteExpostoFormatted > 0 Then
        LogMessage "Formatacao 'ANTE O EXPOSTO' aplicada: " & anteExpostoFormatted & " ocorrencia(s)", LOG_LEVEL_INFO
    End If

    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao CONSIDERANDO/ANTE O EXPOSTO: " & Err.Description, LOG_LEVEL_ERROR
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' FUNCAO AUXILIAR DE FIND/REPLACE - Elimina codigo repetitivo
'================================================================================

Public Function ExecuteFindReplace(doc As Document, searchText As String, replaceText As String, Optional matchCase As Boolean = False, Optional maxIterations As Long = 500) As Long
    ' Retorna quantidade de substituicoes realizadas
    On Error Resume Next
    ExecuteFindReplace = 0

    If doc Is Nothing Then Exit Function
    If searchText = "" Then Exit Function

    Dim rng As Range
    Set rng = doc.Range
    If rng Is Nothing Then Exit Function

    Dim iterCount As Long
    iterCount = 0

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = searchText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True) And iterCount < maxIterations
            iterCount = iterCount + 1
            ExecuteFindReplace = ExecuteFindReplace + 1
        Loop
    End With

    Err.Clear
End Function

'================================================================================
' FORMATACAO DE "IN LOCO" EM ITALICO (REMOVE ASPAS)
'================================================================================

Public Sub FormatInLocoItalic(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim rng As Range
    Dim quotesRemovedCount As Long
    Dim italicAppliedCount As Long
    quotesRemovedCount = 0
    italicAppliedCount = 0

    ' (1) Remove aspas envolvendo a expressao (inclui aspas retas e tipograficas)
    '    Ex.: "in loco" (inclui aspas tipograficas) / "in loco," -> in loco / in loco,
    Set rng = doc.Range

    Dim quoteChars As String
    quoteChars = Chr(34) & ChrW(8220) & ChrW(8221)

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' Word Wildcards nao suportam quantificadores do tipo {0,1}.
        ' Faz 2 passes: (a) com pontuacao dentro das aspas; (b) sem pontuacao.
        .text = "[" & quoteChars & "]([Ii]n loco)([,.;:])[" & quoteChars & "]"
        .Replacement.text = "\1\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True

        Do While .Execute(Replace:=wdReplaceOne)
            quotesRemovedCount = quotesRemovedCount + 1
            If quotesRemovedCount > 200 Then Exit Do  ' Limite de seguranca
        Loop
    End With

    Set rng = doc.Range

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "[" & quoteChars & "]([Ii]n loco)[" & quoteChars & "]"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True

        Do While .Execute(Replace:=wdReplaceOne)
            quotesRemovedCount = quotesRemovedCount + 1
            If quotesRemovedCount > 200 Then Exit Do  ' Limite de seguranca
        Loop
    End With

    ' (2) Garante italico em todas as ocorrencias (com ou sem aspas)
    Set rng = doc.Range

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "in loco"
        .Replacement.text = "^&" ' Mantem o texto encontrado; aplica apenas formatacao
        .Replacement.Font.Italic = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False

        Do While .Execute(Replace:=wdReplaceOne)
            italicAppliedCount = italicAppliedCount + 1
            If italicAppliedCount > 500 Then Exit Do  ' Limite de seguranca
        Loop
    End With

    If quotesRemovedCount > 0 Or italicAppliedCount > 0 Then
        LogMessage "Formatacao 'in loco' aplicada: " & italicAppliedCount & " ocorrencia(s) em italico; aspas removidas: " & quotesRemovedCount & "x", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar 'in loco': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' APLICACAO DE SUBSTITUICOES DE TEXTO
'================================================================================

Public Function ApplyTextReplacements(doc As Document) As Boolean
    Dim errorContext As String
    Dim i As Long  ' Movida para escopo de funcao
    On Error GoTo ErrorHandler

    ' Validacao de documento
    If Not ValidateDocument(doc) Then
        ApplyTextReplacements = False
        Exit Function
    End If

    ' Verifica se ha conteudo suficiente
    If doc.Range.text = "" Or Len(Trim(doc.Range.text)) <= 1 Then
        LogMessage "Documento vazio - substituicoes de texto ignoradas", LOG_LEVEL_INFO
        ApplyTextReplacements = True
        Exit Function
    End If

    Dim rng As Range
    Dim replacementCount As Long
    Dim wasReplaced As Boolean
    Dim totalReplacements As Long
    totalReplacements = 0

    ' Funcionalidade 10: Substitui variantes de "d'Oeste"
    Dim dOesteVariants() As String

    ' Define as variantes possiveis dos 3 primeiros caracteres de "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "d" & Chr(180) & "O"   ' Acento agudo (Chr 180)
    dOesteVariants(2) = "d`O"   ' Acento grave
    dOesteVariants(3) = "d" & ChrW(8220) & "O"   ' Aspas curvas esquerda (Unicode)
    dOesteVariants(4) = "d'o"   ' Minuscula
    dOesteVariants(5) = "d" & Chr(180) & "o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & ChrW(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Maiuscula no D
    dOesteVariants(9) = "D" & Chr(180) & "O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & ChrW(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "D" & Chr(180) & "o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & ChrW(8220) & "o"

    ' Valida o array antes de processar
    On Error Resume Next
    Dim arraySize As Long
    arraySize = UBound(dOesteVariants)
    If Err.Number <> 0 Or arraySize < 0 Then
        LogMessage "Erro ao inicializar array de variantes - substituicoes de texto ignoradas", LOG_LEVEL_WARNING
        Err.Clear
        ApplyTextReplacements = True
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Processa cada variante de forma segura
    For i = 0 To arraySize
        On Error Resume Next
        errorContext = "dOesteVariants(" & i & ")"
        ' Valida a variante antes de usar
        If IsEmpty(dOesteVariants(i)) Or dOesteVariants(i) = "" Then
            GoTo NextVariant
        End If
        ' Cria novo range para cada busca
        Set rng = Nothing
        Set rng = doc.Range
        ' Verifica se o range foi criado com sucesso
        If rng Is Nothing Then GoTo NextVariant
        ' Configura os parametros de busca e substituicao
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = dOesteVariants(i) & "este"
            .Replacement.text = "d'Oeste"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            ' Executa a substituicao e armazena resultado booleano
            wasReplaced = .Execute(Replace:=wdReplaceAll)
            ' Verifica se houve erro
            If Err.Number = 0 Then
                If wasReplaced Then
                    totalReplacements = totalReplacements + 1
                End If
            Else
                If Err.Number <> 0 Then
                    LogMessage "Aviso ao substituir variante #" & i & " ('" & dOesteVariants(i) & "este'): " & Err.Description, LOG_LEVEL_WARNING
                End If
                Err.Clear
            End If
        End With
NextVariant:
        On Error GoTo ErrorHandler
        Err.Clear
    Next i

    If totalReplacements > 0 Then
        LogMessage "Substituicoes de texto aplicadas: " & totalReplacements & " variante(s) substituida(s)", LOG_LEVEL_INFO
    Else
        LogMessage "Substituicoes de texto: nenhuma ocorrencia encontrada", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 11: Substitui " ao Setor, " por " ao setor competente"
    Dim setorCount As Long
    setorCount = ExecuteFindReplace(doc, " ao Setor, ", " ao setor competente", True)
    If setorCount > 0 Then
        LogMessage "Substituicao aplicada: ' ao Setor, ' -> ' ao setor competente' (" & setorCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 12: Substitui " Setor Competente " por " setor competente " (case insensitive)
    Dim competenteCount As Long
    competenteCount = ExecuteFindReplace(doc, " Setor Competente ", " setor competente ", False)
    If competenteCount > 0 Then
        LogMessage "Substituicao aplicada: ' Setor Competente ' -> ' setor competente ' (" & competenteCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 15: No 1 paragrafo apos a ementa, normaliza "... para sugerir" -> "... para indicar"
    Dim art108IndicarCount As Long
    art108IndicarCount = NormalizeArt108ParaIndicarAfterEmenta(doc)
    If art108IndicarCount > 0 Then
        LogMessage "Substituicao aplicada: 'para sugerir' -> 'para indicar' no 1 paragrafo apos a ementa (" & art108IndicarCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 15: Normaliza a abertura do Art. 108 no 1 paragrafo apos a ementa
    Dim art108Count As Long
    art108Count = NormalizeArt108IntroAfterEmenta(doc)
    If art108Count > 0 Then
        LogMessage "Substituicao aplicada: abertura Art. 108 normalizada no 1 paragrafo apos a ementa (" & art108Count & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 13: Normaliza variantes de "tapa-buracos"
    Dim tapaBuracosCount As Long
    tapaBuracosCount = 0
    ' Com aspas
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa buraco" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa buracos" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa-buraco" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa-buracos" & Chr(34), "tapa-buracos", False)
    ' Com aspas mistas (simples e duplas)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(39) & "tapa buraco" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa buraco" & Chr(39), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(39) & "tapa buracos" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa buracos" & Chr(39), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(39) & "tapa-buraco" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa-buraco" & Chr(39), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(39) & "tapa-buracos" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa-buracos" & Chr(39), "tapa-buracos", False)
    ' Sem aspas (ordem importa: primeiro os com hifen para evitar duplicacao)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, "tapa-buraco ", "tapa-buracos ", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, "tapa buracos", "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, "tapa buraco", "tapa-buracos", False)
    If tapaBuracosCount > 0 Then
        LogMessage "Substituicao aplicada: variantes de 'tapa-buracos' normalizadas (" & tapaBuracosCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 14: Substitui "in loco" (com aspas) por in loco (italico, sem aspas)
    FormatInLocoItalic doc

    ' Funcionalidade 16: "�rea P�blica" e "Ro�agem" sempre em minusculas (ChrW seguro)
    Dim areaPublicaCount As Long
    Dim rocagemCount As Long
    areaPublicaCount = ExecuteFindReplace(doc, "" & ChrW(193) & "rea P" & ChrW(250) & "blica", "" & ChrW(225) & "rea p" & ChrW(250) & "blica", True)
    areaPublicaCount = areaPublicaCount + ExecuteFindReplace(doc, "" & ChrW(193) & "rea p" & ChrW(250) & "blica", "" & ChrW(225) & "rea p" & ChrW(250) & "blica", True)
    areaPublicaCount = areaPublicaCount + ExecuteFindReplace(doc, "Area Publica", "" & ChrW(225) & "rea p" & ChrW(250) & "blica", True)
    areaPublicaCount = areaPublicaCount + ExecuteFindReplace(doc, "Area p" & ChrW(250) & "blica", "" & ChrW(225) & "rea p" & ChrW(250) & "blica", True)
    rocagemCount = ExecuteFindReplace(doc, "Ro" & ChrW(231) & "agem", "ro" & ChrW(231) & "agem", True)
    rocagemCount = rocagemCount + ExecuteFindReplace(doc, "Rocagem", "ro" & ChrW(231) & "agem", True)
    If areaPublicaCount > 0 Or rocagemCount > 0 Then
        LogMessage "Substituicao aplicada: 'Area Publica' e 'Rocagem' em minusculas (" & (areaPublicaCount + rocagemCount) & "x)", LOG_LEVEL_INFO
    End If

    ' Normaliza "Bairro" / "bairro:" / "Bairro:" para "bairro" (sem maiuscula, sem dois-pontos)
    Dim bairroCount As Long
    bairroCount = ExecuteFindReplace(doc, "Bairro:", "bairro", True)
    bairroCount = bairroCount + ExecuteFindReplace(doc, "bairro:", "bairro", True)
    bairroCount = bairroCount + ExecuteFindReplace(doc, "Bairro", "bairro", True)
    If bairroCount > 0 Then
        LogMessage "Substituicao aplicada: 'Bairro'/'bairro:'/'Bairro:' -> 'bairro' (" & bairroCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 17: "retorne � esta Casa de Leis com as seguintes respostas" -> "retorne � esta Casa de Leis com as seguintes informa��es" (ChrW seguro)
    Dim casaLeisRespostasCount As Long
    casaLeisRespostasCount = ExecuteFindReplace(doc, "retorne " & ChrW(224) & " esta Casa de Leis com as seguintes respostas", "retorne " & ChrW(224) & " esta Casa de Leis com as seguintes informa" & ChrW(231) & ChrW(245) & "es", False)
    If casaLeisRespostasCount > 0 Then
        LogMessage "Substituicao aplicada: 'retorne a esta Casa de Leis com as seguintes respostas' -> 'retorne a esta Casa de Leis com as seguintes informacoes' (" & casaLeisRespostasCount & "x)", LOG_LEVEL_INFO
    End If

    ' Substitui " Jd " por " Jd. "
    Dim jdCount As Long
    jdCount = ExecuteFindReplace(doc, " Jd ", " Jd. ", True)
    If jdCount > 0 Then
        LogMessage "Substituicao aplicada: ' Jd ' -> ' Jd. ' (" & jdCount & "x)", LOG_LEVEL_INFO
    End If

    ' Substitui " aos n� " / " aos n� " por " ao n� " / " ao n� "
    Dim aosNoCount As Long
    aosNoCount = ExecuteFindReplace(doc, " aos n" & Chr(186) & " ", " ao n" & Chr(186) & " ", True)
    aosNoCount = aosNoCount + ExecuteFindReplace(doc, " aos n" & Chr(176) & " ", " ao n" & Chr(176) & " ", True)
    If aosNoCount > 0 Then
        LogMessage "Substituicao aplicada: ' aos n� ' -> ' ao n� ' (" & aosNoCount & "x)", LOG_LEVEL_INFO
    End If

    ' Substitui " nos n� " / " nos n� " por " no n� " / " no n� "
    Dim nosNoCount As Long
    nosNoCount = ExecuteFindReplace(doc, " nos n" & Chr(186) & " ", " no n" & Chr(186) & " ", True)
    nosNoCount = nosNoCount + ExecuteFindReplace(doc, " nos n" & Chr(176) & " ", " no n" & Chr(176) & " ", True)
    If nosNoCount > 0 Then
        LogMessage "Substituicao aplicada: ' nos n� ' -> ' no n� ' (" & nosNoCount & "x)", LOG_LEVEL_INFO
    End If

    ' Substitui N� por n� exceto no titulo
    ReplaceNoWithNoExceptTitle doc

    ' Substitui par�grafos contendo unicamente a string 'tikinho tk"'
    ReplaceTikinhoTkParagraphs doc
    
    ' Garante espa�o n�o separ�vel ap�s n�/n� antes de algarismos
    EnsureNonBreakingSpaceAfterNo doc

    ' Substitui todos os espa�os n�o separ�veis por espa�os comuns, exceto ap�s n�/n� antes de algarismos
    ReplaceNonBreakingSpacesExceptAfterNo doc

    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    Dim errMsg As String
    errMsg = Err.Description
    If Len(errorContext) > 0 Then
        LogMessage "Erro nas substituicoes de texto (contexto: " & errorContext & "): " & errMsg, LOG_LEVEL_WARNING
    ElseIf i >= 0 And i <= 15 Then
        LogMessage "Erro nas substituicoes de texto (variante: " & CStr(i) & "): " & errMsg, LOG_LEVEL_WARNING
    Else
        LogMessage "Erro nas substituicoes de texto: " & errMsg, LOG_LEVEL_WARNING
    End If
    ' Continua execucao - erros de substituicao nao sao criticos
    ApplyTextReplacements = True
End Function

'================================================================================
' CAPITALIZAR INICIO DE PARAGRAFOS
' Garantia: Qualquer paragrafo iniciado com a primeira letra em minuscula
' deve ter a primeira letra substituida por maiuscula (ex: "nos termos..." -> "Nos termos...").
'================================================================================

Public Function CapitalizeFirstLetterOfParagraphs(doc As Document) As Long
    On Error GoTo ErrorHandler

    CapitalizeFirstLetterOfParagraphs = 0
    If doc Is Nothing Then Exit Function

    Dim para As Paragraph
    Dim paraText As String
    Dim textLen As Long
    Dim k As Long
    Dim ch As String
    Dim capitalizedCount As Long
    Dim paraCounter As Long

    capitalizedCount = 0
    paraCounter = 0

    For Each para In doc.Paragraphs
        paraCounter = paraCounter + 1
        If paraCounter Mod 30 = 0 Then DoEvents

        ' Pula paragrafos sem conteudo visual (imagem/shape)
        If HasVisualContent(para) Then GoTo NextParagraph

        paraText = para.Range.text
        textLen = Len(paraText)
        If textLen = 0 Then GoTo NextParagraph

        ' Procura a primeira letra alfabetica do paragrafo
        For k = 1 To textLen
            ch = Mid$(paraText, k, 1)

            ' Verifica se eh caractere de letra (LCase != UCase)
            If LCase$(ch) <> UCase$(ch) Then
                ' Se a letra for minuscula
                If ch = LCase$(ch) Then
                    On Error Resume Next
                    Dim charRng As Range
                    Set charRng = doc.Range(para.Range.Start + (k - 1), para.Range.Start + k)
                    charRng.text = UCase$(ch)
                    If Err.Number = 0 Then
                        capitalizedCount = capitalizedCount + 1
                        documentDirty = True
                    End If
                    Err.Clear
                    On Error GoTo ErrorHandler
                End If
                ' Interrompe busca no paragrafo ao encontrar a primeira letra
                Exit For
            End If
        Next k

NextParagraph:
    Next para

    If capitalizedCount > 0 Then
        LogMessage "Capitalizacao aplicada: " & capitalizedCount & " paragrafo(s) iniciado(s) com minuscula corrigidos", LOG_LEVEL_INFO
    End If

    CapitalizeFirstLetterOfParagraphs = capitalizedCount
    Exit Function

ErrorHandler:
    LogMessage "Erro ao capitalizar inicio de paragrafos: " & Err.Description, LOG_LEVEL_WARNING
    CapitalizeFirstLetterOfParagraphs = capitalizedCount
End Function

'================================================================================
' NORMALIZA "... PARA SUGERIR" -> "... PARA INDICAR" NO 1 PARAGRAFO APOS EMENTA
' Regras:
' - No primeiro paragrafo textual subsequente a ementa, se INICIAR (case-insensitive)
'   com: "Nos termos do Art. 108 ... dirijo-me a Vossa Excelencia para sugerir"
'   substitui esse trecho inicial por uma versao (case-sensitive) com "para indicar".
' - Tolerante a caracteres nao-ASCII comuns do Word (NBSP, travessoes/hifens).
'================================================================================

Public Function NormalizeArt108ParaIndicarAfterEmenta(doc As Document) As Long
    On Error GoTo ErrorHandler

    NormalizeArt108ParaIndicarAfterEmenta = 0
    If doc Is Nothing Then Exit Function

    Dim ementaIdx As Long
    ementaIdx = FindEmentaParagraphIndex(doc)
    If ementaIdx <= 0 Or ementaIdx >= doc.Paragraphs.count Then Exit Function

    Dim oldPhrase As String
    Dim newPhrase As String
    oldPhrase = "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excel" & ChrW(234) & "ncia para sugerir"
    newPhrase = "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excel" & ChrW(234) & "ncia para indicar"

    Dim oldNormNoSpace As String
    oldNormNoSpace = Replace(NormalizeForComparison(oldPhrase), " ", "")

    Dim i As Long
    For i = ementaIdx + 1 To doc.Paragraphs.count
        Dim para As Paragraph
        Set para = doc.Paragraphs(i)

        If HasVisualContent(para) Then GoTo NextPara

        Dim rawText As String
        rawText = Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")

        Dim trimmedText As String
        trimmedText = LTrim$(rawText)
        If Len(Trim$(trimmedText)) = 0 Then GoTo NextPara

        Dim leadingSpacesLen As Long
        leadingSpacesLen = Len(rawText) - Len(trimmedText)

        ' Normaliza para comparacao tolerante
        Dim cmpText As String
        cmpText = trimmedText
        cmpText = Replace(cmpText, ChrW(160), " ")
        cmpText = Replace(cmpText, vbTab, " ")
        cmpText = Replace(cmpText, ChrW(8209), "-")
        cmpText = Replace(cmpText, ChrW(8211), "-")
        cmpText = Replace(cmpText, ChrW(8212), "-")
        cmpText = Replace(cmpText, ChrW(8722), "-")

        Dim cmpNormNoSpace As String
        cmpNormNoSpace = Replace(NormalizeForComparison(cmpText), " ", "")

        If Len(cmpNormNoSpace) < Len(oldNormNoSpace) Then GoTo NextPara
        If Left$(cmpNormNoSpace, Len(oldNormNoSpace)) <> oldNormNoSpace Then GoTo NextPara

        ' Encontra a posicao de "sugerir" no texto original (para definir o range a substituir)
        Dim posSugerir As Long
        posSugerir = InStr(1, trimmedText, "sugerir", vbTextCompare)
        If posSugerir <= 0 Then GoTo NextPara

        Dim endPos As Long
        endPos = posSugerir + Len("sugerir") - 1

        Dim replaceRng As Range
        Set replaceRng = para.Range.Duplicate
        If replaceRng.End > replaceRng.Start Then replaceRng.End = replaceRng.End - 1

        replaceRng.Start = replaceRng.Start + leadingSpacesLen
        replaceRng.End = replaceRng.Start + endPos
        SafeReplaceText replaceRng, newPhrase

        documentDirty = True
        NormalizeArt108ParaIndicarAfterEmenta = 1
        Exit Function

NextPara:
    Next i

    Exit Function

ErrorHandler:
    NormalizeArt108ParaIndicarAfterEmenta = 0
End Function

'================================================================================
' NORMALIZA ABERTURA DO ART. 108 NO 1 PARAGRAFO APOS EMENTA
    ' Regras:
    ' - No primeiro paragrafo textual subsequente a ementa, se INICIAR (case-insensitive)
    '   com o prefixo do Art. 108 e, em seguida, tiver exatamente:
    '     "Setor, " OU "Setor competente, "
    '   substitui esse trecho inicial por um texto padrao (case-sensitive)
    '   com "setor competente" (minusculo).
'================================================================================

Public Function NormalizeArt108IntroAfterEmenta(doc As Document) As Long
    On Error GoTo ErrorHandler

    NormalizeArt108IntroAfterEmenta = 0
    If doc Is Nothing Then Exit Function

    Dim ementaIdx As Long
    ementaIdx = FindEmentaParagraphIndex(doc)
    If ementaIdx <= 0 Or ementaIdx >= doc.Paragraphs.count Then Exit Function

    Dim prefixBase As String
    Dim newText As String

    ' ASCII-safe: acentos via ChrW
    prefixBase = "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excel" & ChrW(234) & "ncia para indicar que, por interm" & ChrW(233) & "dio do"
    newText = "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excel" & ChrW(234) & "ncia para indicar que, por interm" & ChrW(233) & "dio do setor competente, "

    Dim prefixNormNoSpace As String
    prefixNormNoSpace = Replace(NormalizeForComparison(prefixBase), " ", "")

    Dim i As Long
    For i = ementaIdx + 1 To doc.Paragraphs.count
        Dim para As Paragraph
        Set para = doc.Paragraphs(i)

        ' Apenas paragrafos textuais
        If HasVisualContent(para) Then GoTo NextPara

        Dim rawText As String
        rawText = Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")

        Dim trimmedText As String
        trimmedText = LTrim$(rawText)

        If Len(Trim$(trimmedText)) = 0 Then GoTo NextPara

        Dim leadingSpacesLen As Long
        leadingSpacesLen = Len(rawText) - Len(trimmedText)

        ' Prepara texto para comparacao:
        ' - Troca NBSP/tab por espaco
        ' - Normaliza hifens/travessoes comuns do Word para '-'
        ' - Remove espacos para tolerar variacoes (ex: "Art.108" vs "Art. 108")
        Dim cmpText As String
        cmpText = trimmedText
        cmpText = Replace(cmpText, ChrW(160), " ")
        cmpText = Replace(cmpText, vbTab, " ")
        cmpText = Replace(cmpText, ChrW(8209), "-") ' non-breaking hyphen
        cmpText = Replace(cmpText, ChrW(8211), "-") ' en dash
        cmpText = Replace(cmpText, ChrW(8212), "-") ' em dash
        cmpText = Replace(cmpText, ChrW(8722), "-") ' minus sign

        Dim cmpNormNoSpace As String
        cmpNormNoSpace = Replace(NormalizeForComparison(cmpText), " ", "")

        ' Confirma que o paragrafo inicia com o prefixo do Art. 108 (tolerante a espacos/acentos)
        If Len(cmpNormNoSpace) < Len(prefixNormNoSpace) Then GoTo NextPara
        If Left$(cmpNormNoSpace, Len(prefixNormNoSpace)) <> prefixNormNoSpace Then GoTo NextPara

        ' Confirma que imediatamente apos o prefixo existe exatamente "setor," ou "setorcompetente,"
        Dim afterPrefix As String
        afterPrefix = Mid$(cmpNormNoSpace, Len(prefixNormNoSpace) + 1)
        If Not (Left$(afterPrefix, Len("setor,")) = "setor," Or Left$(afterPrefix, Len("setorcompetente,")) = "setorcompetente,") Then
            GoTo NextPara
        End If

        ' Localiza o trecho "Setor ...," (aceita "Setor," e "Setor competente,")
        Dim setorPos As Long
        setorPos = InStr(1, trimmedText, "Setor", vbTextCompare)
        If setorPos <= 0 Then GoTo NextPara

        ' Valida a forma exata apos a palavra Setor (case-insensitive), permitindo espacos/NBSP:
        ' - ","  OU
        ' - " competente,".
        Dim posAfterSetor As Long
        posAfterSetor = setorPos + Len("Setor")

        Do While posAfterSetor <= Len(trimmedText)
            Dim chS As String
            chS = Mid$(trimmedText, posAfterSetor, 1)
            If chS = " " Or AscW(chS) = 160 Then
                posAfterSetor = posAfterSetor + 1
            Else
                Exit Do
            End If
        Loop

        Dim okForm As Boolean
        okForm = False

        If posAfterSetor <= Len(trimmedText) Then
            If Mid$(trimmedText, posAfterSetor, 1) = "," Then
                okForm = True
            ElseIf LCase$(Mid$(trimmedText, posAfterSetor, Len("competente"))) = "competente" Then
                Dim posAfterCompetente As Long
                posAfterCompetente = posAfterSetor + Len("competente")

                Do While posAfterCompetente <= Len(trimmedText)
                    Dim chC As String
                    chC = Mid$(trimmedText, posAfterCompetente, 1)
                    If chC = " " Or AscW(chC) = 160 Then
                        posAfterCompetente = posAfterCompetente + 1
                    Else
                        Exit Do
                    End If
                Loop

                If posAfterCompetente <= Len(trimmedText) Then
                    If Mid$(trimmedText, posAfterCompetente, 1) = "," Then
                        okForm = True
                    End If
                End If
            End If
        End If

        If Not okForm Then GoTo NextPara

        Dim commaPos As Long
        commaPos = InStr(setorPos, trimmedText, ",", vbBinaryCompare)
        If commaPos <= 0 Then GoTo NextPara

        Dim endPos As Long
        endPos = commaPos + 1

        ' Inclui espacos apos a virgula (espaco normal e NBSP)
        Do While endPos <= Len(trimmedText)
            Dim ch As String
            ch = Mid$(trimmedText, endPos, 1)
            If ch = " " Or AscW(ch) = 160 Then
                endPos = endPos + 1
            Else
                Exit Do
            End If
        Loop

        Dim replaceRng As Range
        Set replaceRng = para.Range.Duplicate
        If replaceRng.End > replaceRng.Start Then replaceRng.End = replaceRng.End - 1 ' exclui marca de paragrafo

        replaceRng.Start = replaceRng.Start + leadingSpacesLen
        replaceRng.End = replaceRng.Start + (endPos - 1)
        SafeReplaceText replaceRng, newText

        documentDirty = True
        NormalizeArt108IntroAfterEmenta = 1

        Exit Function ' apenas o primeiro paragrafo textual apos a ementa

NextPara:
    Next i

    Exit Function

ErrorHandler:
    NormalizeArt108IntroAfterEmenta = 0
End Function

'================================================================================
' APPLY BOLD TO SPECIAL PARAGRAPHS - SIMPLIFIED & OPTIMIZED
'================================================================================

Public Sub ApplyBoldToSpecialParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim cleanText As String
    Dim specialParagraphs As Collection
    Set specialParagraphs = New Collection

    ' FASE 1: Identificar paragrafos especiais (uma unica passada)
    Dim paraCounter As Long
    paraCounter = 0
    For Each para In doc.Paragraphs
        paraCounter = paraCounter + 1
        If paraCounter Mod 25 = 0 Then DoEvents ' Responsividade

        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)

            ' Adiciona Titulo, Justificativa e Anexo (Vereador nao recebe negrito)
            If cleanText = JUSTIFICATIVA_TEXT Or _
               IsAnexoPattern(cleanText) Or _
               (tituloParaIndex > 0 And paraCounter = tituloParaIndex) Then
                specialParagraphs.Add para
            End If
        End If
    Next para

    ' FASE 2: Aplicar negrito E reforcar alinhamento atomicamente
    ' Nao controla ScreenUpdating aqui - deixa a funcao principal controlar

    Dim p As Variant
    Dim pCleanText As String
    For Each p In specialParagraphs
        Set para = p ' Converte Variant para Paragraph

        ' Aplica negrito
        With para.Range.Font
            .Bold = True
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
        End With

        ' REFORCO: Garante alinhamento correto baseado no tipo
        pCleanText = GetCleanParagraphText(para)
        If tituloParaIndex > 0 And para.Range.Start = doc.Paragraphs(tituloParaIndex).Range.Start Then
            ' Titulo: centralizado, caixa alta, sublinhado, sem recuos
            para.Format.alignment = wdAlignParagraphCenter
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 6
            With para.Range.Font
                .Bold = True
                .Underline = wdUnderlineSingle
                .AllCaps = True
            End With
        ElseIf pCleanText = JUSTIFICATIVA_TEXT Then
            ' Justificativa: centralizado (linhas em branco serao inseridas depois)
            para.Format.alignment = wdAlignParagraphCenter
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
        ElseIf IsAnexoPattern(pCleanText) Then
            ' Anexo/Anexos: alinhado a esquerda
            para.Format.alignment = wdAlignParagraphLeft
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
        End If
    Next p

    LogMessage "Negrito e alinhamento aplicados a " & specialParagraphs.count & " paragrafos especiais", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao aplicar negrito a paragrafos especiais: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' FORMAT VEREADOR PARAGRAPHS - Formata paragrafo com "vereador" e adjacentes
' Antes de formatar, remove paragrafos em branco imediatamente acima do "Vereador".
'================================================================================

Public Sub FormatVereadorParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim NextPara As Paragraph
    Dim i As Long
    Dim j As Long
    Dim formattedCount As Long
    Dim blankRemovedCount As Long

    formattedCount = 0
    blankRemovedCount = 0

    ' Procura por paragrafos com "vereador"
    For i = 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For ' Protecao dinamica apos remocoes

        Set para = doc.Paragraphs(i)

        ' OBS: O paragrafo pode conter pontuacao/travessoes/hifens.
        ' A deteccao abaixo ignora tudo que nao for letra e valida se sobrou apenas "vereador".
        If IsVereadorPattern(para.Range.text) Then

            ' Divide paragrafo: se "Vereador" e seguido de hifens/traços + "P" + ate 5 caracteres,
            ' separa o que vem depois em um novo paragrafo abaixo, removendo hifens/traços.
            Dim rawText As String
            rawText = Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")
            Dim rawLower As String
            rawLower = LCase(Trim(rawText))

            Dim verPos As Long
            verPos = InStr(rawLower, "vereador")
            If verPos > 0 Then
                Dim textAfter As String
                textAfter = Mid(rawText, verPos + 8) ' 8 = Len("vereador")
                ' Normaliza hifens/traços para detectar padrao
                Dim normAfter As String
                normAfter = textAfter
                normAfter = Replace(normAfter, ChrW(8209), "-") ' non-breaking hyphen
                normAfter = Replace(normAfter, ChrW(8211), "-") ' en dash
                normAfter = Replace(normAfter, ChrW(8212), "-") ' em dash
                normAfter = Replace(normAfter, ChrW(8722), "-") ' minus sign
                normAfter = Trim(normAfter)

                ' Verifica se comeca com hifen
                If Len(normAfter) > 0 And Left(normAfter, 1) = "-" Then
                    Dim afterHyphen As String
                    afterHyphen = Trim(Mid(normAfter, 2))
                    ' Verifica: "P" + ate 5 caracteres (case sensitive, P maiuscula)
                    If Len(afterHyphen) >= 1 And Len(afterHyphen) <= 6 And Left(afterHyphen, 1) = "P" Then
                        ' Extrai a palavra base (Vereador ou Vereadora)
                        Dim baseWord As String
                        baseWord = GetVereadorNormalizedWord(rawText)
                        If baseWord = "" Then baseWord = "Vereador"

                        ' Monta o texto para o novo paragrafo (sem hifens/traços)
                        Dim newParaText As String
                        newParaText = afterHyphen

                        ' Limpa o paragrafo atual: mantem apenas "Vereador"/"Vereadora"
                        para.Range.Text = baseWord & vbCr

                        ' Insere novo paragrafo abaixo com o texto restante
                        Dim newRange As Range
                        Set newRange = doc.Paragraphs(i).Range
                        newRange.Collapse wdCollapseEnd
                        newRange.InsertAfter newParaText & vbCr

                        ' Formata o novo paragrafo (Arial 12, sem formatacao especial)
                        If i + 1 <= doc.Paragraphs.count Then
                            Dim newPara As Paragraph
                            Set newPara = doc.Paragraphs(i + 1)
                            With newPara.Range.Font
                                .Name = STANDARD_FONT
                                .Size = STANDARD_FONT_SIZE
                                .Bold = False
                                .AllCaps = False
                                .Underline = wdUnderlineNone
                            End With
                        End If

                        documentDirty = True
                        LogMessage "Paragrafo 'Vereador' dividido: '" & baseWord & "' + '" & newParaText & "' (posicao: " & i & ")", LOG_LEVEL_INFO
                    End If
                End If
            End If

            ' Remove paragrafos em branco imediatamente acima do "Vereador".
            ' Percorre de tras para frente (i-1, i-2, ...) enquanto o paragrafo
            ' for vazio e sem conteudo visual, ajustando o indice i apos cada remocao.
            Do While i > 1
                Set prevPara = doc.Paragraphs(i - 1)
                Dim prevClean As String
                prevClean = Trim(Replace(Replace(prevPara.Range.text, vbCr, ""), vbLf, ""))
                If prevClean = "" And Not HasVisualContent(prevPara) Then
                    prevPara.Range.Delete
                    i = i - 1  ' Ajusta indice: o "Vereador" subiu uma posicao
                    blankRemovedCount = blankRemovedCount + 1
                Else
                    Exit Do
                End If
            Loop

            ' Reobtem a referencia ao paragrafo "Vereador" apos possiveis remocoes
            If i > doc.Paragraphs.count Then Exit For
            Set para = doc.Paragraphs(i)

            ApplyVereadorParagraphFormatting para

            ' Formata linha ACIMA (se existir): centraliza, zera recuo, aplica caixa alta e negrito (somente se nao houver conteudo visual)
            ' IMPORTANTE: Nao altera paragrafos que fazem parte do vocativo
            If i > 1 Then
                ' Verifica se o paragrafo acima pertence ao vocativo
                If Not IsVocativoParagraph(i - 1) Then
                    Set prevPara = doc.Paragraphs(i - 1)
                    If Not HasVisualContent(prevPara) Then
                        ' Aplica caixa alta e negrito na fonte
                        With prevPara.Range.Font
                            .AllCaps = True
                            .Bold = True
                            .Name = STANDARD_FONT
                            .size = STANDARD_FONT_SIZE
                        End With
                    End If

                    ' Centraliza e zera recuos (seguro mesmo com conteudo visual)
                    With prevPara.Format
                        .alignment = wdAlignParagraphCenter
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .RightIndent = 0
                    End With
                Else
                    LogMessage "Paragrafo vocativo (indice " & (i - 1) & ") preservado - nao aplicada formatacao de Vereador adjacente", LOG_LEVEL_INFO
                End If
            End If

            ' Formata SEGUNDA linha ACIMA (se existir): negrito, centralizado, recuo esquerda = 0
            ' IMPORTANTE: Nao altera paragrafos que fazem parte do vocativo
            If i > 2 Then
                ' Verifica se o paragrafo acima pertence ao vocativo
                If Not IsVocativoParagraph(i - 2) Then
                    Dim prevPrevPara As Paragraph
                    Set prevPrevPara = doc.Paragraphs(i - 2)
                    If Not HasVisualContent(prevPrevPara) Then
                        With prevPrevPara.Range.Font
                            .Bold = True
                            .Name = STANDARD_FONT
                            .size = STANDARD_FONT_SIZE
                        End With
                    End If
                    With prevPrevPara.Format
                        .alignment = wdAlignParagraphCenter
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .RightIndent = 0
                    End With
                Else
                    LogMessage "Paragrafo vocativo (indice " & (i - 2) & ") preservado - nao aplicada formatacao de Vereador adjacente", LOG_LEVEL_INFO
                End If
            End If

            ' Formata linha ABAIXO (se existir)
            If i < doc.Paragraphs.count Then
                Set NextPara = doc.Paragraphs(i + 1)
                With NextPara.Format
                    .alignment = wdAlignParagraphCenter
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                End With
            End If

            formattedCount = formattedCount + 1
            LogMessage "Paragrafo 'Vereador' formatado (sem negrito) com linhas adjacentes centralizadas (posicao: " & i & ")", LOG_LEVEL_INFO
        End If
    Next i

    If blankRemovedCount > 0 Then
        documentDirty = True
        LogMessage "Paragrafos em branco acima de 'Vereador' removidos: " & blankRemovedCount, LOG_LEVEL_INFO
    End If

    If formattedCount > 0 Then
        LogMessage "Formatacao 'Vereador': " & formattedCount & " ocorrencias formatadas", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar paragrafos 'Vereador': " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' SUBSTITUICAO DE "Vereadora," POR "Senhora Vereadora,"
' Regra: se um paragrafo contem unicamente "Vereadora," e o paragrafo imediatamente
' acima contem unicamente "Senhores Vereadores,", substitui o texto por "Senhora Vereadora,".
'================================================================================

Public Sub SubstituiVereadoraPorSenhoraVereadora(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim i As Long
    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim paraText As String
    Dim prevText As String
    Dim rng As Range
    Dim substitutedCount As Long

    substitutedCount = 0

    For i = 2 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = LCase$(Trim$(Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")))

        If paraText = "vereadora," Then
            Set prevPara = doc.Paragraphs(i - 1)
            prevText = LCase$(Trim$(Replace(Replace(prevPara.Range.text, vbCr, ""), vbLf, "")))

            If prevText = "senhores vereadores," Then
                ' Substitui o conteudo sem remover a marca de paragrafo
                Set rng = para.Range
                rng.MoveEnd wdCharacter, -1
                SafeReplaceText rng, "Senhora Vereadora,"

                substitutedCount = substitutedCount + 1
                LogMessage "Substituicao: 'Vereadora,' -> 'Senhora Vereadora,' (posicao: " & i & ")", LOG_LEVEL_INFO
            End If
        End If
    Next i

    If substitutedCount > 0 Then
        LogMessage "SubstituiVereadoraPorSenhoraVereadora: " & substitutedCount & " substituicao(oes) realizada(s)", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao substituir Vereadora por Senhora Vereadora: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' TIKINHO TK - SUBSTITUICAO DE PARAGRAFO
'================================================================================

Private Function IsTikinhoTk(ByVal text As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(text))
    t = Replace(Replace(t, vbCr, ""), vbLf, "")
    t = Trim$(t)
    
    If t = "tikinho tk""" Then
        IsTikinhoTk = True
        Exit Function
    End If
    If t = "'tikinho tk""'" Then
        IsTikinhoTk = True
        Exit Function
    End If
    If t = "tikinho tk" Then
        IsTikinhoTk = True
        Exit Function
    End If
    If t = "'tikinho tk'" Then
        IsTikinhoTk = True
        Exit Function
    End If
    If t = """tikinho tk""" Then
        IsTikinhoTk = True
        Exit Function
    End If
    
    ' Unicode smart quotes etc.
    Dim stripped As String
    stripped = t
    stripped = Replace(stripped, "'", "")
    stripped = Replace(stripped, """", "")
    stripped = Replace(stripped, ChrW(8216), "")
    stripped = Replace(stripped, ChrW(8217), "")
    stripped = Replace(stripped, ChrW(8220), "")
    stripped = Replace(stripped, ChrW(8221), "")
    stripped = Trim$(stripped)
    
    If stripped = "tikinho tk" Then
        IsTikinhoTk = True
        Exit Function
    End If
    
    IsTikinhoTk = False
End Function


Public Sub ReplaceTikinhoTkParagraphs(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim replacedCount As Long
    
    replacedCount = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = para.Range.text
        
        If IsTikinhoTk(paraText) Then
            Set rng = para.Range
            If rng.Characters.count > 1 Then
                rng.MoveEnd wdCharacter, -1 ' Nao remove a marca de paragrafo
            End If
            SafeReplaceText rng, "TIKINHO TK"
            replacedCount = replacedCount + 1
        End If
    Next i
    
    If replacedCount > 0 Then
        documentDirty = True
        LogMessage "Substituicao de paragrafo 'tikinho tk' realizada (" & replacedCount & "x)", LOG_LEVEL_INFO
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "Erro ao substituir paragrafos tikinho tk: " & Err.Description, LOG_LEVEL_WARNING
End Sub


Public Sub RemoveJustificativaColon(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim replacedCount As Long
    Dim normalizedCount As Long
    
    replacedCount = 0
    normalizedCount = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = para.Range.text
        
        Dim cleanText As String
        cleanText = Trim$(paraText)
        cleanText = Replace(Replace(cleanText, vbCr, ""), vbLf, "")
        cleanText = Trim$(cleanText)
        
        ' Remove colon: "Justificativa:" ou "Justificacao:" -> "Justificativa"
        If LCase$(cleanText) = "justificativa:" Or LCase$(cleanText) = "justificacao:" Then
            Set rng = para.Range
            If rng.Characters.count > 1 Then
                rng.MoveEnd wdCharacter, -1
            End If
            SafeReplaceText rng, "Justificativa"
            replacedCount = replacedCount + 1
        ' Normaliza caixa: "JUSTIFICATIVA" ou "justificativa" -> "Justificativa"
        ' (apenas quando o texto nao esta ja na forma correta)
        ElseIf LCase$(cleanText) = "justificativa" Or LCase$(cleanText) = "justificacao" Then
            If cleanText <> "Justificativa" Then
                Set rng = para.Range
                If rng.Characters.count > 1 Then
                    rng.MoveEnd wdCharacter, -1
                End If
                SafeReplaceText rng, "Justificativa"
                normalizedCount = normalizedCount + 1
            End If
        End If
    Next i
    
    If replacedCount > 0 Then
        documentDirty = True
        LogMessage "Remocao de dois pontos do titulo de Justificativa: " & replacedCount & "x", LOG_LEVEL_INFO
    End If
    
    If normalizedCount > 0 Then
        documentDirty = True
        LogMessage "Normalizacao de caixa do titulo de Justificativa: " & normalizedCount & "x", LOG_LEVEL_INFO
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "Erro ao remover dois pontos da Justificativa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' VEREADOR - FORMATACAO DEDICADA
' Regras:
' - Paragrafo contendo unicamente a palavra "vereador" (case-insensitive), mesmo cercada por hifens/travessoes,
'   deve ficar como "Vereador".
' - Fonte normal (sem negrito/italico/sublinhado/caixa alta), centralizado e com recuos a esquerda = 0.
'================================================================================

Public Sub ApplyVereadorParagraphFormatting(para As Paragraph)
    On Error Resume Next

    ' IMPORTANTE: " - Vereador - " pode disparar autoformatacao de lista (bullets),
    ' gerando recuo padrao (ex: 1,25 cm). Desabilita temporariamente.
    Dim prevAutoBullets As Boolean
    Dim prevAutoNumbers As Boolean
    Dim canToggleAutoFormat As Boolean
    canToggleAutoFormat = False

    Err.Clear
    prevAutoBullets = Application.Options.AutoFormatAsYouTypeApplyBulletedLists
    prevAutoNumbers = Application.Options.AutoFormatAsYouTypeApplyNumberedLists
    If Err.Number = 0 Then
        canToggleAutoFormat = True
        Application.Options.AutoFormatAsYouTypeApplyBulletedLists = False
        Application.Options.AutoFormatAsYouTypeApplyNumberedLists = False
    End If
    Err.Clear

    Dim rngText As Range
    Set rngText = para.Range.Duplicate
    If rngText.End > rngText.Start Then rngText.End = rngText.End - 1 ' exclui marca de paragrafo

    Dim targetWord As String
    targetWord = GetVereadorNormalizedWord(para.Range.text)
    If targetWord = "" Then GoTo Cleanup

    ' Evita apagar imagens/shapes: so reescreve texto quando nao ha conteudo visual.
    If Not HasVisualContent(para) Then
        rngText.text = targetWord
    Else
        ' Caso especial: quando ha conteudo visual, nao reescreva o paragrafo inteiro.
        ' Em vez disso, localiza a palavra e substitui tambem os caracteres nao-letra ao redor
        ' (hifens, travessoes, espacos, pontuacao), evitando duplicacoes como "- - Vereador - -".
        Dim doc As Document
        Set doc = para.Range.Document

        Dim searchWord As String
        If InStr(1, targetWord, "Vereadora", vbTextCompare) > 0 Then
            searchWord = "vereadora"
        Else
            searchWord = "vereador"
        End If

        Dim findRng As Range
        Set findRng = rngText.Duplicate

        With findRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .text = searchWord
        End With

        If findRng.Find.Execute Then
            Dim replaceStart As Long
            Dim replaceEnd As Long
            replaceStart = findRng.Start
            replaceEnd = findRng.End

            Do While replaceStart > rngText.Start
                Dim chLeft As String
                chLeft = doc.Range(replaceStart - 1, replaceStart).text
                If IsAsciiLetterChar(chLeft) Then Exit Do
                replaceStart = replaceStart - 1
            Loop

            Do While replaceEnd < rngText.End
                Dim chRight As String
                chRight = doc.Range(replaceEnd, replaceEnd + 1).text
                If IsAsciiLetterChar(chRight) Then Exit Do
                replaceEnd = replaceEnd + 1
            Loop

            doc.Range(replaceStart, replaceEnd).text = targetWord
        End If
    End If

    ' Estilo e fonte normal
    para.Style = "Normal"

    ' NOTA: ListFormat.RemoveNumbers removido -- marcadores/numeracao nao devem ser editados
    On Error Resume Next
    With para.Range.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .AllCaps = False
        .Name = STANDARD_FONT
        .size = STANDARD_FONT_SIZE
        .Color = wdColorAutomatic
    End With

    ' Centraliza e zera recuos
    With para.Format
        .alignment = wdAlignParagraphCenter
        .leftIndent = 0
        .firstLineIndent = 0
        .RightIndent = 0
    End With

    ' Reforco adicional (em alguns casos, para.Format nao vence estilo/lista)
    With para.Range.ParagraphFormat
        .leftIndent = 0
        .firstLineIndent = 0
        .RightIndent = 0
    End With

Cleanup:
    ' Restaura configuracoes de autoformatacao
    If canToggleAutoFormat Then
        Err.Clear
        Application.Options.AutoFormatAsYouTypeApplyBulletedLists = prevAutoBullets
        Application.Options.AutoFormatAsYouTypeApplyNumberedLists = prevAutoNumbers
        Err.Clear
    End If
End Sub

'================================================================================
' FUNCOES AUXILIARES PARA MANIPULACAO DE LINHAS EM BRANCO
'================================================================================

' Remove linhas vazias ANTES de um paragrafo especifico
' Retorna o novo indice do paragrafo apos remocoes

Public Function RemoveBlankLinesBefore(doc As Document, ByVal targetIndex As Long) As Long
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long

    i = targetIndex - 1
    Do While i >= 1
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText = "" And Not HasVisualContent(para) Then
            para.Range.Delete
            targetIndex = targetIndex - 1
            i = i - 1
        Else
            Exit Do
        End If
    Loop

    RemoveBlankLinesBefore = targetIndex
    Exit Function

ErrorHandler:
    RemoveBlankLinesBefore = targetIndex
End Function

' Remove linhas vazias DEPOIS de um paragrafo especifico

Public Sub RemoveBlankLinesAfter(doc As Document, ByVal targetIndex As Long)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long

    i = targetIndex + 1
    Do While i <= doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText = "" And Not HasVisualContent(para) Then
            para.Range.Delete
        Else
            Exit Do
        End If
    Loop

    Exit Sub

ErrorHandler:
    ' Silently continue
End Sub

' Insere N linhas em branco ANTES de um paragrafo

Public Sub InsertBlankLinesBefore(doc As Document, ByVal targetIndex As Long, ByVal lineCount As Long)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim j As Long

    Set para = doc.Paragraphs(targetIndex)
    For j = 1 To lineCount
        para.Range.InsertParagraphBefore
    Next j

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas antes: " & Err.Description, LOG_LEVEL_WARNING
End Sub

' Insere N linhas em branco DEPOIS de um paragrafo

Public Sub InsertBlankLinesAfter(doc As Document, ByVal targetIndex As Long, ByVal lineCount As Long)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim j As Long

    Set para = doc.Paragraphs(targetIndex)
    For j = 1 To lineCount
        para.Range.InsertParagraphAfter
    Next j

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas depois: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' INSERCAO DE LINHAS EM BRANCO NA JUSTIFICATIVA
'================================================================================

Public Sub InsertJustificativaBlankLines(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim cleanText As String
    Dim i As Long
    Dim justificativaIndex As Long
    Dim paraText As String

    ' Nao controla ScreenUpdating aqui - deixa a funcao principal controlar

    ' FASE 1: Localiza o paragrafo "Justificativa"
    justificativaIndex = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)

            If cleanText = JUSTIFICATIVA_TEXT Then
                justificativaIndex = i
                Exit For
            End If
        End If
    Next i

    If justificativaIndex = 0 Then
        Exit Sub ' Nao encontrou "Justificativa"
    End If

    ' FASE 2-5: Remove linhas vazias e insere exatamente 2 antes e 2 depois
    justificativaIndex = RemoveBlankLinesBefore(doc, justificativaIndex)
    RemoveBlankLinesAfter doc, justificativaIndex
    InsertBlankLinesBefore doc, justificativaIndex, 2
    InsertBlankLinesAfter doc, justificativaIndex + 2, 2  ' +2 por causa das insercoes anteriores

    LogMessage "Linhas em branco ajustadas: 2 antes e 2 depois de 'Justificativa'", LOG_LEVEL_INFO

    ' FASE 6: Processa "Plenario Dr. Tancredo Neves"
    Dim plenarioIndex As Long
    Dim paraTextCmp As String
    Dim paraTextLower As String

    plenarioIndex = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            paraTextCmp = NormalizeForComparison(paraText)

            ' Procura por "Plenario" e "Tancredo Neves" (case insensitive)
            If InStr(paraTextCmp, "plenario") > 0 And _
               InStr(paraTextCmp, "tancredo") > 0 And _
               InStr(paraTextCmp, "neves") > 0 Then
                plenarioIndex = i
                Exit For
            End If
        End If
    Next i

    If plenarioIndex > 0 Then
        ' Remove linhas vazias e insere exatamente 2 antes e 2 depois
        plenarioIndex = RemoveBlankLinesBefore(doc, plenarioIndex)
        RemoveBlankLinesAfter doc, plenarioIndex
        InsertBlankLinesBefore doc, plenarioIndex, 2
        InsertBlankLinesAfter doc, plenarioIndex + 2, 2

        LogMessage "2 linhas em branco inseridas antes e depois de 'Plenario Dr. Tancredo Neves'", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas em branco: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FUNCOES AUXILIARES PARA DETECCAO DE PADROES
'================================================================================

Public Function IsVereadorPattern(text As String) As Boolean
    ' Excecao: "Vereadora," (vocativo) nao deve receber formatacao de vereador
    Dim rawClean As String
    rawClean = LCase$(Trim$(Replace(Replace(text, vbCr, ""), vbLf, "")))
    If rawClean = "vereadora," Then
        IsVereadorPattern = False
        Exit Function
    End If
    IsVereadorPattern = (GetVereadorNormalizedWord(text) <> "")
End Function


Public Function IsVocativoVereadorPresidenteParagraph(text As String) As Boolean
    ' Detecta paragrafos vocativos que aparecem logo abaixo de "Senhor Presidente,"
    ' e nao devem ser separados deste por paragrafos em branco.
    Dim cleanText As String
    cleanText = LCase$(Trim$(Replace(Replace(text, vbCr, ""), vbLf, "")))

    Select Case cleanText
        Case "senhores vereadores,", "senhores vereadores", _
             "senhoras vereadoras,", "senhoras vereadoras", _
             "senhores(as) vereadores(as),", "senhores(as) vereadores(as)", _
             "senhora vereadora,", "senhora vereadora"
            IsVocativoVereadorPresidenteParagraph = True
        Case Else
            IsVocativoVereadorPresidenteParagraph = False
    End Select
End Function


Public Function IsSenhorPresidenteParagraph(text As String) As Boolean
    ' Detecta o paragrafo que contem exclusivamente "Senhor Presidente,"
    Dim cleanText As String
    cleanText = LCase$(Trim$(Replace(Replace(text, vbCr, ""), vbLf, "")))
    IsSenhorPresidenteParagraph = (cleanText = "senhor presidente,")
End Function


Public Function GetVereadorNormalizedWord(text As String) As String
    Dim cleanText As String

    cleanText = Replace(Replace(text, vbCr, ""), vbLf, "")
    cleanText = Trim$(cleanText)
    cleanText = NormalizeLettersOnly(cleanText)

    If cleanText = "vereador" Then
        GetVereadorNormalizedWord = "Vereador"
    ElseIf cleanText = "vereadora" Then
        GetVereadorNormalizedWord = "Vereadora"
    Else
        GetVereadorNormalizedWord = ""
    End If
End Function


Public Function IsAsciiLetterChar(ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsAsciiLetterChar = False
        Exit Function
    End If

    Dim code As Long
    code = AscW(ch)
    If code < 0 Then code = code + 65536

    IsAsciiLetterChar = ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122))
End Function


Public Function NormalizeLettersOnly(text As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim outText As String

    outText = ""

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        ' ASCII letters only (A-Z, a-z)
        If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            outText = outText & LCase$(ch)
        End If
    Next i

    NormalizeLettersOnly = outText
End Function


Public Function IsAnexoPattern(text As String) As Boolean
    Dim cleanText As String
    cleanText = LCase(Trim(text))
    IsAnexoPattern = (cleanText = "anexo" Or cleanText = "anexos")
End Function

'================================================================================
' IS VOCATIVO PARAGRAPH - Verifica se um indice de paragrafo pertence ao vocativo
'================================================================================

Public Function IsVocativoParagraph(paraIndex As Long) As Boolean
    On Error GoTo ErrorHandler

    If vocativoStartIndex <= 0 Or vocativoEndIndex <= 0 Then
        IsVocativoParagraph = False
        Exit Function
    End If

    IsVocativoParagraph = (paraIndex >= vocativoStartIndex And paraIndex <= vocativoEndIndex)
    Exit Function

ErrorHandler:
    IsVocativoParagraph = False
End Function

'================================================================================
' FORMAT DIANTE DO EXPOSTO - Formata "Diante do exposto" no inicio de paragrafos
'================================================================================

Public Sub FormatDianteDoExposto(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    formattedCount = 0

    ' Procura por paragrafos que comecam com "Diante do exposto"
    Dim iterCounter As Long
    iterCounter = 0
    For Each para In doc.Paragraphs
        iterCounter = iterCounter + 1
        If iterCounter Mod 25 = 0 Then DoEvents ' Responsividade

        If Not HasVisualContent(para) Then
            ' Obtem o texto do paragrafo
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            cleanText = LCase(paraText)

            ' Verifica se comeca com "diante do exposto"
            If Left(cleanText, 17) = "diante do exposto" Then
                ' Encontra a posicao exata da frase (primeiros 17 caracteres)
                Dim targetRange As Range
                Set targetRange = para.Range
                targetRange.End = targetRange.Start + 17

                ' Aplica formatacao: negrito e caixa alta
                With targetRange.Font
                    .Bold = True
                    .AllCaps = True
                    .Name = STANDARD_FONT
                    .size = STANDARD_FONT_SIZE
                End With

                formattedCount = formattedCount + 1
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Formatacao 'Diante do exposto': " & formattedCount & " ocorrencias formatadas em negrito e caixa alta", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar 'Diante do exposto': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORMAT REQUEIRO PARAGRAPHS - Formata paragrafos que comecam com "requeiro"
'================================================================================

Public Sub FormatRequeiroParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    formattedCount = 0

    ' Procura por paragrafos que comecam com "requeiro" (case insensitive)
    Dim reqCounter As Long
    reqCounter = 0
    For Each para In doc.Paragraphs
        reqCounter = reqCounter + 1
        If reqCounter Mod 25 = 0 Then DoEvents ' Responsividade

        If Not HasVisualContent(para) Then
            ' Obtem o texto do paragrafo (sem marca de paragrafo)
            paraText = para.Range.text
            If Right(paraText, 1) = vbCr Then
                paraText = Left(paraText, Len(paraText) - 1)
            End If
            paraText = Trim(paraText)
            cleanText = LCase(paraText)

            ' Verifica se comeca com "requeiro" (8 caracteres)
            If Len(paraText) >= 8 Then
                If Left(cleanText, 8) = "requeiro" Then
                    ' Aplica formatacao APENAS a palavra "requeiro": negrito e caixa alta
                    Dim wordRange As Range
                    Dim startPos As Long

                    ' Encontra a posicao inicial do texto (apos espacos/tabs)
                    Set wordRange = para.Range
                    startPos = wordRange.Start

                    ' Move para o inicio do texto visivel
                    Do While startPos < wordRange.End
                        wordRange.Start = startPos
                        If Trim(Left(wordRange.text, 1)) <> "" Then Exit Do
                        startPos = startPos + 1
                    Loop

                    ' Seleciona apenas os 8 caracteres de "requeiro"
                    wordRange.End = wordRange.Start + 8

                    ' Aplica formatacao apenas a palavra
                    With wordRange.Font
                        .Bold = True
                        .AllCaps = True
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                    End With

                    formattedCount = formattedCount + 1
                End If
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Formatacao 'Requeiro': " & formattedCount & " palavras formatadas em negrito e caixa alta", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar paragrafos 'Requeiro': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORMAT "POR TODAS AS RAZOES" PARAGRAPHS - Formata "Por todas as razoes aqui expostas" e "Pelas razoes aqui expostas"
'================================================================================

Public Sub FormatPorTodasRazoesParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    Dim wordRange As Range
    Dim phrase1Len As Long
    Dim phrase2Len As Long

    formattedCount = 0
    phrase1Len = 33 ' "por todas as razoes aqui expostas"
    phrase2Len = 28 ' "pelas razoes aqui expostas"

    ' Procura por paragrafos que comecam com as frases (case insensitive)
    For Each para In doc.Paragraphs
        If Not HasVisualContent(para) Then
            ' Obtem o texto do paragrafo
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            cleanText = LCase(paraText)

            ' Verifica "por todas as razoes aqui expostas"
            If Len(paraText) >= phrase1Len Then
                If Left(cleanText, phrase1Len) = "por todas as razoes aqui expostas" Or _
                   Left(cleanText, phrase1Len) = "por todas as razoes aqui expostas" Then
                    Set wordRange = para.Range.Duplicate
                    wordRange.Collapse wdCollapseStart
                    wordRange.MoveEnd wdCharacter, phrase1Len

                    With wordRange.Font
                        .Bold = True
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                    End With

                    formattedCount = formattedCount + 1
                    GoTo NextPara
                End If
            End If

            ' Verifica "pelas razoes aqui expostas"
            If Len(paraText) >= phrase2Len Then
                If Left(cleanText, phrase2Len) = "pelas razoes aqui expostas" Or _
                   Left(cleanText, phrase2Len) = "pelas razoes aqui expostas" Then
                    Set wordRange = para.Range.Duplicate
                    wordRange.Collapse wdCollapseStart
                    wordRange.MoveEnd wdCharacter, phrase2Len

                    With wordRange.Font
                        .Bold = True
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                    End With

                    formattedCount = formattedCount + 1
                End If
            End If
        End If
NextPara:
    Next para

    If formattedCount > 0 Then
        LogMessage "Formatacao 'Por todas as razoes': " & formattedCount & " frases formatadas em negrito", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar frases 'Por todas as razoes': " & Err.Description, LOG_LEVEL_WARNING
End Sub


'================================================================================
' RESTAURAR BACKUP - Descarta documento atual e restaura o backup mais antigo
'================================================================================
