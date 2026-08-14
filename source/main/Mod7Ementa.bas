Attribute VB_Name = "Mod7Ementa"
Option Explicit

' Mod7Ementa
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)

Public Sub ForceEmentaSpacing(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub
    If ementaParaIndex <= 0 Or ementaParaIndex > doc.Paragraphs.count Then Exit Sub

    ' Garante que a ementa existe e tem conteudo
    Dim ementaText As String
    ementaText = Trim(Replace(Replace(doc.Paragraphs(ementaParaIndex).Range.text, vbCr, ""), vbLf, ""))
    If Len(ementaText) = 0 Then Exit Sub

    Dim idx As Long
    Dim blankCount As Long
    Dim insertCount As Long

    ' =========================================================================
    ' 1. GARANTE 2 PARAGRAFOS EM BRANCO ABAIXO DA EMENTA
    ' =========================================================================
    blankCount = 0
    idx = ementaParaIndex + 1
    Do While idx <= doc.Paragraphs.count
        Dim belowText As String
        belowText = Trim(Replace(Replace(doc.Paragraphs(idx).Range.text, vbCr, ""), vbLf, ""))
        If belowText = "" And Not HasVisualContent(doc.Paragraphs(idx)) Then
            blankCount = blankCount + 1
            idx = idx + 1
        Else
            Exit Do
        End If
    Loop

    If blankCount < 2 Then
        insertCount = 2 - blankCount
        Dim b As Long
        For b = 1 To insertCount
            doc.Paragraphs(ementaParaIndex).Range.InsertParagraphAfter
        Next b
        LogMessage "ForceEmentaSpacing: " & insertCount & " paragrafo(s) em branco inserido(s) abaixo da Ementa (total garantido: 2)", LOG_LEVEL_INFO
    End If

    ' =========================================================================
    ' 2. GARANTE 3 PARAGRAFOS EM BRANCO ACIMA DA EMENTA
    ' Nota: inserir abaixo primeiro nao altera o indice da ementa,
    '       mas inserir abaixo desloca os paragrafos posteriores.
    '       O indice ementaParaIndex permanece o mesmo.
    ' =========================================================================
    blankCount = 0
    idx = ementaParaIndex - 1
    Do While idx >= 1
        Dim aboveText As String
        aboveText = Trim(Replace(Replace(doc.Paragraphs(idx).Range.text, vbCr, ""), vbLf, ""))
        If aboveText = "" And Not HasVisualContent(doc.Paragraphs(idx)) Then
            blankCount = blankCount + 1
            idx = idx - 1
        Else
            Exit Do
        End If
    Loop

    If blankCount < 3 Then
        insertCount = 3 - blankCount
        Dim a As Long
        For a = 1 To insertCount
            doc.Paragraphs(ementaParaIndex).Range.InsertParagraphBefore
        Next a
        ' A ementa deslocou N posicoes para baixo
        ementaParaIndex = ementaParaIndex + insertCount
        LogMessage "ForceEmentaSpacing: " & insertCount & " paragrafo(s) em branco inserido(s) acima da Ementa (total garantido: 3)", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir espacamento da Ementa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORCE DATA SPACING - Garante 2 paragrafos em branco acima da Data
' Executada como ULTIMA etapa para nao ser desfeita por processamento posterior.
'================================================================================

Public Sub ForceDataSpacing(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub
    If dataParaIndex <= 0 Or dataParaIndex > doc.Paragraphs.count Then Exit Sub

    ' Garante que a data existe e tem conteudo
    Dim dataText As String
    dataText = Trim(Replace(Replace(doc.Paragraphs(dataParaIndex).Range.text, vbCr, ""), vbLf, ""))
    If Len(dataText) = 0 Then Exit Sub

    Dim idx As Long
    Dim blankCount As Long
    Dim insertCount As Long

    ' =========================================================================
    ' GARANTE 2 PARAGRAFOS EM BRANCO ACIMA DA DATA
    ' =========================================================================
    blankCount = 0
    idx = dataParaIndex - 1
    Do While idx >= 1
        Dim aboveText As String
        aboveText = Trim(Replace(Replace(doc.Paragraphs(idx).Range.text, vbCr, ""), vbLf, ""))
        If aboveText = "" And Not HasVisualContent(doc.Paragraphs(idx)) Then
            blankCount = blankCount + 1
            idx = idx - 1
        Else
            Exit Do
        End If
    Loop

    If blankCount < 2 Then
        insertCount = 2 - blankCount
        Dim a As Long
        For a = 1 To insertCount
            doc.Paragraphs(dataParaIndex).Range.InsertParagraphBefore
        Next a
        ' A data deslocou N posicoes para baixo
        dataParaIndex = dataParaIndex + insertCount
        LogMessage "ForceDataSpacing: " & insertCount & " paragrafo(s) em branco inserido(s) acima da Data (total garantido: 2)", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir espacamento da Data: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' EMENTA - Remove prefixos "EMENTA:" / "ASSUNTO:" quando forem a primeira palavra
'================================================================================

Public Sub RemoveEmentaLeadingLabelPrefix(doc As Document)
    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = GetEmentaRange(doc)
    If rng Is Nothing Then Exit Sub

    Dim deleteLen As Long
    deleteLen = GetEmentaLeadingLabelDeleteLen(rng.text)
    If deleteLen <= 0 Then Exit Sub

    Dim delRng As Range
    Set delRng = rng.Duplicate
    delRng.Start = rng.Start
    delRng.End = rng.Start + deleteLen
    delRng.Delete

    documentDirty = True
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao remover prefixo da ementa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' EMENTA - Remove sufixo ", neste municipio" quando estiver no final
' Regras:
' - Case-insensitive
' - Remove variantes: ", neste municipio" e ",neste municipio" (inclui "Municipio")
' - Se ja existir ponto final no fim da ementa, mantem
' - Se nao existir, insere ponto final apos a exclusao
'================================================================================

Public Sub RemoveEmentaTrailingMunicipioSuffix(doc As Document)
    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = GetEmentaRange(doc)
    If rng Is Nothing Then Exit Sub

    Dim contentRng As Range
    Set contentRng = rng.Duplicate

    ' Range da ementa sem marca de paragrafo e sem espacos finais (inclui NBSP)
    If contentRng.End > contentRng.Start Then
        If Right$(contentRng.text, 1) = vbCr Then
            contentRng.End = contentRng.End - 1
        End If
    End If
    Do While contentRng.End > contentRng.Start
        If Right$(contentRng.text, 1) = " " Or Right$(contentRng.text, 1) = vbTab Or Right$(contentRng.text, 1) = ChrW(160) Then
            contentRng.End = contentRng.End - 1
        Else
            Exit Do
        End If
    Loop

    Dim rawText As String
    rawText = contentRng.text
    If Len(rawText) = 0 Then Exit Sub

    ' Normaliza NBSP para comparacao (mantem mesmo comprimento)
    Dim normalizedText As String
    normalizedText = Replace(rawText, ChrW(160), " ")
    If Len(normalizedText) = 0 Then Exit Sub

    ' Construcao ASCII-safe de ",neste municipio" (com e sem acento)
    Dim municipio1 As String
    Dim municipio2 As String
    municipio1 = "mun" & ChrW(237) & "cipio" ' munic�pio (com acento)
    municipio2 = "municipio"               ' municipio (sem acento)
    
    Dim suffix1 As String, suffix2 As String
    Dim suffix3 As String, suffix4 As String
    suffix1 = ",neste " & municipio1
    suffix2 = ", neste " & municipio1
    suffix3 = ",neste " & municipio2
    suffix4 = ", neste " & municipio2

    Dim lowerText As String
    lowerText = LCase$(normalizedText)

    ' Regra do ponto final: se ja existir, mantem; senao, adiciona apos exclusao
    Dim hadFinalPeriod As Boolean
    hadFinalPeriod = (Right$(lowerText, 1) = ".")

    Dim lowerBase As String
    lowerBase = lowerText
    If hadFinalPeriod And Len(lowerBase) > 1 Then
        lowerBase = Left$(lowerBase, Len(lowerBase) - 1)
    End If

    Dim deleteSuffix As String
    deleteSuffix = ""

    If Len(lowerBase) >= Len(suffix1) Then
        If Right$(lowerBase, Len(suffix1)) = LCase$(suffix1) Then
            deleteSuffix = suffix1
        End If
    End If

    If deleteSuffix = "" And Len(lowerBase) >= Len(suffix2) Then
        If Right$(lowerBase, Len(suffix2)) = LCase$(suffix2) Then
            deleteSuffix = suffix2
        End If
    End If

    If deleteSuffix = "" And Len(lowerBase) >= Len(suffix3) Then
        If Right$(lowerBase, Len(suffix3)) = LCase$(suffix3) Then
            deleteSuffix = suffix3
        End If
    End If

    If deleteSuffix = "" And Len(lowerBase) >= Len(suffix4) Then
        If Right$(lowerBase, Len(suffix4)) = LCase$(suffix4) Then
            deleteSuffix = suffix4
        End If
    End If

    If deleteSuffix = "" Then Exit Sub

    ' Remove o sufixo no Range real (mantem demais pontuacoes/texto)
    Dim pos As Long
    pos = InStrRev(lowerBase, LCase$(deleteSuffix))
    If pos <= 0 Then Exit Sub
    If (pos + Len(deleteSuffix) - 1) <> Len(lowerBase) Then Exit Sub

    Dim delRng As Range
    Set delRng = contentRng.Duplicate
    delRng.Start = contentRng.Start + pos - 1
    delRng.End = delRng.Start + Len(deleteSuffix)
    delRng.Delete

    ' Recalcula ementa sem marca de paragrafo e sem espacos finais
    Set contentRng = rng.Duplicate
    If contentRng.End > contentRng.Start Then
        If Right$(contentRng.text, 1) = vbCr Then
            contentRng.End = contentRng.End - 1
        End If
    End If
    Do While contentRng.End > contentRng.Start
        If Right$(contentRng.text, 1) = " " Or Right$(contentRng.text, 1) = vbTab Or Right$(contentRng.text, 1) = ChrW(160) Then
            contentRng.End = contentRng.End - 1
        Else
            Exit Do
        End If
    Loop

    ' Aplica regra do ponto final
    If Not hadFinalPeriod Then
        If contentRng.End > contentRng.Start Then
            If Right$(contentRng.text, 1) <> "." Then
                contentRng.Collapse wdCollapseEnd
                contentRng.InsertAfter "."
            End If
        Else
            ' Ementa ficou vazia por algum motivo: nao insere ponto
        End If
    End If

    documentDirty = True
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao remover sufixo da ementa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' EMENTA - Remove aspas envolventes ("...", �...�, �...�, '...')
' Regras:
' - Remove aspas duplas ou simples do inicio e fim do paragrafo da ementa
' - Trata aspas ASCII (" ') e tipograficas (� � � �)
' - So remove se ambos os lados tiverem aspas correspondentes
'================================================================================

Public Sub RemoveEmentaQuotes(doc As Document)
    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = GetEmentaRange(doc)
    If rng Is Nothing Then Exit Sub

    ' Range de conteudo: sem marca de paragrafo final
    Dim contentRng As Range
    Set contentRng = rng.Duplicate
    If contentRng.End > contentRng.Start Then
        If Right$(contentRng.text, 1) = vbCr Then
            contentRng.End = contentRng.End - 1
        End If
    End If
    If contentRng.End <= contentRng.Start Then Exit Sub

    Dim txt As String
    txt = contentRng.text
    If Len(txt) < 2 Then Exit Sub

    Dim firstCh    As String
    Dim lastCh     As String
    Dim prevLastCh As String
    firstCh    = Left$(txt, 1)
    lastCh     = Right$(txt, 1)
    prevLastCh = ""
    If Len(txt) >= 2 Then prevLastCh = Mid$(txt, Len(txt) - 1, 1)

    ' closeOffset: posicao da aspa de fechamento contada a partir do fim do contentRng
    '   1 = ultimo caractere e a aspa
    '   2 = penultimo e a aspa; ultimo e ponto final
    Dim isMatch     As Boolean
    Dim closeOffset As Long
    isMatch     = False
    closeOffset = 1

    If firstCh = Chr(34) And lastCh = Chr(34) Then isMatch = True          ' " ... "
    If firstCh = ChrW(8220) And lastCh = ChrW(8221) Then isMatch = True   ' � ... �
    If firstCh = Chr(39) And lastCh = Chr(39) Then isMatch = True          ' ' ... '
    If firstCh = ChrW(8216) And lastCh = ChrW(8217) Then isMatch = True   ' � ... �

    ' Caso 2: ponto final apos aspa de fechamento ("texto".)
    If Not isMatch And lastCh = "." And Len(txt) >= 3 Then
        If firstCh = Chr(34)    And prevLastCh = Chr(34)    Then isMatch = True: closeOffset = 2
        If firstCh = ChrW(8220) And prevLastCh = ChrW(8221) Then isMatch = True: closeOffset = 2
        If firstCh = Chr(39)    And prevLastCh = Chr(39)    Then isMatch = True: closeOffset = 2
        If firstCh = ChrW(8216) And prevLastCh = ChrW(8217) Then isMatch = True: closeOffset = 2
    End If

    If Not isMatch Then Exit Sub

    ' Remove aspa de fechamento primeiro (nao desloca o inicio do Range)
    Dim closeRng As Range
    Set closeRng = contentRng.Duplicate
    closeRng.Start = contentRng.End - closeOffset
    closeRng.End   = contentRng.End - closeOffset + 1
    closeRng.Delete

    ' Remove aspa de abertura
    Dim openRng As Range
    Set openRng = rng.Duplicate
    openRng.Start = rng.Start
    openRng.End   = rng.Start + 1
    openRng.Delete

    documentDirty = True
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao remover aspas da ementa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' EMENTA - Substitui "Indica ao DAE" ou "Sugere ao DAE" por "Indica ao Poder Executivo Municipal"
' em indica��es (documentos cujo t�tulo come�a com "INDICA��O").
'================================================================================

Public Sub ProcessEmentaIndicacao(doc As Document)
    On Error GoTo ErrorHandler

    Dim titleRng As Range
    Set titleRng = GetTituloRange(doc)
    If titleRng Is Nothing Then Exit Sub

    Dim titleText As String
    titleText = Trim$(titleRng.text)
    If Right$(titleText, 1) = vbCr Then titleText = Left$(titleText, Len(titleText) - 1)
    titleText = Trim$(titleText)

    Dim lowerTitle As String
    lowerTitle = LCase$(titleText)

    ' Verifica se o t�tulo come�a com "INDICA��O" ou "INDICACAO"
    If Left$(lowerTitle, 9) = "indica��o" Or Left$(lowerTitle, 9) = "indicacao" Then
        Dim ementaRng As Range
        Set ementaRng = GetEmentaRange(doc)
        If ementaRng Is Nothing Then Exit Sub

        Dim ementaText As String
        ementaText = ementaRng.text

        ' Remove o par�grafo vbCr final para manipula��o de string
        Dim hasCr As Boolean
        hasCr = (Right$(ementaText, 1) = vbCr)
        
        Dim cleanEmenta As String
        cleanEmenta = ementaText
        If hasCr Then cleanEmenta = Left$(cleanEmenta, Len(cleanEmenta) - 1)

        Dim trimEmenta As String
        trimEmenta = Trim$(cleanEmenta)

        Dim lowerTrimEmenta As String
        lowerTrimEmenta = LCase$(trimEmenta)

        Dim modified As Boolean
        modified = False

        ' Verifica se come�a com "Indica ao DAE" ou "Sugere ao DAE" (13 caracteres)
        If Len(lowerTrimEmenta) >= 13 Then
            If Left$(lowerTrimEmenta, 13) = "indica ao dae" Then
                trimEmenta = "Indica ao Poder Executivo Municipal" & Mid$(trimEmenta, 14)
                modified = True
            ElseIf Left$(lowerTrimEmenta, 13) = "sugere ao dae" Then
                trimEmenta = "Indica ao Poder Executivo Municipal" & Mid$(trimEmenta, 14)
                modified = True
            End If
        End If

        If modified Then
            If hasCr Then
                SafeReplaceText ementaRng, trimEmenta & vbCr
            Else
                SafeReplaceText ementaRng, trimEmenta
            End If
            documentDirty = True
            LogMessage "ProcessEmentaIndicacao: Ementa da Indicacao atualizada de DAE para Poder Executivo Municipal", LOG_LEVEL_INFO
        End If
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao processar ementa da indicacao: " & Err.Description, LOG_LEVEL_WARNING
End Sub


Public Function GetEmentaLeadingLabelDeleteLen(ByVal txt As String) As Long
    On Error GoTo ErrorHandler

    GetEmentaLeadingLabelDeleteLen = 0
    If Len(txt) = 0 Then Exit Function

    Dim i As Long
    i = 1

    ' Ignora espacos/tabs no inicio
    Dim ch As String
    Do While i <= Len(txt)
        ch = Mid$(txt, i, 1)
        If ch = " " Or ch = vbTab Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    Dim wordLen As Long
    If i + 5 <= Len(txt) And LCase$(Mid$(txt, i, 6)) = "ementa" Then
        wordLen = 6
    ElseIf i + 6 <= Len(txt) And LCase$(Mid$(txt, i, 7)) = "assunto" Then
        wordLen = 7
    Else
        Exit Function
    End If

    Dim j As Long
    j = i + wordLen

    ' Ignora espacos/tabs entre a palavra e o ':'
    Do While j <= Len(txt)
        ch = Mid$(txt, j, 1)
        If ch = " " Or ch = vbTab Then
            j = j + 1
        Else
            Exit Do
        End If
    Loop

    ' Exige ':' para considerar prefixo
    If j > Len(txt) Then Exit Function
    If Mid$(txt, j, 1) <> ":" Then Exit Function
    j = j + 1

    ' Ignora espacos/tabs apos ':'
    Do While j <= Len(txt)
        ch = Mid$(txt, j, 1)
        If ch = " " Or ch = vbTab Then
            j = j + 1
        Else
            Exit Do
        End If
    Loop

    ' j aponta para o primeiro caractere a manter (ou para o CR final)
    GetEmentaLeadingLabelDeleteLen = j - 1
    Exit Function

ErrorHandler:
    GetEmentaLeadingLabelDeleteLen = 0
End Function

'================================================================================
' CONFIGURACAO DE PAGINA
'================================================================================
