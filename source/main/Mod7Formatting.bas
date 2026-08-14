Attribute VB_Name = "Mod7Formatting"
Option Explicit

' Mod7Formatting
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)

Public Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
        .Gutter = 0
        .Orientation = wdOrientPortrait
    End With

    ' Configuracao de pagina aplicada (sem log detalhado para performance)
    ApplyPageSetup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na configuracao de pagina: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

'================================================================================
' FORMATACAO DE FONTE (METODO TRADICIONAL - FALLBACK)
'================================================================================

Public Function ApplyStdFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim underlineRemovedCount As Long
    Dim isTitle As Boolean
    Dim hasConsiderando As Boolean
    Dim needsUnderlineRemoval As Boolean
    Dim needsBoldRemoval As Boolean
    Dim paraCount As Long

    ' Cache do count para performance
    paraCount = doc.Paragraphs.count

    For i = paraCount To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Protecao dinamica
        Set para = doc.Paragraphs(i)

        ' Early exit se processou demais (protecao contra documentos gigantes)
        If formattedCount > 50000 Then
            LogMessage "Limite de processamento atingido em ApplyStdFont (50000 paragrafos)", LOG_LEVEL_WARNING
            Exit For
        End If
        hasInlineImage = False
        isTitle = False
        hasConsiderando = False
        needsUnderlineRemoval = False
        needsBoldRemoval = False

        ' SUPER OTIMIZADO: Verificacao previa consolidada - uma unica leitura das propriedades
        Dim paraFont As Font
        Set paraFont = para.Range.Font
        Dim needsFontFormatting As Boolean
        needsFontFormatting = (paraFont.Name <> STANDARD_FONT) Or _
                             (paraFont.size <> STANDARD_FONT_SIZE) Or _
                             (paraFont.Color <> wdColorAutomatic)

        ' Cache das verificacoes de formatacao especial
        needsUnderlineRemoval = (paraFont.Underline <> wdUnderlineNone)
        needsBoldRemoval = (paraFont.Bold = True)

        ' Cache da contagem de InlineShapes para evitar multiplas chamadas
        Dim inlineShapesCount As Long
        inlineShapesCount = para.Range.InlineShapes.count

        ' OTIMIZACAO MAXIMA: Se nao precisa de nenhuma formatacao, pula imediatamente
        If Not needsFontFormatting And Not needsUnderlineRemoval And Not needsBoldRemoval And inlineShapesCount = 0 Then
            formattedCount = formattedCount + 1
            GoTo NextParagraph
        End If

        If inlineShapesCount > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' OTIMIZADO: Verificacao de conteudo visual so quando necessario
        If Not hasInlineImage And (needsFontFormatting Or needsUnderlineRemoval Or needsBoldRemoval) Then
            If HasVisualContent(para) Then
                hasInlineImage = True
                skippedCount = skippedCount + 1
            End If
        End If

        ' OTIMIZADO: Verificacao consolidada de tipo de paragrafo - uma unica leitura do texto
        Dim paraFullText As String
        Dim isSpecialParagraph As Boolean
        isSpecialParagraph = False

        ' So faz verificacao de texto se for necessario para formatacao especial
        If needsUnderlineRemoval Or needsBoldRemoval Then
            paraFullText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' Verifica se e o primeiro paragrafo com texto (titulo) - otimizado
            If i <= 3 And para.Format.alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If

            ' Verifica se o paragrafo comeca com "considerando" - otimizado
            If Len(paraFullText) >= CONSIDERANDO_MIN_LENGTH And LCase(Left(paraFullText, CONSIDERANDO_MIN_LENGTH)) = CONSIDERANDO_PREFIX Then
                hasConsiderando = True
            End If

            ' Verifica se e um paragrafo especial - otimizado
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove pontuacao final para analise com protecao
            Dim punctCounter As Long
            punctCounter = 0
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";") And punctCounter < 50
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
                punctCounter = punctCounter + 1
            Loop
            cleanParaText = Trim(LCase(cleanParaText))

            ' Vereador NAO e mais tratado como paragrafo especial (negrito deve ser removido)
            If cleanParaText = "justificativa" Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
                LogMessage "Paragrafo especial detectado em ApplyStdFont (negrito preservado): " & cleanParaText, LOG_LEVEL_INFO
            End If

            ' O paragrafo ANTERIOR a "vereador" nao precisa mais preservar negrito
            Dim isBeforeVereador As Boolean
            isBeforeVereador = False
        End If

        ' FORMATACAO PRINCIPAL - So executa se necessario
        If needsFontFormatting Then
            If Not hasInlineImage Then
                ' Formatacao rapida para paragrafos sem imagens usando metodo seguro
                If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback para metodo tradicional em caso de erro
                    With paraFont
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                    formattedCount = formattedCount + 1
                End If
            Else
                ' NOVO: Formatacao protegida para paragrafos COM imagens
                If ProtectImagesInRange(para.Range) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback: formatacao basica segura CONSOLIDADA
                    Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, False, False)
                    formattedCount = formattedCount + 1
                End If
            End If
        End If

        ' FORMATACAO ESPECIAL CONSOLIDADA - Remove sublinhado e negrito em uma unica passada
        If needsUnderlineRemoval Or needsBoldRemoval Then
            ' Determina quais formatacoes remover
            Dim removeUnderline As Boolean
            Dim removeBold As Boolean
            removeUnderline = needsUnderlineRemoval And Not isTitle
            removeBold = needsBoldRemoval And Not isTitle And Not hasConsiderando And Not isSpecialParagraph And Not isBeforeVereador

            ' Se precisa remover alguma formatacao
            If removeUnderline Or removeBold Then
                If Not hasInlineImage Then
                    ' Formatacao rapida para paragrafos sem imagens
                    If removeUnderline Then paraFont.Underline = wdUnderlineNone
                    If removeBold Then paraFont.Bold = False
                Else
                    ' Formatacao protegida CONSOLIDADA para paragrafos com imagens
                    Call FormatCharacterByCharacter(para, "", 0, 0, removeUnderline, removeBold)
                End If

                If removeUnderline Then underlineRemovedCount = underlineRemovedCount + 1
            End If
        End If

NextParagraph:
    Next i

    ' Marca documento como modificado se houve formatacao
    If formattedCount > 0 Then documentDirty = True

    ' Log otimizado
    If skippedCount > 0 Then
        LogMessage "Fontes formatadas: " & formattedCount & " paragrafos (incluindo " & skippedCount & " com protecao de imagens)"
    End If

    ApplyStdFont = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao de fonte: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFont = False
End Function

'================================================================================
' FORMATACAO CARACTERE POR CARACTERE CONSOLIDADA
'================================================================================

Public Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
    On Error Resume Next

    Dim j As Long
    Dim charCount As Long
    Dim charRange As Range

    charCount = SafeGetCharacterCount(para.Range) ' Cache da contagem segura

    If charCount > 0 Then ' Verificacao de seguranca
        For j = 1 To charCount
            Set charRange = para.Range.Characters(j)
            If charRange.InlineShapes.count = 0 Then
                With charRange.Font
                    ' Aplica formatacao de fonte se especificada
                    If fontName <> "" Then .Name = fontName
                    If fontSize > 0 Then .size = fontSize
                    If fontColor >= 0 Then .Color = fontColor

                    ' Remove formatacoes especiais se solicitado
                    If removeUnderline Then .Underline = wdUnderlineNone
                    If removeBold Then .Bold = False
                End With
            End If
        Next j
    End If
End Sub

'================================================================================
' FORMATACAO DE PARAGRAFOS
'================================================================================

Public Function ApplyStdParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim paragraphIndent As Single
    Dim firstIndent As Single
    Dim rightMarginPoints As Single
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim paraText As String
    Dim prevPara As Paragraph

    rightMarginPoints = 0

    ' Cache do count para performance
    Dim paraCount As Long
    paraCount = doc.Paragraphs.count

    For i = paraCount To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Protecao dinamica
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        ' Early exit se processou demais
        If formattedCount > 50000 Then
            LogMessage "Limite de processamento atingido em ApplyStdParagraphs (50000 paragrafos)", LOG_LEVEL_WARNING
            Exit For
        End If

        If para.Range.InlineShapes.count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' Protecao adicional: verifica outros tipos de conteudo visual
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' Aplica formatacao de paragrafo para TODOS os paragrafos
        ' (independente se contem imagens ou nao)

        ' Limpeza robusta de espacos multiplos - SEMPRE aplicada
        Dim cleanText As String
        cleanText = para.Range.text

        ' OTIMIZADO: Combinacao de multiplas operacoes de limpeza em um bloco
        If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
            ' Remove multiplos espacos consecutivos com protecao
            Dim cleanCounter As Long
            cleanCounter = 0
            Do While InStr(cleanText, "  ") > 0 And cleanCounter < MAX_LOOP_ITERATIONS
                cleanText = Replace(cleanText, "  ", " ")
                cleanCounter = cleanCounter + 1
            Loop

            ' Remove espacos antes/depois de quebras de linha
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)

            ' Remove tabs extras e converte para espacos com protecao
            cleanCounter = 0
            Do While InStr(cleanText, vbTab & vbTab) > 0 And cleanCounter < MAX_LOOP_ITERATIONS
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
                cleanCounter = cleanCounter + 1
            Loop
            cleanText = Replace(cleanText, vbTab, " ")

            ' Limpeza final de espacos multiplos com protecao
            cleanCounter = 0
            Do While InStr(cleanText, "  ") > 0 And cleanCounter < MAX_LOOP_ITERATIONS
                cleanText = Replace(cleanText, "  ", " ")
                cleanCounter = cleanCounter + 1
            Loop
        End If

        ' Verifica se e um paragrafo especial ANTES de limpar o texto
        Dim isSpecialFormatParagraph As Boolean
        isSpecialFormatParagraph = False

        Dim checkText As String
        checkText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        ' Remove pontuacao final para analise com protecao
        Dim checkCounter As Long
        checkCounter = 0
        Do While Len(checkText) > 0 And (Right(checkText, 1) = "." Or Right(checkText, 1) = "," Or Right(checkText, 1) = ":" Or Right(checkText, 1) = ";") And checkCounter < 50
            checkText = Left(checkText, Len(checkText) - 1)
            checkCounter = checkCounter + 1
        Loop
        checkText = Trim(LCase(checkText))

        ' Verifica se e "Justificativa", "Anexo", "Anexos" ou padrao de vereador
        If checkText = JUSTIFICATIVA_TEXT Or IsAnexoPattern(checkText) Or IsVereadorPattern(checkText) Then
            isSpecialFormatParagraph = True
        End If

        ' Aplica o texto limpo APENAS se nao ha imagens E nao e paragrafo especial
        If cleanText <> para.Range.text And Not hasInlineImage And Not isSpecialFormatParagraph Then
            SafeReplaceText para.Range, cleanText
        End If

        ' Formatacao de paragrafo - SEMPRE aplicada (exceto para paragrafos especiais)
        If Not isSpecialFormatParagraph Then
            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                .RightIndent = rightMarginPoints
                .SpaceBefore = 0
                .SpaceAfter = 0

                If para.alignment = wdAlignParagraphCenter Then
                    .leftIndent = 0
                    .firstLineIndent = 0
                Else
                    firstIndent = .firstLineIndent
                    paragraphIndent = .leftIndent
                    If paragraphIndent >= CentimetersToPoints(5) Then
                        .leftIndent = CentimetersToPoints(9)
                    ElseIf firstIndent < CentimetersToPoints(5) Then
                        .leftIndent = CentimetersToPoints(0)
                        .firstLineIndent = CentimetersToPoints(2.5)
                    End If
                End If
            End With

            If para.alignment = wdAlignParagraphLeft Then
                para.alignment = wdAlignParagraphJustify
            End If
        End If

        formattedCount = formattedCount + 1
    Next i

    ' Marca documento como modificado se houve formatacao
    If formattedCount > 0 Then documentDirty = True

    ' Log atualizado para refletir que todos os paragrafos sao formatados
    If skippedCount > 0 Then
        LogMessage "Paragrafos formatados: " & formattedCount & " (incluindo " & skippedCount & " com protecao de imagens)"
    End If

    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao de paragrafos: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdParagraphs = False
End Function

'================================================================================
' FORMAT SECOND PARAGRAPH - FORMATACAO APENAS DO 2 PARAGRAFO
'================================================================================

Public Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' GARANTIA: ajustes de linhas em branco devem ocorrer apos normalizar quebras de linha (Shift+Enter)
    ' para paragrafos, pois a logica depende de doc.Paragraphs.
    On Error Resume Next
    ReplaceLineBreaksWithParagraphBreaks doc
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long

    ' Usa a ementa identificada pelo sistema de estrutura
    secondParaIndex = ementaParaIndex

    ' Aplica formatacao especifica apenas ao 2 paragrafo
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)

        ' Substitui palavras iniciais conforme regras especificas
        Dim paraFullText As String
        paraFullText = para.Range.text
        paraFullText = Trim(Replace(Replace(paraFullText, vbCr, ""), vbLf, ""))

        Dim lowerStart As String
        Dim wasReplaced As Boolean
        wasReplaced = False

        ' Verifica se inicia com "Solicita" (case insensitive)
        If Len(paraFullText) >= 8 Then
            lowerStart = LCase(Left(paraFullText, 8))
            If lowerStart = "solicita" Then
                SafeReplaceText para.Range, "Requer" & Mid(paraFullText, 9) & vbCr
                LogMessage "Palavra inicial 'Solicita' substituida por 'Requer' no 2 paragrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Verifica se inicia com "Pede" (case insensitive)
        If Not wasReplaced And Len(paraFullText) >= 4 Then
            lowerStart = LCase(Left(paraFullText, 4))
            If lowerStart = "pede" Then
                SafeReplaceText para.Range, "Requer" & Mid(paraFullText, 5) & vbCr
                LogMessage "Palavra inicial 'Pede' substituida por 'Requer' no 2 paragrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Verifica se inicia com "Sugere" (case insensitive)
        If Not wasReplaced And Len(paraFullText) >= 6 Then
            lowerStart = LCase(Left(paraFullText, 6))
            If lowerStart = "sugere" Then
                SafeReplaceText para.Range, "Indica" & Mid(paraFullText, 7) & vbCr
                LogMessage "Palavra inicial 'Sugere' substituida por 'Indica' no 2 paragrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Atualiza o texto do paragrafo se houve substituicao
        If wasReplaced Then
            paraFullText = para.Range.text
        End If

        ' Remove ", neste municipio" se estiver no final do paragrafo
        paraFullText = para.Range.text
        paraFullText = Trim(Replace(Replace(paraFullText, vbCr, ""), vbLf, ""))

        If Len(paraFullText) > 17 Then ' Tamanho minimo para conter ", neste municipio"
            Dim lowerText As String
            lowerText = LCase(paraFullText)

            Dim lowerTextNorm As String
            lowerTextNorm = NormalizeForComparison(lowerText)

            ' Verifica se termina com ", neste municipio"
            If Right(lowerTextNorm, 17) = ", neste municipio" Then
                ' Remove os ultimos 17 caracteres
                SafeReplaceText para.Range, Left(paraFullText, Len(paraFullText) - 17) & vbCr
                LogMessage "String ', neste municipio' removida do 2 paragrafo", LOG_LEVEL_INFO
            End If
        End If

        ' PRIMEIRO: Adiciona 2 linhas em branco ANTES do 2 paragrafo
        Dim insertionPoint As Range
        Set insertionPoint = para.Range
        insertionPoint.Collapse wdCollapseStart

        ' Verifica se ja existem linhas em branco antes
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)

        ' Adiciona linhas em branco conforme necessario para chegar a 2
        If blankLinesBefore < 2 Then
            Dim linesToAdd As Long
            linesToAdd = 2 - blankLinesBefore

            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines

            ' Atualiza o indice do segundo paragrafo (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If

        ' FORMATACAO PRINCIPAL: Aplica formatacao SEMPRE, protegendo apenas as imagens
        With para.Format
            .leftIndent = CentimetersToPoints(9)      ' Recuo a esquerda de 9 cm
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo a direita
            .alignment = wdAlignParagraphJustify      ' Justificado
        End With

        ' SEGUNDO: Adiciona 2 linhas em branco DEPOIS do 2 paragrafo
        Dim insertionPointAfter As Range
        Set insertionPointAfter = para.Range
        insertionPointAfter.Collapse wdCollapseEnd

        ' Verifica se ja existem linhas em branco depois
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)

        ' Adiciona linhas em branco conforme necessario para chegar a 2
        If blankLinesAfter < 2 Then
            Dim linesToAddAfter As Long
            linesToAddAfter = 2 - blankLinesAfter

            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If

        ' Se tem imagens, apenas registra (mas nao pula a formatacao)
        If HasVisualContent(para) Then
            LogMessage "2 paragrafo formatado com protecao de imagem e linhas em branco (posicao: " & secondParaIndex & ")", LOG_LEVEL_INFO
        Else
            LogMessage "2 paragrafo formatado com 2 linhas em branco antes e depois (posicao: " & secondParaIndex & ")", LOG_LEVEL_INFO
        End If
    Else
        LogMessage "2 paragrafo nao encontrado para formatacao", LOG_LEVEL_WARNING
    End If

    FormatSecondParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao do 2 paragrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatSecondParagraph = False
End Function

'================================================================================
' FORMATACAO DO VOCATIVO - Recuo 1a linha 2,5 cm + Texto Justificado
'================================================================================

'================================================================================
' FORMAT POST-EMENTA BODY PARAGRAPHS (2o ao 4o paragrafo apos Ementa)
' Recuo 1a linha 2,5 cm + Texto Justificado
'================================================================================

Public Sub FormatPostEmentaBodyParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub
    If ementaParaIndex <= 0 Or ementaParaIndex > doc.Paragraphs.count Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim nonBlankCount As Long
    Dim formattedCount As Long

    nonBlankCount = 0
    formattedCount = 0

    ' Percorre os paragrafos apos a Ementa, pulando linhas em branco
    For i = ementaParaIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For

        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Ignora paragrafos vazios (linhas em branco)
        If paraText = "" And Not HasVisualContent(para) Then
            GoTo NextPara
        End If

        nonBlankCount = nonBlankCount + 1

        ' Aplica formatacao apenas ao 2o, 3o e 4o paragrafo nao-vazio apos a Ementa
        If nonBlankCount >= 2 And nonBlankCount <= 4 Then
            With para.Range.ParagraphFormat
                .leftIndent = CentimetersToPoints(0)
                .firstLineIndent = CentimetersToPoints(2.5)
                .RightIndent = 0
                .Alignment = wdAlignParagraphJustify
            End With
            formattedCount = formattedCount + 1
        End If

        ' Para apos encontrar o 4o paragrafo nao-vazio
        If nonBlankCount >= 4 Then Exit For

NextPara:
    Next i

    If formattedCount > 0 Then
        LogMessage "Paragrafos 2-4 apos Ementa formatados: " & formattedCount & " (justificado, recuo 2,5 cm)", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro na formatacao dos paragrafos pos-Ementa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' HELPER FUNCTIONS FOR BLANK LINE - Funcoes auxiliares para linhas em branco
'================================================================================
' Nota: CountBlankLinesBefore ja esta definida nas linhas 918-958
' (secao de identificacao de estrutura do documento)

'================================================================================
' SECOND PARAGRAPH LOCATION HELPER - Localiza o segundo paragrafo
'================================================================================

Public Function GetSecondParagraphIndex(doc As Document) As Long
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long

    actualParaIndex = 0

    ' Encontra o 2 paragrafo com conteudo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o paragrafo tem texto ou conteudo visual, conta como paragrafo valido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1

            ' Retorna o indice do 2 paragrafo
            If actualParaIndex = 2 Then
                GetSecondParagraphIndex = i
                Exit Function
            End If
        End If

        ' Protecao: processa ate 20 paragrafos para encontrar o 2
        If i > 20 Then Exit For
    Next i

    GetSecondParagraphIndex = 0  ' Nao encontrado
    Exit Function

ErrorHandler:
    GetSecondParagraphIndex = 0
End Function


Public Function CountBlankLinesAfter(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler

    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String

    count = 0

    ' Verifica paragrafos posteriores (maximo 5 para performance)
    For i = paraIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For

        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o paragrafo esta vazio, conta como linha em branco
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' Se encontrou paragrafo com conteudo, para de contar
            Exit For
        End If

        ' Limite de seguranca
        If count >= 5 Then Exit For
    Next i

    CountBlankLinesAfter = count
    Exit Function

ErrorHandler:
    CountBlankLinesAfter = 0
End Function

'================================================================================
' ENSURE BLANK LINE BELOW CONSIDERANDO - Garante uma linha pulada abaixo de CONSIDERANDO
'================================================================================

Public Function EnsureConsideringBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim p As Paragraph
    Dim cleanText As String
    Dim nextCleanText As String
    Dim addedCount As Long

    addedCount = 0
    LogMessage "Verificando linhas em branco abaixo de CONSIDERANDO...", LOG_LEVEL_INFO

    For i = doc.Paragraphs.count To 1 Step -1
        Set p = doc.Paragraphs(i)
        cleanText = Trim(Replace(Replace(p.Range.text, vbCr, ""), vbLf, ""))
        
        ' Verifica se o paragrafo comeca com CONSIDERANDO
        If Left(UCase(cleanText), 12) = "CONSIDERANDO" Then
            Dim needsBlank As Boolean
            needsBlank = False
            
            If i = doc.Paragraphs.count Then
                ' E o ultimo paragrafo, precisa de linha abaixo
                needsBlank = True
            Else
                ' Verifica se o proximo paragrafo e em branco
                nextCleanText = Trim(Replace(Replace(doc.Paragraphs(i + 1).Range.text, vbCr, ""), vbLf, ""))
                If nextCleanText <> "" Or HasVisualContent(doc.Paragraphs(i + 1)) Then
                    needsBlank = True
                End If
            End If
            
            If needsBlank Then
                ' Insere paragrafo em branco logo apos o paragrafo atual
                p.Range.InsertParagraphAfter
                addedCount = addedCount + 1
            End If
        End If
    Next i

    If addedCount > 0 Then
        LogMessage "Linhas em branco adicionadas abaixo de CONSIDERANDO: " & addedCount, LOG_LEVEL_INFO
    End If

    EnsureConsideringBlankLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao garantir linhas em branco abaixo de CONSIDERANDO: " & Err.Description, LOG_LEVEL_WARNING
    EnsureConsideringBlankLines = False
End Function

'================================================================================
' FORMATACAO DO PRIMEIRO PARAGRAFO
'================================================================================

Public Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim firstParaIndex As Long

    ' Identifica o 1 paragrafo (considerando apenas paragrafos com texto)
    actualParaIndex = 0
    firstParaIndex = 0

    ' Encontra o 1 paragrafo com conteudo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o paragrafo tem texto ou conteudo visual, conta como paragrafo valido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1

            ' Registra o indice do 1 paragrafo
            If actualParaIndex = 1 Then
                firstParaIndex = i
                Exit For ' Ja encontramos o 1 paragrafo
            End If
        End If

        ' Protecao expandida: processa ate 20 paragrafos para encontrar o 1
        If i > 20 Then Exit For
    Next i

    ' Aplica formatacao especifica apenas ao 1 paragrafo
    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(firstParaIndex)

        ' NOVO: Aplica formatacao SEMPRE, protegendo apenas as imagens
        ' Formatacao do 1 paragrafo: caixa alta, negrito e sublinhado
        If HasVisualContent(para) Then
            ' Para paragrafos com imagens, aplica formatacao caractere por caractere
            Dim n As Long
            Dim charCount4 As Long
            charCount4 = SafeGetCharacterCount(para.Range) ' Cache da contagem segura

            If charCount4 > 0 Then ' Verificacao de seguranca
                For n = 1 To charCount4
                    Dim charRange3 As Range
                    Set charRange3 = para.Range.Characters(n)
                    If charRange3.InlineShapes.count = 0 Then
                        With charRange3.Font
                            .AllCaps = True           ' Caixa alta (maiusculas)
                            .Bold = True              ' Negrito
                            .Underline = wdUnderlineSingle ' Sublinhado
                        End With
                    End If
                Next n
            End If
            LogMessage "1 paragrafo formatado com protecao de imagem (posicao: " & firstParaIndex & ")"
        Else
            ' Formatacao normal para paragrafos sem imagens
            With para.Range.Font
                .AllCaps = True           ' Caixa alta (maiusculas)
                .Bold = True              ' Negrito
                .Underline = wdUnderlineSingle ' Sublinhado
            End With
        End If

        ' Aplicar tambem formatacao de paragrafo - SEMPRE
        With para.Format
            .alignment = wdAlignParagraphCenter       ' Centralizado
            .leftIndent = 0                           ' Sem recuo a esquerda
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo a direita
        End With
    Else
        LogMessage "1 paragrafo nao encontrado para formatacao", LOG_LEVEL_WARNING
    End If

    FormatFirstParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao do 1 paragrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatFirstParagraph = False
End Function

'================================================================================
' REMOCAO DE MARCA D'AGUA
'================================================================================

Public Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As shape
    Dim i As Long
    Dim removedCount As Long

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Exists And header.Shapes.count > 0 Then
                For i = header.Shapes.count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header

        For Each header In sec.Footers
            If header.Exists And header.Shapes.count > 0 Then
                For i = header.Shapes.count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header
    Next sec

    If removedCount > 0 Then
        LogMessage "Marcas d'agua removidas: " & removedCount & " itens"
    End If
    ' Log de "nenhuma marca d'agua" removido para performance

    RemoveWatermark = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover marcas d'agua: " & Err.Description, LOG_LEVEL_ERROR
    RemoveWatermark = False
End Function

'================================================================================
' INSERCAO DE NUMEROS DE PAGINA NO RODAPE + INICIAIS DO USUARIO
'================================================================================
' Insere rodape com:
' - Iniciais do usuario a esquerda (Arial 6pt, cinza)
' - "Pagina X de Y" centralizado (Arial 10pt)
'--------------------------------------------------------------------------------

Public Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rngInitials As Range
    Dim rngPage As Range
    Dim rngDash As Range
    Dim rngNum As Range
    Dim fPage As Field
    Dim fTotal As Field
    Dim userInitials As String

    ' Obtem as iniciais do usuario atual
    userInitials = GetUserInitials()

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)

        If footer.Exists Then
            footer.LinkToPrevious = False

            ' Limpa todo o rodape
            footer.Range.Delete

            ' Insere iniciais do usuario a esquerda (Arial 6pt, cinza)
            Set rngInitials = footer.Range
            rngInitials.Collapse Direction:=wdCollapseStart
            rngInitials.text = userInitials
            With rngInitials.Font
                .Name = STANDARD_FONT
                .size = 6
                .Color = RGB(128, 128, 128)
            End With
            rngInitials.ParagraphFormat.alignment = wdAlignParagraphLeft
            rngInitials.InsertParagraphAfter

            ' Insere "Pagina X de Y" centralizado (numero da pagina e total de paginas)
            Set rngPage = footer.Range.Paragraphs.Last.Range
            rngPage.Collapse Direction:=wdCollapseStart
            rngPage.text = "P" & Chr(225) & "gina "
            rngPage.Collapse Direction:=wdCollapseEnd

            ' Campo PAGE (numero da pagina atual)
            Set fPage = rngPage.Fields.Add(Range:=rngPage, Type:=wdFieldPage)

            ' Texto " de "
            Set rngDash = footer.Range.Paragraphs.Last.Range
            rngDash.Collapse Direction:=wdCollapseEnd
            rngDash.text = " de "

            ' Campo NUMPAGES (total de paginas)
            Set rngNum = footer.Range.Paragraphs.Last.Range
            rngNum.Collapse Direction:=wdCollapseEnd
            Set fTotal = rngNum.Fields.Add(Range:=rngNum, Type:=wdFieldNumPages)

            ' Formata todo o paragrafo e centraliza os numeros de pagina
            With footer.Range.Paragraphs.Last.Range
                .ParagraphFormat.alignment = wdAlignParagraphCenter
                .Font.Name = STANDARD_FONT
                .Font.size = FOOTER_FONT_SIZE
                .Font.Color = wdColorAutomatic
            End With
        End If
    Next sec

    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir rodape: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

'================================================================================
' Retorna as iniciais do usuario baseado no nome de usuario do Windows
' Mapeamento:
'   avendemiato -> afv
'   csantos     -> cms
'   alexandre   -> ajc
'   lcurtes     -> lc
'   marta       -> mfcp
'   henrique    -> hmg
'   bruno       -> bra
'--------------------------------------------------------------------------------

Public Function GetUserInitials() As String
    On Error GoTo ErrorHandler

    Dim userName As String
    userName = LCase(Environ("USERNAME"))

    Select Case userName
        Case "avendemiato"
            GetUserInitials = "afv"
        Case "csantos"
            GetUserInitials = "cms"
        Case "alexandre"
            GetUserInitials = "ajc"
        Case "lcurtes"
            GetUserInitials = "lc"
        Case "marta"
            GetUserInitials = "mfcp"
        Case "henrique"
            GetUserInitials = "hmg"
        Case "bruno"
            GetUserInitials = "bra"
        Case Else
            ' Usuario nao mapeado: usa primeiras 3 letras do username
            If Len(userName) >= 3 Then
                GetUserInitials = Left(userName, 3)
            Else
                GetUserInitials = userName
            End If
    End Select

    Exit Function

ErrorHandler:
    GetUserInitials = "usr"
End Function

'================================================================================
' UTILITY: CM TO POINTS
'================================================================================

Public Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then
        CentimetersToPoints = cm * 28.35
    End If
End Function

'================================================================================
' UTILITY: SAFE USERNAME
'================================================================================

Public Function GetSafeUserName() As String
    On Error GoTo ErrorHandler

    Dim rawName As String
    Dim safeName As String
    Dim i As Integer
    Dim c As String

    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then
        On Error Resume Next
        rawName = CreateObject("WScript.Network").username
        On Error GoTo 0
    End If

    If rawName = "" Then
        rawName = "UsuarioDesconhecido"
    End If

    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_\-]" Then
            safeName = safeName & c
        ElseIf c = " " Then
            safeName = safeName & "_"
        End If
    Next i

    If safeName = "" Then safeName = "Usuario"

    GetSafeUserName = safeName
    Exit Function

ErrorHandler:
    GetSafeUserName = "Usuario"
End Function

'================================================================================
' OBTEM A PRIMEIRA PALAVRA DO DOCUMENTO
'================================================================================

Public Function GetFirstWord(doc As Document) As String
    On Error GoTo ErrorHandler

    GetFirstWord = ""

    ' Percorre os primeiros paragrafos ate encontrar texto
    Dim i As Long
    Dim paraText As String

    For i = 1 To doc.Paragraphs.count
        If i > 10 Then Exit For ' Limite de seguranca

        paraText = Trim(Replace(Replace(doc.Paragraphs(i).Range.text, vbCr, ""), vbLf, ""))

        If Len(paraText) > 0 Then
            ' Extrai a primeira palavra (ate o primeiro espaco)
            Dim spacePos As Long
            spacePos = InStr(paraText, " ")

            If spacePos > 0 Then
                GetFirstWord = Left(paraText, spacePos - 1)
            Else
                GetFirstWord = paraText
            End If

            Exit For
        End If
    Next i

    Exit Function

ErrorHandler:
    GetFirstWord = ""
End Function

'================================================================================
' VERIFICACAO DE DADOS SENSIVEIS (LGPD) - MODO ESTRITO
'================================================================================
' Objetivo:
' - Reduzir a checagem a achados realmente graves
' - Maximizar precisao usando validadores deterministicos quando possivel
'   (ex.: CPF/CNPJ com digitos verificadores, cartao com Luhn, CID com padrao estrito)

Public Function ClearAllFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Limpando formatacao..."

    ' SUPER OTIMIZADO: Verificacao unica de conteudo visual no documento
    Dim hasImages As Boolean
    Dim hasShapes As Boolean
    hasImages = (doc.InlineShapes.count > 0)
    hasShapes = (doc.Shapes.count > 0)
    Dim hasAnyVisualContent As Boolean
    hasAnyVisualContent = hasImages Or hasShapes

    Dim paraCount As Long
    Dim styleResetCount As Long

    If hasAnyVisualContent Then
        ' MODO SEGURO OTIMIZADO: Cache de verificacoes visuais por paragrafo
        Dim para As Paragraph
        Dim visualContentCache As Object ' Cache para evitar recalculos
        Set visualContentCache = CreateObject("Scripting.Dictionary")

        Dim clearCounter As Long
        clearCounter = 0
        For Each para In doc.Paragraphs
            clearCounter = clearCounter + 1
            ' DoEvents a cada 15 paragrafos para manter responsividade
            If clearCounter Mod 15 = 0 Then DoEvents

            On Error Resume Next

            ' Cache da verificacao de conteudo visual
            Dim paraKey As String
            paraKey = CStr(para.Range.Start) & "-" & CStr(para.Range.End)

            Dim hasVisualInPara As Boolean
            If visualContentCache.Exists(paraKey) Then
                hasVisualInPara = visualContentCache(paraKey)
            Else
                hasVisualInPara = HasVisualContent(para)
                visualContentCache.Add paraKey, hasVisualInPara
            End If

            If Not hasVisualInPara Then
                ' FORMATACAO CONSOLIDADA: Aplica todas as configuracoes em uma unica operacao
                With para.Range
                    ' Reset completo de fonte em uma unica operacao
                    With .Font
                        .Reset
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                        .Bold = False
                        .Italic = False
                        .Underline = wdUnderlineNone
                    End With

                    ' Reset completo de paragrafo em uma unica operacao
                    With .ParagraphFormat
                        .Reset
                        .alignment = wdAlignParagraphLeft
                        .LineSpacing = 12
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .leftIndent = 0
                        .RightIndent = 0
                        .firstLineIndent = 0
                    End With

                    ' Reset de bordas e sombreamento
                    .Borders.Enable = False
                    .Shading.Texture = wdTextureNone
                End With
                paraCount = paraCount + 1
            Else
                ' OTIMIZADO: Para paragrafos com imagens, formatacao protegida mais rapida
                Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, True, True)
                paraCount = paraCount + 1
            End If

            ' Protecao contra loops infinitos
            If paraCount > 1000 Then Exit For
            On Error GoTo ErrorHandler
        Next para

    Else
        ' MODO ULTRA-RAPIDO: Sem conteudo visual - formatacao global em uma unica operacao
        With doc.Range
            ' Reset completo de fonte
            With .Font
                .Reset
                .Name = STANDARD_FONT
                .size = STANDARD_FONT_SIZE
                .Color = wdColorAutomatic
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
            End With

            ' Reset completo de paragrafo
            With .ParagraphFormat
                .Reset
                .alignment = wdAlignParagraphLeft
                .LineSpacing = 12
                .SpaceBefore = 0
                .SpaceAfter = 0
                .leftIndent = 0
                .RightIndent = 0
                .firstLineIndent = 0
            End With

            On Error Resume Next
            .Borders.Enable = False
            .Shading.Texture = wdTextureNone
            On Error GoTo ErrorHandler
        End With

        paraCount = doc.Paragraphs.count
    End If

    ' OTIMIZADO: Reset de estilos em uma unica passada
    Dim styleCounter As Long
    styleCounter = 0
    For Each para In doc.Paragraphs
        styleCounter = styleCounter + 1
        ' DoEvents a cada 20 paragrafos para manter responsividade
        If styleCounter Mod 20 = 0 Then DoEvents

        On Error Resume Next
        para.Style = "Normal"
        styleResetCount = styleResetCount + 1
        ' Protecao contra loops infinitos
        If styleResetCount > 1000 Then Exit For
        On Error GoTo ErrorHandler
    Next para

    LogMessage "Formatacao limpa: " & paraCount & " paragrafos resetados", LOG_LEVEL_INFO

    ' Cleanup do cache de conteudo visual para evitar memory leak
    If Not visualContentCache Is Nothing Then
        visualContentCache.RemoveAll
        Set visualContentCache = Nothing
    End If

    ClearAllFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao limpar formatacao: " & Err.Description, LOG_LEVEL_WARNING
    ClearAllFormatting = False ' Nao falha o processo por isso
End Function

'================================================================================
' REMOVE PAGE NUMBER LINES - Remove linhas com padrao $NUMERO$/$ANO$/Pagina N
'================================================================================

Public Function RemovePageNumberLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim NextPara As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim removedCount As Long
    Dim i As Long

    removedCount = 0

    ' Percorre de tras para frente para nao afetar indices ao deletar
    For i = doc.Paragraphs.count To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Protecao dinamica

        Set para = doc.Paragraphs(i)
        paraText = para.Range.text
        cleanText = Trim(Replace(Replace(paraText, vbCr, ""), vbLf, ""))

        ' Verifica se a linha termina com o padrao desejado ou e pagina de requerimento
        If IsPageNumberLine(cleanText) Or IsRequerimentoPageLine(cleanText) Or IsIndicacaoPageLine(cleanText) Or IsMocaoPageLine(cleanText) Then
            ' Verifica se existe uma proxima linha
            Dim hasNextLine As Boolean
            Dim nextLineIsEmpty As Boolean
            hasNextLine = False
            nextLineIsEmpty = False

            If i < doc.Paragraphs.count Then
                hasNextLine = True
                Set NextPara = doc.Paragraphs(i + 1)
                Dim nextText As String
                nextText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

                ' Verifica se a proxima linha esta em branco
                If nextText = "" And Not HasVisualContent(NextPara) Then
                    nextLineIsEmpty = True
                End If
            End If

            ' Remove a linha com padrao de paginacao
            para.Range.Delete
            removedCount = removedCount + 1

            ' Se a proxima linha estava em branco, remove tambem
            If hasNextLine And nextLineIsEmpty Then
                ' Atualiza a referencia pois os indices mudaram
                If i <= doc.Paragraphs.count Then
                    Set NextPara = doc.Paragraphs(i)
                    nextText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

                    ' Confirma que ainda esta vazia antes de deletar
                    If nextText = "" And Not HasVisualContent(NextPara) Then
                        NextPara.Range.Delete
                        removedCount = removedCount + 1
                    End If
                End If
            End If
        End If

        ' Protecao contra processamento excessivo
        If removedCount > 500 Then Exit For
    Next i

    If removedCount > 0 Then
        LogMessage "Linhas de paginacao removidas: " & removedCount & " linhas", LOG_LEVEL_INFO
    End If

    RemovePageNumberLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover linhas de paginacao: " & Err.Description, LOG_LEVEL_WARNING
    RemovePageNumberLines = False
End Function

'================================================================================
' REMOVE PARAGRAFOS COMPOSTOS UNICAMENTE POR UNDERLINES (_)
' Paragrafos cujo texto limpo seja uma sequencia de um ou mais caracteres "_"
' (underline/underscore) sao removidos integralmente.
'================================================================================

Public Sub RemoveUnderscoreOnlyParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim para As Paragraph
    Dim cleanText As String
    Dim removedCount As Long
    Dim isUnderscoreOnly As Boolean
    Dim j As Long

    removedCount = 0

    ' Percorre de tras para frente para nao afetar indices ao deletar
    For i = doc.Paragraphs.count To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Protecao dinamica

        Set para = doc.Paragraphs(i)
        cleanText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Verifica se o paragrafo possui texto e e composto unicamente por underlines
        If Len(cleanText) > 0 Then
            isUnderscoreOnly = True
            For j = 1 To Len(cleanText)
                If Mid(cleanText, j, 1) <> "_" Then
                    isUnderscoreOnly = False
                    Exit For
                End If
            Next j

            If isUnderscoreOnly Then
                para.Range.Delete
                removedCount = removedCount + 1
            End If
        End If
    Next i

    If removedCount > 0 Then
        documentDirty = True
        LogMessage "Paragrafos de underline removidos: " & removedCount, LOG_LEVEL_INFO
    End If

    Exit Sub
ErrorHandler:
    LogMessage "Erro ao remover paragrafos de underline: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' LIMPEZA DA ESTRUTURA DO DOCUMENTO
'================================================================================

Public Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim firstTextParaIndex As Long
    Dim emptyLinesRemoved As Long
    Dim leadingSpacesRemoved As Long
    Dim paraCount As Long

    ' Cache da contagem total de paragrafos
    paraCount = doc.Paragraphs.count

    ' Busca otimizada do primeiro paragrafo com texto
    firstTextParaIndex = -1
    For i = 1 To paraCount
        If i > doc.Paragraphs.count Then Exit For ' Protecao dinamica

        Set para = doc.Paragraphs(i)
        Dim paraTextCheck As String
        paraTextCheck = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Encontra o primeiro paragrafo com texto real
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If

        ' Protecao contra documentos muito grandes
        If i > MAX_INITIAL_PARAGRAPHS_TO_SCAN Then Exit For
    Next i

    ' OTIMIZADO: Remove linhas vazias ANTES do primeiro texto em uma unica passada
    If firstTextParaIndex > 1 Then
        ' Processa de tras para frente para evitar problemas com indices
        For i = firstTextParaIndex - 1 To 1 Step -1
            If i > doc.Paragraphs.count Or i < 1 Then Exit For ' Protecao dinamica

            Set para = doc.Paragraphs(i)
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' OTIMIZADO: Verificacao visual so se necessario
            If paraTextEmpty = "" Then
                If Not HasVisualContent(para) Then
                    para.Range.Delete
                    emptyLinesRemoved = emptyLinesRemoved + 1
                    ' Atualiza cache apos remocao
                    paraCount = paraCount - 1
                End If
            End If
        Next i
    End If

    ' Usa Find/Replace que e muito mais rapido que loop por paragrafo
    Dim rng As Range
    Set rng = doc.Range

    ' Remove espacos no inicio de linhas usando Find/Replace
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False

        ' Remove espacos/tabs no inicio de linhas usando Find/Replace simples
        .text = "^p "  ' Quebra seguida de espaco
        .Replacement.text = "^p"

        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            ' Protecao contra loop infinito
            If leadingSpacesRemoved > MAX_LOOP_ITERATIONS Then Exit Do
        Loop

        ' Remove tabs no inicio de linhas
        .text = "^p^t"  ' Quebra seguida de tab
        .Replacement.text = "^p"

        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            If leadingSpacesRemoved > MAX_LOOP_ITERATIONS Then Exit Do
        Loop
    End With

    ' Segunda passada para espacos no inicio do documento (sem ^p precedente)
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False  ' Nao usa wildcards nesta secao

        ' Posiciona no inicio do documento
        rng.Start = 0
        rng.End = 1

        ' Remove espacos/tabs no inicio absoluto do documento
        If rng.text = " " Or rng.text = vbTab Then
            ' Expande o range para pegar todos os espacos iniciais usando metodo seguro
            Do While rng.End <= doc.Range.End And (SafeGetLastCharacter(rng) = " " Or SafeGetLastCharacter(rng) = vbTab)
                rng.End = rng.End + 1
                leadingSpacesRemoved = leadingSpacesRemoved + 1
                If leadingSpacesRemoved > 100 Then Exit Do ' Protecao
            Loop

            If rng.Start < rng.End - 1 Then
                rng.Delete
            End If
        End If
    End With

    ' Log simplificado apenas se houve limpeza significativa
    If emptyLinesRemoved > 0 Then
        LogMessage "Estrutura limpa: " & emptyLinesRemoved & " linhas vazias removidas"
    End If

    CleanDocumentStructure = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza da estrutura: " & Err.Description, LOG_LEVEL_ERROR
    CleanDocumentStructure = False
End Function

'================================================================================
' REMOVE ALL TAB MARKS - Remove todas as marcas de tabulacao do documento
'================================================================================

Public Function RemoveAllTabMarks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim tabsRemoved As Long
    tabsRemoved = 0

    Set rng = doc.Range

    ' Remove todas as tabulacoes substituindo por espaco simples
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^t"  ' ^t representa tabulacao
        .Replacement.text = " "  ' Substitui por espaco simples
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True)
            tabsRemoved = tabsRemoved + 1
            ' Protecao contra loop infinito
            If tabsRemoved > 10000 Then
                LogMessage "Limite de remocao de tabulacoes atingido", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    If tabsRemoved > 0 Then
        LogMessage "Marcas de tabulacao removidas: " & tabsRemoved & " ocorrencias", LOG_LEVEL_INFO
    End If

    RemoveAllTabMarks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover marcas de tabulacao: " & Err.Description, LOG_LEVEL_ERROR
    RemoveAllTabMarks = False
End Function

'================================================================================
' REPLACE LINE BREAKS WITH PARAGRAPH BREAKS - Substitui quebras de linha por quebras de paragrafo
'================================================================================

Public Function ReplaceLineBreaksWithParagraphBreaks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim breaksReplaced As Long
    breaksReplaced = 0

    Set rng = doc.Range

    ' Substitui todas as quebras de linha manuais (^l) por quebras de paragrafo (^p)
    ' ^l = Shift+Enter (quebra de linha manual/soft return)
    ' ^p = Enter (quebra de paragrafo/hard return)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^l"  ' ^l representa quebra de linha manual (Shift+Enter)
        .Replacement.text = "^p"  ' ^p representa quebra de paragrafo (Enter)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True)
            breaksReplaced = breaksReplaced + 1
            ' Protecao contra loop infinito
            If breaksReplaced > 10000 Then
                LogMessage "Limite de substituicao de quebras de linha atingido", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    If breaksReplaced > 0 Then
        LogMessage "Quebras de linha substituidas por quebras de paragrafo: " & breaksReplaced & " ocorrencias", LOG_LEVEL_INFO
    End If

    ReplaceLineBreaksWithParagraphBreaks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao substituir quebras de linha: " & Err.Description, LOG_LEVEL_ERROR
    ReplaceLineBreaksWithParagraphBreaks = False
End Function

'================================================================================
' REMOVE PAGE BREAKS - Remove todas as quebras de pagina do documento
'================================================================================

Public Function RemovePageBreaks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim breaksRemoved As Long
    breaksRemoved = 0

    Set rng = doc.Range

    ' Remove quebras de pagina manuais (^m)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^m"  ' ^m representa quebra de pagina manual
        .Replacement.text = ""  ' Substitui por nada (remove)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True)
            breaksRemoved = breaksRemoved + 1
            ' Protecao contra loop infinito
            If breaksRemoved > 1000 Then
                LogMessage "Limite de remocao de quebras de pagina atingido", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    If breaksRemoved > 0 Then
        LogMessage "Quebras de pagina removidas: " & breaksRemoved & " ocorrencias", LOG_LEVEL_INFO
    End If

    RemovePageBreaks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover quebras de pagina: " & Err.Description, LOG_LEVEL_ERROR
    RemovePageBreaks = False
End Function

'================================================================================
' SAFE CHECK FOR VISUAL CONTENT - VERIFICACAO SEGURA DE CONTEUDO VISUAL
'================================================================================

Public Function HasVisualContent(para As Paragraph) As Boolean
    ' Usa a funcao segura implementada para compatibilidade total
    HasVisualContent = SafeHasVisualContent(para)
End Function

'================================================================================
' FORMATACAO DO TITULO DO DOCUMENTO
'================================================================================

Public Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Limpando espacos..."

    Dim rng As Range
    Dim spacesRemoved As Long
    Dim totalOperations As Long

    ' SUPER OTIMIZADO: Operacoes consolidadas em uma unica configuracao Find
    Set rng = doc.Range

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        ' OTIMIZACAO 1: Remove espacos multiplos (2 ou mais) em uma unica operacao
        ' Usa um loop otimizado que reduz progressivamente os espacos
        Do
            .text = "  "  ' Dois espacos
            .Replacement.text = " "  ' Um espaco

            Dim currentReplaceCount As Long
            currentReplaceCount = 0

            ' Executa ate nao encontrar mais duplos
            Do While .Execute(Replace:=True)
                currentReplaceCount = currentReplaceCount + 1
                spacesRemoved = spacesRemoved + 1
                ' Protecao otimizada - verifica a cada 200 operacoes
                If currentReplaceCount Mod 200 = 0 Then
                    DoEvents
                    If spacesRemoved > 2000 Then Exit Do
                End If
            Loop

            totalOperations = totalOperations + 1
            ' Se nao encontrou mais duplos ou atingiu limite, para
            If currentReplaceCount = 0 Or totalOperations > 10 Then Exit Do
        Loop
    End With

    ' OTIMIZACAO 2: Operacoes de limpeza de quebras de linha consolidadas
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade

        ' Remove multiplos espacos antes de quebras - metodo iterativo
        .text = "  ^p"  ' 2 espacos seguidos de quebra
        .Replacement.text = " ^p"  ' 1 espaco seguido de quebra
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop

        ' Segunda passada para garantir limpeza completa
        .text = " ^p"  ' Espaco antes de quebra
        .Replacement.text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop

        ' Remove multiplos espacos depois de quebras - metodo iterativo
        .text = "^p  "  ' Quebra seguida de 2 espacos
        .Replacement.text = "^p "  ' Quebra seguida de 1 espaco
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With

    ' OTIMIZACAO 3: Limpeza de tabs consolidada e otimizada
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False  ' Usar Find/Replace simples

        ' Remove multiplos tabs iterativamente
        .text = "^t^t"  ' 2 tabs
        .Replacement.text = "^t"  ' 1 tab
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop

        ' Converte tabs para espacos
        .text = "^t"
        .Replacement.text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With

    ' OTIMIZACAO 4: Verificacao final ultra-rapida de espacos duplos remanescentes
    Set rng = doc.Range
    With rng.Find
        .text = "  "
        .Replacement.text = " "
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop  ' Mais rapido que wdFindContinue

        Dim finalCleanCount As Long
        Do While .Execute(Replace:=True) And finalCleanCount < 100
            finalCleanCount = finalCleanCount + 1
            spacesRemoved = spacesRemoved + 1
        Loop
    End With

    ' PROTECAO ESPECIFICA: Garante espaco apos CONSIDERANDO
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

        ' Corrige CONSIDERANDO grudado com a proxima palavra
        .text = "CONSIDERANDOa"
        .Replacement.text = "CONSIDERANDO a"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .text = "CONSIDERANDOe"
        .Replacement.text = "CONSIDERANDO e"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .text = "CONSIDERANDOo"
        .Replacement.text = "CONSIDERANDO o"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .text = "CONSIDERANDOq"
        .Replacement.text = "CONSIDERANDO q"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
    End With

    ' Marca documento como modificado se houve limpeza
    If spacesRemoved > 0 Then documentDirty = True

    LogMessage "Limpeza de espacos concluida: " & spacesRemoved & " correcoes aplicadas (com protecao CONSIDERANDO)", LOG_LEVEL_INFO
    CleanMultipleSpaces = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza de espacos multiplos: " & Err.Description, LOG_LEVEL_WARNING
    CleanMultipleSpaces = False ' Nao falha o processo por isso
End Function

'================================================================================
' LIMITACAO DE LINHAS VAZIAS SEQUENCIAIS
'================================================================================

Public Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' GARANTIA: controle de linhas vazias deve ocorrer apos converter quebras de linha (^l) em paragrafos (^p)
    ' para que o Find/Replace em ^p funcione de forma consistente.
    On Error Resume Next
    ReplaceLineBreaksWithParagraphBreaks doc
    On Error GoTo ErrorHandler

    Application.StatusBar = "Controlando linhas..."

    ' IDENTIFICACAO DO SEGUNDO PARAGRAFO PARA PROTECAO
    Dim secondParaIndex As Long
    secondParaIndex = GetSecondParagraphIndex(doc)

    ' SUPER OTIMIZADO: Usa Find/Replace com wildcard para operacao muito mais rapida
    Dim rng As Range
    Dim linesRemoved As Long
    Dim totalReplaces As Long
    Dim passCount As Long

    passCount = 1 ' Inicializa contador de passadas

    Set rng = doc.Range

    ' METODO ULTRA-RAPIDO: Remove multiplas quebras consecutivas usando wildcard
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade

        ' Remove multiplas quebras consecutivas iterativamente
        .text = "^p^p^p^p"  ' 4 quebras
        .Replacement.text = "^p^p"  ' 2 quebras

        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop

        ' Remove 3 quebras -> 2 quebras
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras

        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
    End With

    ' SEGUNDA PASSADA: Remove quebras duplas restantes (2 quebras -> 1 quebra)
    If totalReplaces > 0 Then passCount = passCount + 1

    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue

        ' Converte quebras duplas em quebras simples
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras

        Dim secondPassCount As Long
        Do While .Execute(Replace:=True) And secondPassCount < 200
            secondPassCount = secondPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With

    ' VERIFICACAO FINAL: Garantir que nao ha mais de 1 linha vazia consecutiva
    If secondPassCount > 0 Then passCount = passCount + 1

    ' Metodo hibrido: Find/Replace para casos simples + loop apenas se necessario
    Set rng = doc.Range
    With rng.Find
        .text = "^p^p^p"  ' 3 quebras (2 linhas vazias + conteudo)
        .Replacement.text = "^p^p"  ' 2 quebras (1 linha vazia + conteudo)
        .MatchWildcards = False

        Dim finalPassCount As Long
        Do While .Execute(Replace:=True) And finalPassCount < 100
            finalPassCount = finalPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With

    If finalPassCount > 0 Then passCount = passCount + 1

    ' FALLBACK OTIMIZADO: Se ainda ha problemas, usa metodo tradicional limitado
    If finalPassCount >= 100 Then
        passCount = passCount + 1 ' Incrementa para o fallback

        Dim para As Paragraph
        Dim i As Long
        Dim emptyLineCount As Long
        Dim paraText As String
        Dim fallbackRemoved As Long

        i = 1
        emptyLineCount = 0

        Do While i <= doc.Paragraphs.count And fallbackRemoved < 50
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' Verifica se o paragrafo esta vazio
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1

                ' Se ja temos mais de 1 linha vazia consecutiva, remove esta
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                    ' Nao incrementa i pois removemos um paragrafo
                Else
                    i = i + 1
                End If
            Else
                ' Se encontrou conteudo, reseta o contador
                emptyLineCount = 0
                i = i + 1
            End If

            ' Responsividade e protecao otimizadas
            If fallbackRemoved Mod 10 = 0 Then DoEvents
            If i > 500 Then Exit Do ' Protecao adicional
        Loop
    End If

    LogMessage "Controle de linhas vazias concluido em " & passCount & " passada(s): " & linesRemoved & " linhas excedentes removidas (maximo 1 sequencial)", LOG_LEVEL_INFO
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro no controle de linhas vazias: " & Err.Description, LOG_LEVEL_WARNING
    LimitSequentialEmptyLines = False ' Nao falha o processo por isso
End Function

'================================================================================
' REMOCAO DE REALCES E BORDAS - REMOVE HIGHLIGHTING AND BORDERS
'================================================================================

Public Function RemoveAllHighlightsAndBorders(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Removendo realces e bordas..."

    Dim para As Paragraph
    Dim highlightCount As Long
    Dim borderCount As Long
    Dim processedCount As Long

    highlightCount = 0
    borderCount = 0
    processedCount = 0

    ' Remove realce de todo o documento primeiro (mais rapido)
    On Error Resume Next
    doc.Range.HighlightColorIndex = 0 ' Remove realce
    If Err.Number = 0 Then
        highlightCount = 1
        LogMessage "Realce removido do documento completo", LOG_LEVEL_INFO
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    ' Remove bordas de todos os paragrafos
    For Each para In doc.Paragraphs
        On Error Resume Next

        ' Remove bordas do paragrafo
        With para.Borders
            .Enable = False
        End With

        If Err.Number = 0 Then
            borderCount = borderCount + 1
        End If
        Err.Clear

        processedCount = processedCount + 1

        ' Responsividade
        If processedCount Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = "Removendo bordas: " & processedCount & " de " & doc.Paragraphs.count
        End If

        On Error GoTo ErrorHandler
    Next para

    LogMessage "Realces e bordas removidos: " & highlightCount & " realces, " & borderCount & " paragrafos com bordas", LOG_LEVEL_INFO
    RemoveAllHighlightsAndBorders = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover realces e bordas: " & Err.Description, LOG_LEVEL_WARNING
    RemoveAllHighlightsAndBorders = False ' Nao falha o processo por isso
End Function

'================================================================================
' REMOCAO DE PAGINAS VAZIAS NO FINAL - REMOVE EMPTY PAGES AT END
'================================================================================

Public Function RemoveEmptyPagesAtEnd(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Verificando paginas vazias no final..."

    ' Verifica se ha paginas vazias no final
    Dim totalPages As Long
    Dim lastPageRange As Range
    Dim lastPageText As String
    Dim pagesRemoved As Long
    Dim maxAttempts As Long
    Dim attemptCount As Long

    pagesRemoved = 0
    maxAttempts = 5 ' Maximo de tentativas para evitar loop infinito
    attemptCount = 0

    Do
        attemptCount = attemptCount + 1

        ' Obtem numero total de paginas
        On Error Resume Next
        totalPages = doc.ComputeStatistics(wdStatisticPages)
        If Err.Number <> 0 Then
            LogMessage "Nao foi possivel obter estatisticas de paginas: " & Err.Description, LOG_LEVEL_WARNING
            Err.Clear
            Exit Do
        End If
        Err.Clear
        On Error GoTo ErrorHandler

        ' Se ha apenas 1 pagina, nao remove nada
        If totalPages <= 1 Then
            Exit Do
        End If

        ' Obtem o range da ultima pagina
        Set lastPageRange = doc.Range
        lastPageRange.Start = doc.Range.End - 1
        lastPageRange.End = doc.Range.End

        ' Expande para incluir toda a ultima pagina
        lastPageRange.Expand wdParagraph

        ' Obtem texto da ultima pagina (ultimos paragrafos)
        Dim lastParaIndex As Long
        Dim para As Paragraph
        Dim hasContent As Boolean

        hasContent = False
        lastParaIndex = doc.Paragraphs.count

        ' Verifica os ultimos paragrafos em busca de conteudo
        Dim checkCount As Long
        checkCount = 0

        Do While lastParaIndex > 0 And checkCount < 20
            Set para = doc.Paragraphs(lastParaIndex)
            lastPageText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' Se encontrou conteudo de texto
            If Len(lastPageText) > 0 Then
                hasContent = True
                Exit Do
            End If

            ' Se encontrou imagem ou objeto
            If para.Range.InlineShapes.count > 0 Then
                hasContent = True
                Exit Do
            End If

            lastParaIndex = lastParaIndex - 1
            checkCount = checkCount + 1
        Loop

        ' Se a ultima pagina NAO tem conteudo, remove paragrafos vazios do final
        If Not hasContent Then
            Dim removedInThisPass As Long
            removedInThisPass = 0

            ' Remove paragrafos vazios do final (minimo necessario)
            lastParaIndex = doc.Paragraphs.count
            Do While lastParaIndex > 0 And removedInThisPass < 10
                Set para = doc.Paragraphs(lastParaIndex)
                lastPageText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

                ' Se e paragrafo vazio sem conteudo visual
                If Len(lastPageText) = 0 And para.Range.InlineShapes.count = 0 Then
                    para.Range.Delete
                    removedInThisPass = removedInThisPass + 1
                    pagesRemoved = pagesRemoved + 1
                    lastParaIndex = lastParaIndex - 1
                Else
                    ' Encontrou conteudo, para de remover
                    Exit Do
                End If

                ' Protecao contra loop infinito
                If removedInThisPass Mod 3 = 0 Then DoEvents
            Loop

            ' Se nao removeu nada nesta passada, termina
            If removedInThisPass = 0 Then
                Exit Do
            End If
        Else
            ' Ultima pagina tem conteudo, nao remove
            Exit Do
        End If

        ' Protecao contra tentativas excessivas
        If attemptCount >= maxAttempts Then
            LogMessage "Atingido numero maximo de tentativas de remocao de paginas vazias", LOG_LEVEL_WARNING
            Exit Do
        End If
    Loop

    If pagesRemoved > 0 Then
        LogMessage "Paginas vazias removidas do final: " & pagesRemoved & " paragrafo(s) vazio(s) removido(s)", LOG_LEVEL_INFO
    Else
        LogMessage "Nenhuma pagina vazia no final do documento", LOG_LEVEL_INFO
    End If

    RemoveEmptyPagesAtEnd = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover paginas vazias: " & Err.Description, LOG_LEVEL_WARNING
    RemoveEmptyPagesAtEnd = False ' Nao falha o processo por isso
End Function

'================================================================================
' FORMAT NUMBERED PARAGRAPHS INDENT - Aplica recuo em paragrafos iniciados com numero
'================================================================================

Public Function FormatNumberedParagraphsIndent(doc As Document) As Boolean
    ' Rotina desabilitada: remocao de todas as formatacoes especificamente definidas para listas numeradas
    FormatNumberedParagraphsIndent = True
End Function

'================================================================================
' FORMAT BULLETED PARAGRAPHS INDENT - Aplica recuo em paragrafos com marcadores
'================================================================================

Public Function FormatBulletedParagraphsIndent(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim firstChar As String
    Dim formattedCount As Long
    Dim defaultIndent As Single
    Dim i As Long

    formattedCount = 0

    ' Recuo padrao de lista com marcadores (36 pontos = 1.27 cm)
    defaultIndent = 36

    ' Array com os marcadores mais comuns
    Dim bulletMarkers() As String
    bulletMarkers = Split("*,-,>,+,~", ",")

    ' Percorre todos os paragrafos
    Dim bulletCounter As Long
    bulletCounter = 0
    For Each para In doc.Paragraphs
        bulletCounter = bulletCounter + 1
        If bulletCounter Mod 30 = 0 Then DoEvents ' Responsividade

        paraText = Trim(para.Range.text)

        ' Verifica se o paragrafo nao esta vazio
        If Len(paraText) > 0 Then
            ' Pega o primeiro caractere
            firstChar = Left(paraText, 1)

            ' Verifica se o primeiro caractere e um marcador comum
            Dim isBullet As Boolean
            isBullet = False

            For i = LBound(bulletMarkers) To UBound(bulletMarkers)
                If firstChar = bulletMarkers(i) Then
                    isBullet = True
                    Exit For
                End If
            Next i

            If isBullet Then
                ' Verifica se o paragrafo nao tem formatacao de lista ja aplicada
                If para.Range.ListFormat.ListType = 0 Then
                    ' Aplica o recuo a esquerda igual ao de uma lista com marcadores
                    With para.Format
                        .leftIndent = defaultIndent
                        .firstLineIndent = 0
                    End With
                    formattedCount = formattedCount + 1
                End If
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Paragrafos iniciados com marcador formatados com recuo de lista: " & formattedCount, LOG_LEVEL_INFO
    End If

    FormatBulletedParagraphsIndent = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de paragrafos com marcadores: " & Err.Description, LOG_LEVEL_WARNING
    FormatBulletedParagraphsIndent = False
End Function

'================================================================================
' REMOVER LINHAS EM BRANCO EXTRAS - Remove linhas duplicadas e aplica ajustes
'================================================================================

Public Sub RemoverLinhasEmBrancoExtras(doc As Document)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim removedCount As Long
    Dim replacedCount As Long

    removedCount = 0
    replacedCount = 0

    LogMessage "Removendo linhas em branco extras e aplicando ajustes...", LOG_LEVEL_INFO

    ' --- Espacamento simples em todos os paragrafos ---
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        On Error Resume Next
        With p.Format
            .LineSpacingRule = wdLineSpaceSingle
            .LineSpacing = 12
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        On Error GoTo ErrorHandler
    Next p

    ' --- Remove linhas em branco extras e espacos unicos ---
    For i = doc.Paragraphs.count To 2 Step -1
        Dim txtAtual As String, txtAnterior As String
        Dim pRange As Range
        
        Set pRange = doc.Paragraphs(i).Range
        If pRange.text = " " & vbCr Then
            pRange.MoveEnd wdCharacter, -1
            pRange.Delete
        End If
        
        Set pRange = doc.Paragraphs(i - 1).Range
        If pRange.text = " " & vbCr Then
            pRange.MoveEnd wdCharacter, -1
            pRange.Delete
        End If
        
        txtAtual = Trim(Replace(doc.Paragraphs(i).Range.text, vbCr, ""))
        txtAnterior = Trim(Replace(doc.Paragraphs(i - 1).Range.text, vbCr, ""))

        If txtAtual = "" And txtAnterior = "" Then
            On Error Resume Next
            doc.Paragraphs(i).Range.Delete
            If Err.Number = 0 Then removedCount = removedCount + 1
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next i

    ' --- Substituicoes no texto padrao ---
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False

        On Error Resume Next
        .text = "por intermedio do Setor,"
        .Replacement.text = "por interm" & ChrW(233) & "dio do Setor competente,"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "por interm" & ChrW(233) & "dio do Setor,"
        .Replacement.text = "por interm" & ChrW(233) & "dio do Setor competente,"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Indica ao Poder Executivo Municipal efetue"
        .Replacement.text = "Indica ao Poder Executivo Municipal que efetue"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Indica ao Poder Executivo Municipal e aos �rg�os competentes"
        .Replacement.text = "Indica ao Poder Executivo Municipal"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Indica ao Poder Executivo Municipal e aos " & ChrW(243) & "rg" & ChrW(227) & "os competentes"
        .Replacement.text = "Indica ao Poder Executivo Municipal"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Fomos procurados por municipes, solicitando essa providencia, pois segundo eles,"
        .Replacement.text = "Fomos procurados por mun" & ChrW(237) & "cipes solicitando essa provid" & ChrW(234) & "ncia, pois, segundo eles,"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Fomos procurados por mun" & ChrW(237) & "cipes, solicitando essa provid" & ChrW(234) & "ncia, pois segundo eles,"
        .Replacement.text = "Fomos procurados por mun" & ChrW(237) & "cipes solicitando essa provid" & ChrW(234) & "ncia, pois, segundo eles,"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1
        On Error GoTo ErrorHandler
    End With

    ' --- Reatualiza indices estruturais apos delecoes fisicas ---
    ' Previne stale pointers: o loop de delecao acima pode ter deslocado
    ' tituloJustificativaIndex, ementaParaIndex e dataParaIndex.
    If removedCount > 0 Then IdentifyDocumentStructure doc

    ' --- Ajustes por paragrafo ---
    Dim para As Paragraph
    Dim adjustCounter As Long
    adjustCounter = 0
    For Each para In doc.Paragraphs
        adjustCounter = adjustCounter + 1
        If adjustCounter Mod 30 = 0 Then DoEvents ' Responsividade

        Dim cleanTxt As String
        cleanTxt = NormalizeForComparison(Trim(Replace(para.Range.text, vbCr, "")))
        cleanTxt = Replace(cleanTxt, "-", "")

        On Error Resume Next

        ' Paragrafo do Plenario (local e data) deve ficar sem espacamento antes/depois
        If InStr(cleanTxt, "plenario") > 0 And InStr(cleanTxt, "tancredo neves") > 0 Then
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
        End If

        ' Centraliza nome, cargo e partido (apenas apos a Justificativa, se existir)
        ' Excecao: vocativos como "Senhor Presidente,", "Senhores Vereadores,",
        ' "Senhora Vereadora," e "Senhores Vereadores (as)," ficam logo abaixo da
        ' ementa e nao devem ser centralizados por esta rotina de assinatura.
        If tituloJustificativaIndex = 0 Or adjustCounter > tituloJustificativaIndex Then
            If Not (Left(cleanTxt, 7) = "senhor " _
                 Or Left(cleanTxt, 7) = "senhora" _
                 Or Left(cleanTxt, 8) = "senhores" _
                 Or Left(cleanTxt, 12) = "considerando" _
                 Or Left(cleanTxt, 8) = "requeiro") Then
                If Left(cleanTxt, 8) = "vereador" _
                   Or Left(cleanTxt, 9) = "vereadora" _
                   Or InStr(cleanTxt, "presidente") > 0 _
                   Or InStr(cleanTxt, "prefeito") > 0 Then
        
                    ' Cargo
                    With para.Format
                        .leftIndent = 0
                        .RightIndent = 0
                        .firstLineIndent = 0
                        .alignment = wdAlignParagraphCenter
                    End With
        
                    ' Nome (paragrafo anterior)
                    If Not para.Previous Is Nothing Then
                        With para.Previous.Format
                            .leftIndent = 0
                            .RightIndent = 0
                            .firstLineIndent = 0
                            .alignment = wdAlignParagraphCenter
                        End With
                        para.Previous.Range.Font.Bold = True
                    End If
        
                    ' Partido (paragrafo seguinte)
                    If Not para.Next Is Nothing Then
                        With para.Next.Format
                            .leftIndent = 0
                            .RightIndent = 0
                            .firstLineIndent = 0
                            .alignment = wdAlignParagraphCenter
                        End With
                    End If
                End If
            End If
        End If

        On Error GoTo ErrorHandler
    Next para

    LogMessage "Linhas em branco removidas: " & removedCount & ", substituicoes: " & replacedCount, LOG_LEVEL_INFO
    
    ' CORRECAO CRITICA (Index Staleness):
    ' Como linhas em branco foram deletadas fisicamente, os indices globais (titulo, ementa, justificativa) 
    ' agora apontam para o limbo (desalinhados). For�amos a reconstru��o do cache antes de prosseguir.
    If removedCount > 0 Then
        LogMessage "Reconstruindo cache arquitetural devido as delecoes fisicas...", LOG_LEVEL_INFO
        ClearParagraphCache
        BuildParagraphCache doc
    End If
    
    Exit Sub

ErrorHandler:
    LogMessage "Erro em RemoverLinhasEmBrancoExtras: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' ENHANCED IMAGE PROTECTION - Protecao aprimorada durante formatacao
'================================================================================

Public Function ProtectImagesInRange(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler

    ' Verifica se ha imagens no range antes de aplicar formatacao
    If targetRange.InlineShapes.count > 0 Then
        ' OTIMIZACAO VERDADEIRA: O Word moderno mantem a integridade da imagem 
        ' ao formatarmos a fonte no range completo (sem iterar por milhares de caracteres na COM interface)
        On Error Resume Next
        With targetRange.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic
        End With
        On Error GoTo ErrorHandler
    Else
        ' Range sem imagens - formatacao normal completa
        With targetRange.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic
        End With
    End If

    ProtectImagesInRange = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na protecao de imagens: " & Err.Description, LOG_LEVEL_WARNING
    ProtectImagesInRange = False
End Function

'================================================================================
' BACKUP VIEW SETTINGS - Faz backup das configuracoes de visualizacao originais
'================================================================================

Public Function BackupViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Salvando visualizacao..."

    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow

    ' Backup das configuracoes de visualizacao
    With originalViewSettings
        .ViewType = docWindow.View.Type
        ' Reguas sao controladas pelo Window, nao pelo View
        On Error Resume Next
        .ShowHorizontalRuler = docWindow.DisplayRulers
        .ShowVerticalRuler = docWindow.DisplayVerticalRuler
        On Error GoTo ErrorHandler
        .ShowFieldCodes = docWindow.View.ShowFieldCodes
        .ShowBookmarks = docWindow.View.ShowBookmarks
        .ShowParagraphMarks = docWindow.View.ShowParagraphs
        .ShowSpaces = docWindow.View.ShowSpaces
        .ShowTabs = docWindow.View.ShowTabs
        .ShowHiddenText = docWindow.View.ShowHiddenText
        .ShowAll = docWindow.View.ShowAll
        .ShowDrawings = docWindow.View.ShowDrawings
        .ShowObjectAnchors = docWindow.View.ShowObjectAnchors
        .ShowTextBoundaries = docWindow.View.ShowTextBoundaries
        .ShowHighlight = docWindow.View.ShowHighlight
        ' .ShowAnimation removida - pode nao existir em todas as versoes
        .DraftFont = docWindow.View.Draft
        .WrapToWindow = docWindow.View.WrapToWindow
        .ShowPicturePlaceHolders = docWindow.View.ShowPicturePlaceHolders
        .ShowFieldShading = docWindow.View.FieldShading
        .TableGridlines = docWindow.View.TableGridlines
        ' .EnlargeFontsLessThan removida - pode nao existir em todas as versoes
    End With

    LogMessage "Backup das configuracoes de visualizacao concluido"
    BackupViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup das configuracoes de visualizacao: " & Err.Description, LOG_LEVEL_WARNING
    BackupViewSettings = False
End Function

'================================================================================
' RESTORE VIEW SETTINGS - Restaura as configuracoes de visualizacao originais
'================================================================================

Public Function RestoreViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Restaurando visualizacao..."

    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow

    ' Restaura todas as configuracoes originais, EXCETO o zoom
    With docWindow.View
        .Type = originalViewSettings.ViewType
        .ShowFieldCodes = originalViewSettings.ShowFieldCodes
        .ShowBookmarks = originalViewSettings.ShowBookmarks
        .ShowParagraphs = originalViewSettings.ShowParagraphMarks
        .ShowSpaces = originalViewSettings.ShowSpaces
        .ShowTabs = originalViewSettings.ShowTabs
        .ShowHiddenText = originalViewSettings.ShowHiddenText
        .ShowAll = originalViewSettings.ShowAll
        .ShowDrawings = originalViewSettings.ShowDrawings
        .ShowObjectAnchors = originalViewSettings.ShowObjectAnchors
        .ShowTextBoundaries = originalViewSettings.ShowTextBoundaries
        .ShowHighlight = originalViewSettings.ShowHighlight
        ' .ShowAnimation removida para compatibilidade
        .Draft = originalViewSettings.DraftFont
        .WrapToWindow = originalViewSettings.WrapToWindow
        .ShowPicturePlaceHolders = originalViewSettings.ShowPicturePlaceHolders
        .FieldShading = originalViewSettings.ShowFieldShading
        .TableGridlines = originalViewSettings.TableGridlines
        ' .EnlargeFontsLessThan removida para compatibilidade

        ' ZOOM e mantido em 140% - unica configuracao que permanece alterada
        .Zoom.Percentage = 140
    End With

    ' Configuracoes especificas do Window (para reguas)
    docWindow.DisplayRulers = originalViewSettings.ShowHorizontalRuler
    docWindow.DisplayVerticalRuler = originalViewSettings.ShowVerticalRuler

    LogMessage "Configuracoes de visualizacao originais restauradas (zoom mantido em 140%)"
    RestoreViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar configuracoes de visualizacao: " & Err.Description, LOG_LEVEL_WARNING
    RestoreViewSettings = False
End Function

'================================================================================
' CLEANUP VIEW SETTINGS - Limpeza das variaveis de configuracoes de visualizacao
'================================================================================

Public Sub CleanupViewSettings()
    On Error Resume Next

    ' Reinicializa a estrutura de configuracoes
    With originalViewSettings
        .ViewType = 0
        .ShowVerticalRuler = False
        .ShowHorizontalRuler = False
        .ShowFieldCodes = False
        .ShowBookmarks = False
        .ShowParagraphMarks = False
        .ShowSpaces = False
        .ShowTabs = False
        .ShowHiddenText = False
        .ShowOptionalHyphens = False
        .ShowAll = False
        .ShowDrawings = False
        .ShowObjectAnchors = False
        .ShowTextBoundaries = False
        .ShowHighlight = False
        ' .ShowAnimation removida para compatibilidade
        .DraftFont = False
        .WrapToWindow = False
        .ShowPicturePlaceHolders = False
        .ShowFieldShading = 0
        .TableGridlines = False
        ' .EnlargeFontsLessThan removida para compatibilidade
    End With

    LogMessage "Variaveis de configuracoes de visualizacao limpas"
End Sub

'================================================================================
' SUBSTITUICAO DO PARAGRAFO DE LOCAL E DATA
'================================================================================

Public Sub ReplacePlenarioDateParagraph(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim matchCount As Integer
    Dim terms() As String

    Dim plenarioAcento As String
    plenarioAcento = "Plen" & ChrW(225) & "rio"  ' Plenario (com acento, ASCII-safe no fonte)

    ' Define os termos de busca
    Dim termsCsv As String
    termsCsv = "Palacio 15 de Junho,Plenario," & plenarioAcento & ",Dr. Tancredo Neves," & _
               " de janeiro de , de fevereiro de, de marco de, de abril de," & _
               " de maio de, de junho de, de julho de, de agosto de," & _
               " de setembro de, de outubro de, de novembro de, de dezembro de"
    terms = Split(termsCsv, ",")

    ' Processa cada paragrafo
    Dim plenCounter As Long
    plenCounter = 0
    For Each para In doc.Paragraphs
        plenCounter = plenCounter + 1
        If plenCounter Mod 30 = 0 Then DoEvents ' Responsividade

        matchCount = 0

        ' Pula paragrafos muito longos
        If Len(para.Range.text) <= 80 Then
            paraText = para.Range.text

            ' Conta matches
            Dim term As Variant
            For Each term In terms
                If InStr(1, paraText, CStr(term), vbTextCompare) > 0 Then
                    matchCount = matchCount + 1
                End If
                If matchCount >= 2 Then
                    ' Encontrou 2+ matches, faz a substituicao
                    ' Usa Delete + InsertAfter para preservar o marcador de paragrafo
                    Dim replaceTarget As Range
                    Set replaceTarget = para.Range
                    replaceTarget.MoveEnd unit:=wdCharacter, count:=-1 ' Exclui o marcador de paragrafo
                    replaceTarget.Delete
                    replaceTarget.InsertAfter plenarioAcento & " ""Dr. Tancredo Neves"", $DATAATUALEXTENSO$."
                    ' Aplica formatacao: centralizado e sem recuos
                    With para.Range.ParagraphFormat
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .alignment = wdAlignParagraphCenter
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                    End With
                    LogMessage "Paragrafo de plenario substituido e formatado", LOG_LEVEL_INFO
                    Exit For
                End If
            Next term
        End If
    Next para

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao processar paragrafos: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' Sub: ExecutarInstalador
' Descricao: Executa o z7_stdproposers_installer.cmd a partir da interface do Word
' Uso: Pode ser chamado de um botao na ribbon ou atalho de teclado
'================================================================================

Public Sub ExecutarInstalador()
    On Error GoTo ErrorHandler

    Dim installerPath As String
    Dim shellCmd As String
    Dim fso As Object
    Dim response As VbMsgBoxResult

    LogMessage "ExecutarInstalador acionado manualmente pelo usuario", LOG_LEVEL_INFO

    ' Pergunta confirmacao ao usuario
    Dim msgInstaller As String
    msgInstaller = "Deseja executar o instalador do Z7_STDPROPOSERS?" & vbCrLf & vbCrLf & _
                   "Isso ira:" & vbCrLf & _
                   " Baixar a versao mais recente do GitHub" & vbCrLf & _
                   " Instalar/atualizar o sistema" & vbCrLf & _
                   " Fechar o Word ao final da instalacao" & vbCrLf & vbCrLf & _
                   "Continuar?"
    response = MsgBox(msgInstaller, vbYesNo + vbQuestion, "Z7_STDPROPOSERS - Executar Instalador")

    If response <> vbYes Then
        LogMessage "ExecutarInstalador cancelado pelo usuario no prompt de confirmacao", LOG_LEVEL_INFO
        Exit Sub
    End If

    LogMessage "ExecutarInstalador: usuario confirmou execucao do instalador", LOG_LEVEL_INFO

    ' Caminho do instalador
    installerPath = Environ("USERPROFILE") & "\AppData\Local\Z7\Apps\Z7_StdProposers\installer.cmd"

    ' Verifica se o instalador existe
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(installerPath) Then
        LogMessage "ERRO CRITICO: Instalador nao encontrado para execucao manual em: " & installerPath, LOG_LEVEL_ERROR
        MsgBox "Instalador nao encontrado em:" & vbCrLf & installerPath & vbCrLf & vbCrLf & _
               "Baixe manualmente de: https://github.com/chrmsantos/Z7_StdProposers/raw/main/z7_stdproposers_installer.cmd", _
               vbExclamation, "Z7_STDPROPOSERS - Instalador Nao Encontrado"
        Exit Sub
    End If

    LogMessage "ExecutarInstalador: instalador localizado em " & installerPath & ". Preparando salvamento de documentos...", LOG_LEVEL_INFO

    ' Salva todos os documentos abertos antes de executar o instalador
    Dim doc As Object
    For Each doc In Application.Documents
        If doc.Saved = False Then
            On Error Resume Next
            LogMessage "Salvando documento pendente antes da instalacao: " & doc.Name, LOG_LEVEL_INFO
            doc.Save
            On Error GoTo ErrorHandler
        End If
    Next doc

    ' Executa o instalador em uma nova janela de comando
    shellCmd = "cmd.exe /c """ & installerPath & """"
    LogMessage "Disparando shell cmd para iniciar atualizacao manual: " & shellCmd, LOG_LEVEL_INFO
    CreateObject("WScript.Shell").Run shellCmd, 1, False

    ' Mensagem informativa
    MsgBox "O instalador foi iniciado em uma nova janela." & vbCrLf & vbCrLf & _
           "O Word sera fechado ao final da instalacao.", _
           vbInformation, "Z7_STDPROPOSERS - Instalador Iniciado"

    ' Fecha o Word apos 2 segundos (tempo para o instalador iniciar)
    Application.OnTime Now + TimeValue("00:00:02"), "FecharWord"

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar instalador: " & Err.Description, vbCritical, "Z7_STDPROPOSERS - Erro"
    LogMessage "Erro ao executar instalador: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' Sub: FecharWord
' Descricao: Fecha o Word (usado apos executar o instalador)
'================================================================================

Public Sub FecharWord()
    On Error Resume Next
    Application.Quit SaveChanges:=wdSaveChanges
End Sub

'================================================================================
' APLICACAO DE FORMATACAO FINAL UNIVERSAL
'================================================================================

Public Sub ApplyUniversalFinalFormatting(doc As Document)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraCount As Long
    Dim formattedCount As Long

    paraCount = doc.Paragraphs.count
    formattedCount = 0

    LogMessage "Aplicando formatacao final universal: Arial 12, espacamento 1.0, 1 linha entre paragrafos...", LOG_LEVEL_INFO

    ' APLICACAO EM LOTE (BLOCK FORMATTING) - ~100x mais rapido
    On Error Resume Next
    doc.AutoHyphenation = False
    With doc.Range
        .Font.Name = "Arial"
        .Font.size = 12
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.Hyphenation = False
    End With
    If Err.Number = 0 Then formattedCount = doc.Paragraphs.count
    Err.Clear
    On Error GoTo ErrorHandler

    ' Processa os paragrafos SOMENTE para logica de negocios condicional
    Dim universalCounter As Long
    universalCounter = 0
    For Each para In doc.Paragraphs
        universalCounter = universalCounter + 1
        If universalCounter Mod 20 = 0 Then DoEvents ' Responsividade

        On Error Resume Next

        ' GARANTIA: paragrafo contendo apenas "vereador" (case-insensitive), mesmo com pontuacao/hifens,
        ' deve sempre ficar como "Vereador", fonte normal, centralizado e com recuos 0 ao final do processamento.
        If tituloJustificativaIndex = 0 Or universalCounter > tituloJustificativaIndex Then
            If IsVereadorPattern(para.Range.text) Then
                ApplyVereadorParagraphFormatting para
            End If
        End If

        On Error GoTo ErrorHandler
    Next para

    LogMessage "Formatacao final aplicada: " & formattedCount & " de " & paraCount & " paragrafos", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao aplicar formatacao final universal: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' ADICAO DE ESPACAMENTO ESPECIAL (EMENTA, JUSTIFICATIVA, DATA)
'================================================================================

Public Sub AddSpecialElementsSpacing(doc As Document)
    On Error GoTo ErrorHandler

    Dim elementsProcessed As Long
    elementsProcessed = 0

    LogMessage "Adicionando espacamento especial (2 paragrafos em branco) para ementa, justificativa e data...", LOG_LEVEL_INFO

    Dim idx As Long
    Dim deletedCount As Long
    Dim newParaIdx As Long

    ' =========================================================================
    ' REGRA 1: EMENTA - 2 paragrafos em branco acima e abaixo
    ' Ajusta indices em ordem reversa (de baixo para cima): Data, Justificativa, Ementa
    ' para que ajustes em elementos posteriores nao afetem indices anteriores
    ' =========================================================================

    ' 1a. Remove paragrafos em branco EXISTENTES abaixo da Ementa
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        RemoveBlankLinesAfter doc, ementaParaIndex
    End If

    ' 1b. Remove paragrafos em branco EXISTENTES acima da Ementa
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        deletedCount = 0
        idx = ementaParaIndex - 1
        Do While idx >= 1
            If idx > doc.Paragraphs.count Then Exit Do
            Dim ementaAboveText As String
            ementaAboveText = Trim(Replace(Replace(doc.Paragraphs(idx).Range.text, vbCr, ""), vbLf, ""))
            If ementaAboveText = "" And Not HasVisualContent(doc.Paragraphs(idx)) Then
                doc.Paragraphs(idx).Range.Delete
                ementaParaIndex = ementaParaIndex - 1
                deletedCount = deletedCount + 1
                idx = idx - 1
            Else
                Exit Do
            End If
        Loop
    End If

    ' Ajusta indices posteriores por exclusoes acima
    If deletedCount > 0 Then
        If tituloJustificativaIndex >= 1 Then
            tituloJustificativaIndex = tituloJustificativaIndex - deletedCount
        End If
        If dataParaIndex >= 1 Then
            dataParaIndex = dataParaIndex - deletedCount
        End If
    End If

    ' 1c. Define posicao atual para insercao (pos da Ementa apos limpeza)
    Dim ementaIdx As Long
    ementaIdx = ementaParaIndex

    ' 1d. Insere 2 paragrafos em branco ACIMA da Ementa
    If ementaIdx > 0 And ementaIdx <= doc.Paragraphs.count Then
        Dim j As Long
        For j = 1 To 2
            doc.Paragraphs(ementaIdx).Range.InsertParagraphBefore
        Next j
        ementaIdx = ementaIdx + 2
        ementaParaIndex = ementaParaIndex + 2

        ' Ajusta indices posteriores pelas insercoes antes da Ementa
        If tituloJustificativaIndex >= 1 Then
            tituloJustificativaIndex = tituloJustificativaIndex + 2
        End If
        If dataParaIndex >= 1 Then
            dataParaIndex = dataParaIndex + 2
        End If
    End If

    ' 1e. Garante SpaceBefore=0 e SpaceAfter=0 na Ementa
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(ementaParaIndex).Format
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' 1f. Insere 2 paragrafos em branco ABAIXO da Ementa
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        doc.Paragraphs(ementaParaIndex).Range.InsertParagraphAfter
        doc.Paragraphs(ementaParaIndex).Range.InsertParagraphAfter
        Err.Clear
        On Error GoTo ErrorHandler
        elementsProcessed = elementsProcessed + 1
    End If

    ' =========================================================================
    ' REGRA 2: TITULO DA JUSTIFICATIVA - 2 paragrafos em branco acima e abaixo
    ' Processo deve rodar DEPOIS de ajustar a Ementa (para capturar os indices)
    ' =========================================================================

    ' 2a. Remove paragrafos em branco EXISTENTES acima do Titulo da Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        deletedCount = 0
        idx = tituloJustificativaIndex - 1
        Do While idx >= 1
            If idx > doc.Paragraphs.count Then Exit Do
            Dim justAboveText As String
            justAboveText = Trim(Replace(Replace(doc.Paragraphs(idx).Range.text, vbCr, ""), vbLf, ""))
            If justAboveText = "" And Not HasVisualContent(doc.Paragraphs(idx)) Then
                doc.Paragraphs(idx).Range.Delete
                tituloJustificativaIndex = tituloJustificativaIndex - 1
                deletedCount = deletedCount + 1
                idx = idx - 1
            Else
                Exit Do
            End If
        Loop
        If deletedCount > 0 Then
            If dataParaIndex >= 1 Then
                dataParaIndex = dataParaIndex - deletedCount
            End If
        End If
    End If

    ' 2b. Remove paragrafos em branco EXISTENTES abaixo do Titulo da Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        RemoveBlankLinesAfter doc, tituloJustificativaIndex
    End If

    ' 2c. Insere 2 paragrafos em branco ACIMA do Titulo da Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        For j = 1 To 2
            doc.Paragraphs(tituloJustificativaIndex).Range.InsertParagraphBefore
        Next j
        tituloJustificativaIndex = tituloJustificativaIndex + 2
        If dataParaIndex >= 1 Then
            dataParaIndex = dataParaIndex + 2
        End If
    End If

    ' 2d. Garante SpaceBefore=0 e SpaceAfter=0 no Titulo da Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(tituloJustificativaIndex).Format
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' 2e. Insere 2 paragrafos em branco ABAIXO do Titulo da Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        doc.Paragraphs(tituloJustificativaIndex).Range.InsertParagraphAfter
        doc.Paragraphs(tituloJustificativaIndex).Range.InsertParagraphAfter
        Err.Clear
        On Error GoTo ErrorHandler
        elementsProcessed = elementsProcessed + 1
    End If

    ' =========================================================================
    ' REGRA 3: DATA - 2 paragrafos em branco acima
    ' =========================================================================

    ' 3a. Remove paragrafos em branco EXISTENTES acima da Data
    If dataParaIndex > 0 And dataParaIndex <= doc.Paragraphs.count Then
        deletedCount = 0
        idx = dataParaIndex - 1
        Do While idx >= 1
            If idx > doc.Paragraphs.count Then Exit Do
            Dim dataAboveText As String
            dataAboveText = Trim(Replace(Replace(doc.Paragraphs(idx).Range.text, vbCr, ""), vbLf, ""))
            If dataAboveText = "" And Not HasVisualContent(doc.Paragraphs(idx)) Then
                doc.Paragraphs(idx).Range.Delete
                dataParaIndex = dataParaIndex - 1
                deletedCount = deletedCount + 1
                idx = idx - 1
            Else
                Exit Do
            End If
        Loop
    End If

    ' 3b. Insere 2 paragrafos em branco ACIMA da Data
    If dataParaIndex > 0 And dataParaIndex <= doc.Paragraphs.count Then
        For j = 1 To 2
            doc.Paragraphs(dataParaIndex).Range.InsertParagraphBefore
        Next j
        dataParaIndex = dataParaIndex + 2
        elementsProcessed = elementsProcessed + 1
    End If

    ' 3c. Garante SpaceBefore=0 e SpaceAfter=0 na Data
    If dataParaIndex > 0 And dataParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(dataParaIndex).Format
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    LogMessage "Espacamento especial aplicado a " & elementsProcessed & " elementos", LOG_LEVEL_INFO
    LogMessage "  Ementa: 2 paragrafos em branco acima e abaixo", LOG_LEVEL_INFO
    LogMessage "  Titulo Justificativa: 2 paragrafos em branco acima e abaixo", LOG_LEVEL_INFO
    LogMessage "  Data: 2 paragrafos em branco acima", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao adicionar espacamento especial: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' SUBSTITUI "N�" POR "n�" EXCENTUANDO O TITULO DO DOCUMENTO
'================================================================================

Public Sub ReplaceNoWithNoExceptTitle(doc As Document)
    On Error GoTo ErrorHandler

    Dim titleRng As Range
    Set titleRng = GetTituloRange(doc)

    Dim searchRng As Range
    Set searchRng = doc.Range

    Dim findTexts(0 To 2) As String
    findTexts(0) = "N" & Chr(186)  ' N� (maiusculo, ordinal)
    findTexts(1) = "N" & Chr(176)  ' N� (maiusculo, grau)
    findTexts(2) = "n" & Chr(186)  ' n� (minusculo, ordinal)

    Dim replaceText As String
    replaceText = "n" & Chr(176)   ' n� (minusculo, grau)

    Dim i As Integer
    Dim foundCount As Long
    foundCount = 0

    For i = 0 To 2
        Set searchRng = doc.Range
        With searchRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = findTexts(i)
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = True
            .MatchWholeWord = False
            
            Do While .Execute
                ' Verifica se o range encontrado esta dentro do titulo
                Dim inTitle As Boolean
                inTitle = False
                
                If Not titleRng Is Nothing Then
                    If searchRng.Start >= titleRng.Start And searchRng.End <= titleRng.End Then
                        inTitle = True
                    End If
                End If
                
                If Not inTitle Then
                    searchRng.text = replaceText
                    foundCount = foundCount + 1
                End If
                
                searchRng.Collapse wdCollapseEnd
            Loop
        End With
    Next i

    If foundCount > 0 Then
        LogMessage "Substituicao aplicada: 'N�'/'N�'/'n�' por 'n�' (" & foundCount & "x), exceto no titulo", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao substituir 'N�' por 'n�': " & Err.Description, LOG_LEVEL_WARNING
End Sub


'================================================================================
' REMOVE NUMBERING FROM BLANK PARAGRAPHS - Remove formatacao de numero/lista de paragrafos vazios
'================================================================================

Public Function RemoveNumberingFromBlankParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim removedCount As Long
    Dim totalParas As Long
    Dim i As Long

    removedCount = 0
    totalParas = doc.Paragraphs.count

    LogMessage "Removendo formatacao de numero de paragrafos em branco...", LOG_LEVEL_INFO

    For i = 1 To totalParas
        If i Mod 50 = 0 Then DoEvents ' Responsividade

        Set para = doc.Paragraphs(i)
        
        ' Verifica se o paragrafo esta vazio (apenas quebra de paragrafo e espacos)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText = "" And Not HasVisualContent(para) Then
            ' ListType = 0 (wdListNoNumbering) indica que nao tem lista
            If para.Range.ListFormat.ListType <> 0 Then
                para.Range.ListFormat.RemoveNumbers
                removedCount = removedCount + 1
            End If
        End If
    Next i

    If removedCount > 0 Then
        LogMessage "Formatacao de numero removida de " & removedCount & " paragrafo(s) em branco", LOG_LEVEL_INFO
    End If

    RemoveNumberingFromBlankParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover formatacao de numero de paragrafos em branco: " & Err.Description, LOG_LEVEL_WARNING
    RemoveNumberingFromBlankParagraphs = False
End Function


'================================================================================
' REPLACE NON BREAKING SPACES EXCEPT AFTER NO - Substitui espacos nao separaveis por comuns, exceto apos n�/n�
'================================================================================

Public Sub ReplaceNonBreakingSpacesExceptAfterNo(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim nbsp As String: nbsp = Chr(160)
    Dim tempMarker As String: tempMarker = "@@NBSP_MARKER@@"
    Dim i As Long
    Dim digit As String
    Dim noOrdinalLower As String: noOrdinalLower = "n" & Chr(186)
    Dim noOrdinalUpper As String: noOrdinalUpper = "N" & Chr(186)
    Dim noDegreeLower As String: noDegreeLower = "n" & Chr(176)
    Dim noDegreeUpper As String: noDegreeUpper = "N" & Chr(176)
    
    ' 1. Protege os espacos nao separaveis que devem ser mantidos (apos n�/n� e antes de algarismos)
    For i = 0 To 9
        digit = CStr(i)
        ExecuteFindReplace doc, noOrdinalLower & nbsp & digit, noOrdinalLower & tempMarker & digit, True
        ExecuteFindReplace doc, noOrdinalUpper & nbsp & digit, noOrdinalUpper & tempMarker & digit, True
        ExecuteFindReplace doc, noDegreeLower & nbsp & digit, noDegreeLower & tempMarker & digit, True
        ExecuteFindReplace doc, noDegreeUpper & nbsp & digit, noDegreeUpper & tempMarker & digit, True
    Next i
    
    ' 2. Substitui todos os outros espacos nao separaveis por espacos comuns
    ExecuteFindReplace doc, nbsp, " ", False
    
    ' 3. Restaura os espacos nao separaveis protegidos
    For i = 0 To 9
        digit = CStr(i)
        ExecuteFindReplace doc, noOrdinalLower & tempMarker & digit, noOrdinalLower & nbsp & digit, True
        ExecuteFindReplace doc, noOrdinalUpper & tempMarker & digit, noOrdinalUpper & nbsp & digit, True
        ExecuteFindReplace doc, noDegreeLower & tempMarker & digit, noDegreeLower & nbsp & digit, True
        ExecuteFindReplace doc, noDegreeUpper & tempMarker & digit, noDegreeUpper & nbsp & digit, True
    Next i
    
    Exit Sub
ErrorHandler:
    LogMessage "Erro ao substituir espacos nao separaveis: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' ENSURE NON BREAKING SPACE AFTER NO - Garante espaco nao separavel apos n�/n�
'================================================================================

Public Sub EnsureNonBreakingSpaceAfterNo(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim digit As String
    Dim noOrdinalLower As String: noOrdinalLower = "n" & Chr(186)
    Dim noOrdinalUpper As String: noOrdinalUpper = "N" & Chr(186)
    Dim noDegreeLower As String: noDegreeLower = "n" & Chr(176)
    Dim noDegreeUpper As String: noDegreeUpper = "N" & Chr(176)
    Dim nbsp As String: nbsp = Chr(160)
    
    For i = 0 To 9
        digit = CStr(i)
        
        ' 1. n� seguida de algarismo sem espaco -> com espaco nao separavel
        ExecuteFindReplace doc, noOrdinalLower & digit, noOrdinalLower & nbsp & digit, True
        ExecuteFindReplace doc, noOrdinalUpper & digit, noOrdinalUpper & nbsp & digit, True
        ExecuteFindReplace doc, noDegreeLower & digit, noDegreeLower & nbsp & digit, True
        ExecuteFindReplace doc, noDegreeUpper & digit, noDegreeUpper & nbsp & digit, True
        
        ' 2. n� seguida de algarismo com espaco comum -> com espaco nao separavel
        ExecuteFindReplace doc, noOrdinalLower & " " & digit, noOrdinalLower & nbsp & digit, True
        ExecuteFindReplace doc, noOrdinalUpper & " " & digit, noOrdinalUpper & nbsp & digit, True
        ExecuteFindReplace doc, noDegreeLower & " " & digit, noDegreeLower & nbsp & digit, True
        ExecuteFindReplace doc, noDegreeUpper & " " & digit, noDegreeUpper & nbsp & digit, True
    Next i
    
    Exit Sub
ErrorHandler:
    LogMessage "Erro ao garantir espaco nao separavel apos n�: " & Err.Description, LOG_LEVEL_WARNING
End Sub

