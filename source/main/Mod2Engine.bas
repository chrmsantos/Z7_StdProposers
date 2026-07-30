Attribute VB_Name = "Mod2Engine"
Option Explicit

' Mod2Engine.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)
' =============================================================================
'================================================================================
' FUNCOES DE VALIDACAO E COMPATIBILIDADE
'================================================================================
Public Function ValidateDocument(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ValidateDocument = False

    If doc Is Nothing Then
        LogMessage "Documento e Nothing", LOG_LEVEL_ERROR
        Exit Function
    End If

    If doc.Paragraphs.count = 0 Then
        LogMessage "Documento nao tem paragrafos", LOG_LEVEL_WARNING
        Exit Function
    End If

    ValidateDocument = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na validacao do documento: " & Err.Description, LOG_LEVEL_ERROR
    ValidateDocument = False
End Function

'================================================================================
' IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS DA PROPOSITURA
'================================================================================

'--------------------------------------------------------------------------------
' IsTituloElement - Identifica se o paragrafo e o titulo da propositura
'--------------------------------------------------------------------------------
' Criterios:
' - 1 linha contendo texto
' - Negrito, sublinhado, caixa alta
' - Recuo = 0
' - Mais de 15 caracteres
' - Termina com "$NUMERO$/$ANO$"
'--------------------------------------------------------------------------------
Public Function IsTituloElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsTituloElement = False

    ' Validacao de seguranca
    If para Is Nothing Then Exit Function
    If para.Range Is Nothing Then Exit Function

    ' Obtem texto limpo
    Dim paraText As String
    paraText = Trim(para.Range.text)
    ' Aceita comprimentos a partir de 10 caracteres (ex: "MOCAO 1/26")
    If Len(paraText) < 10 Then Exit Function

    ' Limpa quebras de linha e espacos extras
    Dim cleanText As String
    cleanText = Replace(Replace(paraText, vbCr, ""), vbLf, "")
    
    Dim cleanTextNorm As String
    cleanTextNorm = Trim(cleanText)
    If Right(cleanTextNorm, 1) = "." Then cleanTextNorm = Left(cleanTextNorm, Len(cleanTextNorm) - 1)
    cleanTextNorm = Trim(cleanTextNorm)

    ' Verifica se termina com a string de template ou com o padrao de Numero/Ano (ex: 123/2026, 45/26)
    Dim hasValidSuffix As Boolean
    hasValidSuffix = False
    
    If Right(cleanTextNorm, Len(REQUIRED_STRING)) = REQUIRED_STRING Then
        hasValidSuffix = True
    Else
        If cleanTextNorm Like "*#/#" Or cleanTextNorm Like "*#/##" Or cleanTextNorm Like "*#/###" Or cleanTextNorm Like "*#/####" Then
            ' Garante que o caractere anterior e posterior a barra sao numericos
            Dim lastSlash As Long
            lastSlash = InStrRev(cleanTextNorm, "/")
            If lastSlash > 1 And lastSlash < Len(cleanTextNorm) Then
                Dim leftChar As String
                Dim rightChar As String
                leftChar = Mid(cleanTextNorm, lastSlash - 1, 1)
                rightChar = Mid(cleanTextNorm, lastSlash + 1, 1)
                If IsNumeric(leftChar) And IsNumeric(rightChar) Then
                    hasValidSuffix = True
                End If
            End If
        End If
    End If
    
    If Not hasValidSuffix Then Exit Function

    ' Verifica formatacao do paragrafo
    With para.Format
        ' Permite recuo de ate 36 pontos (1.27 cm) e aceita alinhamento a esquerda ou centralizado
        If .leftIndent > 36 Then Exit Function
        If .alignment <> wdAlignParagraphLeft And .alignment <> wdAlignParagraphCenter Then Exit Function
    End With

    ' Verifica formatacao do texto (negrito e se inicia com palavra-chave)
    Dim isBold As Boolean
    isBold = (para.Range.Font.Bold = msoTrue)
    
    Dim normText As String
    normText = NormalizeForComparison(cleanTextNorm)
    
    Dim startsWithKeyword As Boolean
    startsWithKeyword = (Left(normText, 9) = "indicacao" Or _
                         Left(normText, 12) = "requerimento" Or _
                         Left(normText, 5) = "mocao")

    ' Se nao for negrito e nao iniciar com palavra-chave, nao e titulo
    If Not isBold And Not startsWithKeyword Then Exit Function

    IsTituloElement = True
    Exit Function

ErrorHandler:
    IsTituloElement = False
End Function

'--------------------------------------------------------------------------------
' IsEmentaElement - Identifica se o paragrafo e a ementa
'--------------------------------------------------------------------------------
' Criterios:
' - Paragrafo unico imediatamente abaixo do titulo
' - Recuo a esquerda > 6 pontos
' - Contem texto
'--------------------------------------------------------------------------------
Public Function IsEmentaElement(para As Paragraph, prevParaIsTitulo As Boolean) As Boolean
    On Error GoTo ErrorHandler

    IsEmentaElement = False

    ' Validacao de seguranca
    If para Is Nothing Then Exit Function
    If Not prevParaIsTitulo Then Exit Function

    ' Verifica se contem texto
    Dim paraText As String
    paraText = Trim(para.Range.text)
    If Len(paraText) = 0 Then Exit Function

    ' Criterio do recuo a esquerda: deve ser acima de 7 cm (aprox. 198.45 pontos)
    Dim minIndent As Single
    minIndent = Application.CentimetersToPoints(7)
    
    If para.Format.leftIndent > minIndent Then
        IsEmentaElement = True
    End If
    Exit Function

ErrorHandler:
    IsEmentaElement = False
End Function

'--------------------------------------------------------------------------------
' IsVocativoElement - Identifica se o paragrafo e o vocativo
'--------------------------------------------------------------------------------
Public Function IsVocativoElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsVocativoElement = False

    ' Validacao de seguranca
    If para Is Nothing Then Exit Function

    ' Verifica se contem texto
    Dim paraText As String
    paraText = Trim(para.Range.text)
    If Len(paraText) = 0 Then Exit Function

    ' Normaliza para comparacao (remove acentos, caixa baixa, quebras de linha)
    Dim normText As String
    normText = NormalizeForComparison(Trim$(Replace(Replace(paraText, vbCr, ""), vbLf, "")))
    
    ' Remove pontuacao comum no final do vocativo (como virgula ou dois pontos)
    Do While Len(normText) > 0 And InStr(".,;:", Right(normText, 1)) > 0
        normText = Left(normText, Len(normText) - 1)
    Loop
    normText = Trim$(normText)
    normText = Replace(normText, "-", " ")

    ' Limite de tamanho para evitar falsos positivos no corpo do texto (vocativos sao curtos)
    If Len(normText) > 100 Then Exit Function

    ' Vocativos nao devem conter numeros (ex: "Art. 1", "Lei 12.345")
    Dim i As Long
    For i = 1 To Len(normText)
        If Mid(normText, i, 1) Like "[0-9]" Then
            Exit Function
        End If
    Next i

    ' Lista de prefixos/palavras-chave comuns que indicam vocativo
    Dim isKeywordMatch As Boolean
    isKeywordMatch = False

    If Left(normText, 7) = "senhor " Or _
       Left(normText, 7) = "senhora" Or _
       Left(normText, 8) = "senhores" Or _
       Left(normText, 8) = "senhoras" Or _
       Left(normText, 3) = "sr " Or _
       Left(normText, 4) = "sra " Or _
       Left(normText, 4) = "srs " Or _
       Left(normText, 14) = "excelentissimo" Or _
       Left(normText, 14) = "excelentissima" Or _
       Left(normText, 5) = "exmo " Or _
       Left(normText, 5) = "exma " Or _
       Left(normText, 5) = "nobre" Or _
       Left(normText, 6) = "nobres" Or _
       Left(normText, 7) = "ilustre" Or _
       Left(normText, 8) = "ilustres" Or _
       normText = "vereador" Or _
       normText = "vereadora" Or _
       normText = "presidente" Or _
       normText = "prefeito" Or _
       normText = "secretario" Or _
       normText = "mesa diretora" Or _
       InStr(normText, "membros da mesa") > 0 Or _
       InStr(normText, "senhores membros") > 0 Then
        isKeywordMatch = True
    End If

    If isKeywordMatch Then
        ' Garante que nao comeca com termos tipicos de proposicao
        If Not (Left(normText, 8) = "requeiro" Or _
                Left(normText, 6) = "indico" Or _
                Left(normText, 8) = "solicito" Or _
                Left(normText, 12) = "considerando" Or _
                Left(normText, 3) = "art" Or _
                Left(normText, 7) = "decreto" Or _
                Left(normText, 9) = "resolucao") Then
            IsVocativoElement = True
        End If
    End If

    Exit Function

ErrorHandler:
    IsVocativoElement = False
End Function

'--------------------------------------------------------------------------------
' IsExactVocativoString - Verifica se o texto corresponde a um vocativo exato
'--------------------------------------------------------------------------------
Public Function IsExactVocativoString(text As String) As Boolean
    On Error GoTo ErrorHandler
    
    IsExactVocativoString = False
    
    Dim norm As String
    norm = NormalizeForComparison(Trim$(Replace(Replace(text, vbCr, ""), vbLf, "")))
    
    ' Remove pontuacao no final
    Do While Len(norm) > 0 And InStr(".,;:", Right(norm, 1)) > 0
        norm = Left(norm, Len(norm) - 1)
    Loop
    norm = Trim$(norm)
    
    If norm = "senhor presidente" Or _
       norm = "senhora vereadora" Or _
       norm = "senhoras vereadoras" Or _
       norm = "senhores vereadores" Or _
       norm = "senhores vereadores(as)" Or _
       norm = "excelentissimo senhor prefeito municipal" Then
        IsExactVocativoString = True
    End If
    
    Exit Function

ErrorHandler:
    IsExactVocativoString = False
End Function

'--------------------------------------------------------------------------------
' IsJustificativaTitleElement - Identifica o titulo "Justificativa"
'--------------------------------------------------------------------------------
Public Function IsJustificativaTitleElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsJustificativaTitleElement = False

    ' Validacao de seguranca
    If para Is Nothing Then Exit Function

    ' Verifica se o texto contem "justificativa" ou "justificacao"
    Dim cleanText As String
    cleanText = GetCleanParagraphText(para)
    
    ' Evita correspondencia em paragrafos normais longos
    If Len(cleanText) > 40 Then Exit Function

    ' Normaliza para comparacao sem acentos
    Dim normText As String
    normText = NormalizeForComparison(cleanText)

    ' Aceita se contiver "justificativa" ou "justificacao"
    If InStr(normText, "justificativa") > 0 Or InStr(normText, "justificacao") > 0 Then
        IsJustificativaTitleElement = True
    End If
    
    Exit Function

ErrorHandler:
    IsJustificativaTitleElement = False
End Function

'--------------------------------------------------------------------------------
' IsDataElement - Identifica o paragrafo de data (Plenario)
'--------------------------------------------------------------------------------
' Criterios:
' - Contem "Plenario "Dr. Tancredo Neves", $DATAATUALEXTENSO$."
'--------------------------------------------------------------------------------
Public Function IsDataElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsDataElement = False

    ' Validacao de seguranca
    If para Is Nothing Then Exit Function

    ' Normaliza para comparacao (remove acentos) para aceitar "Plenario" e "Plenario" com acento
    Dim paraTextCmp As String
    paraTextCmp = NormalizeForComparison(Trim$(Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")))

    ' Limite de tamanho para evitar falsos positivos no corpo do texto
    If Len(paraTextCmp) = 0 Or Len(paraTextCmp) > 120 Then Exit Function

    ' 1. Placeholder de data
    If InStr(paraTextCmp, "$dataatualextenso$") > 0 Then
        IsDataElement = True
        Exit Function
    End If

    ' 2. Plenario Tancredo Neves (caso original)
    If InStr(paraTextCmp, "plenario") > 0 And InStr(paraTextCmp, "tancredo neves") > 0 Then
        IsDataElement = True
        Exit Function
    End If

    ' 3. Contem padrao de data com um mes valido em portugues
    Dim months As Variant
    months = Array("janeiro", "fevereiro", "marco", "abril", "maio", "junho", _
                   "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")
    
    Dim hasMonthPattern As Boolean
    hasMonthPattern = False
    
    Dim m As Variant
    For Each m In months
        If InStr(paraTextCmp, "de " & m & " de") > 0 Then
            hasMonthPattern = True
            Exit For
        End If
    Next m

    If hasMonthPattern Then
        ' Se tiver o padrao do mes, verificamos se contem alguma pista de local ou se simplemente
        ' comeca com letra maiuscula/numero/placeholder e e um paragrafo curto.
        Dim hasKeyword As Boolean
        hasKeyword = (InStr(paraTextCmp, "plenario") > 0) Or _
                     (InStr(paraTextCmp, "sessoes") > 0) Or _
                     (InStr(paraTextCmp, "sessao") > 0) Or _
                     (InStr(paraTextCmp, "camara") > 0) Or _
                     (InStr(paraTextCmp, "sala") > 0) Or _
                     (InStr(paraTextCmp, "gabinete") > 0) Or _
                     (InStr(paraTextCmp, "prefeitura") > 0) Or _
                     (InStr(paraTextCmp, "dado e passado") > 0) Or _
                     (InStr(paraTextCmp, "palacio") > 0)
        
        If hasKeyword Then
            IsDataElement = True
            Exit Function
        Else
            ' Se nao tem keyword, mas e um paragrafo curto contendo a data, e que comeca com letra maiuscula, numero ou $
            Dim origText As String
            origText = Trim$(para.Range.text)
            If Len(origText) > 0 Then
                Dim firstChar As String
                firstChar = Left$(origText, 1)
                Dim firstCharUpper As String
                firstCharUpper = UCase$(firstChar)
                If (firstCharUpper Like "[A-Z0-9À-ÚÇ]") Or (firstCharUpper = "$") Then
                    IsDataElement = True
                    Exit Function
                End If
            End If
        End If
    End If

    Exit Function

ErrorHandler:
    IsDataElement = False
End Function

'--------------------------------------------------------------------------------
' IsTituloAnexoElement - Identifica o titulo "Anexo" ou "Anexos"
'--------------------------------------------------------------------------------
' Criterios:
' - Paragrafo unicamente com palavra "Anexo" ou "Anexos"
' - Negrito, recuo 0, alinhado a esquerda
'--------------------------------------------------------------------------------
Public Function IsTituloAnexoElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsTituloAnexoElement = False

    ' Validacao de seguranca
    If para Is Nothing Then Exit Function

    ' Verifica texto
    Dim cleanText As String
    cleanText = GetCleanParagraphText(para)
    If cleanText <> ANEXO_TEXT_SINGULAR And cleanText <> ANEXO_TEXT_PLURAL Then Exit Function

    ' Verifica formatacao
    With para.Format
        If .leftIndent <> 0 Then Exit Function
        If .alignment <> wdAlignParagraphLeft Then Exit Function
    End With

    ' Verifica negrito
    If para.Range.Font.Bold <> msoTrue Then Exit Function

    IsTituloAnexoElement = True
    Exit Function

ErrorHandler:
    IsTituloAnexoElement = False
End Function

'--------------------------------------------------------------------------------
' CountBlankLinesBefore - Conta linhas em branco antes de um paragrafo
'--------------------------------------------------------------------------------
Public Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler

    CountBlankLinesBefore = 0

    If paraIndex <= 1 Then Exit Function
    If paraIndex > doc.Paragraphs.count Then Exit Function

    Dim i As Long
    Dim blankCount As Long
    blankCount = 0

    ' Volta ate encontrar paragrafo nao-vazio ou ate 5 linhas
    For i = paraIndex - 1 To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For

        Dim paraText As String
        paraText = Trim(doc.Paragraphs(i).Range.text)

        If Len(paraText) = 0 Then
            blankCount = blankCount + 1
        Else
            Exit For
        End If

        ' Limita a 5 linhas para evitar loops longos
        If blankCount >= 5 Then Exit For
    Next i

    CountBlankLinesBefore = blankCount
    Exit Function

ErrorHandler:
    CountBlankLinesBefore = 0
End Function

'--------------------------------------------------------------------------------
' IsAssinaturaStart - Identifica o inicio da assinatura
'--------------------------------------------------------------------------------
' Criterios:
' - 3 paragrafos textuais
' - 2 linhas em branco antes
' - Centralizados
' - Sem linhas em branco entre si
' - Pode ter imagens logo abaixo (sem linhas em branco)
'--------------------------------------------------------------------------------
Public Function IsAssinaturaStart(doc As Document, paraIndex As Long) As Boolean
    On Error GoTo ErrorHandler

    IsAssinaturaStart = False

    ' Validacao de seguranca
    If paraIndex <= 0 Or paraIndex > doc.Paragraphs.count Then Exit Function

    ' Verifica se ha linhas em branco antes (pelo menos 2)
    If CountBlankLinesBefore(doc, paraIndex) < ASSINATURA_BLANK_LINES_BEFORE Then Exit Function

    ' Verifica se ha 3 paragrafos consecutivos centralizados com texto
    Dim i As Long
    Dim consecutiveCount As Long
    consecutiveCount = 0

    For i = paraIndex To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For

        Dim para As Paragraph
        Set para = doc.Paragraphs(i)

        Dim paraText As String
        paraText = Trim(para.Range.text)

        ' Se encontrou paragrafo vazio, para a contagem
        If Len(paraText) = 0 Then
            Exit For
        End If

        ' Verifica se esta centralizado
        If para.Format.alignment = wdAlignParagraphCenter Then
            consecutiveCount = consecutiveCount + 1
        Else
            Exit For
        End If

        ' Se ja encontrou 3, e uma assinatura
        If consecutiveCount >= ASSINATURA_PARAGRAPH_COUNT Then
            IsAssinaturaStart = True
            Exit Function
        End If

        ' Limite de seguranca
        If i - paraIndex > 10 Then Exit For
    Next i

    Exit Function

ErrorHandler:
    IsAssinaturaStart = False
End Function

'--------------------------------------------------------------------------------
' IdentifyDocumentStructure - Identifica todos os elementos estruturais
'--------------------------------------------------------------------------------
' Esta funcao percorre o documento e identifica:
' - Titulo, Ementa, Proposicao, Justificativa, Data, Assinatura, Anexo
'--------------------------------------------------------------------------------
Public Sub IdentifyDocumentStructure(doc As Document)
    On Error GoTo ErrorHandler

    LogMessage "Identificando estrutura do documento...", LOG_LEVEL_INFO

    ' Reseta todos os indices
    tituloParaIndex = 0
    ementaParaIndex = 0
    vocativoStartIndex = 0
    vocativoEndIndex = 0
    proposicaoStartIndex = 0
    proposicaoEndIndex = 0
    tituloJustificativaIndex = 0
    justificativaStartIndex = 0
    justificativaEndIndex = 0
    dataParaIndex = 0
    assinaturaStartIndex = 0
    assinaturaEndIndex = 0
    tituloAnexoIndex = 0
    anexoStartIndex = 0
    anexoEndIndex = 0

    Dim i As Long
    Dim tempJ As Long
    Dim para As Paragraph
    Dim foundTitulo As Boolean
    Dim foundJustificativa As Boolean
    Dim foundData As Boolean
    Dim customDataParaIndex As Long
    Dim startSearchIdx As Long
    Dim foundExactVocativo As Boolean
    Dim tempVocStart As Long
    Dim tempVocEnd As Long
    Dim isBlankPara As Boolean
    Dim foundCustomAssinatura As Boolean
    Dim p1 As Long
    Dim p2 As Long
    Dim p3 As Long
    Dim textCount As Long
    Dim isTextPara As Boolean
    Dim containsAnexo As Boolean
    Dim nextParaIdx As Long
    Dim anexoCleanText As String

    foundTitulo = False
    foundJustificativa = False
    foundData = False

    ' 1. Identifica Titulo, Ementa, Data e Assinatura por regras posicionais dos parágrafos com elementos visíveis
    Dim visibleIndices() As Long
    Dim visibleCount As Long
    ReDim visibleIndices(1 To cacheSize)
    visibleCount = 0
    
    Dim idx As Long
    For idx = 1 To cacheSize
        If idx > doc.Paragraphs.count Then Exit For
        If Len(paragraphCache(idx).cleanText) > 0 Or paragraphCache(idx).hasImages Then
            visibleCount = visibleCount + 1
            visibleIndices(visibleCount) = idx
        End If
    Next idx

    ' Atribui índices baseados nas posições relativas dos parágrafos visíveis (Título e Ementa apenas, Assinatura e Data serão decididas com base nas novas regras abaixo)
    If visibleCount >= 1 Then
        tituloParaIndex = visibleIndices(1)
        foundTitulo = True
    End If
    If visibleCount >= 2 Then
        ementaParaIndex = visibleIndices(2)
    End If

    ' Regra especifica para Data (Plenario): o antepenultimo ou o anterior ao antepenultimo paragrafo e a data
    ' desde que haja paragrafo(s) em branco abaixo e acima deste.
    customDataParaIndex = 0
    
    If cacheSize >= 3 Then
        ' Verifica o antepenultimo paragrafo (cacheSize - 2)
        If (Len(paragraphCache(cacheSize - 2).cleanText) > 0) And (Not paragraphCache(cacheSize - 2).hasImages) Then
            ' Deve haver paragrafo(s) em branco acima (cacheSize - 3) e abaixo (cacheSize - 1)
            If (Len(paragraphCache(cacheSize - 3).cleanText) = 0) And (Len(paragraphCache(cacheSize - 1).cleanText) = 0) Then
                customDataParaIndex = cacheSize - 2
            End If
        End If
    End If
    
    If customDataParaIndex = 0 And cacheSize >= 4 Then
        ' Verifica o anterior ao antepenultimo paragrafo (cacheSize - 3)
        If (Len(paragraphCache(cacheSize - 3).cleanText) > 0) And (Not paragraphCache(cacheSize - 3).hasImages) Then
            ' Deve haver paragrafo(s) em branco acima (cacheSize - 4) e abaixo (cacheSize - 2)
            If (Len(paragraphCache(cacheSize - 4).cleanText) = 0) And (Len(paragraphCache(cacheSize - 2).cleanText) = 0) Then
                customDataParaIndex = cacheSize - 3
            End If
        End If
    End If

    If customDataParaIndex > 0 Then
        dataParaIndex = customDataParaIndex
        foundData = True
    ElseIf visibleCount >= 5 Then
        dataParaIndex = visibleIndices(visibleCount - 2)
        foundData = True
    End If

    ' Regra especifica para Assinatura:
    ' Os três parágrafos textuais posteriores à data, desde que não contenham a palavra "anexo" ou "anexos" (case insensitive), são a Assinatura.
    foundCustomAssinatura = False
    p1 = 0
    p2 = 0
    p3 = 0
    
    If dataParaIndex > 0 Then
        textCount = 0
        
        For i = dataParaIndex + 1 To cacheSize
            If i > doc.Paragraphs.count Then Exit For
            
            isTextPara = (Len(paragraphCache(i).cleanText) > 0) And (Not paragraphCache(i).hasImages)
            
            If isTextPara Then
                textCount = textCount + 1
                If textCount = 1 Then
                    p1 = i
                ElseIf textCount = 2 Then
                    p2 = i
                ElseIf textCount = 3 Then
                    p3 = i
                    Exit For
                End If
            End If
        Next i
        
        If textCount = 3 Then
            ' Verifica se nenhum deles contem "anexo" ou "anexos"
            containsAnexo = (InStr(NormalizeForComparison(paragraphCache(p1).text), "anexo") > 0) Or _
                            (InStr(NormalizeForComparison(paragraphCache(p2).text), "anexo") > 0) Or _
                            (InStr(NormalizeForComparison(paragraphCache(p3).text), "anexo") > 0)
                            
            If Not containsAnexo Then
                assinaturaStartIndex = p1
                assinaturaEndIndex = p3
                foundCustomAssinatura = True
            End If
        End If
    End If
    
    ' Se nao encontrou pela regra especifica, usa o fallback posicional anterior
    If Not foundCustomAssinatura Then
        If visibleCount >= 3 Then
            assinaturaEndIndex = visibleIndices(visibleCount)
        End If
        If visibleCount >= 4 Then
            assinaturaStartIndex = visibleIndices(visibleCount - 1)
        End If
    End If

    ' Regra especifica para Anexos:
    ' O parágrafo posterior à assinatura que contenha apenas "anexo" ou "anexos" (case insensitive) é o Anexo.
    If assinaturaEndIndex > 0 Then
        nextParaIdx = 0
        
        For i = assinaturaEndIndex + 1 To cacheSize
            If i > doc.Paragraphs.count Then Exit For
            
            If (Len(paragraphCache(i).cleanText) > 0) And (Not paragraphCache(i).hasImages) Then
                nextParaIdx = i
                Exit For
            End If
        Next i
        
        If nextParaIdx > 0 Then
            anexoCleanText = NormalizeForComparison(Trim$(Replace(Replace(paragraphCache(nextParaIdx).text, vbCr, ""), vbLf, "")))
            
            Do While Len(anexoCleanText) > 0 And InStr(".,;:", Right(anexoCleanText, 1)) > 0
                anexoCleanText = Left(anexoCleanText, Len(anexoCleanText) - 1)
            Loop
            anexoCleanText = Trim$(anexoCleanText)
            
            If anexoCleanText = "anexo" Or anexoCleanText = "anexos" Then
                tituloAnexoIndex = nextParaIdx
                anexoStartIndex = nextParaIdx + 1
                anexoEndIndex = cacheSize
            End If
        End If
    End If

    If foundTitulo Then
        ' Encontra o Vocativo (incluindo multiplos vocativos sequenciais ou nao)
            If ementaParaIndex > 0 Then
                startSearchIdx = ementaParaIndex + 1
            Else
                startSearchIdx = tituloParaIndex + 1
            End If
            
            foundExactVocativo = False
            tempVocStart = 0
            tempVocEnd = 0
            
            For i = startSearchIdx To cacheSize
                If i > doc.Paragraphs.count Then Exit For
                Set para = doc.Paragraphs(i)
                
                isBlankPara = (Len(paragraphCache(i).cleanText) = 0) And (Not paragraphCache(i).hasImages)
                
                If Not isBlankPara Then
                    ' Se for outro elemento estrutural conhecido, encerra a busca de vocativos exatos
                    If IsJustificativaTitleElement(para) Or (dataParaIndex > 0 And i = dataParaIndex) Or IsTituloAnexoElement(para) Then
                        Exit For
                    End If
                    
                    ' Se for um vocativo exato da lista
                    If IsExactVocativoString(para.Range.text) Then
                        If tempVocStart = 0 Then
                            tempVocStart = i
                        End If
                        tempVocEnd = i
                        foundExactVocativo = True
                    Else
                        ' Se for qualquer outro paragrafo com texto (por exemplo, inicio da proposicao),
                        ' encerra a busca de vocativos exatos
                        Exit For
                    End If
                End If
            Next i
            
            If foundExactVocativo Then
                vocativoStartIndex = tempVocStart
                vocativoEndIndex = tempVocEnd
            Else
                ' Fallback para deteccao heuristica tradicional
                Dim vocStart As Long
                vocStart = 0
                
                For i = startSearchIdx To cacheSize
                    If i > doc.Paragraphs.count Then Exit For
                    Set para = doc.Paragraphs(i)
                    If Len(Trim(para.Range.text)) > 0 Then
                        vocStart = i
                        Exit For
                    End If
                Next i
                
                If vocStart > 0 Then
                    ' Apenas identifica como vocativo se realmente corresponder aos padroes
                    If IsVocativoElement(doc.Paragraphs(vocStart)) Then
                        vocativoStartIndex = vocStart
                        vocativoEndIndex = vocStart
                        
                        For i = vocStart To cacheSize
                            If i > doc.Paragraphs.count Then Exit For
                            Set para = doc.Paragraphs(i)
                            
                            If Len(Trim(para.Range.text)) = 0 Then
                                Exit For
                            End If
                            
                            ' Evita avancar sobre outros elementos estruturais conhecidos
                            If IsJustificativaTitleElement(para) Or (dataParaIndex > 0 And i = dataParaIndex) Or IsTituloAnexoElement(para) Then
                                Exit For
                            End If
                            
                            ' Para a leitura se a linha nao parecer com vocativo (ex: contem numeros ou termos de proposicao)
                            Dim nextNorm As String
                            nextNorm = NormalizeForComparison(Trim$(Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")))
                            Do While Len(nextNorm) > 0 And InStr(".,;:", Right(nextNorm, 1)) > 0
                                nextNorm = Left(nextNorm, Len(nextNorm) - 1)
                            Loop
                            nextNorm = Trim$(nextNorm)
                            
                            If Len(nextNorm) > 80 Then Exit For
                            
                            ' Se contem digitos, nao e vocativo
                            Dim hasDigits As Boolean
                            hasDigits = False
                            For tempJ = 1 To Len(nextNorm)
                                If Mid(nextNorm, tempJ, 1) Like "[0-9]" Then
                                    hasDigits = True
                                    Exit For
                                End If
                            Next tempJ
                            If hasDigits Then Exit For
                            
                            ' Se parecer com proposicao, para
                            If Left(nextNorm, 8) = "requeiro" Or _
                               Left(nextNorm, 6) = "indico" Or _
                               Left(nextNorm, 8) = "solicito" Or _
                               Left(nextNorm, 12) = "considerando" Or _
                               Left(nextNorm, 3) = "art" Or _
                               Left(nextNorm, 7) = "decreto" Or _
                               Left(nextNorm, 9) = "resolucao" Then
                                Exit For
                            End If
                            
                            vocativoEndIndex = i
                        Next i
                    End If
                End If
            End If
        
        ' Encontra o inicio da Proposicao
        Dim startPropSearch As Long
        If vocativoEndIndex > 0 Then
            startPropSearch = vocativoEndIndex + 1
        ElseIf ementaParaIndex > 0 Then
            startPropSearch = ementaParaIndex + 1
        Else
            startPropSearch = tituloParaIndex + 1
        End If
        
        For i = startPropSearch To cacheSize
            If i > doc.Paragraphs.count Then Exit For
            Set para = doc.Paragraphs(i)
            If Len(Trim(para.Range.text)) > 0 Then
                If Not IsJustificativaTitleElement(para) And (dataParaIndex = 0 Or i <> dataParaIndex) And Not IsTituloAnexoElement(para) Then
                    proposicaoStartIndex = i
                    Exit For
                End If
            End If
        Next i
    End If

    ' Percorre todos os paragrafos para marcar no cache e encontrar outros elementos
    For i = 1 To cacheSize
        If i > doc.Paragraphs.count Then Exit For

        Set para = doc.Paragraphs(i)

        ' Atualiza cache com identificacao
        With paragraphCache(i)
            ' Reseta flags
            .isTitulo = False
            .isEmenta = False
            .isVocativo = False
            .isProposicaoContent = False
            .isTituloJustificativa = False
            .isJustificativaContent = False
            .isData = False
            .isAssinatura = False
            .isTituloAnexo = False
            .isAnexoContent = False

            ' Marca flags que ja identificamos acima
            If i = tituloParaIndex Then
                .isTitulo = True
            ElseIf i = ementaParaIndex Then
                .isEmenta = True
            ElseIf i = dataParaIndex Then
                .isData = True
            ElseIf i = assinaturaStartIndex Or i = assinaturaEndIndex Then
                .isAssinatura = True
            ElseIf i >= vocativoStartIndex And i <= vocativoEndIndex And vocativoStartIndex > 0 Then
                .isVocativo = True
            ElseIf i = tituloAnexoIndex Then
                .isTituloAnexo = True
            
            ' Se nao for nenhum dos acima, e e um paragrafo novo/diferente, tenta identificar os demais elementos
            Else
                ' 3. Identifica TITULO DA JUSTIFICATIVA
                If Not foundJustificativa And IsJustificativaTitleElement(para) Then
                    .isTituloJustificativa = True
                    tituloJustificativaIndex = i
                    foundJustificativa = True
                    ' Proposicao termina antes da Justificativa
                    If proposicaoStartIndex > 0 Then
                        proposicaoEndIndex = i - 1
                    End If
                    justificativaStartIndex = i + 1 ' Justificativa comeca logo apos o titulo
                    LogMessage "Titulo da Justificativa identificado no paragrafo " & i, LOG_LEVEL_INFO

                ' 4. Identifica DATA (Plenario) - já foi identificada, mas em caso de loop definimos justificativaEndIndex se i = dataParaIndex
                ElseIf i = dataParaIndex Then
                    ' Justificativa termina antes da Data
                    If justificativaStartIndex > 0 Then
                        justificativaEndIndex = i - 1
                    End If

                ' 6. Identifica TITULO DO ANEXO
                ElseIf tituloAnexoIndex = 0 And IsTituloAnexoElement(para) Then
                    .isTituloAnexo = True
                    tituloAnexoIndex = i
                    anexoStartIndex = i + 1 ' Anexo comeca logo apos o titulo
                    LogMessage "Titulo do Anexo identificado no paragrafo " & i, LOG_LEVEL_INFO
                End If
            End If

            ' Marca conteudo do ANEXO
            If anexoStartIndex > 0 And i >= anexoStartIndex Then
                .isAnexoContent = True
                anexoEndIndex = i
            End If
        End With

        ' Atualiza progresso a cada 50 paragrafos
        If i Mod 50 = 0 Then
            DoEvents
        End If
    Next i

    ' Se nao encontrou fim da proposicao, define ate antes da justificativa ou data
    If proposicaoStartIndex > 0 And proposicaoEndIndex = 0 Then
        If tituloJustificativaIndex > 0 Then
            proposicaoEndIndex = tituloJustificativaIndex - 1
        ElseIf dataParaIndex > 0 Then
            proposicaoEndIndex = dataParaIndex - 1
        Else
            proposicaoEndIndex = cacheSize
        End If
    End If

    ' Se nao encontrou fim da justificativa, define ate antes da data
    If justificativaStartIndex > 0 And justificativaEndIndex = 0 Then
        If dataParaIndex > 0 Then
            justificativaEndIndex = dataParaIndex - 1
        Else
            justificativaEndIndex = cacheSize
        End If
    End If

    ' Marca conteudo de proposicao e justificativa pos-processo usando os limites corretos
    For i = 1 To cacheSize
        If i > doc.Paragraphs.count Then Exit For
        With paragraphCache(i)
            If proposicaoStartIndex > 0 And proposicaoEndIndex > 0 Then
                If i >= proposicaoStartIndex And i <= proposicaoEndIndex Then
                    If Not .isVocativo Then
                        .isProposicaoContent = True
                    End If
                End If
            End If
            If justificativaStartIndex > 0 And justificativaEndIndex > 0 Then
                If i >= justificativaStartIndex And i <= justificativaEndIndex Then
                    .isJustificativaContent = True
                End If
            End If
        End With
    Next i

    ' Relatorio de identificacao
    LogMessage "=== ESTRUTURA DO DOCUMENTO IDENTIFICADA ===", LOG_LEVEL_INFO
    LogMessage "Titulo: paragrafo " & tituloParaIndex, LOG_LEVEL_INFO
    LogMessage "Ementa: paragrafo " & ementaParaIndex, LOG_LEVEL_INFO
    LogMessage "Vocativo: paragrafos " & vocativoStartIndex & " a " & vocativoEndIndex, LOG_LEVEL_INFO
    LogMessage "Proposicao: paragrafos " & proposicaoStartIndex & " a " & proposicaoEndIndex, LOG_LEVEL_INFO
    LogMessage "Titulo Justificativa: paragrafo " & tituloJustificativaIndex, LOG_LEVEL_INFO
    LogMessage "Justificativa: paragrafos " & justificativaStartIndex & " a " & justificativaEndIndex, LOG_LEVEL_INFO
    LogMessage "Data: paragrafo " & dataParaIndex, LOG_LEVEL_INFO
    LogMessage "Assinatura: paragrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex, LOG_LEVEL_INFO
    LogMessage "Titulo Anexo: paragrafo " & tituloAnexoIndex, LOG_LEVEL_INFO
    LogMessage "Anexo: paragrafos " & anexoStartIndex & " a " & anexoEndIndex, LOG_LEVEL_INFO
    LogMessage "==========================================", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao identificar estrutura do documento: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' CONSTRUCAO DO CACHE DE PARAGRAFOS - Otimizacao principal
'================================================================================
Public Sub BuildParagraphCache(doc As Document)
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    LogMessage "Iniciando construcao do cache de paragrafos...", LOG_LEVEL_INFO

    cacheSize = doc.Paragraphs.count
    ReDim paragraphCache(1 To cacheSize)

    Dim i As Long
    Dim para As Paragraph
    Dim rawText As String

    For i = 1 To cacheSize
        ' DoEvents a cada 20 paragrafos para manter responsividade
        If i Mod 20 = 0 Then DoEvents

        Set para = doc.Paragraphs(i)

        ' Captura o texto bruto uma unica vez
        On Error Resume Next
        rawText = para.Range.text
        On Error GoTo ErrorHandler

        With paragraphCache(i)
            .index = i
            .text = rawText
            .cleanText = NormalizarTexto(rawText)
            .hasImages = HasVisualContent(para)
            .isSpecial = DetectSpecialParagraph(.cleanText, .specialType)
            .needsFormatting = (Len(.cleanText) > 0) And (Not .hasImages)
        End With

        ' Atualiza progresso a cada 100 paragrafos
        If i Mod 100 = 0 Then
            UpdateProgress "Indexando: " & i & "/" & cacheSize, 5 + (i * 5 \ cacheSize)
        End If
    Next i

    cacheEnabled = True

    Dim elapsed As Single
    elapsed = Timer - startTime

    LogMessage "Cache construido: " & cacheSize & " paragrafos em " & Format(elapsed, "0.00") & "s", LOG_LEVEL_INFO

    ' Identifica a estrutura do documento apos construir o cache
    IdentifyDocumentStructure doc

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao construir cache: " & Err.Description, LOG_LEVEL_ERROR
    cacheEnabled = False
End Sub

'================================================================================
' LIMPEZA DO CACHE
'================================================================================
Public Sub ClearParagraphCache()
    On Error Resume Next
    Erase paragraphCache
    cacheSize = 0
    cacheEnabled = False

    ' Limpa tambem os indices de identificacao
    tituloParaIndex = 0
    ementaParaIndex = 0
    vocativoStartIndex = 0
    vocativoEndIndex = 0
    proposicaoStartIndex = 0
    proposicaoEndIndex = 0
    tituloJustificativaIndex = 0
    justificativaStartIndex = 0
    justificativaEndIndex = 0
    dataParaIndex = 0
    assinaturaStartIndex = 0
    assinaturaEndIndex = 0
    tituloAnexoIndex = 0
    anexoStartIndex = 0
    anexoEndIndex = 0
End Sub

'================================================================================
' LOCALIZA O PARAGRAFO DA EMENTA DE FORMA ROBUSTA
'================================================================================
Public Function FindEmentaParagraphIndex(doc As Document) As Long
    On Error Resume Next
    FindEmentaParagraphIndex = 0

    If doc Is Nothing Then Exit Function

    ' Preferencia: usa indice identificado pelo sistema de estrutura (quando disponivel)
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        FindEmentaParagraphIndex = ementaParaIndex
        Exit Function
    End If

    ' Fallback: heuristica no inicio do documento
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String

    For i = 1 To doc.Paragraphs.count
        If i > 35 Then Exit For

        Set para = doc.Paragraphs(i)
        paraText = Trim$(para.Range.text)
        If Len(paraText) <= 1 Then GoTo NextPara

        ' Ementa tipicamente tem recuo a esquerda maior que o minimo
        If para.Format.leftIndent > EMENTA_MIN_LEFT_INDENT Then
            FindEmentaParagraphIndex = i
            Exit Function
        End If

NextPara:
    Next i
End Function

'================================================================================
' VERIFICACAO DE VERSAO DO WORD
'================================================================================
Public Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler

    Dim version As Double
    ' Uso de Val para garantir conversao correta independente do locale
    version = Val(Application.version)

    If version < MIN_SUPPORTED_VERSION Then
        CheckWordVersion = False
        LogMessage "Versao detectada: " & CStr(version) & " - Minima suportada: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
    Else
        CheckWordVersion = True
        LogMessage "Versao do Word compativel: " & CStr(version), LOG_LEVEL_INFO
    End If

    Exit Function

ErrorHandler:
    ' Se nao conseguir detectar a versao, assume incompatibilidade por seguranca
    CheckWordVersion = False
    LogMessage "Erro ao detectar versao do Word: " & Err.Description, LOG_LEVEL_ERROR
End Function

'================================================================================
' FORMATACAO DE FONTE OTIMIZADA COM CACHE
'================================================================================
Public Function ApplyStdFontOptimized(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If Not cacheEnabled Then
        ' Fallback para metodo tradicional se cache nao estiver disponivel
        ApplyStdFontOptimized = ApplyStdFont(doc)
        Exit Function
    End If

    Dim i As Long
    Dim para As Paragraph
    Dim cache As paragraphCache
    Dim formattedCount As Long
    Dim startTime As Double

    startTime = Timer
    formattedCount = 0

    LogMessage "Aplicando fonte padrao (modo otimizado com cache)...", LOG_LEVEL_INFO

    ' Valida cache antes de processar
    If cacheSize < 1 Then
        LogMessage "Cache vazio - usando metodo tradicional", LOG_LEVEL_INFO
        ApplyStdFontOptimized = ApplyStdFont(doc)
        Exit Function
    End If

    ' Valida limites do array
    On Error Resume Next
    Dim cacheUpperBound As Long
    cacheUpperBound = UBound(paragraphCache)
    If Err.Number <> 0 Or cacheUpperBound < 1 Then
        Err.Clear
        On Error GoTo ErrorHandler
        LogMessage "Array de cache invalido - usando metodo tradicional", LOG_LEVEL_WARNING
        ApplyStdFontOptimized = ApplyStdFont(doc)
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Ajusta cacheSize se necessario
    If cacheSize > cacheUpperBound Then
        cacheSize = cacheUpperBound
    End If

    ' SINGLE PASS - Processa todos os paragrafos em uma passagem usando cache
    For i = 1 To cacheSize
        cache = paragraphCache(i)

        ' Pula paragrafos vazios ou com imagens
        If Not cache.needsFormatting Then
            GoTo NextParagraph
        End If

        ' Validacao do indice do paragrafo no documento
        If cache.index < 1 Or cache.index > doc.Paragraphs.count Then
            LogMessage "Erro: Indice de paragrafo invalido (" & cache.index & ")", LOG_LEVEL_WARNING
            GoTo NextParagraph
        End If

        Set para = doc.Paragraphs(cache.index)

        ' Aplica fonte padrao
        On Error Resume Next
        With para.Range.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic

            ' Remove sublinhado exceto para titulo (primeiro paragrafo com texto)
            If i > 3 Then
                .Underline = wdUnderlineNone
            End If

            ' Remove negrito exceto para paragrafos especiais
            If Not cache.isSpecial Or cache.specialType = "vereador" Then
                .Bold = False
            End If
        End With

        If Err.Number = 0 Then
            formattedCount = formattedCount + 1
        Else
            LogMessage "Erro ao formatar paragrafo " & i & ": " & Err.Description, LOG_LEVEL_WARNING
            Err.Clear
        End If
        On Error GoTo ErrorHandler

NextParagraph:
        ' Atualiza progresso a cada 500 paragrafos
        If i Mod 500 = 0 Then
            DoEvents ' Permite cancelamento
        End If
    Next i

    Dim elapsed As Single
    elapsed = Timer - startTime

    ' Marca documento como modificado se houve formatacao
    If formattedCount > 0 Then documentDirty = True

    LogMessage "Fonte padrao aplicada: " & formattedCount & " paragrafos em " & Format(elapsed, "0.00") & "s", LOG_LEVEL_INFO
    ApplyStdFontOptimized = True
    Exit Function

ErrorHandler:
    LogMessage "Erro em ApplyStdFontOptimized: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFontOptimized = False
End Function

'================================================================================
' VALIDACAO DO TIPO DE PROPOSITURA
'================================================================================
' Verifica se a primeira palavra do documento e um tipo valido de propositura
' Tipos validos: indicacao, requerimento, mocao (com tolerancia a erros de grafia)
'================================================================================
Public Function ValidateProposituraType(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ValidateProposituraType = True

    ' Obtem a primeira palavra do documento
    Dim firstWord As String
    firstWord = GetFirstWord(doc)

    If Len(firstWord) = 0 Then
        LogMessage "Documento vazio ou sem texto no inicio", LOG_LEVEL_WARNING
        Exit Function
    End If

    ' Converte para minusculas e remove acentos para comparacao
    Dim normalizedWord As String
    normalizedWord = NormalizeForComparison(firstWord)

    ' Verifica se corresponde a um tipo valido (com tolerancia a erros)
    If IsValidProposituraWord(normalizedWord) Then
        LogMessage "Tipo de propositura identificado: " & firstWord, LOG_LEVEL_INFO
        ValidateProposituraType = True
        Exit Function
    End If

    ' Nao e um tipo reconhecido - pergunta ao usuario
    Dim userResponse As VbMsgBoxResult
    Dim msgTipo As String
    msgTipo = "A primeira palavra do titulo e: """ & firstWord & """" & vbCrLf & vbCrLf & _
              "Nao parece ser uma propositura de Indicacao, Requerimento ou Mocao," & vbCrLf & _
              "ou ha algum erro de grafia na primeira palavra do titulo." & vbCrLf & vbCrLf & _
              "Deseja prosseguir com o processamento mesmo assim?"
    userResponse = MsgBox(msgTipo, vbYesNo + vbQuestion, "Z7_STDPROPOSERS - Tipo de Propositura")

    If userResponse = vbYes Then
        LogMessage "Usuario optou por prosseguir com tipo nao reconhecido: " & firstWord, LOG_LEVEL_WARNING
        ValidateProposituraType = True
    Else
        LogMessage "Usuario cancelou - tipo nao reconhecido: " & firstWord, LOG_LEVEL_INFO
        ValidateProposituraType = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao validar tipo de propositura: " & Err.Description, LOG_LEVEL_WARNING
    ValidateProposituraType = True ' Em caso de erro, permite prosseguir
End Function

'================================================================================
' VERIFICA SE A PALAVRA E UM TIPO VALIDO DE PROPOSITURA
'================================================================================
Public Function IsValidProposituraWord(normalizedWord As String) As Boolean
    IsValidProposituraWord = False

    ' Padroes validos (normalizados, sem acentos)
    ' indicacao, requerimento, mocao

    ' Verifica correspondencia exata primeiro
    If normalizedWord = "indicacao" Or _
       normalizedWord = "requerimento" Or _
       normalizedWord = "mocao" Then
        IsValidProposituraWord = True
        Exit Function
    End If

    ' Verifica com tolerancia a pequenos erros (distancia de Levenshtein <= 2)
    If LevenshteinDistance(normalizedWord, "indicacao") <= 2 Then
        IsValidProposituraWord = True
        Exit Function
    End If

    If LevenshteinDistance(normalizedWord, "requerimento") <= 2 Then
        IsValidProposituraWord = True
        Exit Function
    End If

    If LevenshteinDistance(normalizedWord, "mocao") <= 2 Then
        IsValidProposituraWord = True
        Exit Function
    End If
End Function

'================================================================================
' VALIDACAO DE ESTRUTURA DO DOCUMENTO
'================================================================================
Public Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next

    ' Verificacao basica e rapida
    If doc.Range.End > 0 And doc.Sections.count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' DETECTA PADRAO NUMERICO DE CPF (XXX.XXX.XXX-XX)
'================================================================================
Public Function ContainsCPFPattern(text As String) As Boolean
    On Error Resume Next
    ContainsCPFPattern = False

    Dim i As Long
    Dim segment As String
    Dim digitCount As Long
    Dim hasSeparator As Boolean

    ' Busca sequencia de 11 digitos com separadores tipicos de CPF
    For i = 1 To Len(text) - 13
        segment = Mid(text, i, 14)

        ' Verifica padrao XXX.XXX.XXX-XX
        If Mid(segment, 4, 1) = "." And Mid(segment, 8, 1) = "." And Mid(segment, 12, 1) = "-" Then
            digitCount = CountDigitsInString(segment)
            If digitCount = 11 Then
                ContainsCPFPattern = True
                Exit Function
            End If
        End If
    Next i

    ' Busca sequencia de 11 digitos consecutivos
    digitCount = 0
    For i = 1 To Len(text)
        If Mid(text, i, 1) Like "[0-9]" Then
            digitCount = digitCount + 1
            If digitCount = 11 Then
                ' Verifica se nao e parte de um numero maior
                If i < Len(text) Then
                    If Not Mid(text, i + 1, 1) Like "[0-9]" Then
                        ContainsCPFPattern = True
                        Exit Function
                    End If
                End If
            End If
        Else
            digitCount = 0
        End If
    Next i
End Function

'================================================================================
' DETECTA PADRAO NUMERICO DE RG
'================================================================================
Public Function ContainsRGPattern(text As String) As Boolean
    On Error Resume Next
    ContainsRGPattern = False

    Dim i As Long
    Dim segment As String
    Dim digitCount As Long

    ' RG geralmente tem 7-9 digitos com separadores
    ' Padrao comum: XX.XXX.XXX-X ou similar
    For i = 1 To Len(text) - 11
        segment = Mid(text, i, 12)

        ' Verifica padrao XX.XXX.XXX-X
        If Mid(segment, 3, 1) = "." And Mid(segment, 7, 1) = "." And Mid(segment, 11, 1) = "-" Then
            digitCount = CountDigitsInString(segment)
            If digitCount >= 8 And digitCount <= 10 Then
                ContainsRGPattern = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
'================================================================================
' IS PAGE NUMBER LINE - Verifica se texto termina com padrao de paginacao
'================================================================================
Public Function IsPageNumberLine(text As String) As Boolean
    On Error GoTo ErrorHandler

    IsPageNumberLine = False

    ' Verifica se esta vazio
    If Len(text) < 10 Then Exit Function

    ' Converte para minusculas para comparacao case-insensitive
    Dim lowerText As String
    lowerText = LCase(text)

    ' Verifica se contem o padrao base
    If InStr(lowerText, "$numero$/$ano$/p") = 0 Then Exit Function

    ' Procura pelos padroes possiveis no final
    Dim patterns() As String
    ReDim patterns(0 To 1)
    patterns(0) = "$numero$/$ano$/pagina"
    patterns(1) = "$numero$/$ano$/página"

    Dim pattern As String
    Dim i As Long

    For i = 0 To UBound(patterns)
        pattern = patterns(i)

        ' Verifica se o padrao esta presente
        Dim patternPos As Long
        patternPos = InStr(lowerText, pattern)

        If patternPos > 0 Then
            ' Extrai o texto apos o padrao
            Dim afterPattern As String
            afterPattern = Trim(Mid(text, patternPos + Len(pattern)))

            ' Remove espacos
            afterPattern = Trim(afterPattern)

            ' Verifica se o que sobrou e apenas 1 ou 2 digitos
            If Len(afterPattern) >= 1 And Len(afterPattern) <= 2 Then
                If IsNumeric(afterPattern) Then
                    IsPageNumberLine = True
                    Exit Function
                End If
            End If
        End If
    Next i

    Exit Function

ErrorHandler:
    IsPageNumberLine = False
End Function


'================================================================================
' IS REQUERIMENTO PAGE LINE - Verifica se texto e exclusivamente a linha de pagina
' de um requerimento no formato "REQUERIMENTO n° $NUMERO$/$ANO$/Página X" ou
' "REQUERIMENTO n° $NUMERO$/$ANO$/Página XX" (com tolerâncias de grafia).
'================================================================================
Public Function IsRequerimentoPageLine(text As String) As Boolean
    On Error GoTo ErrorHandler
    IsRequerimentoPageLine = False
    
    Dim lowerText As String
    lowerText = LCase(Trim(text))
    
    ' A string deve ter pelo menos o tamanho de "requerimento $numero$/$ano$/pagina X"
    If Len(lowerText) < 30 Then Exit Function
    
    ' O paragrafo deve comecar com "requerimento"
    If Left(lowerText, 12) <> "requerimento" Then Exit Function
    
    ' Procura pela presenca de "$numero$/$ano$/p"
    Dim pPos As Long
    pPos = InStr(lowerText, "$numero$/$ano$/p")
    If pPos = 0 Then Exit Function
    
    ' Procura por "página" ou "pagina" a partir de pPos
    Dim pagPos As Long
    pagPos = InStr(pPos, lowerText, "página")
    If pagPos = 0 Then
        pagPos = InStr(pPos, lowerText, "pagina")
    End If
    If pagPos = 0 Then Exit Function
    
    ' Extrai o numero da pagina apos o padrao
    Dim pageNumText As String
    pageNumText = Trim(Mid(lowerText, pagPos + 6))
    
    ' Verifica se sobrou apenas 1 ou 2 digitos numericos
    If Len(pageNumText) >= 1 And Len(pageNumText) <= 2 Then
        If IsNumeric(pageNumText) Then
            ' Verifica o trecho intermediario entre "requerimento" e "$numero$/$ano$/"
            Dim midPart As String
            midPart = Trim(Mid(lowerText, 13, pPos - 13))
            
            Dim validMid As Boolean
            validMid = False
            
            ' Variacoes aceitaveis do indicador de numero (removendo todos os espacos)
            Dim collapsedMid As String
            collapsedMid = Replace(midPart, " ", "")
            
            If collapsedMid = "" Or _
               collapsedMid = "n" Or _
               collapsedMid = "n." Or _
               collapsedMid = "no" Or _
               collapsedMid = "no." Or _
               collapsedMid = "n.o" Or _
               collapsedMid = "n°" Or _
               collapsedMid = "nº" Or _
               collapsedMid = "n.º" Or _
               collapsedMid = "n.o." Or _
               collapsedMid = "no.o" Then
                validMid = True
            End If
            
            If validMid Then
                IsRequerimentoPageLine = True
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    IsRequerimentoPageLine = False
End Function


' Mod4Media.bas
'================================================================================
' GERENCIAMENTO DE CAMINHO DA IMAGEM DE CABECALHO
'================================================================================
Public Function GetHeaderImagePath() As String
    On Error GoTo ErrorHandler
    Dim headerImagePath As String

    ' Constroi caminho absoluto para a imagem desejada
    headerImagePath = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH

    ' Verifica se o arquivo existe
    If Dir(headerImagePath) = "" Then
        LogMessage "Imagem de cabecalho nao encontrada em: " & headerImagePath, LOG_LEVEL_WARNING
        GetHeaderImagePath = ""
        Exit Function
    End If

    GetHeaderImagePath = headerImagePath
    Exit Function

ErrorHandler:
    LogMessage "Erro ao localizar imagem de cabecalho: " & Err.Description, LOG_LEVEL_ERROR
    GetHeaderImagePath = ""
End Function

'================================================================================
' INSERCAO DE IMAGEM DE CABECALHO
'================================================================================
Public Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    ' Define o caminho da imagem do cabecalho
    imgFile = GetHeaderImagePath()

    If imgFile = "" Then
        Application.StatusBar = "Aviso: Imagem nao encontrada"
        InsertHeaderstamp = False
        Exit Function
    End If

    ' Size using standard constants
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete

            ' Define fonte padrao para o cabecalho: Arial 12
            With header.Range.Font
                .Name = STANDARD_FONT  ' Arial
                .size = STANDARD_FONT_SIZE  ' 12
            End With

            Set shp = header.Shapes.AddPicture(fileName:=imgFile, LinkToFile:=False, SaveWithDocument:=msoTrue)

            If shp Is Nothing Then
                LogMessage "Failed to insert header image at section " & sectionsProcessed + 1, LOG_LEVEL_WARNING
            Else
                With shp
                    .LockAspectRatio = msoTrue
                    .Width = imgWidth
                    .Height = imgHeight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Left = (doc.PageSetup.PageWidth - .Width) / 2
                    .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                    .ZOrder msoSendToBack
                End With

                imgFound = True
                sectionsProcessed = sectionsProcessed + 1
            End If
        End If
    Next sec

    If imgFound Then
        ' Log detalhado removido para performance
        InsertHeaderstamp = True
    Else
    LogMessage "No header was inserted", LOG_LEVEL_WARNING
        InsertHeaderstamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Error inserting header: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderstamp = False
End Function

'================================================================================
' IMAGE PROTECTION SYSTEM - SISTEMA DE PROTECAO DE IMAGENS
'================================================================================

'================================================================================
' BACKUP DE IMAGENS
'================================================================================
Public Function BackupAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Protegendo imagens..."

    imageCount = 0
    ReDim savedImages(0)

    Dim para As Paragraph
    Dim i As Long
    Dim j As Long
    Dim shape As InlineShape
    Dim tempImageInfo As ImageInfo

    ' Conta todas as imagens primeiro (com DoEvents para responsividade)
    Dim totalImages As Long
    For i = 1 To doc.Paragraphs.count
        If i Mod 30 = 0 Then DoEvents ' Responsividade
        Set para = doc.Paragraphs(i)
        totalImages = totalImages + para.Range.InlineShapes.count
    Next i

    ' Adiciona shapes flutuantes
    totalImages = totalImages + doc.Shapes.count

    ' Redimensiona array se necessario
    If totalImages > 0 Then
        ReDim savedImages(totalImages - 1)

        ' Backup de imagens inline - apenas propriedades criticas
        For i = 1 To doc.Paragraphs.count
            If i Mod 30 = 0 Then DoEvents ' Responsividade
            Set para = doc.Paragraphs(i)

            For j = 1 To para.Range.InlineShapes.count
                Set shape = para.Range.InlineShapes(j)

                ' Salva apenas propriedades essenciais para protecao
                With tempImageInfo
                    .paraIndex = i
                    .ImageIndex = j
                    .ImageType = "Inline"
                    .Position = shape.Range.Start
                    .Width = shape.Width
                    .Height = shape.Height
                    Set .AnchorRange = shape.Range.Duplicate
                    .ImageData = "InlineShape_Protected"
                End With

                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1

                ' Evita overflow
                If imageCount >= UBound(savedImages) + 1 Then Exit For
            Next j

            ' Evita overflow
            If imageCount >= UBound(savedImages) + 1 Then Exit For
        Next i

        ' Backup de shapes flutuantes - apenas propriedades criticas
        Dim floatingShape As shape
        For i = 1 To doc.Shapes.count
            Set floatingShape = doc.Shapes(i)

            If floatingShape.Type = msoPicture Then
                ' Redimensiona array se necessario
                If imageCount >= UBound(savedImages) + 1 Then
                    ReDim Preserve savedImages(imageCount)
                End If

                With tempImageInfo
                    .paraIndex = -1 ' Indica que e flutuante
                    .ImageIndex = i
                    .ImageType = "Floating"
                    .WrapType = floatingShape.WrapFormat.Type
                    .Width = floatingShape.Width
                    .Height = floatingShape.Height
                    .LeftPosition = floatingShape.Left
                    .TopPosition = floatingShape.Top
                    .ImageData = "FloatingShape_Protected"
                End With

                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1
            End If
        Next i
    End If

    LogMessage "Backup de propriedades de imagens concluido: " & imageCount & " imagens catalogadas"
    BackupAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup de propriedades de imagens: " & Err.Description, LOG_LEVEL_WARNING
    BackupAllImages = False
End Function

'================================================================================
' RESTAURACAO DE IMAGENS
'================================================================================
Public Function RestoreAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If imageCount = 0 Then
        RestoreAllImages = True
        Exit Function
    End If

    Application.StatusBar = "Verificando integridade das imagens..."

    Dim i As Long
    Dim verifiedCount As Long
    Dim correctedCount As Long

    For i = 0 To imageCount - 1
        On Error Resume Next

        With savedImages(i)
            If .ImageType = "Inline" Then
                ' Verifica se a imagem inline ainda existe na posicao esperada
                If .paraIndex <= doc.Paragraphs.count Then
                    Dim para As Paragraph
                    Set para = doc.Paragraphs(.paraIndex)

                    ' Se ainda ha imagens inline no paragrafo, considera verificada
                    If para.Range.InlineShapes.count > 0 Then
                        verifiedCount = verifiedCount + 1
                    End If
                End If

            ElseIf .ImageType = "Floating" Then
                ' Verifica e corrige propriedades de shapes flutuantes se ainda existem
                If .ImageIndex <= doc.Shapes.count Then
                    Dim targetShape As shape
                    Set targetShape = doc.Shapes(.ImageIndex)

                    ' Verifica se as propriedades foram alteradas e corrige se necessario
                    Dim needsCorrection As Boolean
                    needsCorrection = False

                    If Abs(targetShape.Width - .Width) > 1 Then needsCorrection = True
                    If Abs(targetShape.Height - .Height) > 1 Then needsCorrection = True
                    If Abs(targetShape.Left - .LeftPosition) > 1 Then needsCorrection = True
                    If Abs(targetShape.Top - .TopPosition) > 1 Then needsCorrection = True

                    If needsCorrection Then
                        ' Restaura propriedades originais
                        With targetShape
                            .Width = savedImages(i).Width
                            .Height = savedImages(i).Height
                            .Left = savedImages(i).LeftPosition
                            .Top = savedImages(i).TopPosition
                            .WrapFormat.Type = savedImages(i).WrapType
                        End With
                        correctedCount = correctedCount + 1
                    End If

                    verifiedCount = verifiedCount + 1
                End If
            End If
        End With

        On Error GoTo ErrorHandler
    Next i

    If correctedCount > 0 Then
        LogMessage "Verificacao de imagens concluida: " & verifiedCount & " verificadas, " & correctedCount & " corrigidas"
    Else
        LogMessage "Verificacao de imagens concluida: " & verifiedCount & " imagens integras"
    End If

    RestoreAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar imagens: " & Err.Description, LOG_LEVEL_WARNING
    RestoreAllImages = False
End Function

'================================================================================
' FORMAT IMAGE PARAGRAPHS INDENTS - Formata recuos de paragrafos com imagens
'================================================================================
Public Function FormatImageParagraphsIndents(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim formattedCount As Long
    formattedCount = 0

    ' Percorre todos os paragrafos
    Dim imgCounter As Long
    imgCounter = 0
    For Each para In doc.Paragraphs
        imgCounter = imgCounter + 1
        If imgCounter Mod 30 = 0 Then DoEvents ' Responsividade

        ' Verifica se o paragrafo contem imagens inline
        If para.Range.InlineShapes.count > 0 Then
            ' Zera o recuo a esquerda e centraliza
            With para.Format
                .leftIndent = 0
                .firstLineIndent = 0
                .alignment = wdAlignParagraphCenter
            End With
            formattedCount = formattedCount + 1
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Recuos de paragrafos com imagens formatados: " & formattedCount & " paragrafos", LOG_LEVEL_INFO
    End If

    FormatImageParagraphsIndents = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de imagens: " & Err.Description, LOG_LEVEL_WARNING
    FormatImageParagraphsIndents = False
End Function

'================================================================================
' BACKUP LIST FORMATS - Salva formatacoes de lista antes do processamento
'================================================================================
Public Function BackupListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim tempListInfo As ListFormatInfo

    listFormatCount = 0
    ReDim savedListFormats(0)

    ' Conta quantos paragrafos tem formatacao de lista (com DoEvents)
    Dim totalLists As Long
    Dim countIter As Long
    totalLists = 0
    countIter = 0
    For Each para In doc.Paragraphs
        countIter = countIter + 1
        If countIter Mod 30 = 0 Then DoEvents ' Responsividade
        If para.Range.ListFormat.ListType <> 0 Then
            totalLists = totalLists + 1
        End If
    Next para

    If totalLists = 0 Then
        LogMessage "Nenhuma lista encontrada no documento", LOG_LEVEL_INFO
        BackupListFormats = True
        Exit Function
    End If

    ' Aloca array com tamanho adequado
    ReDim savedListFormats(totalLists - 1)

    ' Salva informacoes de cada paragrafo com lista (com DoEvents)
    Dim saveIter As Long
    saveIter = 0
    i = 1
    For Each para In doc.Paragraphs
        saveIter = saveIter + 1
        If saveIter Mod 30 = 0 Then DoEvents ' Responsividade

        If para.Range.ListFormat.ListType <> 0 Then
            With tempListInfo
                .paraIndex = i
                .HasList = True
                .ListType = para.Range.ListFormat.ListType

                ' Salva o nivel da lista se aplicavel
                On Error Resume Next
                .ListLevelNumber = para.Range.ListFormat.ListLevelNumber
                If Err.Number <> 0 Then
                    .ListLevelNumber = 1
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

                ' Salva a string da lista (marcador ou numero)
                On Error Resume Next
                .ListString = para.Range.ListFormat.ListString
                If Err.Number <> 0 Then
                    .ListString = ""
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End With

            savedListFormats(listFormatCount) = tempListInfo
            listFormatCount = listFormatCount + 1

            If listFormatCount >= UBound(savedListFormats) + 1 Then Exit For
        End If
        i = i + 1
    Next para

    LogMessage "Formatacoes de lista salvas: " & listFormatCount & " paragrafos com lista", LOG_LEVEL_INFO
    BackupListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao salvar formatacoes de lista: " & Err.Description, LOG_LEVEL_WARNING
    BackupListFormats = False
End Function

'================================================================================
' RESTORE LIST FORMATS - Restaura formatacoes de lista apos o processamento
'================================================================================
Public Function RestoreListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If listFormatCount = 0 Then
        RestoreListFormats = True
        Exit Function
    End If

    Dim i As Long
    Dim restoredCount As Long
    Dim para As Paragraph

    restoredCount = 0

    For i = 0 To listFormatCount - 1
        On Error Resume Next

        With savedListFormats(i)
            If .HasList And .paraIndex <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(.paraIndex)

                ' Remove qualquer formatacao de lista existente primeiro
                para.Range.ListFormat.RemoveNumbers

                ' Aplica a formatacao de lista original
                Select Case .ListType
                    Case 2 ' wdListBullet
                        ' Lista com marcadores
                        para.Range.ListFormat.ApplyBulletDefault

                    Case 3, 4 ' wdListSimpleNumbering, wdListListNumOnly
                        ' Lista numerada simples
                        para.Range.ListFormat.ApplyNumberDefault

                    Case 5 ' wdListMixedNumbering
                        ' Lista com numeracao mista
                        para.Range.ListFormat.ApplyNumberDefault

                    Case 6 ' wdListOutlineNumbering
                        ' Lista com numeracao de topicos
                        para.Range.ListFormat.ApplyOutlineNumberDefault

                    Case Else
                        ' Tenta aplicar formatacao padrao
                        If InStr(.ListString, ".") > 0 Or IsNumeric(Left(.ListString, 1)) Then
                            para.Range.ListFormat.ApplyNumberDefault
                        Else
                            para.Range.ListFormat.ApplyBulletDefault
                        End If
                End Select

                ' Tenta restaurar o nivel da lista
                If .ListLevelNumber > 0 And .ListLevelNumber <= 9 Then
                    para.Range.ListFormat.ListLevelNumber = .ListLevelNumber
                End If

                If Err.Number = 0 Then
                    restoredCount = restoredCount + 1
                Else
                    LogMessage "Aviso: Nao foi possivel restaurar lista no paragrafo " & .paraIndex & ": " & Err.Description, LOG_LEVEL_WARNING
                    Err.Clear
                End If
            End If
        End With

        On Error GoTo ErrorHandler
    Next i

    If restoredCount > 0 Then
        LogMessage "Formatacoes de lista restauradas: " & restoredCount & " paragrafos", LOG_LEVEL_INFO
    End If

    ' Limpa o array
    ReDim savedListFormats(0)
    listFormatCount = 0

    RestoreListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar formatacoes de lista: " & Err.Description, LOG_LEVEL_WARNING
    RestoreListFormats = False
End Function

'================================================================================
' CENTER IMAGE AFTER PLENARIO - Centraliza imagem entre 5a e 7a linha apos Plenario
'================================================================================
Public Function CenterImageAfterPlenario(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim plenarioIndex As Long
    Dim paraText As String
    Dim paraTextCmp As String
    Dim lineCount As Long
    Dim centeredCount As Long

    plenarioIndex = 0
    centeredCount = 0

    ' Localiza o paragrafo "Plenario Dr. Tancredo Neves"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(para.Range.text)
        paraTextCmp = NormalizeForComparison(paraText)

        ' Procura por "Plenario" e "Tancredo Neves" com $DATAATUALEXTENSO$
          If InStr(paraTextCmp, "plenario") > 0 And _
              InStr(paraTextCmp, "tancredo neves") > 0 And _
           InStr(paraText, "$DATAATUALEXTENSO$") > 0 Then
            plenarioIndex = i
            Exit For
        End If
    Next i

    ' Se nao encontrou o paragrafo do Plenario, retorna
    If plenarioIndex = 0 Then
        LogMessage "Paragrafo do Plenario nao encontrado para centralizar imagem", LOG_LEVEL_INFO
        CenterImageAfterPlenario = True
        Exit Function
    End If

    ' Verifica as linhas 5, 6 e 7 apos o Plenario (contando em branco e textuais)
    lineCount = 0
    For i = plenarioIndex + 1 To doc.Paragraphs.count
        lineCount = lineCount + 1

        ' Verifica apenas entre a 5 e 7 linha
        If lineCount >= 5 And lineCount <= 7 Then
            Set para = doc.Paragraphs(i)

            ' Se o paragrafo contem imagem, centraliza
            If para.Range.InlineShapes.count > 0 Then
                para.alignment = wdAlignParagraphCenter
                centeredCount = centeredCount + 1
                LogMessage "Imagem centralizada na linha " & lineCount & " apos Plenario", LOG_LEVEL_INFO
            End If
        End If

        ' Para apos a 7 linha
        If lineCount > 7 Then
            Exit For
        End If
    Next i

    If centeredCount > 0 Then
        LogMessage "Imagens centralizadas apos Plenario: " & centeredCount, LOG_LEVEL_INFO
    End If

    CenterImageAfterPlenario = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao centralizar imagem apos Plenario: " & Err.Description, LOG_LEVEL_WARNING
    CenterImageAfterPlenario = False
End Function


'================================================================================
' LIMPEZA DE PROTECAO DE IMAGENS
'================================================================================
Public Sub CleanupImageProtection()
    On Error Resume Next

    ' Limpa arrays de imagens
    If imageCount > 0 Then
        Dim i As Long
        For i = 0 To imageCount - 1
            Set savedImages(i).AnchorRange = Nothing
        Next i
    End If

    imageCount = 0
    ReDim savedImages(0)


    LogMessage "Variaveis de protecao de imagens limpas"
End Sub

'================================================================================
' BACKUP DE PARAGRAFOS CENTRALIZADOS
' Salva os indices dos paragrafos que estao centralizados antes do processamento
'================================================================================
Public Function BackupCenteredParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    centeredParaCount = 0
    ReDim savedCenteredParas(0)

    Dim para As Paragraph
    Dim i As Long
    Dim totalCentered As Long
    Dim tempInfo As CenteredParaInfo
    Dim builtinStyleId As Long

    ' Primeira passagem: conta quantos paragrafos sao centralizados
    ' Paragrafos com estilo Heading/Titulo sao excluidos: a centralizacao deles vem
    ' do proprio estilo e deve ser corrigida pelo pipeline, nao preservada.
    totalCentered = 0
    i = 0
    For Each para In doc.Paragraphs
        i = i + 1
        If i Mod 50 = 0 Then DoEvents
        If para.Format.alignment = wdAlignParagraphCenter Then
            On Error Resume Next
            builtinStyleId = para.Style.BuiltinStyle
            If Err.Number <> 0 Then builtinStyleId = 0 : Err.Clear
            On Error GoTo ErrorHandler
            ' Pula estilos Titulo 1-9 (wdStyleHeading1=2..wdStyleHeading9=10) e Titulo (wdStyleTitle=63)
            If Not ((builtinStyleId >= wdStyleHeading1 And builtinStyleId <= wdStyleHeading9) _
                    Or builtinStyleId = wdStyleTitle) Then
                totalCentered = totalCentered + 1
            End If
        End If
    Next para

    If totalCentered = 0 Then
        LogMessage "Nenhum paragrafo centralizado encontrado para backup", LOG_LEVEL_INFO
        BackupCenteredParagraphs = True
        Exit Function
    End If

    ReDim savedCenteredParas(totalCentered - 1)

    ' Segunda passagem: salva indices e chave de texto (mesma exclusao de estilos Heading)
    i = 0
    Dim saveIdx As Long
    saveIdx = 0
    For Each para In doc.Paragraphs
        i = i + 1
        If i Mod 50 = 0 Then DoEvents
        If para.Format.alignment = wdAlignParagraphCenter Then
            On Error Resume Next
            builtinStyleId = para.Style.BuiltinStyle
            If Err.Number <> 0 Then builtinStyleId = 0 : Err.Clear
            On Error GoTo ErrorHandler
            If Not ((builtinStyleId >= wdStyleHeading1 And builtinStyleId <= wdStyleHeading9) _
                    Or builtinStyleId = wdStyleTitle) Then
                Dim rawText As String
                rawText = para.Range.text
                rawText = Replace(Replace(rawText, vbCr, ""), vbLf, "")
                If Len(rawText) > 50 Then rawText = Left(rawText, 50)
                tempInfo.paraIndex = i
                tempInfo.originalText = rawText
                savedCenteredParas(saveIdx) = tempInfo
                saveIdx = saveIdx + 1
                centeredParaCount = saveIdx
                If saveIdx >= UBound(savedCenteredParas) + 1 Then Exit For
            End If
        End If
    Next para

    LogMessage "Backup de paragrafos centralizados: " & centeredParaCount & " paragrafos", LOG_LEVEL_INFO
    BackupCenteredParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup de paragrafos centralizados: " & Err.Description, LOG_LEVEL_WARNING
    BackupCenteredParagraphs = False
End Function

'================================================================================
' RESTAURACAO DE PARAGRAFOS CENTRALIZADOS
' Re-centraliza e zera recuo a esquerda dos paragrafos que estavam centralizados
'================================================================================
Public Function RestoreCenteredParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If centeredParaCount = 0 Then
        RestoreCenteredParagraphs = True
        Exit Function
    End If

    Dim i As Long
    Dim restoredCount As Long
    Dim para As Paragraph
    Dim paraTotal As Long
    paraTotal = doc.Paragraphs.count
    restoredCount = 0

    For i = 0 To centeredParaCount - 1
        On Error Resume Next

        Dim targetIdx As Long
        targetIdx = savedCenteredParas(i).paraIndex

        Dim matched As Boolean
        matched = False

        ' Tentativa 1: indice original + verificacao por chave de texto
        If targetIdx >= 1 And targetIdx <= paraTotal Then
            Set para = doc.Paragraphs(targetIdx)
            Dim curText As String
            curText = para.Range.text
            curText = Replace(Replace(curText, vbCr, ""), vbLf, "")
            If Len(curText) > 50 Then curText = Left(curText, 50)
            If curText = savedCenteredParas(i).originalText Then
                matched = True
            End If
        End If

        ' Tentativa 2: busca por chave de texto em janela de +-5 paragrafos
        If Not matched And savedCenteredParas(i).originalText <> "" Then
            Dim searchStart As Long
            Dim searchEnd As Long
            searchStart = targetIdx - 5
            If searchStart < 1 Then searchStart = 1
            searchEnd = targetIdx + 5
            If searchEnd > paraTotal Then searchEnd = paraTotal

            Dim j As Long
            For j = searchStart To searchEnd
                If j >= 1 And j <= paraTotal Then
                    Set para = doc.Paragraphs(j)
                    Dim searchText As String
                    searchText = para.Range.text
                    searchText = Replace(Replace(searchText, vbCr, ""), vbLf, "")
                    If Len(searchText) > 50 Then searchText = Left(searchText, 50)
                    If searchText = savedCenteredParas(i).originalText Then
                        matched = True
                        Exit For
                    End If
                End If
            Next j
        End If

        ' Aplica centralizacao e zeragem de recuo
        If matched Then
            With para.Format
                .alignment = wdAlignParagraphCenter
                .leftIndent = 0
                .firstLineIndent = 0
            End With
            restoredCount = restoredCount + 1
        End If

        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrorHandler
    Next i

    If restoredCount > 0 Then
        LogMessage "Paragrafos centralizados restaurados: " & restoredCount & " de " & centeredParaCount, LOG_LEVEL_INFO
    End If

    RestoreCenteredParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar paragrafos centralizados: " & Err.Description, LOG_LEVEL_WARNING
    RestoreCenteredParagraphs = False
End Function

'================================================================================
' LIMPEZA DO BACKUP DE PARAGRAFOS CENTRALIZADOS
'================================================================================
Public Sub CleanupCenteredParaBackup()
    On Error Resume Next
    centeredParaCount = 0
    ReDim savedCenteredParas(0)
    LogMessage "Variaveis de backup de paragrafos centralizados limpas"
End Sub

