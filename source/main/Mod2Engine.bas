Attribute VB_Name = "Mod2Engine"
Option Explicit

' Mod2Engine.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Versao: 7.7.7-rc2
' Data: 2026-05-19
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
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
' IS DOCUMENT HEALTHY - Validacao profunda da integridade do documento
'================================================================================
Public Function IsDocumentHealthy(doc As Document) As Boolean
    On Error Resume Next

    IsDocumentHealthy = False

    ' Verifica acessibilidade basica
    If doc Is Nothing Then Exit Function
    If doc.Range Is Nothing Then Exit Function
    If doc.Paragraphs.count = 0 Then Exit Function

    ' Verifica se documento esta corrompido
    Dim testAccess As Long
    testAccess = doc.Range.End
    If Err.Number <> 0 Then Exit Function

    ' Testa acesso a paragrafos
    Dim testPara As Paragraph
    Set testPara = doc.Paragraphs(1)
    If Err.Number <> 0 Then Exit Function

    IsDocumentHealthy = True
End Function

'================================================================================
' IS OPERATION TIMEOUT - Verifica timeout de operacoes longas
'================================================================================
Public Function IsOperationTimeout(startTime As Date) As Boolean
    IsOperationTimeout = (DateDiff("s", startTime, Now) > MAX_OPERATION_TIMEOUT_SECONDS)
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
' VERIFICACAO DE VERSAO DO WORD
'================================================================================
Public Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler

    Dim version As Double
    ' Uso de CDbl para garantir conversao correta em todas as versoes
    version = CDbl(Application.version)

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
' VALIDACAO DE CONSISTENCIA EMENTA x PROPOSICAO
' Compara elementos-chave entre ementa e texto da proposicao
' - Enderecos
' - Nomes completos (quando explicitamente citados na ementa)
'================================================================================
Public Function ValidateAddressConsistency(doc As Document) As Boolean
    ' [REMOVIDO A PEDIDO DO USUARIO] 
    ' Verifica consistencia ementa x proposicao desativada para evitar interferencia visual e avisos falsos.
    LogMessage "Validacao ementa x proposicao foi desativada.", LOG_LEVEL_INFO
    ValidateAddressConsistency = True
End Function

'================================================================================
' OBTEM TEXTO DA EMENTA
'================================================================================
Public Function GetEmentaText(doc As Document) As String
    On Error Resume Next
    GetEmentaText = ""

    If doc Is Nothing Then Exit Function

    Dim idx As Long
    idx = FindEmentaParagraphIndex(doc)
    If idx <= 0 Or idx > doc.Paragraphs.count Then Exit Function

    GetEmentaText = Trim$(doc.Paragraphs(idx).Range.text)
End Function

'================================================================================
' OBTEM TEXTO DA PROPOSICAO (CORPO DO DOCUMENTO)
'================================================================================
Public Function GetProposicaoText(doc As Document) As String
    On Error Resume Next
    GetProposicaoText = ""

    If doc Is Nothing Then Exit Function

    Dim ementaIdx As Long
    ementaIdx = FindEmentaParagraphIndex(doc)
    If ementaIdx <= 0 Or ementaIdx >= doc.Paragraphs.count Then Exit Function

    Dim result As String
    result = ""

    Dim collectedParas As Long
    collectedParas = 0

    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim paraNorm As String

    For i = ementaIdx + 1 To doc.Paragraphs.count
        If collectedParas >= 12 Then Exit For
        If Len(result) > 4000 Then Exit For

        Set para = doc.Paragraphs(i)
        paraText = Trim$(para.Range.text)
        If Len(paraText) <= 1 Then GoTo NextPara

        paraNorm = NormalizeForComparison(paraText)

        ' Para ao encontrar "Justificativa" ou assinatura
        If InStr(paraNorm, "justificativa") > 0 Then Exit For
        If InStr(paraNorm, "vereador") > 0 Then Exit For
        If InStr(paraNorm, "vereadora") > 0 Then Exit For

        ' Ignora linha de data do plenario (nao faz parte do corpo)
        If InStr(paraNorm, "plenario") > 0 And InStr(paraNorm, "tancredo") > 0 And InStr(paraNorm, "neves") > 0 Then
            GoTo NextPara
        End If

        result = result & " " & paraText
        collectedParas = collectedParas + 1

NextPara:
    Next i

    GetProposicaoText = Trim$(result)
End Function

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
' VERIFICA CONSISTENCIA DE ENDERECOS
'================================================================================
Public Function CheckAddressConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckAddressConsistency = ""

    Dim ementaFlat As String
    Dim proposicaoFlat As String
    Dim ementaNorm As String
    Dim proposicaoNorm As String

    ementaFlat = CleanTextForComparison(ementaText)
    proposicaoFlat = CleanTextForComparison(proposicaoText)
    ementaNorm = NormalizeForComparison(ementaFlat)
    proposicaoNorm = NormalizeForComparison(proposicaoFlat)

    Dim addrDict As Object
    Set addrDict = ExtractAddressesFromText(ementaFlat, ementaNorm)
    If addrDict Is Nothing Then Exit Function
    If addrDict.count = 0 Then Exit Function

    Dim k As Variant
    For Each k In addrDict.Keys
        Dim addrDisplay As String
        addrDisplay = CStr(addrDict(k))

        If Not CheckAddressInTextAdvanced(CStr(k), proposicaoNorm) Then
            CheckAddressConsistency = "ENDERECO: '" & addrDisplay & "' da ementa nao encontrado no texto."
            Exit Function
        End If
    Next k
End Function

'================================================================================
' EXTRAI PALAVRAS DO ENDERECO
'================================================================================
Public Function ExtractAddressWords(text As String, startPos As Long, keyword As String) As String
    On Error Resume Next
    ExtractAddressWords = ""

    Dim afterKeyword As String
    Dim words() As String
    Dim result As String
    Dim i As Long

    ' Pega texto apos a palavra-chave
    afterKeyword = Mid(text, startPos + Len(keyword), 60)
    afterKeyword = CleanTextForComparison(afterKeyword)

    ' Divide em palavras
    words = Split(afterKeyword, " ")

    result = ""
    For i = 0 To UBound(words)
        If i > 2 Then Exit For ' Maximo 3 palavras

        Dim word As String
        word = Trim(words(i))

        ' Ignora artigos e preposicoes
        If Len(word) > 2 Then
            If result <> "" Then result = result & " "
            result = result & word
        End If
    Next i

    ExtractAddressWords = result
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE NOMES COMPLETOS (EMENTA -> TEXTO)
'================================================================================
Public Function CheckPersonNameConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckPersonNameConsistency = ""

    Dim ementaFlat As String
    Dim proposicaoFlat As String
    Dim ementaNorm As String
    Dim proposicaoNorm As String

    ementaFlat = CleanTextForComparison(ementaText)
    proposicaoFlat = CleanTextForComparison(proposicaoText)
    ementaNorm = NormalizeForComparison(ementaFlat)
    proposicaoNorm = NormalizeForComparison(proposicaoFlat)

    Dim namesDict As Object
    Set namesDict = ExtractPersonNamesFromText(ementaFlat, ementaNorm)
    If namesDict Is Nothing Then Exit Function
    If namesDict.count = 0 Then Exit Function

    Dim k As Variant
    For Each k In namesDict.Keys
        If Not CheckNameInTextAdvanced(CStr(k), proposicaoNorm) Then
            CheckPersonNameConsistency = "NOME: '" & CStr(namesDict(k)) & "' da ementa nao encontrado no texto."
            Exit Function
        End If
    Next k
End Function

'================================================================================
' EXTRAI NOMES COMPLETOS PROVAVEIS A PARTIR DE MARCADORES (SR, SENHOR, NOME:)
' Retorna Dictionary: key=nome normalizado, value=nome para exibicao
'================================================================================
Public Function ExtractPersonNamesFromText(flatText As String, flatNorm As String) As Object
    On Error Resume Next
    Set ExtractPersonNamesFromText = Nothing

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    If Len(flatText) < 10 Then
        Set ExtractPersonNamesFromText = dict
        Exit Function
    End If

    Dim markers As Variant
    markers = Array("sr ", "sr. ", "sra ", "sra. ", "senhor ", "senhora ", "srta ", "srta. ", "nome: ", "nome : ")

    Dim m As Variant
    For Each m In markers
        Dim pos As Long
        pos = 1
        Do
            pos = InStr(pos, flatNorm, CStr(m), vbTextCompare)
            If pos <= 0 Then Exit Do

            Dim afterPos As Long
            afterPos = pos + Len(CStr(m))

            Dim snippet As String
            snippet = Mid$(flatText, afterPos, 140)

            Dim candidate As String
            candidate = ExtractNameFromSnippet(snippet)
            If Len(candidate) > 0 Then
                Dim key As String
                key = NormalizeForComparison(candidate)
                key = Trim$(key)
                If Len(key) > 0 Then
                    If Not dict.Exists(key) Then dict.Add key, candidate
                End If
            End If

            pos = afterPos
        Loop
    Next m

    Set ExtractPersonNamesFromText = dict
End Function

Public Function ExtractNameFromSnippet(snippet As String) As String
    On Error Resume Next
    ExtractNameFromSnippet = ""

    Dim s As String
    s = snippet
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)
    If Len(s) < 3 Then Exit Function

    ' Corta em pontuacao forte
    Dim cutPos As Long
    cutPos = 0
    cutPos = FirstOfAny(s, Array(",", ".", ";", ":"))
    If cutPos > 0 Then s = Left$(s, cutPos - 1)

    s = Trim$(s)
    If Len(s) < 3 Then Exit Function

    Dim words() As String
    words = Split(s, " ")

    Dim nameOut As String
    nameOut = ""

    Dim realWords As Long
    realWords = 0

    Dim i As Long
    For i = 0 To UBound(words)
        If i > 6 Then Exit For

        Dim w As String
        w = Trim$(words(i))
        If w = "" Then GoTo ContinueLoop

        ' Remove pontuacao em volta
        w = StripEdgePunctuation(w)
        If w = "" Then GoTo ContinueLoop

        Dim wNorm As String
        wNorm = NormalizeForComparison(w)

        ' Palavras que indicam fim do nome
        If wNorm = "residente" Or wNorm = "morador" Or wNorm = "portador" Or wNorm = "inscrito" Or wNorm = "cpf" Or wNorm = "rg" Or wNorm = "na" Or wNorm = "no" Or wNorm = "em" Then
            Exit For
        End If

        ' Conectores permitidos
        If wNorm = "da" Or wNorm = "de" Or wNorm = "do" Or wNorm = "das" Or wNorm = "dos" Or wNorm = "e" Then
            If nameOut <> "" Then nameOut = nameOut & " " & wNorm
            GoTo ContinueLoop
        End If

        ' Precisa conter letras
        If Not ContainsLetter(w) Then Exit For

        If nameOut <> "" Then nameOut = nameOut & " "
        nameOut = nameOut & w
        realWords = realWords + 1

ContinueLoop:
    Next i

    ' Nome completo: pelo menos 2 palavras reais
    If realWords >= 2 Then
        ExtractNameFromSnippet = Trim$(nameOut)
    End If
End Function

Public Function CheckNameInTextAdvanced(nameNorm As String, textNorm As String) As Boolean
    On Error Resume Next
    CheckNameInTextAdvanced = False

    If Len(nameNorm) < 5 Then Exit Function
    If Len(textNorm) < 10 Then Exit Function

    Dim words() As String
    words = Split(nameNorm, " ")

    Dim total As Long
    Dim found As Long
    total = 0
    found = 0

    Dim i As Long
    For i = 0 To UBound(words)
        Dim w As String
        w = Trim$(words(i))
        If Len(w) < 3 Then GoTo NextWord
        If w = "da" Or w = "de" Or w = "do" Or w = "das" Or w = "dos" Or w = "e" Then GoTo NextWord

        total = total + 1
        If InStr(1, textNorm, w, vbTextCompare) > 0 Then found = found + 1

NextWord:
    Next i

    If total = 0 Then Exit Function
    ' Exige que a maioria dos componentes (>=80%) apareca no texto
    CheckNameInTextAdvanced = (found / total) >= 0.8
End Function

'================================================================================
' EXTRAI ENDERECOS PROVAVEIS (EMENTA)
' Retorna Dictionary: key=assinatura normalizada (tokens), value=endereco para exibicao
'================================================================================
Public Function ExtractAddressesFromText(flatText As String, flatNorm As String) As Object
    On Error Resume Next
    Set ExtractAddressesFromText = Nothing

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    If Len(flatText) < 10 Then
        Set ExtractAddressesFromText = dict
        Exit Function
    End If

    Dim keywords As Variant
    keywords = Array("rua ", "avenida ", "av. ", "av ", "travessa ", "alameda ", "praca ", "estrada ", "rodovia ")

    Dim kw As Variant
    For Each kw In keywords
        Dim pos As Long
        pos = 1
        Do
            pos = InStr(pos, flatNorm, CStr(kw), vbTextCompare)
            If pos <= 0 Then Exit Do

            Dim snippet As String
            snippet = Mid$(flatText, pos, 120)

            Dim addrDisplay As String
            addrDisplay = ExtractAddressDisplay(snippet)

            Dim addrKey As String
            addrKey = BuildAddressKey(addrDisplay)

            If Len(addrKey) > 0 Then
                If Not dict.Exists(addrKey) Then dict.Add addrKey, addrDisplay
            End If

            pos = pos + Len(CStr(kw))
        Loop
    Next kw

    Set ExtractAddressesFromText = dict
End Function

Public Function ExtractAddressDisplay(snippet As String) As String
    On Error Resume Next
    ExtractAddressDisplay = ""

    Dim s As String
    s = snippet
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)
    If Len(s) < 5 Then Exit Function

    ' Corta em pontuacao forte
    Dim cutPos As Long
    cutPos = FirstOfAny(s, Array(".", ";"))
    If cutPos > 0 Then s = Left$(s, cutPos - 1)

    ' Limita em virgula dupla (comum em listas)
    Dim commaPos As Long
    commaPos = InStr(1, s, ",", vbTextCompare)
    If commaPos > 0 And commaPos < 10 Then
        ' Mantem a primeira virgula (rua x, n 123...) mas corta se a frase virar lista
        Dim comma2 As Long
        comma2 = InStr(commaPos + 1, s, ",", vbTextCompare)
        If comma2 > 0 Then s = Left$(s, comma2 - 1)
    End If

    ExtractAddressDisplay = Trim$(s)
End Function

Public Function BuildAddressKey(addrDisplay As String) As String
    On Error Resume Next
    BuildAddressKey = ""
    If Len(addrDisplay) < 5 Then Exit Function

    Dim norm As String
    norm = NormalizeForComparison(CleanTextForComparison(addrDisplay))

    ' Normaliza abreviacoes comuns
    norm = Replace(norm, "av.", "avenida")
    norm = Replace(norm, "av ", "avenida ")

    Dim parts() As String
    parts = Split(norm, " ")

    Dim key As String
    key = ""

    Dim keepWords As Long
    keepWords = 0
    Dim sawNumber As Boolean
    sawNumber = False

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim w As String
        w = Trim$(parts(i))
        If w = "" Then GoTo NextPart

        If w = "rua" Or w = "avenida" Or w = "travessa" Or w = "alameda" Or w = "praca" Or w = "estrada" Or w = "rodovia" Then GoTo NextPart
        If w = "da" Or w = "de" Or w = "do" Or w = "das" Or w = "dos" Then GoTo NextPart
        If w = "no" Or w = "na" Or w = "em" Or w = "n" Or w = "n." Or w = "numero" Then GoTo NextPart
        If w = "bairro" Or w = "cep" Or w = "km" Then GoTo NextPart

        If IsNumeric(w) Then
            If Not sawNumber Then
                If key <> "" Then key = key & " "
                key = key & w
                sawNumber = True
            End If
            GoTo NextPart
        End If

        If Len(w) >= 3 Then
            If key <> "" Then key = key & " "
            key = key & w
            keepWords = keepWords + 1
        End If

        If keepWords >= 4 Then Exit For

NextPart:
    Next i

    ' Evita chaves muito curtas (alto risco de falso positivo)
    If keepWords >= 2 Or (keepWords >= 1 And sawNumber) Then
        BuildAddressKey = Trim$(key)
    End If
End Function

Public Function CheckAddressInTextAdvanced(addrKey As String, textNorm As String) As Boolean
    On Error Resume Next
    CheckAddressInTextAdvanced = False

    If Len(addrKey) < 5 Then Exit Function
    If Len(textNorm) < 10 Then Exit Function

    Dim parts() As String
    parts = Split(addrKey, " ")

    Dim totalWords As Long
    Dim foundWords As Long
    Dim requiresNumber As Boolean
    Dim foundNumber As Boolean

    totalWords = 0
    foundWords = 0
    requiresNumber = False
    foundNumber = False

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim w As String
        w = Trim$(parts(i))
        If w = "" Then GoTo NextW

        If IsNumeric(w) Then
            requiresNumber = True
            If InStr(1, textNorm, w, vbTextCompare) > 0 Then foundNumber = True
            GoTo NextW
        End If

        If Len(w) < 3 Then GoTo NextW
        totalWords = totalWords + 1
        If InStr(1, textNorm, w, vbTextCompare) > 0 Then foundWords = foundWords + 1

NextW:
    Next i

    If requiresNumber And Not foundNumber Then Exit Function
    If totalWords = 0 Then Exit Function

    ' Exige pelo menos 60% das palavras-chave do endereco
    CheckAddressInTextAdvanced = (foundWords / totalWords) >= 0.6
End Function

Public Function FirstOfAny(text As String, symbols As Variant) As Long
    On Error Resume Next
    FirstOfAny = 0

    Dim i As Long
    Dim best As Long
    best = 0

    For i = LBound(symbols) To UBound(symbols)
        Dim p As Long
        p = InStr(1, text, CStr(symbols(i)), vbTextCompare)
        If p > 0 Then
            If best = 0 Or p < best Then best = p
        End If
    Next i

    FirstOfAny = best
End Function

Public Function StripEdgePunctuation(word As String) As String
    On Error Resume Next
    StripEdgePunctuation = word

    Dim w As String
    w = Trim$(word)
    If w = "" Then Exit Function

    Dim changed As Boolean
    changed = True
    Do While changed
        changed = False

        If Len(w) = 0 Then Exit Do
        Dim firstChar As String
        firstChar = Left$(w, 1)
        If firstChar = Chr(40) Or firstChar = "[" Or firstChar = "{" Or firstChar = Chr(34) Or firstChar = Chr(39) Then
            w = Mid$(w, 2)
            w = Trim$(w)
            changed = True
        End If

        If Len(w) = 0 Then Exit Do
        Dim lastChar As String
        lastChar = Right$(w, 1)
        If lastChar = Chr(41) Or lastChar = "]" Or lastChar = "}" Or lastChar = "," Or lastChar = "." Or lastChar = ";" Or lastChar = ":" Or lastChar = Chr(34) Or lastChar = Chr(39) Then
            w = Left$(w, Len(w) - 1)
            w = Trim$(w)
            changed = True
        End If
    Loop

    StripEdgePunctuation = w
End Function

Public Function ContainsLetter(text As String) As Boolean
    On Error Resume Next
    ContainsLetter = False

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            ContainsLetter = True
            Exit Function
        End If
    Next i
End Function

'================================================================================
' VERIFICA SE ENDERECO EXISTE NO TEXTO
'================================================================================
Public Function CheckAddressInText(address As String, text As String) As Boolean
    On Error Resume Next
    CheckAddressInText = False

    Dim normalizedAddress As String
    Dim normalizedText As String
    Dim words() As String
    Dim word As Variant
    Dim foundCount As Long
    Dim totalWords As Long

    normalizedAddress = CleanTextForComparison(address)
    normalizedText = CleanTextForComparison(text)

    words = Split(normalizedAddress, " ")
    foundCount = 0
    totalWords = 0

    For Each word In words
        If Len(Trim(CStr(word))) > 2 Then
            totalWords = totalWords + 1
            If InStr(1, normalizedText, CStr(word), vbTextCompare) > 0 Then
                foundCount = foundCount + 1
            End If
        End If
    Next word

    ' Considera consistente se encontrou pelo menos 70% das palavras
    If totalWords > 0 Then
        CheckAddressInText = (foundCount / totalWords) >= 0.7
    End If
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE VALORES MONETARIOS
'================================================================================
Public Function CheckMonetaryConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckMonetaryConsistency = ""

    Dim rsPos As Long
    Dim valueInEmenta As String
    Dim normalizedProposicao As String

    ' Procura por R$ na ementa
    rsPos = InStr(1, ementaText, "R$", vbTextCompare)

    If rsPos > 0 Then
        ' Extrai valor (R$ seguido de numeros)
        valueInEmenta = ExtractMonetaryValue(ementaText, rsPos)

        If Len(valueInEmenta) > 0 Then
            ' Normaliza proposicao para comparacao
            normalizedProposicao = CleanTextForComparison(proposicaoText)
            normalizedProposicao = Replace(normalizedProposicao, ".", "")
            normalizedProposicao = Replace(normalizedProposicao, ",", "")

            ' Remove pontuacao do valor para comparacao
            Dim normalizedValue As String
            normalizedValue = Replace(valueInEmenta, ".", "")
            normalizedValue = Replace(normalizedValue, ",", "")
            normalizedValue = Replace(normalizedValue, " ", "")

            ' Verifica se valor numerico existe na proposicao
            If InStr(1, normalizedProposicao, normalizedValue, vbTextCompare) = 0 Then
                CheckMonetaryConsistency = "VALOR: 'R$ " & valueInEmenta & "' da ementa nao encontrado no texto."
            End If
        End If
    End If
End Function

'================================================================================
' EXTRAI VALOR MONETARIO
'================================================================================
Public Function ExtractMonetaryValue(text As String, rsPos As Long) As String
    On Error Resume Next
    ExtractMonetaryValue = ""

    Dim afterRS As String
    Dim i As Long
    Dim c As String
    Dim result As String
    Dim foundDigit As Boolean

    afterRS = Mid(text, rsPos + 2, 30) ' Pega ate 30 caracteres apos R$
    afterRS = Trim(afterRS)

    result = ""
    foundDigit = False

    For i = 1 To Len(afterRS)
        c = Mid(afterRS, i, 1)

        ' Aceita digitos, ponto, virgula e espaco
        If c Like "[0-9]" Then
            result = result & c
            foundDigit = True
        ElseIf (c = "." Or c = "," Or c = " ") And foundDigit Then
            result = result & c
        ElseIf foundDigit Then
            Exit For ' Terminou o numero
        End If
    Next i

    ExtractMonetaryValue = Trim(result)
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE NUMEROS DE REFERENCIA
'================================================================================
Public Function CheckNumberConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckNumberConsistency = ""

    Dim numberPatterns() As Variant
    numberPatterns = Array("n. ", "n.o ", "no ", "numero ")

    Dim pattern As Variant
    Dim patternPos As Long
    Dim numberInEmenta As String
    Dim normalizedProposicao As String

    normalizedProposicao = CleanTextForComparison(proposicaoText)

    For Each pattern In numberPatterns
        patternPos = InStr(1, LCase(ementaText), CStr(pattern), vbTextCompare)

        If patternPos > 0 Then
            ' Extrai numero apos o padrao
            numberInEmenta = ExtractReferenceNumber(ementaText, patternPos + Len(pattern))

            If Len(numberInEmenta) > 0 Then
                ' Verifica se numero existe na proposicao
                If InStr(1, normalizedProposicao, numberInEmenta, vbTextCompare) = 0 Then
                    CheckNumberConsistency = "NUMERO: '" & numberInEmenta & "' da ementa nao encontrado no texto."
                    Exit Function
                End If
            End If
        End If
    Next pattern
End Function

'================================================================================
' EXTRAI NUMERO DE REFERENCIA
'================================================================================
Public Function ExtractReferenceNumber(text As String, startPos As Long) As String
    On Error Resume Next
    ExtractReferenceNumber = ""

    Dim afterPattern As String
    Dim i As Long
    Dim c As String
    Dim result As String

    afterPattern = Mid(text, startPos, 20)
    afterPattern = Trim(afterPattern)

    result = ""

    For i = 1 To Len(afterPattern)
        c = Mid(afterPattern, i, 1)

        If c Like "[0-9]" Then
            result = result & c
        ElseIf c = "." Or c = "/" Or c = "-" Then
            ' Aceita separadores comuns em numeros de referencia
            If Len(result) > 0 Then result = result & c
        ElseIf Len(result) > 0 Then
            Exit For ' Terminou o numero
        End If
    Next i

    ' Remove separadores no final
    Do While Right(result, 1) = "." Or Right(result, 1) = "/" Or Right(result, 1) = "-"
        result = Left(result, Len(result) - 1)
    Loop

    ExtractReferenceNumber = result
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE BAIRROS
'================================================================================
Public Function CheckBairroConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckBairroConsistency = ""

    Dim bairroPatterns() As Variant
    bairroPatterns = Array("bairro ", "no bairro ", "do bairro ")

    Dim pattern As Variant
    Dim patternPos As Long
    Dim bairroInEmenta As String
    Dim normalizedProposicao As String

    normalizedProposicao = CleanTextForComparison(proposicaoText)

    For Each pattern In bairroPatterns
        patternPos = InStr(1, LCase(ementaText), CStr(pattern), vbTextCompare)

        If patternPos > 0 Then
            ' Extrai nome do bairro (ate 30 caracteres)
            bairroInEmenta = ExtractBairroName(ementaText, patternPos + Len(pattern))

            If Len(bairroInEmenta) > 2 Then
                ' Verifica se bairro existe na proposicao (com tolerancia)
                If Not CheckBairroInText(bairroInEmenta, normalizedProposicao) Then
                    CheckBairroConsistency = "BAIRRO: '" & bairroInEmenta & "' da ementa nao encontrado no texto."
                    Exit Function
                End If
            End If
        End If
    Next pattern
End Function

'================================================================================
' EXTRAI NOME DO BAIRRO
'================================================================================
Public Function ExtractBairroName(text As String, startPos As Long) As String
    On Error Resume Next
    ExtractBairroName = ""

    Dim afterPattern As String
    Dim words() As String
    Dim result As String
    Dim i As Long

    afterPattern = Mid(text, startPos, 40)
    afterPattern = CleanTextForComparison(afterPattern)

    words = Split(afterPattern, " ")
    result = ""

    For i = 0 To UBound(words)
        If i > 2 Then Exit For ' Maximo 3 palavras

        Dim word As String
        word = Trim(words(i))

        ' Para se encontrar pontuacao ou palavras-chave que indicam fim
        If InStr(word, ",") > 0 Or InStr(word, ".") > 0 Then Exit For
        If LCase(word) = "neste" Or LCase(word) = "desta" Then Exit For

        If Len(word) > 1 Then
            If result <> "" Then result = result & " "
            result = result & word
        End If
    Next i

    ExtractBairroName = result
End Function

'================================================================================
' VERIFICA SE BAIRRO EXISTE NO TEXTO
'================================================================================
Public Function CheckBairroInText(bairro As String, text As String) As Boolean
    On Error Resume Next
    CheckBairroInText = False

    ' Busca exata primeiro
    If InStr(1, text, bairro, vbTextCompare) > 0 Then
        CheckBairroInText = True
        Exit Function
    End If

    ' Busca por palavras individuais
    Dim words() As String
    Dim word As Variant
    Dim foundCount As Long

    words = Split(bairro, " ")
    foundCount = 0

    For Each word In words
        If Len(Trim(CStr(word))) > 2 Then
            If InStr(1, text, CStr(word), vbTextCompare) > 0 Then
                foundCount = foundCount + 1
            End If
        End If
    Next word

    ' Considera encontrado se achou pelo menos metade das palavras
    CheckBairroInText = (foundCount >= (UBound(words) + 1) / 2)
End Function

'================================================================================
' VERIFICA DADOS SENSIVEIS ESPECIAIS (Art. 5, II LGPD)
' Origem racial/etnica, conviccao religiosa, opiniao politica, filiacao sindical,
' dados de saude, vida sexual, dados geneticos ou biometricos
'================================================================================
Public Function CheckSensitiveSpecialData(docText As String) As String
    On Error Resume Next
    CheckSensitiveSpecialData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Origem racial ou etnica
    If InStr(lowerText, "raca:") > 0 Or InStr(lowerText, "etnia:") > 0 Or _
       InStr(lowerText, "cor da pele") > 0 Or InStr(lowerText, "origem etnica") > 0 Or _
       InStr(lowerText, "afrodescendente") > 0 Or InStr(lowerText, "indigena") > 0 Then
        findings = findings & "  - Origem racial/etnica detectada" & vbCrLf
    End If

    ' Conviccao religiosa
    If InStr(lowerText, "religiao:") > 0 Or InStr(lowerText, "crenca:") > 0 Or _
       InStr(lowerText, "conviccao religiosa") > 0 Or InStr(lowerText, "fe:") > 0 Or _
       InStr(lowerText, "praticante de") > 0 Then
        findings = findings & "  - Conviccao religiosa detectada" & vbCrLf
    End If

    ' Opiniao politica
    If InStr(lowerText, "opiniao politica") > 0 Or InStr(lowerText, "filiacao partidaria") > 0 Or _
       InStr(lowerText, "partido politico:") > 0 Or InStr(lowerText, "ideologia:") > 0 Then
        findings = findings & "  - Opiniao politica detectada" & vbCrLf
    End If

    ' Filiacao sindical
    If InStr(lowerText, "sindicato:") > 0 Or InStr(lowerText, "filiacao sindical") > 0 Or _
       InStr(lowerText, "sindicalizado") > 0 Or InStr(lowerText, "membro do sindicato") > 0 Then
        findings = findings & "  - Filiacao sindical detectada" & vbCrLf
    End If

    ' Vida sexual
    If InStr(lowerText, "orientacao sexual") > 0 Or InStr(lowerText, "identidade de genero") > 0 Or _
       InStr(lowerText, "vida sexual") > 0 Or InStr(lowerText, "preferencia sexual") > 0 Then
        findings = findings & "  - Dado sobre vida sexual detectado" & vbCrLf
    End If

    ' Dados geneticos
    If InStr(lowerText, "dna") > 0 Or InStr(lowerText, "genetico") > 0 Or _
       InStr(lowerText, "exame genetico") > 0 Or InStr(lowerText, "teste de paternidade") > 0 Then
        findings = findings & "  - Dado genetico detectado" & vbCrLf
    End If

    ' Dados biometricos
    If InStr(lowerText, "biometria") > 0 Or InStr(lowerText, "biometrico") > 0 Or _
       InStr(lowerText, "impressao digital") > 0 Or InStr(lowerText, "reconhecimento facial") > 0 Or _
       InStr(lowerText, "iris") > 0 Then
        findings = findings & "  - Dado biometrico detectado" & vbCrLf
    End If

    CheckSensitiveSpecialData = findings
End Function

'================================================================================
' VERIFICA DADOS DE MENORES DE IDADE (Art. 14 LGPD)
'================================================================================
Public Function CheckMinorData(docText As String) As String
    On Error Resume Next
    CheckMinorData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Mencoes a menores
    If InStr(lowerText, "menor de idade") > 0 Or InStr(lowerText, "crianca") > 0 Or _
       InStr(lowerText, "adolescente") > 0 Then
        findings = findings & "  - Referencia a menor de idade detectada" & vbCrLf
    End If

    ' Dados escolares de menores
    If (InStr(lowerText, "aluno") > 0 Or InStr(lowerText, "estudante") > 0) And _
       (InStr(lowerText, "escola") > 0 Or InStr(lowerText, "colegio") > 0) Then
        If InStr(lowerText, "fundamental") > 0 Or InStr(lowerText, "infantil") > 0 Then
            findings = findings & "  - Dados escolares de menor detectados" & vbCrLf
        End If
    End If

    ' Responsavel legal
    If InStr(lowerText, "responsavel legal") > 0 Or InStr(lowerText, "representante legal") > 0 Or _
       InStr(lowerText, "tutor:") > 0 Or InStr(lowerText, "curador:") > 0 Then
        findings = findings & "  - Mencao a responsavel legal (possivel menor)" & vbCrLf
    End If

    ' ECA - Estatuto da Crianca e Adolescente
    If InStr(lowerText, "eca") > 0 Or InStr(lowerText, "estatuto da crianca") > 0 Or _
       InStr(lowerText, "conselho tutelar") > 0 Then
        findings = findings & "  - Referencia ao ECA detectada" & vbCrLf
    End If

    CheckMinorData = findings
End Function

'================================================================================
' VERIFICA DADOS JUDICIAIS E CRIMINAIS
'================================================================================
Public Function CheckJudicialData(docText As String) As String
    On Error Resume Next
    CheckJudicialData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Antecedentes criminais
    If InStr(lowerText, "antecedentes criminais") > 0 Or InStr(lowerText, "folha corrida") > 0 Or _
       InStr(lowerText, "certidao criminal") > 0 Then
        findings = findings & "  - Antecedentes criminais detectados" & vbCrLf
    End If

    ' Processos judiciais
     If InStr(lowerText, "processo n") > 0 And _
         (InStr(lowerText, "vara") > 0 Or InStr(lowerText, "tribunal") > 0 Or InStr(lowerText, "juizo") > 0) Then
        findings = findings & "  - Numero de processo judicial detectado" & vbCrLf
    End If

    ' Inquerito policial
    If InStr(lowerText, "inquerito policial") > 0 Or InStr(lowerText, "boletim de ocorrencia") > 0 Or _
       InStr(lowerText, "b.o.") > 0 Then
        findings = findings & "  - Inquerito/BO detectado" & vbCrLf
    End If

    ' Condenacao
    If InStr(lowerText, "condenado") > 0 Or InStr(lowerText, "sentenciado") > 0 Or _
       InStr(lowerText, "apenado") > 0 Or InStr(lowerText, "reeducando") > 0 Then
        findings = findings & "  - Informacao de condenacao detectada" & vbCrLf
    End If

    ' Medida protetiva
    If InStr(lowerText, "medida protetiva") > 0 Or InStr(lowerText, "lei maria da penha") > 0 Then
        findings = findings & "  - Medida protetiva detectada" & vbCrLf
    End If

    CheckJudicialData = findings
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
' VERIFICA DADOS PESSOAIS
'================================================================================
Public Function CheckPersonalData(docText As String) As String
    On Error Resume Next
    CheckPersonalData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Filiacao
    If InStr(lowerText, "nome da mae") > 0 Or InStr(lowerText, "mae:") > 0 Or _
       InStr(lowerText, "filiacao:") > 0 Or InStr(lowerText, "filho de") > 0 Or _
       InStr(lowerText, "filha de") > 0 Then
        findings = findings & "  - Filiacao detectada" & vbCrLf
    End If

    ' Data de nascimento
    If InStr(lowerText, "nascimento:") > 0 Or InStr(lowerText, "nascido em") > 0 Or _
       InStr(lowerText, "nascida em") > 0 Or InStr(lowerText, "data de nascimento") > 0 Then
        findings = findings & "  - Data de nascimento detectada" & vbCrLf
    End If

    ' Naturalidade
    If InStr(lowerText, "naturalidade:") > 0 Or InStr(lowerText, "natural de") > 0 Then
        findings = findings & "  - Naturalidade detectada" & vbCrLf
    End If

    ' Estado civil
    If InStr(lowerText, "estado civil:") > 0 Then
        findings = findings & "  - Estado civil detectado" & vbCrLf
    End If

    ' Nacionalidade
    If InStr(lowerText, "nacionalidade:") > 0 Then
        findings = findings & "  - Nacionalidade detectada" & vbCrLf
    End If

    ' Profissao/Ocupacao
    If InStr(lowerText, "profissao:") > 0 Or InStr(lowerText, "ocupacao:") > 0 Then
        findings = findings & "  - Profissao/Ocupacao detectada" & vbCrLf
    End If

    ' Endereco residencial
     If InStr(lowerText, "residente") > 0 And _
         (InStr(lowerText, "rua ") > 0 Or InStr(lowerText, "avenida ") > 0) Then
        findings = findings & "  - Endereco residencial detectado" & vbCrLf
    End If

    ' Sexo/Genero
    If InStr(lowerText, "sexo:") > 0 Or InStr(lowerText, "genero:") > 0 Then
        findings = findings & "  - Sexo/Genero detectado" & vbCrLf
    End If

    ' Escolaridade
    If InStr(lowerText, "escolaridade:") > 0 Or InStr(lowerText, "grau de instrucao") > 0 Then
        findings = findings & "  - Escolaridade detectada" & vbCrLf
    End If

    CheckPersonalData = findings
End Function

'================================================================================
' VERIFICA DADOS DE CONTATO
'================================================================================
Public Function CheckContactData(docText As String) As String
    On Error Resume Next
    CheckContactData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Email
    If ContainsEmailPattern(docText) Then
        findings = findings & "  - Email detectado" & vbCrLf
    End If

    ' Telefone
    If InStr(lowerText, "telefone:") > 0 Or InStr(lowerText, "tel:") > 0 Or _
       InStr(lowerText, "celular:") > 0 Or InStr(lowerText, "fone:") > 0 Or _
       ContainsPhonePattern(docText) Then
        findings = findings & "  - Telefone detectado" & vbCrLf
    End If

    ' WhatsApp
    If InStr(lowerText, "whatsapp") > 0 Or InStr(lowerText, "zap:") > 0 Then
        findings = findings & "  - WhatsApp detectado" & vbCrLf
    End If

    CheckContactData = findings
End Function

'================================================================================
' DETECTA PADRAO DE EMAIL
'================================================================================
Public Function ContainsEmailPattern(text As String) As Boolean
    On Error Resume Next
    ContainsEmailPattern = False

    ' Busca por @ seguido de dominio
    Dim atPos As Long
    atPos = InStr(text, "@")

    If atPos > 1 Then
        ' Verifica se tem caracteres antes e depois do @
        Dim beforeAt As String
        Dim afterAt As String

        beforeAt = Mid(text, atPos - 1, 1)
        If atPos < Len(text) - 3 Then
            afterAt = Mid(text, atPos + 1, 4)
            ' Verifica se parece um dominio (letras seguidas de ponto)
            If InStr(afterAt, ".") > 0 Then
                ContainsEmailPattern = True
            End If
        End If
    End If
End Function

'================================================================================
' DETECTA PADRAO DE TELEFONE
'================================================================================
Public Function ContainsPhonePattern(text As String) As Boolean
    On Error Resume Next
    ContainsPhonePattern = False

    Dim i As Long
    Dim segment As String
    Dim digitCount As Long

    ' Busca padrao (XX) XXXXX-XXXX ou similar
    For i = 1 To Len(text) - 13
        segment = Mid(text, i, 15)

        ' Verifica se comeca com parenteses
        If Mid(segment, 1, 1) = Chr(40) Then
            digitCount = CountDigitsInString(segment)
            ' Telefone brasileiro tem 10-11 digitos
            If digitCount >= 10 And digitCount <= 11 Then
                ContainsPhonePattern = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
' VERIFICA DADOS DE VEICULOS
'================================================================================
Public Function CheckVehicleData(docText As String) As String
    On Error Resume Next
    CheckVehicleData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Placa
    If InStr(lowerText, "placa:") > 0 Or InStr(lowerText, "placa n") > 0 Or _
       ContainsPlacaPattern(docText) Then
        findings = findings & "  - Placa de veiculo detectada" & vbCrLf
    End If

    ' Renavam
    If InStr(lowerText, "renavam") > 0 Then
        findings = findings & "  - RENAVAM detectado" & vbCrLf
    End If

    ' Chassi
    If InStr(lowerText, "chassi") > 0 Then
        findings = findings & "  - Chassi detectado" & vbCrLf
    End If

    CheckVehicleData = findings
End Function

'================================================================================
' DETECTA PADRAO DE PLACA (ABC-1234 ou ABC1D23)
'================================================================================
Public Function ContainsPlacaPattern(text As String) As Boolean
    On Error Resume Next
    ContainsPlacaPattern = False

    Dim i As Long
    Dim segment As String
    Dim c As String

    ' Busca padrao antigo: ABC-1234 ou ABC1234
    For i = 1 To Len(text) - 6
        segment = UCase(Mid(text, i, 8))

        ' Verifica 3 letras + hifen ou digito + 4 digitos
        If Mid(segment, 1, 1) Like "[A-Z]" And _
           Mid(segment, 2, 1) Like "[A-Z]" And _
           Mid(segment, 3, 1) Like "[A-Z]" Then

            ' Padrao com hifen: ABC-1234
            If Mid(segment, 4, 1) = "-" Then
                If Mid(segment, 5, 1) Like "[0-9]" And _
                   Mid(segment, 6, 1) Like "[0-9]" And _
                   Mid(segment, 7, 1) Like "[0-9]" And _
                   Mid(segment, 8, 1) Like "[0-9]" Then
                    ContainsPlacaPattern = True
                    Exit Function
                End If
            End If

            ' Padrao Mercosul: ABC1D23
            If Mid(segment, 4, 1) Like "[0-9]" And _
               Mid(segment, 5, 1) Like "[A-Z]" And _
               Mid(segment, 6, 1) Like "[0-9]" And _
               Mid(segment, 7, 1) Like "[0-9]" Then
                ContainsPlacaPattern = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
' VERIFICA DADOS FINANCEIROS
'================================================================================
Public Function CheckFinancialData(docText As String) As String
    On Error Resume Next
    CheckFinancialData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Conta bancaria
    If InStr(lowerText, "conta:") > 0 Or InStr(lowerText, "conta corrente") > 0 Or _
       InStr(lowerText, "conta poupanca") > 0 Or InStr(lowerText, "n. da conta") > 0 Then
        findings = findings & "  - Conta bancaria detectada" & vbCrLf
    End If

    ' Agencia
    If InStr(lowerText, "agencia:") > 0 Or InStr(lowerText, "ag:") > 0 Then
        findings = findings & "  - Agencia bancaria detectada" & vbCrLf
    End If

    ' PIX
    If InStr(lowerText, "pix:") > 0 Or InStr(lowerText, "chave pix") > 0 Then
        findings = findings & "  - Chave PIX detectada" & vbCrLf
    End If

    ' Salario/Renda
    If InStr(lowerText, "salario:") > 0 Or InStr(lowerText, "renda:") > 0 Or _
       InStr(lowerText, "remuneracao:") > 0 Then
        findings = findings & "  - Informacao de renda detectada" & vbCrLf
    End If

    CheckFinancialData = findings
End Function

'================================================================================
' VERIFICA DADOS DE SAUDE
'================================================================================
Public Function CheckHealthData(docText As String) As String
    On Error Resume Next
    CheckHealthData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Cartao SUS
    If InStr(lowerText, "cartao sus") > 0 Or InStr(lowerText, "cns:") > 0 Or _
       InStr(lowerText, "cartao nacional de saude") > 0 Then
        findings = findings & "  - Cartao SUS detectado" & vbCrLf
    End If

    ' CID (Classificacao Internacional de Doencas)
    If InStr(lowerText, "cid:") > 0 Or InStr(lowerText, "cid-10") > 0 Or _
       InStr(lowerText, "cid 10") > 0 Then
        findings = findings & "  - Codigo CID detectado (DADO SENSIVEL ESPECIAL)" & vbCrLf
    End If

    ' Laudo medico
    If InStr(lowerText, "laudo medico") > 0 Or InStr(lowerText, "atestado medico") > 0 Then
        findings = findings & "  - Laudo/Atestado medico detectado (DADO SENSIVEL ESPECIAL)" & vbCrLf
    End If

    ' Deficiencia (dado sensivel especial)
    If InStr(lowerText, "deficiencia:") > 0 Or InStr(lowerText, "pcd") > 0 Or _
       InStr(lowerText, "pessoa com deficiencia") > 0 Then
        findings = findings & "  - Informacao de deficiencia detectada (DADO SENSIVEL ESPECIAL)" & vbCrLf
    End If

    ' Tipo sanguineo
    If InStr(lowerText, "tipo sanguineo") > 0 Or InStr(lowerText, "fator rh") > 0 Then
        findings = findings & "  - Tipo sanguineo detectado" & vbCrLf
    End If

    ' Alergia
    If InStr(lowerText, "alergia:") > 0 Or InStr(lowerText, "alergico a") > 0 Then
        findings = findings & "  - Informacao de alergia detectada" & vbCrLf
    End If

    CheckHealthData = findings
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
    patterns(1) = "$numero$/$ano$/pagina"

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

