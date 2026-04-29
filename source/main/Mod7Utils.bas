' Mod7Utils.bas
Option Explicit

'================================================================================
' FUNCOES AUXILIARES DE LIMPEZA DE TEXTO
'================================================================================
Public Function GetCleanParagraphText(para As Paragraph) As String
    On Error Resume Next

    Dim txt As String
    txt = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

    ' Remove pontuacao final com protecao contra loop infinito
    Dim safetyCounter As Long
    safetyCounter = 0
    Do While Len(txt) > 0 And InStr(".,;:", Right(txt, 1)) > 0 And safetyCounter < MAX_LOOP_ITERATIONS
        txt = Left(txt, Len(txt) - 1)
        safetyCounter = safetyCounter + 1
    Loop

        GetCleanParagraphText = RemovePunctuation(Trim(LCase(txt)))
End Function

Public Function RemovePunctuation(text As String) As String
    Dim result As String
    result = text

    ' Remove pontuacao final com protecao contra loop infinito
    Dim safetyCounter As Long
    safetyCounter = 0
    Do While Len(result) > 0 And InStr(".,;:", Right(result, 1)) > 0 And safetyCounter < 100
        result = Left(result, Len(result) - 1)
        safetyCounter = safetyCounter + 1
    Loop

    RemovePunctuation = Trim(result)
End Function

'================================================================================
' ACESSO SEGURO A PROPRIEDADES
'================================================================================
Public Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo FallbackMethod

    ' Metodo preferido - mais rapido
    SafeGetCharacterCount = targetRange.Characters.count
    Exit Function

FallbackMethod:
    On Error GoTo ErrorHandler
    ' Metodo alternativo para versoes com problemas de .Characters.Count
    SafeGetCharacterCount = Len(targetRange.text)
    Exit Function

ErrorHandler:
    ' Ultimo recurso - valor padrao seguro
    SafeGetCharacterCount = 0
    LogMessage "Erro ao obter contagem de caracteres: " & Err.Description, LOG_LEVEL_WARNING
End Function

Public Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler

    ' Aplica formatacao de fonte de forma segura
    With targetRange.Font
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .size = fontSize
        .Color = wdColorAutomatic
    End With

    SafeSetFont = True
    Exit Function

ErrorHandler:
    SafeSetFont = False
    LogMessage "Erro ao aplicar fonte: " & Err.Description & " - Range: " & Left(targetRange.text, 20), LOG_LEVEL_WARNING
End Function

Public Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
    On Error GoTo ErrorHandler

    With para.Format
        If alignment >= 0 Then .alignment = alignment
        If leftIndent >= 0 Then .leftIndent = leftIndent
        If firstLineIndent >= 0 Then .firstLineIndent = firstLineIndent
    End With

    SafeSetParagraphFormat = True
    Exit Function

ErrorHandler:
    SafeSetParagraphFormat = False
    LogMessage "Erro ao aplicar formatacao de paragrafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

Public Function SafeHasVisualContent(para As Paragraph) As Boolean
    On Error GoTo SafeMode

    ' Verificacao padrao mais robusta
    Dim hasImages As Boolean
    Dim hasShapes As Boolean

    ' Verifica imagens inline de forma segura
    hasImages = (para.Range.InlineShapes.count > 0)

    ' Verifica shapes flutuantes de forma segura
    hasShapes = False
    If Not hasImages Then
        Dim shp As shape
        For Each shp In para.Range.ShapeRange
            hasShapes = True
            Exit For
        Next shp
    End If

    SafeHasVisualContent = hasImages Or hasShapes
    Exit Function

SafeMode:
    On Error GoTo ErrorHandler
    ' Metodo alternativo mais simples
    SafeHasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function

ErrorHandler:
    ' Em caso de erro, assume que nao ha conteudo visual
    SafeHasVisualContent = False
End Function

'================================================================================
' ACESSO SEGURO A CARACTERES
'================================================================================
Public Function SafeGetLastCharacter(rng As Range) As String
    On Error GoTo ErrorHandler

    Dim charCount As Long
    charCount = SafeGetCharacterCount(rng)

    If charCount > 0 Then
        SafeGetLastCharacter = rng.Characters(charCount).text
    Else
        SafeGetLastCharacter = ""
    End If
    Exit Function

ErrorHandler:
    ' Metodo alternativo usando Right()
    On Error GoTo FinalFallback
    SafeGetLastCharacter = Right(rng.text, 1)
    Exit Function

FinalFallback:
    SafeGetLastCharacter = ""
End Function

'================================================================================
' UTILITY: GET PROTECTION TYPE
'================================================================================
Public Function GetProtectionType(doc As Document) As String
    On Error Resume Next

    Select Case doc.protectionType
        Case wdNoProtection: GetProtectionType = "Sem protecao"
        Case 1: GetProtectionType = "Protegido contra revisoes"
        Case 2: GetProtectionType = "Protegido contra comentarios"
        Case 3: GetProtectionType = "Protegido contra formularios"
        Case 4: GetProtectionType = "Protegido contra leitura"
        Case Else: GetProtectionType = "Tipo desconhecido (" & doc.protectionType & ")"
    End Select
End Function

'================================================================================
' UTILITY: GET DOCUMENT SIZE
'================================================================================
Public Function GetDocumentSize(doc As Document) As String
    On Error Resume Next

    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").value * 2

    If Err.Number <> 0 Then
        GetDocumentSize = "Desconhecido"
        Exit Function
    End If

    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' UTILITY: GET WINDOWS VERSION
'================================================================================
Public Function GetWindowsVersion() As String
    On Error Resume Next

    Dim osVersion As String
    osVersion = Environ("OS")

    If osVersion = "" Then osVersion = "Windows"

    GetWindowsVersion = osVersion
End Function

'================================================================================
' UTILITY: GET WORD VERSION NAME
'================================================================================
Public Function GetWordVersionName() As String
    On Error Resume Next

    Dim ver As String
    ver = Application.version

    Select Case ver
        Case "16.0": GetWordVersionName = "Word 2016/2019/2021/365"
        Case "15.0": GetWordVersionName = "Word 2013"
        Case "14.0": GetWordVersionName = "Word 2010"
        Case "12.0": GetWordVersionName = "Word 2007"
        Case "11.0": GetWordVersionName = "Word 2003"
        Case Else: GetWordVersionName = "Word " & ver
    End Select
End Function

'================================================================================
' UTILITY: GET USER INITIALS
'================================================================================
' NORMALIZA TEXTO PARA COMPARACAO (remove acentos e converte para minusculas)
'================================================================================
Public Function NormalizeForComparison(text As String) As String
    Dim result As String
    result = LCase(text)

    ' Remove acentos comuns do portugues
    result = Replace(result, Chr(225), "a") ' a com acento agudo
    result = Replace(result, Chr(227), "a") ' a com til
    result = Replace(result, Chr(226), "a") ' a com circunflexo
    result = Replace(result, Chr(224), "a") ' a com acento grave
    result = Replace(result, Chr(233), "e") ' e com acento agudo
    result = Replace(result, Chr(234), "e") ' e com circunflexo
    result = Replace(result, Chr(237), "i") ' i com acento agudo
    result = Replace(result, Chr(243), "o") ' o com acento agudo
    result = Replace(result, Chr(245), "o") ' o com til
    result = Replace(result, Chr(244), "o") ' o com circunflexo
    result = Replace(result, Chr(250), "u") ' u com acento agudo
    result = Replace(result, Chr(231), "c") ' c cedilha

    NormalizeForComparison = result
End Function

'================================================================================
' CALCULA A DISTANCIA DE LEVENSHTEIN ENTRE DUAS STRINGS
'================================================================================
Public Function LevenshteinDistance(s1 As String, s2 As String) As Long
    Dim len1 As Long, len2 As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim d() As Long

    len1 = Len(s1)
    len2 = Len(s2)

    ' Casos triviais
    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If

    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If

    ' Matriz de distancias
    ReDim d(0 To len1, 0 To len2)

    ' Inicializa primeira coluna e linha
    For i = 0 To len1
        d(i, 0) = i
    Next i

    For j = 0 To len2
        d(0, j) = j
    Next j

    ' Calcula distancias
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            ' Minimo entre insercao, delecao e substituicao
            d(i, j) = MinOfThree(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = d(len1, len2)
End Function

'================================================================================
' RETORNA O MINIMO DE TRES VALORES
'================================================================================
Public Function MinOfThree(a As Long, b As Long, c As Long) As Long
    MinOfThree = a
    If b < MinOfThree Then MinOfThree = b
    If c < MinOfThree Then MinOfThree = c
End Function

'================================================================================
' CONTA DIGITOS EM UMA STRING
'================================================================================
Public Function CountDigitsInString(text As String) As Long
    On Error Resume Next
    CountDigitsInString = 0

    Dim i As Long
    Dim count As Long

    count = 0
    For i = 1 To Len(text)
        If Mid(text, i, 1) Like "[0-9]" Then
            count = count + 1
        End If
    Next i

    CountDigitsInString = count
End Function

'================================================================================
' GET CLIPBOARD DATA - Obtem dados da area de transferencia
'================================================================================
Public Function GetClipboardData() As Variant
    On Error GoTo ErrorHandler

    ' Placeholder para dados da area de transferencia
    ' Em uma implementacao completa, seria necessario usar APIs do Windows
    ' ou metodos mais avancados para capturar dados binarios
    GetClipboardData = "ImageDataPlaceholder"
    Exit Function

ErrorHandler:
    GetClipboardData = Empty
End Function

