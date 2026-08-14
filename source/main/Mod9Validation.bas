Attribute VB_Name = "Mod9Validation"
Option Explicit

' Mod9Validation
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)

Public Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogSection "VERIFICACOES INICIAIS"
    LogStepStart "Validacao de documento"

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Documento inacessivel"
        LogMessage "Documento nao acessivel para verificacao", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Erro: Tipo nao suportado"
        LogMessage "Tipo de documento nao suportado: " & doc.Type, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    ' Verifica se a primeira palavra e um tipo valido de propositura
    If Not ValidateProposituraType(doc) Then
        LogMessage "Usuario cancelou processamento - tipo de propositura nao reconhecido", LOG_LEVEL_INFO
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        Application.StatusBar = "Erro: Documento protegido"
        LogMessage "Documento protegido detectado: " & protectionType, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.ReadOnly Then
        Application.StatusBar = "Erro: Somente leitura"
        LogMessage "Documento em modo somente leitura: " & doc.FullName, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        Application.StatusBar = "Erro: Espaco insuficiente"
        LogMessage "Espaco em disco insuficiente para operacao segura", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        LogMessage "Estrutura do documento validada com avisos", LOG_LEVEL_WARNING
    End If

    ' Verifica presenca de possiveis dados sensiveis
    If Not CheckSensitiveData(doc) Then
        LogMessage "Aviso de dados sensiveis foi exibido ao usuario", LOG_LEVEL_INFO
    End If

    LogStepComplete "Validacao de documento", "Todas as verificacoes passaram"
    LogMessage "Verificacoes de seguranca concluidas com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Erro na verificacao"
    LogMessage "Erro durante verificacoes: " & Err.Description, LOG_LEVEL_ERROR
    PreviousChecking = False
End Function

'================================================================================
' VERIFICACAO DE ESPACO EM DISCO
'================================================================================

Public Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Verificacao simplificada - assume espaco suficiente se nao conseguir verificar
    Dim fso As Object
    Dim drive As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If

    ' Verificacao basica - 10MB minimo
    If drive.AvailableSpace < 10485760 Then ' 10MB em bytes
        LogMessage "Espaco em disco muito baixo", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If

    Exit Function

ErrorHandler:
    ' Se nao conseguir verificar, assume que ha espaco suficiente
    CheckDiskSpace = True
End Function

'================================================================================
' ROTINA PRINCIPAL DE FORMATACAO
'================================================================================

Public Function CheckSensitiveData(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim docText As String
    docText = ""
    If doc Is Nothing Then
        CheckSensitiveData = True
        Exit Function
    End If

    docText = doc.Range.text
    If Len(docText) < 10 Then
        CheckSensitiveData = True
        Exit Function
    End If

    Dim cpfValidCount As Long
    Dim cnpjValidCount As Long
    Dim cardValidCount As Long
    Dim cidCount As Long

    cpfValidCount = CountValidCPFInText(docText)
    cnpjValidCount = CountValidCNPJInText(docText)
    cardValidCount = CountLikelyCreditCardsInText(docText)
    cidCount = CountCID10InText(docText)

    If (cpfValidCount + cnpjValidCount + cardValidCount + cidCount) > 0 Then
        Dim findings As String
        findings = ""

          If cpfValidCount > 0 Then findings = findings & "  - Possivel CPF valido (formato + digitos verificadores) encontrado (" & cpfValidCount & "x)" & vbCrLf
          If cnpjValidCount > 0 Then findings = findings & "  - Possivel CNPJ valido (formato + digitos verificadores) encontrado (" & cnpjValidCount & "x)" & vbCrLf
        If cardValidCount > 0 Then findings = findings & "  - Possivel numero de cartao (Luhn) detectado (" & cardValidCount & "x)" & vbCrLf
          If cidCount > 0 Then findings = findings & "  - Possivel CID (saude) encontrado (" & cidCount & "x)" & vbCrLf

        Dim msg As String
          msg = "ATENCAO: POSSIVEIS DADOS SENSIVEIS IDENTIFICADOS (LGPD)" & vbCrLf & vbCrLf & _
              findings & vbCrLf & _
              "Recomenda-se revisar e, se aplicavel, remover/anonimizar antes de prosseguir."

        MsgBox msg, vbExclamation, "Verificacao LGPD"
        LogMessage "LGPD (estrito): CPF=" & cpfValidCount & ", CNPJ=" & cnpjValidCount & ", Cartao=" & cardValidCount & ", CID=" & cidCount, LOG_LEVEL_WARNING
        CheckSensitiveData = False
        Exit Function
    End If

    LogMessage "Verificacao LGPD (estrito): nenhum achado grave", LOG_LEVEL_INFO
    CheckSensitiveData = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na verificacao LGPD (estrito): " & Err.Description, LOG_LEVEL_WARNING
    CheckSensitiveData = True
End Function


Public Function CountValidCPFInText(text As String) As Long
    On Error GoTo ErrorHandler
    CountValidCPFInText = 0

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = "\b\d{3}\.\d{3}\.\d{3}-\d{2}\b|\b\d{11}\b"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches Is Nothing Then Exit Function

    Dim m As Object
    For Each m In matches
        Dim digits As String
        digits = OnlyDigits(CStr(m.Value))
        If Len(digits) = 11 Then
            If IsValidCPF(digits) Then CountValidCPFInText = CountValidCPFInText + 1
        End If
    Next m

    Exit Function

ErrorHandler:
    CountValidCPFInText = 0
End Function


Public Function CountValidCNPJInText(text As String) As Long
    On Error GoTo ErrorHandler
    CountValidCNPJInText = 0

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    ' Formatos com pontuacao (mais confiaveis) e formato puro apenas quando antecedido por "cnpj"
    re.Pattern = "\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b|\bcnpj\s*[:\-]?\s*\d{14}\b"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches Is Nothing Then Exit Function

    Dim m As Object
    For Each m In matches
        Dim digits As String
        digits = OnlyDigits(CStr(m.Value))
        If Len(digits) = 14 Then
            If IsValidCNPJ(digits) Then CountValidCNPJInText = CountValidCNPJInText + 1
        End If
    Next m

    Exit Function

ErrorHandler:
    CountValidCNPJInText = 0
End Function


Public Function CountLikelyCreditCardsInText(text As String) As Long
    On Error GoTo ErrorHandler
    CountLikelyCreditCardsInText = 0

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    ' Busca sequencias tipicas com separadores (espaco ou hifen) para reduzir falsos positivos.
    re.Pattern = "\b(?:\d[ -]){12,18}\d\b"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches Is Nothing Then Exit Function

    Dim m As Object
    For Each m In matches
        Dim digits As String
        digits = OnlyDigits(CStr(m.Value))

        ' Comprimento tipico de cartao: 13 a 19
        If Len(digits) >= 13 And Len(digits) <= 19 Then
            ' Evita contar CPF/CNPJ como cartao
            If Len(digits) <> 11 And Len(digits) <> 14 Then
                If IsLuhnValid(digits) And Not IsAllSameDigit(digits) Then
                    CountLikelyCreditCardsInText = CountLikelyCreditCardsInText + 1
                End If
            End If
        End If
    Next m

    Exit Function

ErrorHandler:
    CountLikelyCreditCardsInText = 0
End Function


Public Function CountCID10InText(text As String) As Long
    On Error GoTo ErrorHandler
    CountCID10InText = 0

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    ' Padrao estrito: exige literal "CID" seguido de codigo tipo A00 ou A00.0
    re.Pattern = "\bCID(?:-?10)?\s*[:\-]?\s*[A-TV-Z][0-9]{2}(?:\.[0-9A-TV-Z]{1,2})?\b"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches Is Nothing Then Exit Function

    CountCID10InText = matches.count
    Exit Function

ErrorHandler:
    CountCID10InText = 0
End Function


Public Function OnlyDigits(text As String) As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    outText = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "[0-9]" Then outText = outText & ch
    Next i

    OnlyDigits = outText
End Function


Public Function IsAllSameDigit(digits As String) As Boolean
    Dim i As Long
    If Len(digits) <= 1 Then
        IsAllSameDigit = False
        Exit Function
    End If

    Dim firstChar As String
    firstChar = Mid$(digits, 1, 1)

    For i = 2 To Len(digits)
        If Mid$(digits, i, 1) <> firstChar Then
            IsAllSameDigit = False
            Exit Function
        End If
    Next i

    IsAllSameDigit = True
End Function


Public Function IsValidCPF(cpfDigits As String) As Boolean
    On Error GoTo ErrorHandler
    IsValidCPF = False

    If Len(cpfDigits) <> 11 Then Exit Function
    If IsAllSameDigit(cpfDigits) Then Exit Function

    Dim i As Long
    Dim sum As Long
    Dim rest As Long
    Dim d1 As Long
    Dim d2 As Long

    ' Primeiro digito verificador
    sum = 0
    For i = 1 To 9
        sum = sum + (CLng(Mid$(cpfDigits, i, 1)) * (11 - i))
    Next i
    rest = sum Mod 11
    If rest < 2 Then
        d1 = 0
    Else
        d1 = 11 - rest
    End If

    ' Segundo digito verificador
    sum = 0
    For i = 1 To 10
        sum = sum + (CLng(Mid$(cpfDigits, i, 1)) * (12 - i))
    Next i
    rest = sum Mod 11
    If rest < 2 Then
        d2 = 0
    Else
        d2 = 11 - rest
    End If

    IsValidCPF = (CLng(Mid$(cpfDigits, 10, 1)) = d1 And CLng(Mid$(cpfDigits, 11, 1)) = d2)
    Exit Function

ErrorHandler:
    IsValidCPF = False
End Function


Public Function IsValidCNPJ(cnpjDigits As String) As Boolean
    On Error GoTo ErrorHandler
    IsValidCNPJ = False

    If Len(cnpjDigits) <> 14 Then Exit Function
    If IsAllSameDigit(cnpjDigits) Then Exit Function

    Dim weights1 As Variant
    Dim weights2 As Variant
    weights1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    weights2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)

    Dim i As Long
    Dim sum As Long
    Dim rest As Long
    Dim d1 As Long
    Dim d2 As Long

    sum = 0
    For i = 1 To 12
        sum = sum + (CLng(Mid$(cnpjDigits, i, 1)) * CLng(weights1(i - 1)))
    Next i
    rest = sum Mod 11
    If rest < 2 Then
        d1 = 0
    Else
        d1 = 11 - rest
    End If

    sum = 0
    For i = 1 To 13
        sum = sum + (CLng(Mid$(cnpjDigits, i, 1)) * CLng(weights2(i - 1)))
    Next i
    rest = sum Mod 11
    If rest < 2 Then
        d2 = 0
    Else
        d2 = 11 - rest
    End If

    IsValidCNPJ = (CLng(Mid$(cnpjDigits, 13, 1)) = d1 And CLng(Mid$(cnpjDigits, 14, 1)) = d2)
    Exit Function

ErrorHandler:
    IsValidCNPJ = False
End Function


Public Function IsLuhnValid(digits As String) As Boolean
    On Error GoTo ErrorHandler
    IsLuhnValid = False

    Dim sum As Long
    Dim i As Long
    Dim digit As Long
    Dim alt As Boolean

    sum = 0
    alt = False

    For i = Len(digits) To 1 Step -1
        digit = CLng(Mid$(digits, i, 1))
        If alt Then
            digit = digit * 2
            If digit > 9 Then digit = digit - 9
        End If
        sum = sum + digit
        alt = Not alt
    Next i

    IsLuhnValid = (sum Mod 10 = 0)
    Exit Function

ErrorHandler:
    IsLuhnValid = False
End Function

'================================================================================
' VERIFICA DOCUMENTOS DE IDENTIFICACAO
'================================================================================

Public Function CheckDocumentIdentifiers(docText As String) As String
    On Error Resume Next
    CheckDocumentIdentifiers = ""

    Dim lowerText As String
    Dim findings As String
    Dim cpfCount As Long
    Dim rgCount As Long

    lowerText = LCase(docText)
    findings = ""

    ' Verifica mencoes a CPF
    cpfCount = 0
    If InStr(lowerText, "cpf:") > 0 Then cpfCount = cpfCount + 1
    If InStr(lowerText, "cpf n") > 0 Then cpfCount = cpfCount + 1
    If InStr(lowerText, "cpf/mf") > 0 Then cpfCount = cpfCount + 1
    If InStr(lowerText, "inscrito no cpf") > 0 Then cpfCount = cpfCount + 1

    ' Detecta padrao numerico de CPF (XXX.XXX.XXX-XX)
    If ContainsCPFPattern(docText) Then cpfCount = cpfCount + 1

    If cpfCount > 0 Then
        findings = findings & "  - CPF detectado" & vbCrLf
    End If

    ' Verifica mencoes a RG
    rgCount = 0
    If InStr(lowerText, "rg:") > 0 Then rgCount = rgCount + 1
    If InStr(lowerText, "rg n") > 0 Then rgCount = rgCount + 1
    If InStr(lowerText, "identidade n") > 0 Then rgCount = rgCount + 1
    If InStr(lowerText, "carteira de identidade") > 0 Then rgCount = rgCount + 1

    ' Detecta padrao numerico de RG
    If ContainsRGPattern(docText) Then rgCount = rgCount + 1

    If rgCount > 0 Then
        findings = findings & "  - RG/Identidade detectado" & vbCrLf
    End If

    ' CNH
    If InStr(lowerText, "cnh:") > 0 Or InStr(lowerText, "cnh n") > 0 Or _
       InStr(lowerText, "habilitacao n") > 0 Then
        findings = findings & "  - CNH detectada" & vbCrLf
    End If

    ' CTPS
    If InStr(lowerText, "ctps") > 0 Or InStr(lowerText, "carteira de trabalho") > 0 Then
        findings = findings & "  - CTPS detectada" & vbCrLf
    End If

    ' Titulo de eleitor
    If InStr(lowerText, "titulo de eleitor") > 0 Or InStr(lowerText, "titulo eleitoral") > 0 Then
        findings = findings & "  - Titulo de eleitor detectado" & vbCrLf
    End If

    ' PIS/PASEP
    If InStr(lowerText, "pis:") > 0 Or InStr(lowerText, "pis/pasep") > 0 Or _
       InStr(lowerText, "pasep:") > 0 Then
        findings = findings & "  - PIS/PASEP detectado" & vbCrLf
    End If

    CheckDocumentIdentifiers = findings
End Function

'================================================================================
' LIMPEZA DE FORMATACAO
'================================================================================
