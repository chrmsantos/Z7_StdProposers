Attribute VB_Name = "Mod12AIStructure"
Option Explicit

' Mod12AIStructure.bas
' =============================================================================
' Z7_STDPROPOSERS - Identificacao de Estrutura do Documento via IA (OpenRouter)
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
' Este modulo fornece identificacao de elementos estruturais de proposituras
' legislativas utilizando a API da OpenRouter (mesma infraestrutura de
' Mod11RevisionText.bas). A IA analisa o texto completo do documento e
' identifica: Titulo, Ementa, Vocativo, Corpo, Titulo da Justificativa,
' Justificativa, Data, Assinatura, Titulo do Anexo e Anexo.
'
' Em caso de falha na chamada a IA, o chamador deve recorrer a implementacao
' heuristica (IdentifyDocumentStructureHeuristics em Mod2Engine.bas).
' =============================================================================

' =============================================================================
' CONSTANTES
' =============================================================================
Private Const AI_STRUCT_URL As String = _
    "https://openrouter.ai/api/v1/chat/completions"

Private Const AI_STRUCT_PREFIX As String = "AI_STRUCTURE"
Private Const AI_STRUCT_DEFAULT_MODEL As String = _
    "deepseek/deepseek-v4-pro"

' Timeouts em milissegundos (resolve, connect, send, receive)
Private Const AI_STRUCT_RESOLVE_TIMEOUT As Long = 5000
Private Const AI_STRUCT_CONNECT_TIMEOUT As Long = 10000
Private Const AI_STRUCT_SEND_TIMEOUT As Long = 30000
Private Const AI_STRUCT_RECEIVE_TIMEOUT As Long = 60000

' Maximo de paragrafos para enviar a IA (protecao de contexto)
Private Const MAX_PARAGRAPHS_FOR_AI As Long = 400

' Comprimento maximo de cada paragrafo enviado (evita tokens excessivos)
Private Const MAX_PARAGRAPH_TEXT_LENGTH As Long = 500

' =============================================================================
' DECLARACOES DA API WINDOWS (DPAPI) - mesma infraestrutura de Mod11
' =============================================================================
#If VBA7 Then
Private Type AI_DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type
Private Declare PtrSafe Function AI_CryptUnprotectData _
    Lib "crypt32.dll" Alias "CryptUnprotectData" ( _
    ByRef pDataIn As AI_DATA_BLOB, _
    ByVal ppszDataDescr As LongPtr, _
    ByVal pOptionalEntropy As LongPtr, _
    ByVal pvReserved As LongPtr, _
    ByVal pPromptStruct As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As AI_DATA_BLOB _
    ) As Long
Private Declare PtrSafe Sub AI_CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As LongPtr)
Private Declare PtrSafe Function AI_LocalFree _
    Lib "kernel32" Alias "LocalFree" ( _
    ByVal hMem As LongPtr _
    ) As LongPtr
#Else
Private Type AI_DATA_BLOB
    cbData As Long
    pbData As Long
End Type
Private Declare Function AI_CryptUnprotectData _
    Lib "crypt32.dll" Alias "CryptUnprotectData" ( _
    ByRef pDataIn As AI_DATA_BLOB, _
    ByVal ppszDataDescr As Long, _
    ByVal pOptionalEntropy As Long, _
    ByVal pvReserved As Long, _
    ByVal pPromptStruct As Long, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As AI_DATA_BLOB _
    ) As Long
Private Declare Sub AI_CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)
Private Declare Function AI_LocalFree _
    Lib "kernel32" Alias "LocalFree" ( _
    ByVal hMem As Long _
    ) As Long
#End If

' =============================================================================
' FUNCAO PRINCIPAL: IdentifyDocumentStructureWithAI
' =============================================================================
Public Function IdentifyDocumentStructureWithAI(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    IdentifyDocumentStructureWithAI = False
    If doc Is Nothing Then Exit Function
    If doc.Paragraphs.count = 0 Then Exit Function

    Dim startTime As Double
    startTime = Timer
    LogMessage AI_STRUCT_PREFIX & ": Iniciando identificacao via IA...", LOG_LEVEL_INFO

    Dim docText As String
    docText = MontarTextoDocumentoParaIA(doc)
    If Len(docText) = 0 Then
        LogMessage AI_STRUCT_PREFIX & ": Texto vazio", LOG_LEVEL_WARNING
        Exit Function
    End If

    Dim apiKey As String
    apiKey = AI_CarregarChaveAPI()
    If Len(apiKey) = 0 Then
        LogMessage AI_STRUCT_PREFIX & ": Chave API nao disponivel", LOG_LEVEL_WARNING
        Exit Function
    End If

    Dim modelo As String
    modelo = AI_CarregarModelo()
    LogMessage AI_STRUCT_PREFIX & ": Modelo: " & modelo, LOG_LEVEL_INFO

    Dim jsonPayload As String
    jsonPayload = MontarJSONPayload(modelo, _
        EscaparJSONAI(MontarPromptEstrutura()), EscaparJSONAI(docText))

    Dim resposta As String
    resposta = AI_ChamarAPI(apiKey, jsonPayload)
    If Len(resposta) = 0 Then
        LogMessage AI_STRUCT_PREFIX & ": Resposta IA vazia", LOG_LEVEL_WARNING
        Exit Function
    End If

    If Not ParsearRespostaEstruturaIA(resposta, doc) Then
        LogMessage AI_STRUCT_PREFIX & ": Falha ao parsear resposta", LOG_LEVEL_WARNING
        Exit Function
    End If

    If Not ValidarIndicesEstrutura(doc) Then
        LogMessage AI_STRUCT_PREFIX & ": Indices invalidos - fallback", LOG_LEVEL_WARNING
        Exit Function
    End If

    MarcarFlagsEstrutura doc

    LogMessage AI_STRUCT_PREFIX & ": Estrutura IA em " & _
               Format(Timer - startTime, "0.00") & "s", LOG_LEVEL_INFO
    LogMessage "=== ESTRUTURA(IA): T=" & tituloParaIndex & " E=" & ementaParaIndex & _
               " V=" & vocativoStartIndex & "-" & vocativoEndIndex & _
               " C=" & corpoStartIndex & "-" & corpoEndIndex & _
               " TJ=" & tituloJustificativaIndex & _
               " J=" & justificativaStartIndex & "-" & justificativaEndIndex & _
               " D=" & dataParaIndex & _
               " A=" & assinaturaStartIndex & "-" & assinaturaEndIndex & _
               " AN=" & anexoStartIndex & "-" & anexoEndIndex & " ===", LOG_LEVEL_INFO

    IdentifyDocumentStructureWithAI = True
    Exit Function

ErrorHandler:
    LogMessage AI_STRUCT_PREFIX & ": Erro IA: " & Err.Description, LOG_LEVEL_ERROR
    IdentifyDocumentStructureWithAI = False
End Function

' =============================================================================
' MONTA TEXTO DO DOCUMENTO COM INDICES DE PARAGRAFOS
' =============================================================================
Private Function MontarTextoDocumentoParaIA(doc As Document) As String

' =============================================================================
' PROMPT DE SISTEMA
' =============================================================================
Private Function MontarPromptEstrutura() As String
    MontarPromptEstrutura = _
        "Voce e um especialista em analise de documentos legislativos brasileiros. " & _
        "Identifique a estrutura de uma propositura legislativa. " & _
        "O documento tem marcadores [Pn] indicando o numero de cada paragrafo." & vbCrLf & vbCrLf & _
        "Retorne APENAS o JSON abaixo, sem explicacoes:" & vbCrLf & vbCrLf & _
        "{""titulo"":[primeiro,ultimo],""ementa"":[primeiro,ultimo]," & _
        """vocativo"":[primeiro,ultimo],""corpo"":[primeiro,ultimo]," & _
        """titulo_da_justificativa":[unico],""justificativa"":[primeiro,ultimo]," & _
        """data"":[unico],""assinatura"":[primeiro,ultimo]," & _
        """titulo_do_anexo"":[unico ou null],""anexo"":[primeiro,ultimo ou null]}" & vbCrLf & vbCrLf & _
        "REGRAS:" & vbCrLf & _
        "1. Valores sao NUMEROS dos paragrafos (sem 'P')." & vbCrLf & _
        "2. Se nao existir, use 0 (zero)." & vbCrLf & _
        "3. 'corpo' e o texto principal entre vocativo e justificativa." & vbCrLf & _
        "4. Assinatura: 3 paragrafos centralizados no final." & vbCrLf & _
        "5. Data: contem nome do plenario e data de emissao."

' =============================================================================
' MONTA JSON DO PAYLOAD
' =============================================================================
Private Function MontarJSONPayload(ByVal modelo As String, _
    ByVal systemJSON As String, ByVal userJSON As String) As String
    MontarJSONPayload = "{""model"":""" & modelo & """,""temperature"":0.1," & _
        """messages"":[{""role"":""system"",""content"":""" & systemJSON & _
        """},{""role"":""user"",""content"":""" & userJSON & """}]}"
End Function

' =============================================================================
' CHAMADA HTTP A API (OPENROUTER)
' =============================================================================
Private Function AI_ChamarAPI(ByVal apiKey As String, _
    ByVal jsonPayload As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts AI_STRUCT_RESOLVE_TIMEOUT, AI_STRUCT_CONNECT_TIMEOUT, _
        AI_STRUCT_SEND_TIMEOUT, AI_STRUCT_RECEIVE_TIMEOUT
    http.Open "POST", AI_STRUCT_URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.setRequestHeader "HTTP-Referer", "https://localhost"
    http.setRequestHeader "X-Title", "Word - Estrutura Documento"
    http.send AI_StringParaUTF8(jsonPayload)

    If http.Status = 200 Then
        AI_ChamarAPI = AI_BytesParaStringUTF8(http.ResponseBody)
    Else
        Dim errResp As String
        errResp = AI_BytesParaStringUTF8(http.ResponseBody)
        LogMessage AI_STRUCT_PREFIX & ": HTTP " & http.Status & " - " & Left(errResp, 200), LOG_LEVEL_ERROR
        AI_ChamarAPI = ""
    End If
    Set http = Nothing
    Exit Function
ErrorHandler:

' =============================================================================
' PARSEIA RESPOSTA JSON DA IA
' =============================================================================
Private Function ParsearRespostaEstruturaIA(ByVal resposta As String, _
    doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ParsearRespostaEstruturaIA = False

    Dim content As String
    content = AI_ExtrairContentJSON(resposta)
    If Len(content) = 0 Then
        LogMessage AI_STRUCT_PREFIX & ": content vazio na resposta", LOG_LEVEL_WARNING
        Exit Function
    End If

    LogMessage AI_STRUCT_PREFIX & ": Resposta: " & Left(content, 300), LOG_LEVEL_DEBUG

    ' Reseta todos os indices
    tituloParaIndex = 0: ementaParaIndex = 0
    vocativoStartIndex = 0: vocativoEndIndex = 0
    corpoStartIndex = 0: corpoEndIndex = 0
    tituloJustificativaIndex = 0
    justificativaStartIndex = 0: justificativaEndIndex = 0
    dataParaIndex = 0
    assinaturaStartIndex = 0: assinaturaEndIndex = 0
    tituloAnexoIndex = 0
    anexoStartIndex = 0: anexoEndIndex = 0

    tituloParaIndex = AI_ExtrairIndiceUnico(content, "titulo")
    ementaParaIndex = AI_ExtrairIndiceUnico(content, "ementa")
    vocativoStartIndex = AI_ExtrairArrayPrimeiro(content, "vocativo")
    vocativoEndIndex = AI_ExtrairArrayUltimo(content, "vocativo")
    corpoStartIndex = AI_ExtrairArrayPrimeiro(content, "corpo")
    corpoEndIndex = AI_ExtrairArrayUltimo(content, "corpo")
    tituloJustificativaIndex = AI_ExtrairIndiceUnico(content, "titulo_da_justificativa")
    justificativaStartIndex = AI_ExtrairArrayPrimeiro(content, "justificativa")
    justificativaEndIndex = AI_ExtrairArrayUltimo(content, "justificativa")
    dataParaIndex = AI_ExtrairIndiceUnico(content, "data")
    assinaturaStartIndex = AI_ExtrairArrayPrimeiro(content, "assinatura")
    assinaturaEndIndex = AI_ExtrairArrayUltimo(content, "assinatura")
    tituloAnexoIndex = AI_ExtrairIndiceUnico(content, "titulo_do_anexo")
    anexoStartIndex = AI_ExtrairArrayPrimeiro(content, "anexo")
    anexoEndIndex = AI_ExtrairArrayUltimo(content, "anexo")

' =============================================================================
' VALIDA INDICES DE ESTRUTURA
' =============================================================================
Private Function ValidarIndicesEstrutura(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    ValidarIndicesEstrutura = False

    Dim maxPara As Long
    maxPara = doc.Paragraphs.count

    ' Pelo menos titulo deve ter sido identificado
    If tituloParaIndex <= 0 Or tituloParaIndex > maxPara Then Exit Function
    If ementaParaIndex > maxPara Then Exit Function
    If vocativoStartIndex > maxPara Or vocativoEndIndex > maxPara Then Exit Function
    If vocativoStartIndex > 0 And vocativoEndIndex > 0 Then
        If vocativoStartIndex > vocativoEndIndex Then Exit Function
    End If
    If corpoStartIndex > maxPara Or corpoEndIndex > maxPara Then Exit Function
    If corpoStartIndex > 0 And corpoEndIndex > 0 Then
        If corpoStartIndex > corpoEndIndex Then Exit Function
    End If
    If tituloJustificativaIndex > maxPara Then Exit Function
    If justificativaStartIndex > maxPara Or justificativaEndIndex > maxPara Then Exit Function
    If justificativaStartIndex > 0 And justificativaEndIndex > 0 Then
        If justificativaStartIndex > justificativaEndIndex Then Exit Function
    End If
    If dataParaIndex > maxPara Then Exit Function
    If assinaturaStartIndex > maxPara Or assinaturaEndIndex > maxPara Then Exit Function
    If assinaturaStartIndex > 0 And assinaturaEndIndex > 0 Then
        If assinaturaStartIndex > assinaturaEndIndex Then Exit Function
    End If
    If tituloAnexoIndex > maxPara Then Exit Function
    If anexoStartIndex > maxPara Or anexoEndIndex > maxPara Then Exit Function
    If anexoStartIndex > 0 And anexoEndIndex > 0 Then
        If anexoStartIndex > anexoEndIndex Then Exit Function
    End If

    ValidarIndicesEstrutura = True

' =============================================================================
' MARCA FLAGS DE ESTRUTURA NO CACHE
' =============================================================================
Private Sub MarcarFlagsEstrutura(doc As Document)
    On Error GoTo ErrorHandler

    Dim i As Long
    For i = 1 To cacheSize
        If i > doc.Paragraphs.count Then Exit For
        With paragraphCache(i)
            .isTitulo = False: .isEmenta = False: .isVocativo = False
            .isCorpoContent = False: .isTituloJustificativa = False
            .isJustificativaContent = False: .isData = False
            .isAssinatura = False: .isTituloAnexo = False: .isAnexoContent = False

            If i = tituloParaIndex Then .isTitulo = True
            If i = ementaParaIndex Then .isEmenta = True
            If i = dataParaIndex Then .isData = True
            If i = tituloJustificativaIndex Then .isTituloJustificativa = True

            If assinaturaStartIndex > 0 And assinaturaEndIndex > 0 Then
                If i >= assinaturaStartIndex And i <= assinaturaEndIndex Then .isAssinatura = True
            End If
            If vocativoStartIndex > 0 And vocativoEndIndex > 0 Then
                If i >= vocativoStartIndex And i <= vocativoEndIndex Then .isVocativo = True
            End If
            If corpoStartIndex > 0 And corpoEndIndex > 0 Then
                If i >= corpoStartIndex And i <= corpoEndIndex Then
                    If Not .isVocativo Then .isCorpoContent = True
                End If
            End If
            If justificativaStartIndex > 0 And justificativaEndIndex > 0 Then
                If i >= justificativaStartIndex And i <= justificativaEndIndex Then .isJustificativaContent = True
            End If
            If tituloAnexoIndex > 0 And i = tituloAnexoIndex Then .isTituloAnexo = True
            If anexoStartIndex > 0 And i >= anexoStartIndex Then .isAnexoContent = True
        End With
    Next i

' =============================================================================
' FUNCOES AUXILIARES DE PARSE JSON
' =============================================================================

Private Function AI_ExtrairIndiceUnico(ByVal json As String, ByVal chave As String) As Long
    On Error GoTo ErrorHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True: regex.Global = False

    regex.Pattern = """" & chave & """\s*:\s*\[\s*(\d+)\s*\]"
    If regex.Test(json) Then AI_ExtrairIndiceUnico = CLng(regex.Execute(json)(0).SubMatches(0)): Exit Function
    regex.Pattern = """" & chave & """\s*:\s*\[\s*(\d+)\s*,"
    If regex.Test(json) Then AI_ExtrairIndiceUnico = CLng(regex.Execute(json)(0).SubMatches(0)): Exit Function
    regex.Pattern = """" & chave & """\s*:\s*(\d+)"
    If regex.Test(json) Then AI_ExtrairIndiceUnico = CLng(regex.Execute(json)(0).SubMatches(0)): Exit Function

    AI_ExtrairIndiceUnico = 0
    Exit Function
ErrorHandler: AI_ExtrairIndiceUnico = 0
End Function

Private Function AI_ExtrairArrayPrimeiro(ByVal json As String, ByVal chave As String) As Long
    On Error GoTo ErrorHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True: regex.Global = False
    regex.Pattern = """" & chave & """\s*:\s*\[\s*(\d+)"
    If regex.Test(json) Then AI_ExtrairArrayPrimeiro = CLng(regex.Execute(json)(0).SubMatches(0)): Exit Function
    AI_ExtrairArrayPrimeiro = 0
    Exit Function
ErrorHandler: AI_ExtrairArrayPrimeiro = 0
End Function

Private Function AI_ExtrairArrayUltimo(ByVal json As String, ByVal chave As String) As Long
    On Error GoTo ErrorHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True: regex.Global = False
    regex.Pattern = """" & chave & """\s*:\s*\[\s*\d+\s*,\s*(\d+)\s*\]"
    If regex.Test(json) Then AI_ExtrairArrayUltimo = CLng(regex.Execute(json)(0).SubMatches(0)): Exit Function

' =============================================================================
' CARREGAR CHAVE API (DPAPI) - mesma logica de Mod11RevisionText
' =============================================================================
Private Function AI_CarregarChaveAPI() As String
    On Error GoTo ErrorHandler
    Dim caminhoArquivo As String
    Dim bytArquivo() As Byte
    Dim ff As Integer, tamArquivo As Long
    Dim blobIn As AI_DATA_BLOB, blobOut As AI_DATA_BLOB, resultado As Long
    #If VBA7 Then
    Dim pMem As LongPtr
    #Else
    Dim pMem As Long
    #End If

    caminhoArquivo = GetZ7StdProposersDataPath() & "\openrouter.key"
    If Dir(caminhoArquivo) = "" Then
        LogMessage AI_STRUCT_PREFIX & ": Arquivo de chave nao encontrado", LOG_LEVEL_WARNING
        AI_CarregarChaveAPI = "": Exit Function
    End If

    ff = FreeFile
    Open caminhoArquivo For Binary Access Read As #ff
    tamArquivo = LOF(ff)
    If tamArquivo = 0 Then Close #ff: AI_CarregarChaveAPI = "": Exit Function
    ReDim bytArquivo(0 To tamArquivo - 1)
    Get #ff, , bytArquivo
    Close #ff

    blobIn.cbData = tamArquivo
    blobIn.pbData = VarPtr(bytArquivo(0))
    resultado = AI_CryptUnprotectData(blobIn, 0, 0, 0, 0, 0, blobOut)
    If resultado = 0 Then
        LogMessage AI_STRUCT_PREFIX & ": Falha DPAPI", LOG_LEVEL_ERROR
        AI_CarregarChaveAPI = "": Exit Function
    End If

    Dim chaveDecrypt() As Byte
    ReDim chaveDecrypt(0 To blobOut.cbData - 1)
    AI_CopyMemory chaveDecrypt(0), ByVal blobOut.pbData, blobOut.cbData
    AI_LocalFree blobOut.pbData
    AI_CarregarChaveAPI = AI_BytesParaStringUTF8(chaveDecrypt)

' =============================================================================
' CARREGAR MODELO IA
' =============================================================================
Private Function AI_CarregarModelo() As String
    On Error GoTo ErrorHandler
    Dim caminhoArquivo As String, ff As Integer, conteudo As String
    caminhoArquivo = GetZ7StdProposersDataPath() & "\selected_model.txt"
    If Dir(caminhoArquivo) <> "" Then
        ff = FreeFile
        Open caminhoArquivo For Input As #ff
        If Not EOF(ff) Then Line Input #ff, conteudo
        Close #ff
        conteudo = Trim(conteudo)
        If Len(conteudo) > 0 Then AI_CarregarModelo = conteudo: Exit Function
    End If
    AI_CarregarModelo = AI_STRUCT_DEFAULT_MODEL
    Exit Function
ErrorHandler: AI_CarregarModelo = AI_STRUCT_DEFAULT_MODEL
End Function

' =============================================================================
' ESCAPAR STRING PARA JSON
' =============================================================================
Private Function EscaparJSONAI(ByVal texto As String) As String
    On Error Resume Next
    Dim resultado As String, i As Long, ch As Long, cleanResult As String
    resultado = Replace(Replace(Replace(texto, "\", "\\"), """", "\"""), vbCrLf, "\n")
    resultado = Replace(Replace(resultado, vbCr, "\n"), vbLf, "\n")
    resultado = Replace(resultado, vbTab, "\t")
    For i = 1 To Len(resultado)
        ch = AscW(Mid(resultado, i, 1))
        If ch >= 32 Or ch < 0 Then cleanResult = cleanResult & Mid(resultado, i, 1)
    Next i

' =============================================================================
' CONVERSAO UTF-8
' =============================================================================
Private Function AI_StringParaUTF8(ByVal texto As String) As Byte()
    On Error GoTo ErrorHandler
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText texto: stream.Position = 0
    stream.Type = 1: stream.Position = 3  ' Skip BOM
    AI_StringParaUTF8 = stream.Read
    stream.Close: Set stream = Nothing
    Exit Function
ErrorHandler:
    Set stream = Nothing
End Function

Private Function AI_BytesParaStringUTF8(ByRef bytes() As Byte) As String
    On Error GoTo ErrorHandler
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open
    stream.Write bytes: stream.Position = 0
    stream.Type = 2: stream.Charset = "utf-8"
    AI_BytesParaStringUTF8 = stream.ReadText
    stream.Close: Set stream = Nothing
    Exit Function
ErrorHandler:
    Set stream = Nothing

' =============================================================================
' EXTRAI CONTENT DO JSON DE RESPOSTA (OPENROUTER)
' =============================================================================
Private Function AI_ExtrairContentJSON(ByVal json As String) As String
    On Error GoTo ErrorHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = """content""\s*:\s*""((?:[^""\\]|\\.)*)"""
    regex.IgnoreCase = True: regex.Global = False
    Dim matches As Object
    Set matches = regex.Execute(json)
    If matches.count = 0 Then AI_ExtrairContentJSON = "": Exit Function
    AI_ExtrairContentJSON = Trim(AI_DesescaparJSON(matches(0).SubMatches(0)))
    Exit Function
ErrorHandler: AI_ExtrairContentJSON = ""
End Function

' =============================================================================
' DESESCAPA SEQUENCIAS JSON
' =============================================================================
Private Function AI_DesescaparJSON(ByVal texto As String) As String
    On Error Resume Next
    Dim i As Long, caractere As String, sequencia As String
    Dim resultado As String, codigo As String, numero As Long
    resultado = "": i = 1
    Do While i <= Len(texto)
        caractere = Mid(texto, i, 1)
        If caractere = "\" And i < Len(texto) Then
            sequencia = Mid(texto, i + 1, 1)
            Select Case sequencia
                Case """": resultado = resultado & """": i = i + 2
                Case "\": resultado = resultado & "\": i = i + 2
                Case "/": resultado = resultado & "/": i = i + 2
                Case "n": resultado = resultado & vbCrLf: i = i + 2
                Case "r": i = i + 2
                Case "t": resultado = resultado & vbTab: i = i + 2
                Case "u"
                    If i + 5 <= Len(texto) Then
                        codigo = Mid(texto, i + 2, 4)
                        On Error Resume Next: numero = CLng("&H" & codigo): On Error GoTo 0
                        If numero > 0 Then resultado = resultado & ChrW(numero): i = i + 6 Else resultado = resultado & caractere: i = i + 1
                    Else: resultado = resultado & caractere: i = i + 1
                    End If
                Case Else: resultado = resultado & sequencia: i = i + 2
            End Select
        Else
            resultado = resultado & caractere: i = i + 1
        End If
    Loop
    AI_DesescaparJSON = resultado
End Function

    AI_BytesParaStringUTF8 = ""
End Function

    EscaparJSONAI = cleanResult
End Function

    Exit Function
ErrorHandler:
    LogMessage AI_STRUCT_PREFIX & ": Erro ao carregar chave: " & Err.Description, LOG_LEVEL_ERROR
    AI_CarregarChaveAPI = ""
End Function

    regex.Pattern = """" & chave & """\s*:\s*\[\s*(\d+)\s*\]"
    If regex.Test(json) Then AI_ExtrairArrayUltimo = CLng(regex.Execute(json)(0).SubMatches(0)): Exit Function
    AI_ExtrairArrayUltimo = 0
    Exit Function
ErrorHandler: AI_ExtrairArrayUltimo = 0
End Function

    Exit Sub
ErrorHandler:
    LogMessage AI_STRUCT_PREFIX & ": Erro ao marcar flags: " & Err.Description, LOG_LEVEL_ERROR
End Sub

    Exit Function
ErrorHandler:
    ValidarIndicesEstrutura = False
End Function


    ParsearRespostaEstruturaIA = True
    Exit Function
ErrorHandler:
    LogMessage AI_STRUCT_PREFIX & ": Erro ao parsear: " & Err.Description, LOG_LEVEL_ERROR
    ParsearRespostaEstruturaIA = False
End Function

    Set http = Nothing
    LogMessage AI_STRUCT_PREFIX & ": Erro HTTP: " & Err.Description, LOG_LEVEL_ERROR
    AI_ChamarAPI = ""
End Function

End Function

    On Error GoTo ErrorHandler

    Dim sb As String
    Dim i As Long
    Dim paraText As String
    Dim paraCount As Long

    paraCount = doc.Paragraphs.count
    If paraCount > MAX_PARAGRAPHS_FOR_AI Then paraCount = MAX_PARAGRAPHS_FOR_AI

    For i = 1 To paraCount
        On Error Resume Next
        paraText = doc.Paragraphs(i).Range.text
        On Error GoTo ErrorHandler
        paraText = Replace(Replace(paraText, vbCr, ""), vbLf, "")
        If Len(paraText) > MAX_PARAGRAPH_TEXT_LENGTH Then
            paraText = Left(paraText, MAX_PARAGRAPH_TEXT_LENGTH) & "..."
        End If
        sb = sb & "[P" & i & "] " & paraText & vbCrLf
    Next i

    MontarTextoDocumentoParaIA = sb
    Exit Function
ErrorHandler:
    MontarTextoDocumentoParaIA = ""
End Function


