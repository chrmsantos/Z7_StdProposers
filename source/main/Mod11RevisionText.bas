Attribute VB_Name = "Mod11RevisionText"
Option Explicit

' Mod11RevisionText.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
' REVISAO DE TEXTOS DO WORD COM IA - OPENROUTER
'
' A IA revisa SOMENTE O CONTEUDO DO TEXTO.
' O VBA e responsavel por preservar a formatacao do Word.
'
' Entradas publicas:
'   - TestarRevisaoTextoSelecionado: revisa texto selecionado
'   - CorrigirProposituraComIA:     revisa documento inteiro
'   - DiagnosticarOpenRouter:       diagnostico de conectividade com a API
'
' Arquivos externos esperados (via config_prompt.py / z7_api_key.py):
'   - %USERPROFILE%\AppData\Local\Z7\Tmp\StdProposers\openrouter.key
'   - %USERPROFILE%\AppData\Local\Z7\Tmp\StdProposers\selected_model.txt
' =============================================================================

' =============================================================================
' CONSTANTES
' =============================================================================
Private Const OPENROUTER_URL As String = _
    "https://openrouter.ai/api/v1/chat/completions"

Private Const MODELO_IA_DEFAULT As String = _
    "deepseek/deepseek-v4-pro"

' Timeouts em milissegundos (resolve, connect, send, receive)
Private Const HTTP_RESOLVE_TIMEOUT_MS As Long = 5000
Private Const HTTP_CONNECT_TIMEOUT_MS As Long = 10000
Private Const HTTP_SEND_TIMEOUT_MS As Long = 30000
Private Const HTTP_RECEIVE_TIMEOUT_MS As Long = 120000

' Tamanho minimo de paragrafo para enviar a IA (evita chamadas desnecessarias)
Private Const PARAGRAPH_MIN_LENGTH As Long = 5

' Prefixo do log para esta operacao
Private Const LOG_PREFIX As String = "REVISAO_IA"

' =============================================================================
' DECLARACOES DA API WINDOWS (DPAPI)
' =============================================================================
#If VBA7 Then
Private Type DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type
Private Declare PtrSafe Function CryptUnprotectData _
    Lib "crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, _
    ByVal ppszDataDescr As LongPtr, _
    ByVal pOptionalEntropy As LongPtr, _
    ByVal pvReserved As LongPtr, _
    ByVal pPromptStruct As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB _
    ) As Long
Private Declare PtrSafe Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As LongPtr)
Private Declare PtrSafe Function LocalFree _
    Lib "kernel32" ( _
    ByVal hMem As LongPtr _
    ) As LongPtr
#Else
Private Type DATA_BLOB
    cbData As Long
    pbData As Long
End Type
Private Declare Function CryptUnprotectData _
    Lib "crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, _
    ByVal ppszDataDescr As Long, _
    ByVal pOptionalEntropy As Long, _
    ByVal pvReserved As Long, _
    ByVal pPromptStruct As Long, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB _
    ) As Long
Private Declare Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _

' =============================================================================
' CARREGAR MODELO IA DO CONFIG_PROMPT
' =============================================================================
' Le selected_model.txt gravado pelo config_prompt.py.
' Retorna MODELO_IA_DEFAULT se o arquivo nao existir.
' =============================================================================
Private Function CarregarModeloIA() As String
    On Error GoTo ErrorHandler

    Dim caminhoArquivo As String
    Dim ff As Integer
    Dim conteudo As String

    caminhoArquivo = GetZ7StdProposersDataPath() & _
        "\selected_model.txt"

    If Dir(caminhoArquivo) <> "" Then
        ff = FreeFile
        Open caminhoArquivo For Input As #ff
        If Not EOF(ff) Then
            Line Input #ff, conteudo
        End If
        Close #ff
        conteudo = Trim(conteudo)
        If Len(conteudo) > 0 Then
            CarregarModeloIA = conteudo
            LogMessage LOG_PREFIX & ": Modelo IA carregado: " & conteudo, LOG_LEVEL_INFO
            Exit Function
        End If
    End If

    CarregarModeloIA = MODELO_IA_DEFAULT
    LogMessage LOG_PREFIX & ": Modelo IA padrao: " & MODELO_IA_DEFAULT, LOG_LEVEL_INFO
    Exit Function

ErrorHandler:
    LogMessage LOG_PREFIX & ": Erro ao carregar modelo IA: " & Err.Description, LOG_LEVEL_WARNING
    CarregarModeloIA = MODELO_IA_DEFAULT
End Function

' =============================================================================
' CARREGAR CHAVE API DO CONFIG_PROMPT (DPAPI)
' =============================================================================
' Le e descriptografa openrouter.key gravado pelo z7_api_key.py.
' Retorna string vazia se o arquivo nao existir ou falhar.
' =============================================================================
Private Function CarregarChaveAPI() As String
    On Error GoTo ErrorHandler

    Dim caminhoArquivo As String
    Dim bytArquivo() As Byte
    Dim ff As Integer
    Dim tamArquivo As Long
    Dim blobIn As DATA_BLOB
    Dim blobOut As DATA_BLOB
    Dim resultado As Long
    Dim chaveDecrypt As String

    #If VBA7 Then
    Dim pMem As LongPtr
    #Else
    Dim pMem As Long
    #End If

    caminhoArquivo = GetZ7StdProposersDataPath() & _
        "\openrouter.key"

    If Dir(caminhoArquivo) = "" Then
        LogMessage LOG_PREFIX & ": Arquivo de chave nao encontrado", LOG_LEVEL_WARNING
        CarregarChaveAPI = ""
        Exit Function
    End If

    ' Le o arquivo binario completo
    ff = FreeFile
    Open caminhoArquivo For Binary Access Read As #ff
    tamArquivo = LOF(ff)
    If tamArquivo = 0 Then
        Close #ff
        LogMessage LOG_PREFIX & ": Arquivo de chave vazio", LOG_LEVEL_WARNING
        CarregarChaveAPI = ""
        Exit Function
    End If
    ReDim bytArquivo(0 To tamArquivo - 1)
    Get #ff, , bytArquivo
    Close #ff

    ' Preenche o blob de entrada
    blobIn.cbData = tamArquivo
    blobIn.pbData = VarPtr(bytArquivo(0))

    ' Descriptografa com DPAPI
    resultado = CryptUnprotectData( _
        blobIn, 0, 0, 0, 0, 0, blobOut)

    If resultado = 0 Then
        LogMessage LOG_PREFIX & ": Falha ao descriptografar chave API (DPAPI)", LOG_LEVEL_ERROR
        CarregarChaveAPI = ""
        Exit Function
    End If

    ' Copia o resultado para string
    If blobOut.cbData > 0 Then
        Dim bytResultado() As Byte
        ReDim bytResultado(0 To blobOut.cbData - 1)
        CopyMemory bytResultado(0), _
            ByVal blobOut.pbData, blobOut.cbData
        chaveDecrypt = _
            StrConv(bytResultado, vbUnicode)
        LocalFree blobOut.pbData
    End If

    CarregarChaveAPI = Trim(chaveDecrypt)
    LogMessage LOG_PREFIX & ": Chave API carregada com sucesso", LOG_LEVEL_INFO
    Exit Function

ErrorHandler:
    LogMessage LOG_PREFIX & ": Erro ao carregar chave API: " & Err.Description, LOG_LEVEL_ERROR
    CarregarChaveAPI = ""
End Function

    ByVal Length As Long)
Private Declare Function LocalFree _
    Lib "kernel32" ( _
    ByVal hMem As Long _
    ) As Long
#End If

' =============================================================================
' VALIDAR CHAVE API (ANTES DA PRIMEIRA CHAMADA)
' =============================================================================
' Retorna True se a chave existir.
' Exibe mensagem ao usuario se a chave nao estiver configurada.
' =============================================================================
Private Function ValidarChaveAPI(ByVal apiKey As String) As Boolean
    If Len(Trim(apiKey)) = 0 Then
        MsgBox _
            "A chave do OpenRouter nao foi configurada." & _
            vbCrLf & vbCrLf & _
            "Configure-a na interface config_prompt.", _
            vbCritical, _
            "Chave nao configurada"
        LogMessage LOG_PREFIX & ": Chave API nao configurada", LOG_LEVEL_ERROR
        ValidarChaveAPI = False
        Exit Function
    End If
    ValidarChaveAPI = True
End Function

' =============================================================================
' 1 - REVISAR TEXTO SELECIONADO
' =============================================================================
'
' Selecione qualquer texto no Word e execute esta macro.
' Ela revisa e substitui o texto automaticamente.
' A formatacao do Word e preservada.
'
' =============================================================================
Public Sub TestarRevisaoTextoSelecionado()
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim textoOriginal As String
    Dim textoCorrigido As String
    Dim startTime As Double

    startTime = Timer

    ' -----------------------------------------------------------------
    ' VALIDACOES INICIAIS
    ' -----------------------------------------------------------------
    If Selection Is Nothing Then
        LogMessage LOG_PREFIX & ": Nenhuma selecao ativa", LOG_LEVEL_WARNING
        Exit Sub
    End If

    Set rng = Selection.Range.Duplicate
    textoOriginal = ExtrairTextoLimpo(rng)

    If Len(Trim(textoOriginal)) = 0 Then
        LogMessage LOG_PREFIX & ": Selecao vazia", LOG_LEVEL_WARNING
        Exit Sub
    End If

    LogSection "REVISAO IA - TEXTO SELECIONADO"
    LogStepStart "Revisao de texto selecionado"
    LogMetric "Caracteres selecionados", Len(textoOriginal)

    ' -----------------------------------------------------------------
    ' ENVIA PARA A IA
    ' -----------------------------------------------------------------
    Application.StatusBar = _
        LOG_PREFIX & ": Enviando texto selecionado para a IA..."
    DoEvents

    textoCorrigido = ProcessarTextoComIA(textoOriginal)
    Application.StatusBar = False

    If Len(Trim(textoCorrigido)) = 0 Then
        LogStepSkipped "Revisao de texto selecionado", "IA retornou resposta vazia"
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' SUBSTITUI SE HOUVE ALTERACAO
    ' -----------------------------------------------------------------
    If NormalizarComparacao(textoCorrigido) <> _
       NormalizarComparacao(textoOriginal) Then
        SubstituirTextoPreservandoFormatacao _
            rng, _
            textoCorrigido
        LogStepComplete "Revisao de texto selecionado", _
            "Texto substituido | " & _
            Format(Timer - startTime, "0.0") & "s"
    Else
        LogStepComplete "Revisao de texto selecionado", _
            "Nenhuma alteracao | " & _
            Format(Timer - startTime, "0.0") & "s"
    End If

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    LogMessage LOG_PREFIX & ": Erro na revisao: " & _
        Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    MsgBox _
        "Erro durante a revisao:" & vbCrLf & vbCrLf & _
        "Numero: " & Err.Number & vbCrLf & _
        "Descricao: " & Err.Description, _
        vbCritical, _
        "Erro"
End Sub



' =============================================================================
' 2 - REVISAR DOCUMENTO INTEIRO
' =============================================================================
Public Sub CorrigirProposituraComIA()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Dim totalParas As Long
    Dim i As Long
    Dim rng As Range
    Dim textoOriginal As String
    Dim textoCorrigido As String
    Dim alterados As Long
    Dim analisados As Long
    Dim ignorados As Long
    Dim startTime As Double
    Dim chaveAPI As String

    startTime = Timer

    ' -----------------------------------------------------------------
    ' VALIDACOES INICIAIS
    ' -----------------------------------------------------------------
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento aberto"
        LogMessage LOG_PREFIX & ": Nenhum documento aberto", LOG_LEVEL_ERROR
        MsgBox "Nenhum documento esta aberto para revisao.", _
            vbCritical, "Erro"
        Exit Sub
    End If

    totalParas = doc.Paragraphs.count
    If totalParas = 0 Then
        LogMessage LOG_PREFIX & ": Documento vazio", LOG_LEVEL_WARNING
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' VALIDA CHAVE ANTES DE INICIAR (EVITA TRAVAMENTO NO LOOP)
    ' -----------------------------------------------------------------
    chaveAPI = CarregarChaveAPI()
    If Not ValidarChaveAPI(chaveAPI) Then Exit Sub

    ' -----------------------------------------------------------------
    ' INICIALIZA LOGGING E PROGRESSO
    ' -----------------------------------------------------------------
    LogSection "REVISAO IA - DOCUMENTO INTEIRO"
    LogStepStart "Revisao do documento com IA"
    LogMetric "Total de paragrafos", totalParas
    LogMetric "Modelo IA", CarregarModeloIA()

    InitializeProgress totalParas + 1

    Application.ScreenUpdating = False
    Application.StatusBar = _
        LOG_PREFIX & ": Iniciando revisao com IA..."
    DoEvents

    ' -----------------------------------------------------------------
    ' PERCORRE OS PARAGRAFOS
    ' -----------------------------------------------------------------
    For i = 1 To totalParas

        ' Verifica cancelamento pelo usuario
        If formattingCancelled Then
            LogMessage LOG_PREFIX & ": Revisao cancelada no paragrafo " & i, LOG_LEVEL_WARNING
            GoTo CleanUp
        End If

        IncrementProgress "Revisando paragrafo " & i & " de " & totalParas & _
            " | Alterados: " & alterados

        Set rng = doc.Paragraphs(i).Range.Duplicate

        ' --------------------------------------------------------
        ' NAO INCLUI A MARCA DE PARAGRAFO
        ' --------------------------------------------------------
        If rng.End > rng.Start Then
            rng.End = rng.End - 1
        End If

        textoOriginal = ExtrairTextoLimpo(rng)

        ' --------------------------------------------------------
        ' IGNORA PARAGRAFOS VAZIOS OU MUITO CURTOS
        ' --------------------------------------------------------
        If Len(Trim(textoOriginal)) <= PARAGRAPH_MIN_LENGTH Then
            ignorados = ignorados + 1
        Else
            analisados = analisados + 1

            ' ----------------------------------------------------
            ' ENVIA O TEXTO REAL PARA A IA
            ' ----------------------------------------------------
            textoCorrigido = _
                ProcessarTextoComIA(textoOriginal)

            ' ----------------------------------------------------
            ' SE HOUVE ALTERACAO
            ' ----------------------------------------------------
            If Len(Trim(textoCorrigido)) > 0 Then
                If NormalizarComparacao(textoCorrigido) <> _
                   NormalizarComparacao(textoOriginal) Then
                    SubstituirTextoPreservandoFormatacao _
                        rng, _
                        textoCorrigido
                    alterados = alterados + 1
                End If
            End If
        End If
    Next i

CleanUp:
    ' -----------------------------------------------------------------
    ' FINALIZACAO
    ' -----------------------------------------------------------------
    Application.StatusBar = False
    Application.ScreenUpdating = True

    LogMetric "Paragrafos analisados", analisados
    LogMetric "Paragrafos ignorados", ignorados
    LogMetric "Paragrafos alterados", alterados
    LogStepComplete "Revisao do documento com IA", _
        alterados & " de " & analisados & " alterados | " & _
        Format(Timer - startTime, "0.0") & "s"

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    LogMessage LOG_PREFIX & ": Erro na revisao do documento: " & _
        Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    MsgBox _
        "Ocorreu um erro durante a revisao." & vbCrLf & vbCrLf & _
        "Numero: " & Err.Number & vbCrLf & _
        "Descricao: " & Err.Description, _
        vbCritical, _
        "Erro"
End Sub

' =============================================================================
' PROCESSAR TEXTO COM IA
' =============================================================================
Private Function ProcessarTextoComIA( _
    ByVal textoInput As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim promptSystem As String
    Dim textoJSON As String
    Dim systemJSON As String
    Dim jsonPayload As String
    Dim resposta As String
    Dim conteudo As String
    Dim apiKey As String
    Dim modeloIA As String

    ' -----------------------------------------------------------------
    ' CARREGA CHAVE E MODELO DO CONFIG_PROMPT
    ' -----------------------------------------------------------------
    apiKey = CarregarChaveAPI()
    If Len(Trim(apiKey)) = 0 Then
        ProcessarTextoComIA = ""
        Exit Function
    End If

    modeloIA = CarregarModeloIA()

    ' -----------------------------------------------------------------
    ' PROMPT E JSON
    ' -----------------------------------------------------------------
    promptSystem = MontarPromptRevisao()
    textoJSON = EscaparJSON(textoInput)
    systemJSON = EscaparJSON(promptSystem)
    jsonPayload = MontarJSONRequest(modeloIA, systemJSON, textoJSON)

    ' -----------------------------------------------------------------
    ' HTTP COM TIMEOUTS
    ' -----------------------------------------------------------------
    Set http = CreateObject( _
        "MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts _
        HTTP_RESOLVE_TIMEOUT_MS, _
        HTTP_CONNECT_TIMEOUT_MS, _
        HTTP_SEND_TIMEOUT_MS, _
        HTTP_RECEIVE_TIMEOUT_MS

    http.Open "POST", OPENROUTER_URL, False
    http.setRequestHeader "Content-Type", _
        "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", _
        "Bearer " & apiKey
    http.setRequestHeader "HTTP-Referer", _
        "https://localhost"
    http.setRequestHeader "X-Title", _
        "Word - Revisao de Textos"

    ' -----------------------------------------------------------------
    ' ENVIA E RECEBE
    ' -----------------------------------------------------------------
    http.send StringParaUTF8(jsonPayload)

    If http.Status = 200 Then
        resposta = BytesParaStringUTF8(http.ResponseBody)
        conteudo = ExtrairContentJSON(resposta)
        ProcessarTextoComIA = LimparRespostaIA(conteudo)
    Else
        Dim respostaErro As String
        respostaErro = BytesParaStringUTF8(http.ResponseBody)
        LogMessage LOG_PREFIX & ": HTTP " & http.Status & " - " & _
            Left(respostaErro, 200), LOG_LEVEL_ERROR
        MsgBox _
            "A OpenRouter retornou um erro." & _
            vbCrLf & vbCrLf & _
            "Codigo HTTP: " & http.Status & _
            vbCrLf & vbCrLf & _
            "Resposta:" & vbCrLf & _
            Left(respostaErro, 500), _
            vbCritical, "Erro OpenRouter"
        ProcessarTextoComIA = ""
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    Set http = Nothing
    LogMessage LOG_PREFIX & ": Erro ao comunicar com a IA: " & _
        Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    MsgBox _
        "Erro ao comunicar com a IA:" & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, "Erro"
    ProcessarTextoComIA = ""
End Function

' =============================================================================
' MONTA JSON DO REQUEST
' =============================================================================
Private Function MontarJSONRequest( _
    ByVal modeloIA As String, _
    ByVal systemJSON As String, _
    ByVal textoJSON As String) As String

    Dim json As String

    json = "{"
    json = json & """model"":""" & modeloIA & ""","
    json = json & """temperature"":0.2,"
    json = json & """messages"":["
    json = json & "{""role"":""system"",""content"":"""
    json = json & systemJSON
    json = json & """},"
    json = json & "{""role"":""user"",""content"":"""
    json = json & textoJSON
    json = json & """}"
    json = json & "]"
    json = json & "}"

    MontarJSONRequest = json
End Function


' =============================================================================
' PROMPT DA IA
' =============================================================================
Private Function MontarPromptRevisao() As String
    Dim p As String

    p = "Voce e um revisor especialista em lingua portuguesa, " & _
        "redacao oficial e tecnica legislativa."
    p = p & vbLf & vbLf & _
        "O texto recebido pode tratar de QUALQUER ASSUNTO."
    p = p & vbLf & vbLf & _
        "Revise cuidadosamente o texto recebido."
    p = p & vbLf & vbLf & _
        "CORRIJA erros de ortografia, acentuacao, digitacao, " & _
        "concordancia, regencia e pontuacao."
    p = p & vbLf & _
        "Corrija palavras grudadas, palavras incompletas e erros " & _
        "claros de digitacao."
    p = p & vbLf & _
        "Identifique frases confusas, truncadas ou mal construidas."
    p = p & vbLf & _
        "Quando uma frase estiver sem sentido ou mal construida, " & _
        "reescreva-a de forma clara, natural e coerente."
    p = p & vbLf & vbLf & _
        "A REESCRITA deve preservar o significado original."
    p = p & vbLf & vbLf & _
        "Use portugues correto, claro, natural e formal."
    p = p & vbLf & _
        "Quando o texto for destinado a documento publico ou " & _
        "legislativo, utilize linguagem adequada e redacao oficial."
    p = p & vbLf & vbLf & _
        "NAO invente informacoes."
    p = p & vbLf & _
        "NAO invente fatos."
    p = p & vbLf & _
        "NAO invente nomes, datas, locais ou numeros."
    p = p & vbLf & _
        "NAO acrescente argumentos que nao estejam no texto."
    p = p & vbLf & _
        "NAO altere a intencao do autor."
    p = p & vbLf & vbLf & _
        "Preserve nomes proprios, orgaos publicos, cargos, " & _
        "numeros, datas, locais e referencias legais."
    p = p & vbLf & vbLf & _
        "Se uma frase estiver correta, nao altere desnecessariamente."
    p = p & vbLf & _
        "Se houver erro, corrija."
    p = p & vbLf & _
        "Se houver trecho confuso, reescreva."
    p = p & vbLf & _
        "Se houver trecho sem sentido, reconstrua utilizando " & _
        "somente as informacoes existentes no proprio texto."
    p = p & vbLf & vbLf & _
        "IMPORTANTE: voce esta revisando SOMENTE O TEXTO."
    p = p & vbLf & _
        "Nao forneca instrucoes de formatacao."
    p = p & vbLf & _
        "Nao adicione negrito, titulos, marcadores ou observacoes."
    p = p & vbLf & vbLf & _
        "RETORNE SOMENTE O TEXTO FINAL REVISADO."
    p = p & vbLf & _
        "Nao explique as alteracoes."
    p = p & vbLf & _
        "Nao faca comentarios."
    p = p & vbLf & _
        "Nao coloque aspas no inicio ou no final."

    MontarPromptRevisao = p
End Function

' =============================================================================
' EXTRAI TEXTO DO RANGE
' =============================================================================
Private Function ExtrairTextoLimpo( _
    ByVal rng As Range) As String
    On Error Resume Next

    Dim texto As String
    texto = rng.text
    texto = Replace(texto, vbCr, "")
    texto = Replace(texto, vbLf, "")
    texto = Replace(texto, Chr(7), "")    ' BEL character (form fields)
    texto = Replace(texto, Chr(160), " ") ' Non-breaking space

    ExtrairTextoLimpo = Trim(texto)
End Function


' =============================================================================
' SUBSTITUI TEXTO PRESERVANDO FORMATACAO
' =============================================================================
'
' Salva a formatacao de CADA caractere individualmente, agrupa em
' "runs" (trechos com formatacao identica) e aplica por run atraves
' de mapeamento proporcional de posicoes.
'
' Propriedades salvas por caractere:
'   Font.Name, Font.Size, Font.Bold, Font.Italic,
'   Font.Underline, Font.Color, Font.HighlightColorIndex,
'   Font.StrikeThrough, Font.SmallCaps, Font.AllCaps,
'   Font.Superscript, Font.Subscript, Font.Spacing,
'   Font.Position
'
' Propriedades de paragrafo (valor unico):
'   Style, Alignment, SpaceBefore, SpaceAfter,
'   LineSpacing, LeftIndent, RightIndent, FirstLineIndent
'
' =============================================================================
Private Sub SubstituirTextoPreservandoFormatacao( _
    ByVal rngOriginal As Range, _
    ByVal novoTexto As String)
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim i As Long
    Dim lenOld As Long
    Dim lenNew As Long
    Dim charRng As Range
    Dim runRng As Range

    Dim arrFontName() As String
    Dim arrFontSize() As Single
    Dim arrBold() As Long
    Dim arrItalic() As Long
    Dim arrUnderline() As Long
    Dim arrColor() As Long
    Dim arrHighlight() As Long
    Dim arrStrikeThrough() As Long
    Dim arrSmallCaps() As Long
    Dim arrAllCaps() As Long
    Dim arrSuperscript() As Long
    Dim arrSubscript() As Long
    Dim arrSpacing() As Single
    Dim arrPosition() As Single

    Dim runCount As Long
    Dim runEndPos() As Long

    Dim estilo As Variant
    Dim alinhamento As WdParagraphAlignment
    Dim espAntes As Single
    Dim espDepois As Single
    Dim espLinha As Single
    Dim recuoEsquerdo As Single
    Dim recuoDireito As Single
    Dim primeiraLinha As Single

    Dim oldStart As Long
    Dim oldEnd As Long
    Dim newStart As Long
    Dim newEnd As Long

    ' -----------------------------------------------------------------
    ' DUPLICA RANGE
    ' -----------------------------------------------------------------
    Set rng = rngOriginal.Duplicate
    lenOld = rng.End - rng.Start
    If lenOld <= 0 Then Exit Sub

    ' -----------------------------------------------------------------
    ' SALVA FORMATACAO DO PARAGRAFO
    ' -----------------------------------------------------------------
    On Error Resume Next
    estilo = rng.Style
    alinhamento = rng.ParagraphFormat.alignment
    espAntes = rng.ParagraphFormat.SpaceBefore
    espDepois = rng.ParagraphFormat.SpaceAfter
    espLinha = rng.ParagraphFormat.LineSpacing
    recuoEsquerdo = rng.ParagraphFormat.leftIndent
    recuoDireito = rng.ParagraphFormat.RightIndent
    primeiraLinha = rng.ParagraphFormat.firstLineIndent
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------
    ' SALVA FORMATACAO DE CADA CARACTERE
    ' -----------------------------------------------------------------
    ReDim arrFontName(1 To lenOld)
    ReDim arrFontSize(1 To lenOld)
    ReDim arrBold(1 To lenOld)
    ReDim arrItalic(1 To lenOld)
    ReDim arrUnderline(1 To lenOld)
    ReDim arrColor(1 To lenOld)
    ReDim arrHighlight(1 To lenOld)
    ReDim arrStrikeThrough(1 To lenOld)
    ReDim arrSmallCaps(1 To lenOld)
    ReDim arrAllCaps(1 To lenOld)
    ReDim arrSuperscript(1 To lenOld)
    ReDim arrSubscript(1 To lenOld)
    ReDim arrSpacing(1 To lenOld)
    ReDim arrPosition(1 To lenOld)

    On Error Resume Next
    For i = 1 To lenOld
        Set charRng = rng.Duplicate
        charRng.Start = rng.Start + i - 1
        charRng.End = rng.Start + i
        arrFontName(i) = charRng.Font.Name
        arrFontSize(i) = charRng.Font.size
        arrBold(i) = charRng.Font.Bold
        arrItalic(i) = charRng.Font.Italic
        arrUnderline(i) = charRng.Font.Underline
        arrColor(i) = charRng.Font.Color
        arrHighlight(i) = charRng.Font.HighlightColorIndex
        arrStrikeThrough(i) = charRng.Font.StrikeThrough
        arrSmallCaps(i) = charRng.Font.SmallCaps
        arrAllCaps(i) = charRng.Font.AllCaps
        arrSuperscript(i) = charRng.Font.Superscript
        arrSubscript(i) = charRng.Font.Subscript
        arrSpacing(i) = charRng.Font.Spacing
        arrPosition(i) = charRng.Font.Position
    Next i
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------
    ' CONSTRUI RUNS DE FORMATACAO
    ' -----------------------------------------------------------------
    runCount = 1
    ReDim runEndPos(1 To lenOld)

    For i = 2 To lenOld
        If arrFontName(i) <> arrFontName(i - 1) Or _
           arrFontSize(i) <> arrFontSize(i - 1) Or _
           arrBold(i) <> arrBold(i - 1) Or _
           arrItalic(i) <> arrItalic(i - 1) Or _
           arrUnderline(i) <> arrUnderline(i - 1) Or _
           arrColor(i) <> arrColor(i - 1) Or _
           arrHighlight(i) <> arrHighlight(i - 1) Or _
           arrStrikeThrough(i) <> arrStrikeThrough(i - 1) Or _
           arrSmallCaps(i) <> arrSmallCaps(i - 1) Or _
           arrAllCaps(i) <> arrAllCaps(i - 1) Or _
           arrSuperscript(i) <> arrSuperscript(i - 1) Or _
           arrSubscript(i) <> arrSubscript(i - 1) Or _
           arrSpacing(i) <> arrSpacing(i - 1) Or _
           arrPosition(i) <> arrPosition(i - 1) Then
            runEndPos(runCount) = i - 1
            runCount = runCount + 1
        End If
    Next i
    runEndPos(runCount) = lenOld


    ' -----------------------------------------------------------------
    ' SUBSTITUI O TEXTO
    ' -----------------------------------------------------------------
    rng.text = novoTexto
    lenNew = rng.End - rng.Start
    If lenNew <= 0 Then Exit Sub

    ' -----------------------------------------------------------------
    ' RESTAURA FORMATACAO POR RUN (mapeamento proporcional)
    ' -----------------------------------------------------------------
    On Error Resume Next
    For i = 1 To runCount
        If i = 1 Then
            oldStart = 1
        Else
            oldStart = runEndPos(i - 1) + 1
        End If
        oldEnd = runEndPos(i)

        newStart = CLng((oldStart - 1) * lenNew / lenOld) + 1
        newEnd = CLng(oldEnd * lenNew / lenOld)
        If newStart < 1 Then newStart = 1
        If newEnd > lenNew Then newEnd = lenNew
        If newEnd < newStart Then newEnd = newStart

        Set runRng = rng.Duplicate
        runRng.Start = rng.Start + newStart - 1
        runRng.End = rng.Start + newEnd

        runRng.Font.Name = arrFontName(oldStart)
        runRng.Font.size = arrFontSize(oldStart)
        runRng.Font.Bold = arrBold(oldStart)
        runRng.Font.Italic = arrItalic(oldStart)
        runRng.Font.Underline = arrUnderline(oldStart)
        runRng.Font.Color = arrColor(oldStart)
        runRng.Font.HighlightColorIndex = arrHighlight(oldStart)
        runRng.Font.StrikeThrough = arrStrikeThrough(oldStart)
        runRng.Font.SmallCaps = arrSmallCaps(oldStart)
        runRng.Font.AllCaps = arrAllCaps(oldStart)
        runRng.Font.Superscript = arrSuperscript(oldStart)
        runRng.Font.Subscript = arrSubscript(oldStart)
        runRng.Font.Spacing = arrSpacing(oldStart)
        runRng.Font.Position = arrPosition(oldStart)
    Next i
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------
    ' RESTAURA FORMATACAO DO PARAGRAFO
    ' -----------------------------------------------------------------
    On Error Resume Next
    rng.ParagraphFormat.alignment = alinhamento
    rng.ParagraphFormat.SpaceBefore = espAntes
    rng.ParagraphFormat.SpaceAfter = espDepois
    rng.ParagraphFormat.LineSpacing = espLinha
    rng.ParagraphFormat.leftIndent = recuoEsquerdo
    rng.ParagraphFormat.RightIndent = recuoDireito
    rng.ParagraphFormat.firstLineIndent = primeiraLinha
    If Not IsEmpty(estilo) Then
        rng.Style = estilo
    End If
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    ' NAO INTERROMPE A REVISAO POR CAUSA DE FORMATACAO
    LogMessage LOG_PREFIX & ": Aviso - falha ao preservar formatacao: " & _
        Err.Description, LOG_LEVEL_WARNING
End Sub


' =============================================================================
' NORMALIZA PARA COMPARACAO
' =============================================================================
Private Function NormalizarComparacao( _
    ByVal texto As String) As String
    texto = Replace(texto, vbCr, "")
    texto = Replace(texto, vbLf, "")
    texto = Replace(texto, Chr(160), " ")
    NormalizarComparacao = Trim(texto)
End Function

' =============================================================================
' LIMPA RESPOSTA DA IA
' =============================================================================
Private Function LimparRespostaIA( _
    ByVal texto As String) As String
    On Error Resume Next

    texto = Trim(texto)

    ' Remove blocos de markdown
    If Left(texto, 3) = "```" Then
        Dim posFim As Long
        posFim = InStr(texto, vbLf)
        If posFim > 0 Then
            texto = Mid(texto, posFim + 1)
        Else
            texto = Replace(texto, "```text", "")
            texto = Replace(texto, "```", "")
        End If
        Dim posFechamento As Long
        posFechamento = InStrRev(texto, "```")
        If posFechamento > 0 Then
            texto = Left(texto, posFechamento - 1)
        End If
        texto = Trim(texto)
    End If

    ' Remove aspas envoltorias
    If Len(texto) >= 2 Then
        If Left(texto, 1) = """" And Right(texto, 1) = """" Then
            texto = Mid(texto, 2, Len(texto) - 2)
        End If
    End If

    LimparRespostaIA = Trim(texto)
End Function

' =============================================================================
' ESCAPA JSON
' =============================================================================
Private Function EscaparJSON( _
    ByVal texto As String) As String
    On Error Resume Next

    texto = Replace(texto, "\", "\\")
    texto = Replace(texto, """", "\""")
    texto = Replace(texto, vbCrLf, "\n")
    texto = Replace(texto, vbCr, "\n")
    texto = Replace(texto, vbLf, "\n")
    texto = Replace(texto, vbTab, "\t")

    EscaparJSON = texto
End Function

' =============================================================================
' STRING PARA UTF-8
' =============================================================================
Private Function StringParaUTF8( _
    ByVal texto As String) As Variant
    On Error GoTo ErrorHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText texto
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    StringParaUTF8 = stream.Read
    stream.Close
    Set stream = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    LogMessage LOG_PREFIX & ": Erro ao converter string para UTF-8", LOG_LEVEL_ERROR
    StringParaUTF8 = Empty
End Function


' =============================================================================
' UTF-8 PARA STRING
' =============================================================================
Private Function BytesParaStringUTF8( _
    ByVal bytes As Variant) As String
    On Error GoTo ErrorHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    BytesParaStringUTF8 = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    LogMessage LOG_PREFIX & ": Erro ao converter bytes UTF-8", LOG_LEVEL_ERROR
    BytesParaStringUTF8 = ""
End Function

' =============================================================================
' EXTRAI CONTENT DO JSON
' =============================================================================
Private Function ExtrairContentJSON( _
    ByVal json As String) As String
    On Error GoTo ErrorHandler

    Dim reg As Object
    Dim matches As Object
    Dim resultado As String

    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = False
    reg.IgnoreCase = True
    reg.pattern = """content""\s*:\s*""((\\.|[^""\\])*)"""

    Set matches = reg.Execute(json)
    If matches.count = 0 Then
        LogMessage LOG_PREFIX & ": Nenhum campo 'content' na resposta JSON", LOG_LEVEL_WARNING
        ExtrairContentJSON = ""
        Exit Function
    End If

    resultado = matches(0).SubMatches(0)
    resultado = DesescaparJSON(resultado)

    ExtrairContentJSON = Trim(resultado)
    Exit Function

ErrorHandler:
    LogMessage LOG_PREFIX & ": Erro ao extrair content do JSON: " & Err.Description, LOG_LEVEL_ERROR
    ExtrairContentJSON = ""
End Function


' =============================================================================
' DESESCAPA JSON
' =============================================================================
Private Function DesescaparJSON( _
    ByVal texto As String) As String
    On Error Resume Next

    Dim i As Long
    Dim caractere As String
    Dim sequencia As String
    Dim resultado As String
    Dim codigo As String
    Dim numero As Long

    resultado = ""
    i = 1

    Do While i <= Len(texto)
        caractere = Mid(texto, i, 1)
        If caractere = "\" And i < Len(texto) Then
            sequencia = Mid(texto, i + 1, 1)
            Select Case sequencia
                Case """"
                    resultado = resultado & """"
                    i = i + 2
                Case "\"
                    resultado = resultado & "\"
                    i = i + 2
                Case "/"
                    resultado = resultado & "/"
                    i = i + 2
                Case "n"
                    resultado = resultado & vbCrLf
                    i = i + 2
                Case "r"
                    i = i + 2
                Case "t"
                    resultado = resultado & vbTab
                    i = i + 2
                Case "u"
                    If i + 5 <= Len(texto) Then
                        codigo = Mid(texto, i + 2, 4)
                        On Error Resume Next
                        numero = CLng("&H" & codigo)
                        On Error GoTo 0
                        If numero > 0 Then
                            resultado = resultado & ChrW(numero)
                            i = i + 6
                        Else
                            resultado = resultado & caractere
                            i = i + 1
                        End If
                    Else
                        resultado = resultado & caractere
                        i = i + 1
                    End If
                Case Else
                    resultado = resultado & sequencia
                    i = i + 2
            End Select
        Else
            resultado = resultado & caractere
            i = i + 1
        End If
    Loop

    DesescaparJSON = resultado
End Function

' =============================================================================
' DIAGNOSTICO DA CHAVE OPENROUTER
' =============================================================================
Public Sub DiagnosticarOpenRouter()
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim resposta As String
    Dim chaveAPI As String

    LogSection "DIAGNOSTICO OPENROUTER"
    LogStepStart "Diagnostico de conectividade"

    chaveAPI = CarregarChaveAPI()
    If Len(Trim(chaveAPI)) = 0 Then
        LogMessage LOG_PREFIX & ": Chave API nao configurada", LOG_LEVEL_ERROR
        MsgBox _
            "A chave do OpenRouter nao foi configurada." & _
            vbCrLf & vbCrLf & _
            "Configure-a na interface config_prompt.", _
            vbCritical, "Chave nao configurada"
        Exit Sub
    End If

    Application.StatusBar = LOG_PREFIX & ": Testando conectividade..."

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts _
        HTTP_RESOLVE_TIMEOUT_MS, _
        HTTP_CONNECT_TIMEOUT_MS, _
        HTTP_SEND_TIMEOUT_MS, _
        HTTP_RECEIVE_TIMEOUT_MS

    http.Open "GET", "https://openrouter.ai/api/v1/models", False
    http.setRequestHeader "Authorization", "Bearer " & chaveAPI
    http.setRequestHeader "Content-Type", "application/json"
    http.send

    resposta = BytesParaStringUTF8(http.ResponseBody)
    Application.StatusBar = False

    If http.Status = 200 Then
        LogStepComplete "Diagnostico de conectividade", _
            "Conexao OK | HTTP " & http.Status
        MsgBox _
            "Conexao com OpenRouter OK!" & vbCrLf & _
            "Codigo HTTP: " & http.Status, _
            vbInformation, "Diagnostico OpenRouter"
    Else
        LogMessage LOG_PREFIX & ": Diagnostico HTTP " & http.Status, LOG_LEVEL_ERROR
        MsgBox _
            "Falha na conexao com OpenRouter." & vbCrLf & vbCrLf & _
            "Codigo HTTP: " & http.Status & vbCrLf & _
            "Resposta: " & Left(resposta, 300), _
            vbCritical, "Diagnostico OpenRouter"
    End If

    Set http = Nothing
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Set http = Nothing
    LogMessage LOG_PREFIX & ": Erro no diagnostico: " & _
        Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    MsgBox _
        "Erro VBA:" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, "Diagnostico"
End Sub

