Attribute VB_Name = "Mod_11_RevisionText"
Option Explicit

' Mod_11_RevisionText.bas
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
'   - CorrigirProposituraComIA:     revisa texto selecionado (substitui no documento)
'   - DiagnosticarOpenRouter:       diagnostico de conectividade com a API
'
' Arquivos externos esperados (via config_prompt.py / z7_api_key.py):
'   - %USERPROFILE%\AppData\Local\Z7\Apps\StdProposers\LocalConfigs\openrouter.key
'   - %USERPROFILE%\AppData\Local\Z7\Apps\StdProposers\LocalConfigs\selected_model.txt
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

' Marcador de paragrafo enviado a IA (ChrW(182) = ¶, Pilcrow Sign)
' Preservado no prompt e parseado no retorno para processamento paragrafo-a-paragrafo
Private Const PARAGRAPH_MARKER As String = "¶"

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
    ByVal Length As Long)
Private Declare Function LocalFree _
    Lib "kernel32" ( _
    ByVal hMem As Long _
    ) As Long
#End If

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
' CARREGAR PROMPT DE REVISAO DO CONFIG_PROMPT
' =============================================================================
' Le revision_prompt.txt gravado pelo config_prompt.py.
' Retorna MontarPromptRevisao() se o arquivo nao existir.
' =============================================================================
Private Function CarregarPromptRevisao() As String
    On Error GoTo ErrorHandler

    Dim caminhoArquivo As String
    Dim ff As Integer
    Dim conteudo As String
    Dim linha As String

    caminhoArquivo = GetZ7StdProposersDataPath() & _
        "\revision_prompt.txt"

    If Dir(caminhoArquivo) <> "" Then
        ff = FreeFile
        Open caminhoArquivo For Input As #ff
        conteudo = ""
        Do While Not EOF(ff)
            Line Input #ff, linha
            If Len(conteudo) > 0 Then
                conteudo = conteudo & vbLf & linha
            Else
                conteudo = linha
            End If
        Loop
        Close #ff
        conteudo = Trim(conteudo)
        If Len(conteudo) > 0 Then
            CarregarPromptRevisao = conteudo
            LogMessage LOG_PREFIX & ": Prompt de revisao carregado do arquivo", LOG_LEVEL_INFO
            Exit Function
        End If
    End If

    CarregarPromptRevisao = MontarPromptRevisao()
    LogMessage LOG_PREFIX & ": Prompt de revisao padrao (embutido)", LOG_LEVEL_INFO
    Exit Function

ErrorHandler:
    LogMessage LOG_PREFIX & ": Erro ao carregar prompt de revisao: " & Err.Description, LOG_LEVEL_WARNING
    CarregarPromptRevisao = MontarPromptRevisao()
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
    textoOriginal = ExtrairTextoParaIA(rng)

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
    Application.StatusBar = RenderProgressBar(20, "Enviando texto para a IA")
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
        SubstituirTextoPreservandoFormatacaoMultiParagrafo _
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
' 2 - CORRIGIR TEXTO SELECIONADO COM IA
' =============================================================================
'
' Selecione qualquer texto no Word e execute esta macro.
' A IA revisa e corrige o texto selecionado.
' A formatacao original e preservada apos a substituicao.
'
' Diferente de TestarRevisaoTextoSelecionado, esta macro exibe metricas
' detalhadas e mensagens de status mais informativas ao usuario.
'
' =============================================================================
Public Sub CorrigirProposituraComIA()
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim textoOriginal As String
    Dim textoCorrigido As String
    Dim startTime As Double
    Dim chaveAPI As String

    startTime = Timer
    undoGroupEnabled = False ' Reset inicial

    ' -----------------------------------------------------------------
    ' VALIDACOES INICIAIS
    ' -----------------------------------------------------------------
    If Selection Is Nothing Then
        LogMessage LOG_PREFIX & ": Nenhuma selecao ativa", LOG_LEVEL_WARNING
        MsgBox "Selecione o texto a ser corrigido antes de executar " & _
            "esta macro.", vbExclamation, "Corrigir com IA"
        Exit Sub
    End If

    Set rng = Selection.Range.Duplicate
    textoOriginal = ExtrairTextoParaIA(rng)

    If Len(Trim(textoOriginal)) = 0 Then
        LogMessage LOG_PREFIX & ": Selecao vazia", LOG_LEVEL_WARNING
        MsgBox "Nenhum texto selecionado para correcao.", _
            vbExclamation, "Corrigir com IA"
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' VALIDA CHAVE ANTES DE ENVIAR
    ' -----------------------------------------------------------------
    chaveAPI = CarregarChaveAPI()
    If Not ValidarChaveAPI(chaveAPI) Then Exit Sub

    ' -----------------------------------------------------------------
    ' INICIALIZA LOGGING
    ' -----------------------------------------------------------------
    LogSection "CORRECAO IA - TEXTO SELECIONADO"
    LogStepStart "Correcao de texto selecionado com IA"
    LogMetric "Caracteres selecionados", Len(textoOriginal)
    LogMetric "Modelo IA", CarregarModeloIA()

    ' -----------------------------------------------------------------
    ' ENVIA O TEXTO SELECIONADO PARA A IA
    ' -----------------------------------------------------------------
    Application.StatusBar = RenderProgressBar(20, "Enviando texto para a IA")
    DoEvents

    textoCorrigido = ProcessarTextoComIA(textoOriginal)

    If Len(Trim(textoCorrigido)) = 0 Then
        Application.StatusBar = False
        LogStepSkipped "Correcao de texto selecionado", _
            "IA retornou resposta vazia"
        MsgBox "A IA nao retornou texto corrigido." & vbCrLf & vbCrLf & _
            "Verifique a conexao e a chave de API.", _
            vbExclamation, "Corrigir com IA"
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' SUBSTITUI SE HOUVE ALTERACAO (PRESERVANDO FORMATACAO)
    ' -----------------------------------------------------------------

    ' -----------------------------------------------------------------
    ' INICIO DO GRUPO DE DESFAZER (UndoRecord) - melhor esforco
    ' Late binding via CallByName: evita erro de compilacao em versoes
    ' do Word onde UndoRecord nao resolve como early-bound.
    ' Agrupa todas as edicoes da substituicao em uma unica acao Ctrl+Z
    ' -----------------------------------------------------------------
    On Error Resume Next
    Dim objUndoStart As Object
    Set objUndoStart = CallByName(Application, "UndoRecord", VbGet)
    If Err.Number = 0 Then
        If Not objUndoStart Is Nothing Then
            CallByName objUndoStart, "StartCustomRecord", VbMethod, "Z7_STDPROPOSERS - Correcao IA"
            If Err.Number = 0 Then
                undoGroupEnabled = True
                LogMessage LOG_PREFIX & ": UndoRecord iniciado", LOG_LEVEL_INFO
            Else
                undoGroupEnabled = False
                Err.Clear
            End If
        Else
            undoGroupEnabled = False
        End If
    Else
        undoGroupEnabled = False
        Err.Clear
    End If
    Set objUndoStart = Nothing
    On Error GoTo ErrorHandler
    ' -----------------------------------------------------------------

    If NormalizarComparacao(textoCorrigido) <> _
       NormalizarComparacao(textoOriginal) Then
        SubstituirTextoPreservandoFormatacaoMultiParagrafo _
            rng, _
            textoCorrigido
        Application.StatusBar = RenderProgressBar(100, "Texto corrigido com sucesso")
        LogStepComplete "Correcao de texto selecionado", _
            "Texto corrigido e substituido | " & _
            Format(Timer - startTime, "0.0") & "s"
    Else
        Application.StatusBar = RenderProgressBar(100, "Nenhuma alteracao necessaria")
        LogStepComplete "Correcao de texto selecionado", _
            "Nenhuma alteracao necessaria | " & _
            Format(Timer - startTime, "0.0") & "s"
    End If

    ' -----------------------------------------------------------------
    ' FIM DO GRUPO DE DESFAZER - SEMPRE fecha o UndoRecord
    ' Late binding via CallByName: evita erro de compilacao
    ' -----------------------------------------------------------------
    On Error Resume Next
    If undoGroupEnabled Then
        Dim objUndoEnd As Object
        Set objUndoEnd = CallByName(Application, "UndoRecord", VbGet)
        If Err.Number = 0 Then
            If Not objUndoEnd Is Nothing Then
                CallByName objUndoEnd, "EndCustomRecord", VbMethod
            End If
        End If
        Err.Clear
        Application.OnRepeat "Z7_STDPROPOSERS - Correcao IA", "CorrigirProposituraComIA"
        undoGroupEnabled = False
        LogMessage LOG_PREFIX & ": UndoRecord finalizado com sucesso", LOG_LEVEL_INFO
    End If
    Err.Clear
    On Error GoTo 0
    ' -----------------------------------------------------------------

    Exit Sub

ErrorHandler:
    ' Captura os dados do erro antes de qualquer manipulacao do UndoRecord
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    ' CRITICO: Garante fechamento do UndoRecord mesmo em erro
    ' Late binding via CallByName: evita erro de compilacao
    If undoGroupEnabled Then
        On Error Resume Next
        Dim errUndo As Object
        Set errUndo = CallByName(Application, "UndoRecord", VbGet)
        If Err.Number = 0 Then
            If Not errUndo Is Nothing Then
                CallByName errUndo, "EndCustomRecord", VbMethod
            End If
        End If
        Err.Clear
        undoGroupEnabled = False
        On Error GoTo 0
    End If

    Application.StatusBar = False
    LogMessage LOG_PREFIX & ": Erro na correcao: " & _
        errNum & " - " & errDesc, LOG_LEVEL_ERROR
    MsgBox _
        "Erro durante a correcao:" & vbCrLf & vbCrLf & _
        "Numero: " & errNum & vbCrLf & _
        "Descricao: " & errDesc, _
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
    promptSystem = CarregarPromptRevisao()
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
        "ESTRUTURA DE PARAGRAFOS: o carcater '" & PARAGRAPH_MARKER & "' " & _
        "(pilcrow) separa paragrafos."
    p = p & vbLf & _
        "Cada '" & PARAGRAPH_MARKER & "' indica o fim de um paragrafo e o " & _
        "inicio do proximo."
    p = p & vbLf & _
        "PRESERVE TODOS os marcadores '" & PARAGRAPH_MARKER & "' " & _
        "EXATAMENTE nas mesmas posicoes."
    p = p & vbLf & _
        "Seu texto de saida DEVE conter o MESMO numero de '" & _
        PARAGRAPH_MARKER & "' que o texto de entrada."
    p = p & vbLf & _
        "NAO remova, NAO adicione, NAO mova os marcadores '" & _
        PARAGRAPH_MARKER & "'."
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
' EXTRAI TEXTO DO RANGE PRESERVANDO PARAGRAFOS (via marcador ¶)
' =============================================================================
' Substitui vbCr por PARAGRAPH_MARKER (¶) para que a IA receba as
' fronteiras entre paragrafos. Caracteres de controle sao removidos.
' =============================================================================
Private Function ExtrairTextoParaIA( _
    ByVal rng As Range) As String
    On Error Resume Next

    Dim texto As String
    texto = rng.text
    texto = Replace(texto, vbCr, PARAGRAPH_MARKER)
    texto = Replace(texto, vbLf, "")
    texto = Replace(texto, Chr(7), "")    ' BEL character (form fields)
    texto = Replace(texto, Chr(160), " ") ' Non-breaking space

    ExtrairTextoParaIA = Trim(texto)
End Function


' =============================================================================
' SUBSTITUI TEXTO PRESERVANDO FORMATACAO (ROBUSTA)
' =============================================================================
'
' Estrategia robusta de preservacao de formatacao:
'
' 1. Salva TODAS as propriedades do paragrafo ANTES da substituicao
'    (incluindo Style, Borders, Shading, KeepWithNext, etc.).
' 2. Detecta e EXCLUI marcas de paragrafo (vbCr) do range de
'    substituicao para que a marca ¶ (que armazena a formatacao
'    do paragrafo) NAO seja destruida pelo rng.Text = novoTexto.
' 3. Salva a formatacao caractere-a-caractere APENAS do range
'    ajustado (sem as marcas de paragrafo), garantindo que o
'    mapeamento proporcional entre lenOld e lenNew seja exato.
' 4. Agrupa em "runs" (trechos com formatacao identica) e aplica
'    por run com mapeamento proporcional de posicoes.
' 5. Restaura a formatacao do paragrafo APOS a formatacao de
'    caractere, garantindo que propriedades como alinhamento,
'    espacamento, bordas e sombreamento sejam reaplicadas.
'
' Propriedades salvas por caractere:
'   Font.Name, Font.Size, Font.Bold, Font.Italic,
'   Font.Underline, Font.Color, Font.HighlightColorIndex,
'   Font.StrikeThrough, Font.SmallCaps, Font.AllCaps,
'   Font.Superscript, Font.Subscript, Font.Spacing,
'   Font.Position
'
' Propriedades de paragrafo salvas:
'   Style, Alignment, SpaceBefore, SpaceAfter,
'   SpaceBeforeAuto, SpaceAfterAuto, LineSpacing,
'   LineSpacingRule, LeftIndent, RightIndent,
'   FirstLineIndent, KeepWithNext, KeepTogether,
'   PageBreakBefore, WidowControl, OutlineLevel,
'   Borders (left/right/top/bottom), Shading
'   (BackgroundPatternColor, ForegroundPatternColor, Texture)
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

    ' --- arrays de formatacao de caractere ---
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

    ' --- propriedades de paragrafo (salvamento completo) ---
    Dim estilo As Variant
    Dim alinhamento As Long
    Dim espAntes As Single
    Dim espDepois As Single
    Dim espAntesAuto As Long
    Dim espDepoisAuto As Long
    Dim espLinha As Single
    Dim espLinhaRegra As Long
    Dim recuoEsquerdo As Single
    Dim recuoDireito As Single
    Dim primeiraLinha As Single
    Dim keepWithNext As Long
    Dim keepTogether As Long
    Dim pageBreakBefore As Long
    Dim widowControl As Long
    Dim outlineLevel As Long

    ' --- bordas do paragrafo ---
    Dim bordaEsqEstilo As Long
    Dim bordaEsqCor As Long
    Dim bordaEsqLargura As Long
    Dim bordaDirEstilo As Long
    Dim bordaDirCor As Long
    Dim bordaDirLargura As Long
    Dim bordaTopEstilo As Long
    Dim bordaTopCor As Long
    Dim bordaTopLargura As Long
    Dim bordaBotEstilo As Long
    Dim bordaBotCor As Long
    Dim bordaBotLargura As Long

    ' --- shading do paragrafo ---
    Dim shadingBgColor As Long
    Dim shadingFgColor As Long
    Dim shadingTexture As Long

    ' --- controle de posicao ---
    Dim oldStart As Long
    Dim oldEnd As Long
    Dim newStart As Long
    Dim newEnd As Long
    Dim posInicioNovoTexto As Long   ' Salva posicao apos substituicao (p/ step 10)
    Dim temMarcaParagrafo As Boolean
    Dim testRng As Range

    ' -----------------------------------------------------------------
    ' 1. DUPLICA RANGE E VERIFICA TAMANHO
    ' -----------------------------------------------------------------
    Set rng = rngOriginal.Duplicate
    lenOld = rng.End - rng.Start
    If lenOld <= 0 Then Exit Sub

    ' -----------------------------------------------------------------
    ' 2. SALVA FORMATACAO COMPLETA DO PARAGRAFO (ANTES DE QUALQUER
    '    ALTERACAO NO TEXTO)
    ' -----------------------------------------------------------------
    On Error Resume Next
    estilo = rng.Style
    alinhamento = rng.ParagraphFormat.alignment
    espAntes = rng.ParagraphFormat.SpaceBefore
    espDepois = rng.ParagraphFormat.SpaceAfter
    espAntesAuto = rng.ParagraphFormat.SpaceBeforeAuto
    espDepoisAuto = rng.ParagraphFormat.SpaceAfterAuto
    espLinha = rng.ParagraphFormat.LineSpacing
    espLinhaRegra = rng.ParagraphFormat.LineSpacingRule
    recuoEsquerdo = rng.ParagraphFormat.leftIndent
    recuoDireito = rng.ParagraphFormat.RightIndent
    primeiraLinha = rng.ParagraphFormat.firstLineIndent
    keepWithNext = rng.ParagraphFormat.KeepWithNext
    keepTogether = rng.ParagraphFormat.KeepTogether
    pageBreakBefore = rng.ParagraphFormat.PageBreakBefore
    widowControl = rng.ParagraphFormat.WidowControl
    outlineLevel = rng.ParagraphFormat.OutlineLevel

    ' Bordas
    bordaEsqEstilo = rng.ParagraphFormat.Borders(wdBorderLeft).LineStyle
    bordaEsqCor = rng.ParagraphFormat.Borders(wdBorderLeft).Color
    bordaEsqLargura = rng.ParagraphFormat.Borders(wdBorderLeft).LineWidth
    bordaDirEstilo = rng.ParagraphFormat.Borders(wdBorderRight).LineStyle
    bordaDirCor = rng.ParagraphFormat.Borders(wdBorderRight).Color
    bordaDirLargura = rng.ParagraphFormat.Borders(wdBorderRight).LineWidth
    bordaTopEstilo = rng.ParagraphFormat.Borders(wdBorderTop).LineStyle
    bordaTopCor = rng.ParagraphFormat.Borders(wdBorderTop).Color
    bordaTopLargura = rng.ParagraphFormat.Borders(wdBorderTop).LineWidth
    bordaBotEstilo = rng.ParagraphFormat.Borders(wdBorderBottom).LineStyle
    bordaBotCor = rng.ParagraphFormat.Borders(wdBorderBottom).Color
    bordaBotLargura = rng.ParagraphFormat.Borders(wdBorderBottom).LineWidth

    ' Shading
    shadingBgColor = rng.ParagraphFormat.Shading.BackgroundPatternColor
    shadingFgColor = rng.ParagraphFormat.Shading.ForegroundPatternColor
    shadingTexture = rng.ParagraphFormat.Shading.Texture
    On Error GoTo ErrorHandler

    LogMessage LOG_PREFIX & ": Formatacao do paragrafo salva (" & _
        "Style=" & TypeName(estilo) & ", Align=" & alinhamento & ")", LOG_LEVEL_INFO

    ' -----------------------------------------------------------------
    ' 3. DETECTA E EXCLUI MARCA DE PARAGRAFO (vbCr) DO FINAL DO RANGE
    '    Esta e a correcao CRITICA: a marca ¶ armazena a formatacao
    '    do paragrafo e NAO pode ser destruida pelo rng.Text = novoTexto.
    ' -----------------------------------------------------------------
    temMarcaParagrafo = False
    On Error Resume Next
    If rng.End > rng.Start Then
        Set testRng = rng.Duplicate
        testRng.Collapse wdCollapseEnd
        testRng.MoveStart wdCharacter, -1
        If Asc(testRng.text) = 13 Then
            temMarcaParagrafo = True
            rng.MoveEnd wdCharacter, -1
            LogMessage LOG_PREFIX & ": Marca de paragrafo detectada e " & _
                "excluida do range de substituicao", LOG_LEVEL_INFO
        End If
    End If
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------
    ' 4. RECALCULA lenOld COM O RANGE AJUSTADO
    ' -----------------------------------------------------------------
    lenOld = rng.End - rng.Start
    If lenOld <= 0 Then
        ' Range so tinha a marca de paragrafo — restaura e sai
        If temMarcaParagrafo Then
            rng.MoveEnd wdCharacter, 1
        End If
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' 5. SALVA FORMATACAO DE CADA CARACTERE (apenas do range ajustado,
    '    SEM as marcas de paragrafo)
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
        arrHighlight(i) = charRng.HighlightColorIndex
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
    ' 6. CONSTRUI RUNS DE FORMATACAO (agrupa caracteres consecutivos
    '    com formatacao identica)
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
    ' 7. SUBSTITUI O TEXTO (a marca de paragrafo NAO e alterada
    '    porque o range foi ajustado no passo 3)
    ' -----------------------------------------------------------------
    rng.text = novoTexto
    lenNew = rng.End - rng.Start
    posInicioNovoTexto = rng.Start   ' Salva posicao para reaplicar font apos Style
    If lenNew <= 0 Then
        If temMarcaParagrafo Then
            rng.MoveEnd wdCharacter, 1
        End If
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' 8. RESTAURA FORMATACAO DE CARACTERE POR RUN
    '    (mapeamento proporcional exato porque lenOld e lenNew
    '    referem-se ambos ao range sem marcas de paragrafo)
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
        runRng.HighlightColorIndex = arrHighlight(oldStart)
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
    ' 9. RESTAURA FORMATACAO COMPLETA DO PARAGRAFO
    '    Reexpande o range para incluir a marca de paragrafo e
    '    reaplica todas as propriedades salvas no passo 2.
    ' -----------------------------------------------------------------
    If temMarcaParagrafo Then
        rng.MoveEnd wdCharacter, 1
    End If

    On Error Resume Next

    rng.ParagraphFormat.alignment = alinhamento
    rng.ParagraphFormat.SpaceBefore = espAntes
    rng.ParagraphFormat.SpaceAfter = espDepois
    rng.ParagraphFormat.SpaceBeforeAuto = espAntesAuto
    rng.ParagraphFormat.SpaceAfterAuto = espDepoisAuto
    rng.ParagraphFormat.LineSpacing = espLinha
    rng.ParagraphFormat.LineSpacingRule = espLinhaRegra
    rng.ParagraphFormat.leftIndent = recuoEsquerdo
    rng.ParagraphFormat.RightIndent = recuoDireito
    rng.ParagraphFormat.firstLineIndent = primeiraLinha
    rng.ParagraphFormat.KeepWithNext = keepWithNext
    rng.ParagraphFormat.KeepTogether = keepTogether
    rng.ParagraphFormat.PageBreakBefore = pageBreakBefore
    rng.ParagraphFormat.WidowControl = widowControl
    rng.ParagraphFormat.OutlineLevel = outlineLevel

    If Not IsEmpty(estilo) Then
        rng.Style = estilo
        rng.ParagraphFormat.alignment = alinhamento
        rng.ParagraphFormat.SpaceBefore = espAntes
        rng.ParagraphFormat.SpaceAfter = espDepois
        rng.ParagraphFormat.SpaceBeforeAuto = espAntesAuto
        rng.ParagraphFormat.SpaceAfterAuto = espDepoisAuto
        rng.ParagraphFormat.LineSpacing = espLinha
        rng.ParagraphFormat.LineSpacingRule = espLinhaRegra
        rng.ParagraphFormat.leftIndent = recuoEsquerdo
        rng.ParagraphFormat.RightIndent = recuoDireito
        rng.ParagraphFormat.firstLineIndent = primeiraLinha
        rng.ParagraphFormat.KeepWithNext = keepWithNext
        rng.ParagraphFormat.KeepTogether = keepTogether
        rng.ParagraphFormat.PageBreakBefore = pageBreakBefore
        rng.ParagraphFormat.WidowControl = widowControl
        rng.ParagraphFormat.OutlineLevel = outlineLevel
    End If

    If bordaEsqEstilo <> 0 Then
        rng.ParagraphFormat.Borders(wdBorderLeft).LineStyle = bordaEsqEstilo
        rng.ParagraphFormat.Borders(wdBorderLeft).Color = bordaEsqCor
        rng.ParagraphFormat.Borders(wdBorderLeft).LineWidth = bordaEsqLargura
    End If
    If bordaDirEstilo <> 0 Then
        rng.ParagraphFormat.Borders(wdBorderRight).LineStyle = bordaDirEstilo
        rng.ParagraphFormat.Borders(wdBorderRight).Color = bordaDirCor
        rng.ParagraphFormat.Borders(wdBorderRight).LineWidth = bordaDirLargura
    End If
    If bordaTopEstilo <> 0 Then
        rng.ParagraphFormat.Borders(wdBorderTop).LineStyle = bordaTopEstilo
        rng.ParagraphFormat.Borders(wdBorderTop).Color = bordaTopCor
        rng.ParagraphFormat.Borders(wdBorderTop).LineWidth = bordaTopLargura
    End If
    If bordaBotEstilo <> 0 Then
        rng.ParagraphFormat.Borders(wdBorderBottom).LineStyle = bordaBotEstilo
        rng.ParagraphFormat.Borders(wdBorderBottom).Color = bordaBotCor
        rng.ParagraphFormat.Borders(wdBorderBottom).LineWidth = bordaBotLargura
    End If

    If shadingTexture <> 0 Then
        rng.ParagraphFormat.Shading.BackgroundPatternColor = shadingBgColor
        rng.ParagraphFormat.Shading.ForegroundPatternColor = shadingFgColor
        rng.ParagraphFormat.Shading.Texture = shadingTexture
    End If

    On Error GoTo 0

    ' -----------------------------------------------------------------
    ' 10. REAPLICAR FORMATACAO DE CARACTERE (FONT, BOLD, ITALIC, ETC.)
    '     A restauracao do Style no passo 9 pode sobrescrever a
    '     formatacao de caractere (Font.Name, Font.Size, etc.)
    '     aplicada no passo 8, pois ao aplicar o estilo em um range
    '     que inclui a marca de paragrafo (¶), o Word propaga as
    '     definicoes de fonte do estilo para todo o texto do paragrafo.
    '     Reaplicamos aqui apos a restauracao do paragrafo para
    '     garantir que tipo/tamanho da fonte e demais atributos de
    '     caractere sejam efetivamente preservados.
    ' -----------------------------------------------------------------
    Dim posInicioChar As Long
    Dim posFimChar As Long
    Dim rngChar As Range

    On Error Resume Next
    For i = 1 To runCount
        If i = 1 Then
            oldStart = 1
        Else
            oldStart = runEndPos(i - 1) + 1
        End If
        oldEnd = runEndPos(i)

        posInicioChar = posInicioNovoTexto + _
            CLng((oldStart - 1) * lenNew / lenOld)
        posFimChar = posInicioNovoTexto + _
            CLng(oldEnd * lenNew / lenOld) - 1
        If posFimChar < posInicioChar Then
            posFimChar = posInicioChar
        End If

        Set rngChar = ActiveDocument.Range( _
            posInicioChar, posFimChar + 1)

        rngChar.Font.Name = arrFontName(oldStart)
        rngChar.Font.size = arrFontSize(oldStart)
        rngChar.Font.Bold = arrBold(oldStart)
        rngChar.Font.Italic = arrItalic(oldStart)
        rngChar.Font.Underline = arrUnderline(oldStart)
        rngChar.Font.Color = arrColor(oldStart)
        rngChar.HighlightColorIndex = arrHighlight(oldStart)
        rngChar.Font.StrikeThrough = arrStrikeThrough(oldStart)
        rngChar.Font.SmallCaps = arrSmallCaps(oldStart)
        rngChar.Font.AllCaps = arrAllCaps(oldStart)
        rngChar.Font.Superscript = arrSuperscript(oldStart)
        rngChar.Font.Subscript = arrSubscript(oldStart)
        rngChar.Font.Spacing = arrSpacing(oldStart)
        rngChar.Font.Position = arrPosition(oldStart)
    Next i
    On Error GoTo 0

    LogMessage LOG_PREFIX & ": Formatacao do paragrafo restaurada com " & _
        "sucesso", LOG_LEVEL_INFO

    Exit Sub

ErrorHandler:
    LogMessage LOG_PREFIX & ": Aviso - falha ao preservar formatacao: " & _
        Err.Number & " - " & Err.Description, LOG_LEVEL_WARNING
    Resume Next
End Sub


' =============================================================================
' SUBSTITUI TEXTO PRESERVANDO FORMATACAO - MULTI-PARAGRAFO
' =============================================================================
'
' Processa selecoes com MULTIPLOS PARAGRAFOS de forma independente.
'
' Estrategia:
'   1. Divide o range original em paragrafos individuais via
'      rng.Paragraphs (cada um com seu proprio vbCr).
'   2. Divide o texto corrigido da IA pelos marcadores PARAGRAPH_MARKER
'      (¶) para obter o conteudo de cada paragrafo corrigido.
'   3. Para CADA paragrafo, chama SubstituirTextoPreservandoFormatacao
'      apenas no range daquele paragrafo especifico, preservando
'      alinhamento, recuos, Style, bordas, shading e formatacao de
'      caractere daquele paragrafo individualmente.
'   4. Se a IA alterar o numero de paragrafos, o excedente do texto
'      original e mantido inalterado, e paragrafos extras da IA sao
'      descartados (comportamento seguro).
'
' NOTA: Esta rotina e SEMPRE chamada dentro de um grupo UndoRecord
'       gerenciado pela rotina chamadora (CorrigirProposituraComIA
'       ou TestarRevisaoTextoSelecionado).
'
' =============================================================================
Private Sub SubstituirTextoPreservandoFormatacaoMultiParagrafo( _
    ByVal rngOriginal As Range, _
    ByVal textoCorrigidoComMarcadores As String)
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------
    ' 1. PARSE DO TEXTO CORRIGIDO: divide por marcador ¶
    ' -----------------------------------------------------------------
    Dim paragrafosCorrigidos() As String
    Dim pCount As Long
    Dim i As Long
    Dim posMarker As Long
    Dim textoRestante As String

    textoRestante = textoCorrigidoComMarcadores
    pCount = 0

    ' Conta paragrafos (numero de ¶ + 1, ou 1 se nao houver marcador)
    Do While Len(textoRestante) > 0
        posMarker = InStr(textoRestante, PARAGRAPH_MARKER)
        If posMarker = 0 Then
            pCount = pCount + 1
            Exit Do
        End If
        pCount = pCount + 1
        textoRestante = Mid(textoRestante, posMarker + 1)
    Loop

    If pCount = 0 Then Exit Sub

    ReDim paragrafosCorrigidos(1 To pCount)

    ' Preenche array
    textoRestante = textoCorrigidoComMarcadores
    i = 1
    Do While Len(textoRestante) > 0
        posMarker = InStr(textoRestante, PARAGRAPH_MARKER)
        If posMarker = 0 Then
            paragrafosCorrigidos(i) = Trim(textoRestante)
            Exit Do
        End If
        paragrafosCorrigidos(i) = Trim(Left(textoRestante, posMarker - 1))
        textoRestante = Mid(textoRestante, posMarker + 1)
        i = i + 1
        If i > pCount Then Exit Do
    Loop

    LogMessage LOG_PREFIX & ": Processando " & pCount & _
        " paragrafo(s) corrigido(s) pela IA", LOG_LEVEL_INFO

    ' -----------------------------------------------------------------
    ' 2. ITERA SOBRE PARAGRAFOS DO RANGE ORIGINAL
    ' -----------------------------------------------------------------
    Dim totalParasOriginal As Long
    totalParasOriginal = rngOriginal.Paragraphs.count

    LogMessage LOG_PREFIX & ": Range original contem " & _
        totalParasOriginal & " paragrafo(s)", LOG_LEVEL_INFO

    ' Se for paragrafo unico e sem marcador, delega direto
    If totalParasOriginal = 1 And pCount = 1 Then
        Dim paraUnico As Range
        Set paraUnico = rngOriginal.Paragraphs(1).Range.Duplicate
        If paraUnico.End > paraUnico.Start Then
            If Asc(paraUnico.Characters.Last.text) = 13 Then
                paraUnico.MoveEnd wdCharacter, -1
            End If
        End If
        If Len(paraUnico.text) > 0 Or Len(paragrafosCorrigidos(1)) > 0 Then
            SubstituirTextoPreservandoFormatacao paraUnico, paragrafosCorrigidos(1)
        End If
        Set paraUnico = Nothing
        Exit Sub
    End If

    ' Multiplos paragrafos: processa cada um independentemente
    Dim paraIndex As Long
    Dim textoCorrigidoParagrafo As String
    Dim paraRange As Range

    For paraIndex = 1 To totalParasOriginal
        Set paraRange = rngOriginal.Paragraphs(paraIndex).Range.Duplicate

        ' Exclui a marca de paragrafo (vbCr) do final do range
        If paraRange.End > paraRange.Start Then
            If Asc(paraRange.Characters.Last.text) = 13 Then
                paraRange.MoveEnd wdCharacter, -1
            End If
        End If

        ' Determina o texto corrigido correspondente
        If paraIndex <= pCount Then
            textoCorrigidoParagrafo = paragrafosCorrigidos(paraIndex)
        Else
            textoCorrigidoParagrafo = paraRange.text
            LogMessage LOG_PREFIX & _
                ": Aviso - IA retornou " & pCount & _
                " paragrafo(s) mas range tem " & totalParasOriginal & _
                "; mantendo paragrafo " & paraIndex & " inalterado", _
                LOG_LEVEL_WARNING
        End If

        ' Substitui texto deste paragrafo preservando SUA formatacao
        If Len(paraRange.text) > 0 Or Len(textoCorrigidoParagrafo) > 0 Then
            SubstituirTextoPreservandoFormatacao _
                paraRange, textoCorrigidoParagrafo
        End If

        Set paraRange = Nothing
    Next paraIndex

    If pCount > totalParasOriginal Then
        LogMessage LOG_PREFIX & _
            ": Aviso - IA retornou " & pCount & _
            " paragrafo(s) mas range tem " & totalParasOriginal & _
            "; excedentes da IA descartados", LOG_LEVEL_WARNING
    End If

    Exit Sub

ErrorHandler:
    LogMessage LOG_PREFIX & _
        ": Erro na substituicao multi-paragrafo: " & _
        Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    Resume Next
End Sub


' =============================================================================
' NORMALIZA PARA COMPARACAO
' =============================================================================
Private Function NormalizarComparacao( _
    ByVal texto As String) As String
    texto = Replace(texto, vbCr, "")
    texto = Replace(texto, vbLf, "")
    texto = Replace(texto, PARAGRAPH_MARKER, "")
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

    Application.StatusBar = RenderProgressBar(30, "Testando conectividade OpenRouter")

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

