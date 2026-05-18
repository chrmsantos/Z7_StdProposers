Attribute VB_Name = "Mod3Pipeline"
Option Explicit

' Mod3Pipeline.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Versao: 7.7.7-rc1
' Data: 2026-05-15
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
'
' =============================================================================
' INDICE DE MODULOS (Ctrl+F para navegar)
' =============================================================================
'
' [MOD.CONST]    CONSTANTES E CONFIGURACAO .......................... ~L13
'                - Constantes do Word (wdXxx, msoXxx)
'                - Constantes de Formatacao (fontes, margens)
'                - Constantes de Sistema (versao, limites)
'                - Constantes de Elementos Estruturais
'
' [MOD.VARS]     VARIAVEIS GLOBAIS .................................. ~L102
'                - Estado da aplicacao (flags, contadores)
'                - Cache de paragrafos (Type paragraphCache)
'                - Protecao de imagens (Type ImageInfo)
'                - Configuracoes de visualizacao (Type ViewSettings)
'
' [MOD.MAIN]     PONTO DE ENTRADA PRINCIPAL ......................... ~L220
'                - PadronizarDocumentoMain() - Orquestrador
'
' [MOD.ERROR]    TRATAMENTO DE ERROS E RECUPERACAO .................. ~L475
'                - ShowUserFriendlyError, EmergencyRecovery
'                - SafeCleanup, ReleaseObjects, CloseAllOpenFiles
'
' [MOD.VALID]    VALIDACAO E COMPATIBILIDADE ........................ ~L579
'                - ValidateDocument, IsDocumentHealthy
'                - IsOperationTimeout, CheckWordVersion
'
' [MOD.TEXT]     PROCESSAMENTO DE TEXTO ............................. ~L638
'                - GetCleanParagraphText, RemovePunctuation
'                - NormalizarTexto, DetectSpecialParagraph
'
' [MOD.STRUCT]   IDENTIFICACAO DE ESTRUTURA ......................... ~L738
'                - IsTituloElement, IsEmentaElement
'                - IsJustificativaTitleElement, IsDataElement
'                - IsAssinaturaStart, IsTituloAnexoElement
'                - IdentifyDocumentStructure
'
' [MOD.CACHE]    SISTEMA DE CACHE ................................... ~L1225
'                - BuildParagraphCache, ClearParagraphCache
'
' [MOD.API]      API PUBLICA DE ACESSO .............................. ~L1311
'                - GetTituloRange, GetEmentaRange
'                - GetProposicaoRange, GetJustificativaRange
'                - GetDataRange, GetAssinaturaRange
'                - GetAnexoRange, GetElementInfo
'
' [MOD.PROGRESS] BARRA DE PROGRESSO ................................. ~L1602
'                - UpdateProgress, InitializeProgress
'                - IncrementProgress
'
' [MOD.SAFE]     ACESSO SEGURO A PROPRIEDADES ....................... ~L1658
'                - SafeGetCharacterCount, SafeSetFont
'                - SafeSetParagraphFormat, SafeFindReplace
'
' [MOD.PATH]     FUNCOES DE CAMINHO ................................. ~L1818
'                - GetProjectRootPath, GetZ7StdProposersBackupsPath
'                - GetZ7StdProposersLogsPath, EnsureZ7StdProposersFolders (Mod6System)
'
' [MOD.LOG]      SISTEMA DE LOGS .................................... ~L1907
'                - InitializeLogging, LogMessage, FlushLogBuffer
'                - LogSection, LogStepStart, LogStepComplete
'                - SafeFinalizeLogging
'
' [MOD.UTIL]     UTILITARIOS GERAIS ................................. ~L2321
'                - GetProtectionType, GetDocumentSize
'                - SanitizeFileName, GetWindowsVersion
'
' [MOD.STATE]    GERENCIAMENTO DE ESTADO ............................ ~L2418
'                - SetAppState, ValidateDocument (pre-checks)
'
' [MOD.FORMAT]   ROTINAS DE FORMATACAO .............................. ~L2577
'                - ApplyDocumentFormatting (orquestrador)
'                - ConfigurarPagina, FormatFont, FormatParagraphs
'                - FormatFirstParagraph, FormatSecondParagraph
'
' [MOD.CLEAN]    LIMPEZA DE FORMATACAO .............................. ~L5775
'                - ClearAllFormatting, RemovePageNumberLines
'                - CleanupDocumentStructure, RemoveTabMarks
'
' [MOD.TITLE]    FORMATACAO DE TITULO ............................... ~L6356
'                - FormatDocumentTitle
'
' [MOD.SPECIAL]  PARAGRAFOS ESPECIAIS ............................... ~L6480
'                - Considerando, Ante o Exposto, In Loco
'                - ApplyBoldToSpecialParagraphs
'                - FormatVereadorParagraphs
'
' [MOD.BLANK]    GERENCIAMENTO DE LINHAS EM BRANCO .................. ~L7011
'                - InsertBlankLinesInJustificativa
'                - EnsureSingleBlankLineBetweenParagraphs
'
' [MOD.PUBLIC]   SUBROTINAS PUBLICAS ................................ ~L7456
'                - AbrirRepositorioGitHub
'                - ConfirmarDesfazimento, DesfazerPadronizacao
'
' [MOD.BACKUP]   SISTEMA DE BACKUP .................................. ~L7621
'                - CreateDocumentBackup, RestoreBackup
'                - CleanupOldBackups
'
' [MOD.SPACES]   LIMPEZA DE ESPACOS ................................. ~L7846
'                - LimparEspacosMultiplos
'                - LimitarLinhasVaziasSequenciais
'
' [MOD.VIEW]     CONFIGURACAO DE VISUALIZACAO ....................... ~L8172
'                - ConfigureDocumentView
'                - RemoveHighlightingAndBorders
'
' [MOD.IMAGE]    PROTECAO DE IMAGENS ................................ ~L8397
'                - BackupAllImages, RestoreAllImages
'                - FormatImageParagraphsIndents
'                - CenterImageAfterPlenario
'
' [MOD.LIST]     FORMATACAO DE LISTAS ............................... ~L8625
'                - BackupListFormats, RestoreListFormats
'                - FormatNumberedParagraphsIndent
'                - FormatBulletedParagraphsIndent
'
' [MOD.UPDATE]   VERIFICACAO DE ATUALIZACAO ......................... ~L9666
'                - CheckForUpdates, ExecutarInstalador

' [MOD.FINAL]    FORMATACAO FINAL ................................... ~L10052
'                - ApplyUniversalFinalFormatting
'                - AddSpecialSpacing
'
' =============================================================================


'================================================================================
' Criterios para identificacao dos elementos da propositura
Public Const TITULO_MIN_LENGTH As Long = 15              ' Comprimento minimo do titulo
Public Const EMENTA_MIN_LEFT_INDENT As Single = 6        ' Recuo minimo a esquerda da ementa (em pontos)
Public Const PLENARIO_TEXT As String = "plenario"        ' Texto identificador da data (parcial)
Public Const ANEXO_TEXT_SINGULAR As String = "anexo"     ' Texto identificador de anexo (singular)
Public Const ANEXO_TEXT_PLURAL As String = "anexos"      ' Texto identificador de anexo (plural)
Public Const ASSINATURA_PARAGRAPH_COUNT As Long = 3      ' Numero de paragrafos da assinatura
Public Const ASSINATURA_BLANK_LINES_BEFORE As Long = 2   ' Linhas em branco antes da assinatura

'================================================================================
' Regras:
' - Copia o texto da ementa para a area de transferencia
' - Salva o documento atual e somente fecha se o salvamento foi bem sucedido
' - Se houver apenas o documento ativo aberto: fecha o Word
' - Se houver outros documentos abertos e algum NAO estiver salvo: minimiza o Word
' - Nao exibe mensagens ao usuario, exceto em caso de erro (e aborta com seguranca)
Public Sub concluir()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = Nothing
    Set doc = ActiveDocument
    If doc Is Nothing Then GoTo ErrorHandler

    ' Captura estado dos outros documentos ANTES de fechar o atual
    Dim docsCountBefore As Long
    docsCountBefore = Application.Documents.count

    Dim hasOtherUnsaved As Boolean
    hasOtherUnsaved = False

    Dim d As Document
    For Each d In Application.Documents
        If Not (d Is doc) Then
            If d.Saved = False Then
                hasOtherUnsaved = True
                Exit For
            End If
        End If
    Next d

    ' Copia a ementa (texto) para a area de transferencia
    If Not CopyEmentaToClipboard(doc) Then
        Err.Raise vbObjectError + 651, "concluir", "Nao foi possivel copiar a ementa para a area de transferencia."
    End If

    ' Salva com verificacao: nao fecha nada antes de garantir que salvou
    If Not SaveDocumentSafely(doc) Then
        Err.Raise vbObjectError + 652, "concluir", "Nao foi possivel salvar o documento com seguranca."
    End If

    ' Fecha o documento somente apos confirmar salvamento
    doc.Close SaveChanges:=wdDoNotSaveChanges

    ' Se era o unico documento aberto, fecha o Word
    If docsCountBefore <= 1 Then
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If

    ' Se ha outros documentos NAO salvos, minimiza o Word
    If hasOtherUnsaved Then
        On Error Resume Next
        Application.WindowState = wdWindowStateMinimize
        If Err.Number <> 0 Then
            Err.Clear
            If Not Application.ActiveWindow Is Nothing Then
                Application.ActiveWindow.WindowState = wdWindowStateMinimize
            End If
        End If
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ErrorHandler:
    ' Interrompe com seguranca: nao fecha documento/Word em erro
    Dim msg As String
    msg = "Erro em 'concluir': " & Err.Description
    On Error Resume Next
    MsgBox msg, vbCritical, "Z7_STDPROPOSERS - Erro"
End Sub

Public Function SaveDocumentSafely(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    SaveDocumentSafely = False
    If doc Is Nothing Then Exit Function

    ' Evita dialogs: se nao tiver caminho, Save pode abrir 'Salvar Como'
    If doc.Path = "" Then Exit Function
    If doc.ReadOnly Then Exit Function

    On Error Resume Next
    doc.Save
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Confirmacao minima: Word marcou como salvo
    If doc.Saved = True Then
        SaveDocumentSafely = True
    End If

    Exit Function

ErrorHandler:
    SaveDocumentSafely = False
End Function

Public Function CopyEmentaToClipboard(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    CopyEmentaToClipboard = False
    If doc Is Nothing Then Exit Function

    Dim ementaText As String
    ementaText = ""

    Dim rng As Range
    Set rng = Nothing
    Set rng = GetEmentaRange(doc)

    If Not rng Is Nothing Then
        ementaText = rng.text
    Else
        ' Fallback (mais tolerante): tenta extrair via heuristica
        ementaText = GetEmentaText(doc)
    End If

    ementaText = Trim$(Replace(Replace(ementaText, vbCr, ""), vbLf, ""))
    If ementaText = "" Then Exit Function

    ' Tenta copiar texto puro para a area de transferencia
    If PutTextInClipboard(ementaText) Then
        CopyEmentaToClipboard = True
        Exit Function
    End If

    ' Fallback: copia o Range (se disponivel)
    If Not rng Is Nothing Then
        On Error Resume Next
        rng.Copy
        If Err.Number = 0 Then
            CopyEmentaToClipboard = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:
    CopyEmentaToClipboard = False
End Function

Public Function PutTextInClipboard(text As String) As Boolean
    On Error GoTo ErrorHandler

    PutTextInClipboard = False
    If text = "" Then Exit Function

    ' 1. Preferencia: MSForms.DataObject (late binding)
    Dim dataObj As Object
    On Error Resume Next
    Set dataObj = CreateObject("MSForms.DataObject")
    If Err.Number = 0 And Not dataObj Is Nothing Then
        Err.Clear
        dataObj.SetText text
        dataObj.PutInClipboard
        PutTextInClipboard = (Err.Number = 0)
        Exit Function
    End If
    Err.Clear

    ' 2. Fallback: htmlfile clipboardData
    Dim html As Object
    Set html = CreateObject("htmlfile")
    If Not html Is Nothing Then
        html.parentWindow.clipboardData.setData "text", text
        PutTextInClipboard = True
        Exit Function
    End If

    Exit Function

ErrorHandler:
    PutTextInClipboard = False
End Function

'================================================================================
' LIMPEZA SEGURA DE RECURSOS
'================================================================================
Public Sub SafeCleanup()
    On Error Resume Next

    ' Nao tenta fechar UndoRecord aqui - ja foi fechado em CleanUp

    ReleaseObjects
End Sub

'================================================================================
' LIBERACAO DE OBJETOS
'================================================================================
Public Sub ReleaseObjects()
    On Error Resume Next

    Dim nullObj As Object
    Set nullObj = Nothing

    Dim memoryCounter As Long
    For memoryCounter = 1 To 3
        DoEvents
    Next memoryCounter
End Sub

'================================================================================
' FECHAMENTO DE ARQUIVOS ABERTOS
'================================================================================
Public Sub CloseAllOpenFiles()
    On Error Resume Next

    Dim fileNumber As Integer
    For fileNumber = 1 To 511
        Close #fileNumber
    Next fileNumber
    Err.Clear
End Sub

'================================================================================
' NORMALIZACAO OTIMIZADA DE TEXTO - Unica passagem
'================================================================================
Public Function NormalizarTexto(text As String) As String
    Dim result As String
    Dim loopGuard As Long
    result = text

    ' Remove caracteres de controle em uma unica passagem
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")

    ' Remove espacos multiplos com protecao contra loop infinito
    loopGuard = 0
    Do While InStr(result, "  ") > 0 And loopGuard < 500
        result = Replace(result, "  ", " ")
        loopGuard = loopGuard + 1
    Loop

    NormalizarTexto = Trim(LCase(result))
End Function

'================================================================================
' DETECCAO DE TIPO DE PARAGRAFO ESPECIAL
'================================================================================
Public Function DetectSpecialParagraph(cleanText As String, ByRef specialType As String) As Boolean
    specialType = ""

    ' Remove pontuacao final para analise
    Dim textForAnalysis As String
    textForAnalysis = cleanText

    Dim safetyCounter As Long
    safetyCounter = 0
    Do While Len(textForAnalysis) > 0 And InStr(".,;:", Right(textForAnalysis, 1)) > 0 And safetyCounter < 50
        textForAnalysis = Left(textForAnalysis, Len(textForAnalysis) - 1)
        safetyCounter = safetyCounter + 1
    Loop
    textForAnalysis = Trim(textForAnalysis)

    ' Verifica tipos especiais
    If Left(textForAnalysis, CONSIDERANDO_MIN_LENGTH) = CONSIDERANDO_PREFIX Then
        specialType = "considerando"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = JUSTIFICATIVA_TEXT Then
        specialType = "justificativa"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = "vereador" Or textForAnalysis = "vereadora" _
        Or textForAnalysis = "presidente" Or textForAnalysis = "prefeito" Or textForAnalysis = "vicepresidente" _
        Or textForAnalysis = "vice-presidente" Or textForAnalysis = "vice presidente" Then
        specialType = "vereador"
        DetectSpecialParagraph = True
    ElseIf Left(textForAnalysis, 17) = "diante do exposto" Then
        specialType = "dianteexposto"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = "requeiro" Then
        specialType = "requeiro"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = "anexo" Or textForAnalysis = "anexos" Then
        specialType = "anexo"
        DetectSpecialParagraph = True
    Else
        DetectSpecialParagraph = False
    End If
End Function

'================================================================================
' CALCULO DE PROGRESSO BASEADO EM ETAPAS
'================================================================================
Public Sub InitializeProgress(steps As Long)
    totalSteps = steps
    currentStep = 0
End Sub

Public Sub IncrementProgress(message As String)
    currentStep = currentStep + 1
    Dim percent As Long
    If totalSteps > 0 Then
        percent = CLng((currentStep * 100) / totalSteps)
    Else
        percent = 0
    End If
    UpdateProgress message, percent
End Sub

'================================================================================
' SAFE FIND/REPLACE OPERATIONS
'================================================================================
Public Function SafeFindReplace(doc As Document, findText As String, replaceText As String, Optional useWildcards As Boolean = False) As Long
    On Error GoTo ErrorHandler

    Dim findCount As Long
    findCount = 0

    ' Configuracao segura de Find/Replace
    With doc.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = useWildcards  ' Parametro controlado
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        ' Executa a substituicao e conta ocorrencias
        Do While .Execute(Replace:=True)
            findCount = findCount + 1
            ' Limite de seguranca para evitar loops infinitos
            If findCount > 10000 Then
                LogMessage "Limite de substituicoes atingido para: " & findText, LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    SafeFindReplace = findCount
    Exit Function

ErrorHandler:
    SafeFindReplace = 0
    LogMessage "Erro na operacao Find/Replace: " & findText & " -> " & replaceText & " | " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' SISTEMA DE REGISTRO DE LOGS
'================================================================================

'--------------------------------------------------------------------------------
' WriteTextUTF8 - Escreve texto em arquivo com encoding UTF-8
'--------------------------------------------------------------------------------
Public Sub WriteTextUTF8(filePath As String, textContent As String, Optional appendMode As Boolean = False)
    On Error GoTo ErrorHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open

    ' Se modo append, le conteudo existente primeiro
    If appendMode And Dir(filePath) <> "" Then
        stream.LoadFromFile filePath
        stream.Position = stream.size
    End If

    ' Escreve o novo conteudo
    stream.WriteText textContent, 1 ' adWriteLine

    ' Salva com UTF-8
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
End Sub

Public Sub EnforceLogRetention(logFolder As String, logPrefix As String, Optional maxFiles As Long = 5)
    On Error GoTo CleanExit

    If maxFiles < 1 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(logFolder) Then GoTo CleanExit

    Dim folder As Object
    Set folder = fso.GetFolder(logFolder)

    Dim fileItem As Object
    Dim prefixLower As String
    prefixLower = LCase(logPrefix)

    Dim items() As String
    Dim itemCount As Long
    itemCount = 0

    For Each fileItem In folder.Files
        If LCase(fileItem.Name) Like prefixLower & "*.log" Then
            ReDim Preserve items(itemCount)
            items(itemCount) = Format(fileItem.DateLastModified, "yyyymmddHHMMSS") & "|" & fileItem.Path
            itemCount = itemCount + 1
        End If
    Next fileItem

    If itemCount <= maxFiles Then GoTo CleanExit

    Dim i As Long, j As Long, temp As String
    For i = 0 To itemCount - 2
        For j = i + 1 To itemCount - 1
            If items(i) < items(j) Then
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next j
    Next i

    Dim idx As Long
    For idx = maxFiles To itemCount - 1
        Dim parts() As String
        parts = Split(items(idx), "|")
        On Error Resume Next
        fso.DeleteFile parts(1), True
        On Error GoTo CleanExit
    Next idx

CleanExit:
    On Error Resume Next
    Set sortedList = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

Public Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim logFolder As String
    Dim docNameClean As String
    Dim fileNum As Integer
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Garante que a estrutura de pastas do projeto existe
    EnsureZ7StdProposersFolders

    ' SEMPRE USA source\logs para todos os documentos
    logFolder = GetZ7StdProposersLogsPath() & "\"

    ' Garante que a pasta de logs existe antes de criar o arquivo
    If Not fso.FolderExists(logFolder) Then
        On Error Resume Next
        fso.CreateFolder logFolder
        On Error GoTo ErrorHandler
    End If

    If Not fso.FolderExists(logFolder) Then
        InitializeLogging = False
        loggingEnabled = False
        Exit Function
    End If

    ' Sanitiza nome do documento para uso em arquivo
    docNameClean = doc.Name
    docNameClean = Replace(docNameClean, ".doc", "")
    docNameClean = Replace(docNameClean, ".docx", "")
    docNameClean = Replace(docNameClean, ".docm", "")
    docNameClean = SanitizeFileName(docNameClean)

    ' Define nome do arquivo de log com timestamp
    logFilePath = logFolder & "z7_stdproposers_" & Format(Now, "yyyymmdd_HHmmss") & "_" & docNameClean & ".log"

    ' Inicializa contadores e controles
    errorCount = 0
    warningCount = 0
    infoCount = 0
    logBufferEnabled = False
    logBuffer = ""
    lastFlushTime = Now
    logFileHandle = 0
    currentLogSessionId = Format(Now, "yyyymmddHHmmss")
    currentOperationId = currentLogSessionId

    ' Cria arquivo de log com informacoes de contexto usando UTF-8
    Dim headerText As String
    headerText = String(80, "=") & vbCrLf
    headerText = headerText & "Z7_STDPROPOSERS - LOG DE PROCESSAMENTO DE DOCUMENTO" & vbCrLf
    headerText = headerText & String(80, "=") & vbCrLf & vbCrLf
    headerText = headerText & "[SESSAO]" & vbCrLf
    headerText = headerText & "  Inicio: " & Format(Now, "dd/mm/yyyy HH:mm:ss") & vbCrLf
    headerText = headerText & "  ID: " & currentLogSessionId & vbCrLf
    headerText = headerText & "  Operacao: " & currentOperationId & vbCrLf & vbCrLf
    headerText = headerText & "[AMBIENTE]" & vbCrLf
    headerText = headerText & "  Usuario: " & Environ("USERNAME") & vbCrLf
    headerText = headerText & "  Computador: " & Environ("COMPUTERNAME") & vbCrLf
    headerText = headerText & "  Dominio: " & Environ("USERDOMAIN") & vbCrLf
    headerText = headerText & "  SO: Windows " & GetWindowsVersion() & vbCrLf
    headerText = headerText & "  Word: " & Application.version & " (" & GetWordVersionName() & ")" & vbCrLf & vbCrLf
    headerText = headerText & "[DOCUMENTO]" & vbCrLf
    headerText = headerText & "  Nome: " & doc.Name & vbCrLf
    headerText = headerText & "  Caminho: " & IIf(doc.Path = "", "(Nao salvo)", doc.Path) & vbCrLf
    headerText = headerText & "  Tamanho: " & GetDocumentSize(doc) & vbCrLf
    headerText = headerText & "  Paragrafos: " & doc.Paragraphs.count & vbCrLf
    headerText = headerText & "  Paginas: " & doc.ComputeStatistics(wdStatisticPages) & vbCrLf
    headerText = headerText & "  Protecao: " & GetProtectionType(doc) & vbCrLf
    headerText = headerText & "  Idioma: " & doc.Range.LanguageID & vbCrLf & vbCrLf
    headerText = headerText & "[CONFIGURACAO]" & vbCrLf
    headerText = headerText & "  Debug: " & IIf(DEBUG_MODE, "Ativado", "Desativado") & vbCrLf
    headerText = headerText & "  Log: " & logFilePath & vbCrLf
    headerText = headerText & "  Backup: " & GetZ7StdProposersBackupsPath() & "\" & vbCrLf & vbCrLf
    headerText = headerText & String(80, "=") & vbCrLf & vbCrLf

    ' Escreve cabecalho em UTF-8
    WriteTextUTF8 logFilePath, headerText, False

    ' Enforces log retention limit for this routine
    EnforceLogRetention logFolder, "z7_stdproposers_", 5

    loggingEnabled = True
    InitializeLogging = True

    LogMessage "Logging inicializado com sucesso", LOG_LEVEL_INFO

    Exit Function

ErrorHandler:
    On Error Resume Next
    logFileHandle = 0
    loggingEnabled = False
    InitializeLogging = False
    Debug.Print "ERRO CRITICO: Falha ao inicializar logging - " & Err.Description
End Function

Public Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler

    If Not loggingEnabled Then Exit Sub

    Dim levelText As String
    Dim levelPrefix As String
    Dim fileNum As Integer
    Dim formattedMessage As String
    Dim timeStamp As String
    Dim elapsedTime As String
    Dim operationId As String

    ' Calcula tempo decorrido desde inicio
    If executionStartTime > 0 Then
        Dim elapsed As Double
        elapsed = (Now - executionStartTime) * 86400 ' Converte para segundos
        elapsedTime = Format(Int(elapsed / 60), "00") & ":" & Format(elapsed Mod 60, "00.0")
    Else
        elapsedTime = "00:00.0"
    End If

    ' Define nivel e incrementa contadores
    Select Case level
        Case LOG_LEVEL_INFO
            levelText = "INFO "
            levelPrefix = "?"
            infoCount = infoCount + 1
        Case LOG_LEVEL_WARNING
            levelText = "WARN "
            levelPrefix = "?"
            warningCount = warningCount + 1
        Case LOG_LEVEL_ERROR
            levelText = "ERROR"
            levelPrefix = "?"
            errorCount = errorCount + 1
        Case Else
            levelText = "DEBUG"
            levelPrefix = "?"
    End Select

    ' Formata mensagem com timestamp, tempo decorrido e nivel
    timeStamp = Format(Now, "HH:mm:ss.") & Format((Timer * 1000) Mod 1000, "000")
    operationId = currentOperationId
    If Len(operationId) = 0 Then operationId = "N/A"
    formattedMessage = timeStamp & " [" & elapsedTime & "] [op=" & operationId & "] " & levelText & " " & levelPrefix & " " & message

    ' Debug mode output para console VBA
    If DEBUG_MODE Then
        Debug.Print formattedMessage
    End If

    ' Buffer para reduzir I/O quando nao for erro critico
    If level = LOG_LEVEL_ERROR Or Len(logBuffer) > 4096 Or (Now - lastFlushTime) > (LOG_BUFFER_FLUSH_SECONDS / 86400) Then
        ' Escreve imediatamente: erros, buffer cheio (>4KB), ou 5+ segundos desde ultimo flush
        FlushLogBuffer

        ' Escreve mensagem em UTF-8
        WriteTextUTF8 logFilePath, formattedMessage, True

        lastFlushTime = Now
    Else
        ' Adiciona ao buffer para flush posterior (otimizacao de performance)
        logBuffer = logBuffer & formattedMessage & vbCrLf
    End If

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    Debug.Print "FALHA NO LOG: " & message & " | Erro: " & Err.Description
End Sub

Public Sub FlushLogBuffer()
    On Error Resume Next

    If Len(logBuffer) = 0 Then Exit Sub

    ' Escreve buffer em UTF-8
    WriteTextUTF8 logFilePath, logBuffer, True

    logBuffer = ""
    lastFlushTime = Now
End Sub

'================================================================================
' FUNCOES AUXILIARES DE LOG
'================================================================================
Public Sub LogSection(sectionName As String)
    On Error Resume Next

    If Not loggingEnabled Then Exit Sub

    FlushLogBuffer

    ' Cria texto de secao
    Dim sectionText As String
    sectionText = vbCrLf & String(80, "-") & vbCrLf
    sectionText = sectionText & "SECAO: " & UCase(sectionName) & vbCrLf
    sectionText = sectionText & String(80, "-")

    ' Escreve em UTF-8
    WriteTextUTF8 logFilePath, sectionText, True

    lastFlushTime = Now
End Sub

Public Sub LogStepStart(stepName As String)
    On Error Resume Next
    LogMessage "? Iniciando: " & stepName, LOG_LEVEL_INFO
End Sub

Public Sub LogStepComplete(stepName As String, Optional details As String = "")
    On Error Resume Next
    Dim msg As String
    msg = "? Concluido: " & stepName
    If Len(details) > 0 Then msg = msg & " | " & details
    LogMessage msg, LOG_LEVEL_INFO
End Sub

Public Sub LogStepSkipped(stepName As String, reason As String)
    On Error Resume Next
    LogMessage "? Ignorado: " & stepName & " | Motivo: " & reason, LOG_LEVEL_INFO
End Sub

Public Sub LogMetric(metricName As String, value As Variant, Optional unit As String = "")
    On Error Resume Next
    Dim msg As String
    msg = "?? " & metricName & ": " & CStr(value)
    If Len(unit) > 0 Then msg = msg & " " & unit
    LogMessage msg, LOG_LEVEL_INFO
End Sub

Public Sub LogContextSnapshot(doc As Document, contextName As String)
    On Error GoTo ErrorHandler

    If Not loggingEnabled Then Exit Sub
    If doc Is Nothing Then Exit Sub

    LogSection "SNAPSHOT - " & contextName
    LogMetric "Paragrafos", doc.Paragraphs.count
    LogMetric "Paginas", doc.ComputeStatistics(wdStatisticPages)
    LogMetric "Caracteres", SafeGetCharacterCount(doc.Range)
    LogMetric "Protecao", GetProtectionType(doc)
    LogMetric "SomenteLeitura", IIf(doc.ReadOnly, "SIM", "NAO")
    LogMetric "DocumentoSalvo", IIf(doc.Saved, "SIM", "NAO")
    Exit Sub

ErrorHandler:
    LogMessage "Falha ao coletar snapshot de contexto: " & Err.Description, LOG_LEVEL_WARNING
End Sub

Public Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler

    If Not loggingEnabled Then Exit Sub

    Dim fileNum As Integer
    Dim statusText As String
    Dim statusIcon As String
    Dim duration As Double
    Dim durationText As String
    Dim totalEvents As Long

    ' Flush pendente no buffer
    FlushLogBuffer

    ' Calcula duracao total
    duration = (Now - executionStartTime) * 86400
    If duration < 60 Then
        durationText = Format(duration, "0.0") & "s"
    ElseIf duration < 3600 Then
        durationText = Format(Int(duration / 60), "0") & "m " & Format(duration Mod 60, "00") & "s"
    Else
        durationText = Format(Int(duration / 3600), "0") & "h " & Format(Int((duration Mod 3600) / 60), "00") & "m"
    End If

    ' Determina status final
    If formattingCancelled Then
        statusText = "CANCELADO PELO USUARIO"
        statusIcon = "?"
    ElseIf errorCount > 0 Then
        statusText = "CONCLUIDO COM ERROS"
        statusIcon = "?"
    ElseIf warningCount > 0 Then
        statusText = "CONCLUIDO COM AVISOS"
        statusIcon = "?"
    Else
        statusText = "CONCLUIDO COM SUCESSO"
        statusIcon = "?"
    End If

    totalEvents = infoCount + warningCount + errorCount

    ' Escreve rodape estruturado em UTF-8
    Dim footerText As String
    footerText = vbCrLf & String(80, "=") & vbCrLf
    footerText = footerText & "RESUMO DA SESSAO" & vbCrLf
    footerText = footerText & String(80, "=") & vbCrLf & vbCrLf
    footerText = footerText & "[STATUS]" & vbCrLf
    footerText = footerText & "  Final: " & statusText & " " & statusIcon & vbCrLf
    footerText = footerText & "  Termino: " & Format(Now, "dd/mm/yyyy HH:mm:ss") & vbCrLf
    footerText = footerText & "  Duracao: " & durationText & vbCrLf & vbCrLf
    footerText = footerText & "[ESTATISTICAS]" & vbCrLf
    footerText = footerText & "  Total de eventos: " & totalEvents & vbCrLf
    footerText = footerText & "  Informacoes: " & infoCount & " (" & Format(infoCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)" & vbCrLf
    footerText = footerText & "  Avisos: " & warningCount & " (" & Format(warningCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)" & vbCrLf
    footerText = footerText & "  Erros: " & errorCount & " (" & Format(errorCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)" & vbCrLf & vbCrLf

    ' Adiciona informacoes de performance
    If totalEvents > 0 Then
        footerText = footerText & "[PERFORMANCE]" & vbCrLf
        footerText = footerText & "  Eventos/segundo: " & Format(totalEvents / IIf(duration > 0, duration, 1), "0.0") & vbCrLf
        footerText = footerText & "  Tempo medio/evento: " & Format((duration / totalEvents) * 1000, "0.0") & "ms" & vbCrLf & vbCrLf
    End If

    ' Recomendacoes se houver problemas
    If errorCount > 0 Or warningCount > 5 Then
        footerText = footerText & "[RECOMENDACOES]" & vbCrLf
        If errorCount > 0 Then
            footerText = footerText & "   Verifique os erros acima e corrija problemas no documento" & vbCrLf
        End If
        If warningCount > 5 Then
            footerText = footerText & "   Multiplos avisos detectados - revise o documento manualmente" & vbCrLf
        End If
        If duration > 60 Then
            footerText = footerText & "   Processamento demorado - considere otimizar o documento" & vbCrLf
        End If
        footerText = footerText & vbCrLf
    End If

    footerText = footerText & String(80, "=") & vbCrLf
    footerText = footerText & "FIM DO LOG" & vbCrLf
    footerText = footerText & String(80, "=")

    ' Escreve footer em UTF-8
    WriteTextUTF8 logFilePath, footerText, True

    ' Limpa variaveis
    loggingEnabled = False
    logBuffer = ""
    logFileHandle = 0

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    loggingEnabled = False
    Debug.Print "ERRO CRITICO ao finalizar logging: " & Err.Description
End Sub

'================================================================================
' UTILITY: SANITIZE FILE NAME
'================================================================================
Public Function SanitizeFileName(fileName As String) As String
    On Error Resume Next

    Dim result As String
    Dim invalidChars As String
    Dim i As Long

    result = fileName
    invalidChars = "\/:*?""<>|"

    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i

    ' Limita tamanho
    If Len(result) > 50 Then
        result = Left(result, 50)
    End If

    SanitizeFileName = result
End Function

'================================================================================
' VERIFICACOES GLOBAIS ANTES DA FORMATACAO
'================================================================================
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

    ' [REMOVIDO A PEDIDO DO USUARIO] Verifica consistencia ementa x corpo
    ' If Not ValidateAddressConsistency(doc) Then
    '     LogMessage "Recomendacao para verificar enderecos foi exibida ao usuario", LOG_LEVEL_INFO
    ' End If

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
Public Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Formatacoes basicas de pagina e estrutura
    If Not ApplyPageSetup(doc) Then
        LogMessage "Falha na configuracao de pagina", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    LogSection "LIMPEZA E FORMATACAO"

    ' Limpeza e formatacoes otimizadas
    LogStepStart "Limpeza de formatacao"
    ClearAllFormatting doc
    LogStepComplete "Limpeza de formatacao"

    LogStepStart "Normalizacao de quebras"
    ReplaceLineBreaksWithParagraphBreaks doc
    RemovePageBreaks doc
    LogStepComplete "Normalizacao de quebras"

    LogStepStart "Limpeza estrutural"
    RemovePageNumberLines doc
    CleanDocumentStructure doc
    RemoveAllTabMarks doc
    LogStepComplete "Limpeza estrutural"

    LogStepStart "Limpeza de prefixo da ementa"
    RemoveEmentaLeadingLabelPrefix doc
    LogStepComplete "Limpeza de prefixo da ementa"

    LogStepStart "Limpeza de sufixo da ementa"
    RemoveEmentaTrailingMunicipioSuffix doc
    LogStepComplete "Limpeza de sufixo da ementa"

    LogStepStart "Remocao de aspas da ementa"
    RemoveEmentaQuotes doc
    LogStepComplete "Remocao de aspas da ementa"

    LogStepStart "Formatacao de titulo"
    FormatDocumentTitle doc
    LogStepComplete "Formatacao de titulo"

    ' Formatacoes principais - Usa versao otimizada se cache disponivel
    LogStepStart "Aplicacao de fonte padrao"
    If cacheEnabled Then
        If Not ApplyStdFontOptimized(doc) Then
            LogMessage "Falha na formatacao de fontes (otimizada) - tentando metodo tradicional", LOG_LEVEL_WARNING
            If Not ApplyStdFont(doc) Then
                LogMessage "Falha na formatacao de fontes", LOG_LEVEL_ERROR
                PreviousFormatting = False
                Exit Function
            End If
        End If
    Else
        If Not ApplyStdFont(doc) Then
            LogMessage "Falha na formatacao de fontes", LOG_LEVEL_ERROR
            PreviousFormatting = False
            Exit Function
        End If
    End If
    LogStepComplete "Aplicacao de fonte padrao", doc.Paragraphs.count & " paragrafos"

    LogStepStart "Aplicacao de formatacao de paragrafos"
    If Not ApplyStdParagraphs(doc) Then
        LogMessage "Falha na formatacao de paragrafos", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    LogStepComplete "Aplicacao de formatacao de paragrafos"

    LogSection "FORMATACOES ESPECIFICAS"

    LogStepStart "Formatacao de paragrafos 1 e 2"
    FormatFirstParagraph doc
    FormatSecondParagraph doc
    LogStepComplete "Formatacao de paragrafos 1 e 2"

    LogStepStart "Formatacao de considerandos"
    FormatConsiderandoParagraphs doc
    LogStepComplete "Formatacao de considerandos"

    LogStepStart "Aplicacao de substituicoes de texto"
    ApplyTextReplacements doc
    LogStepComplete "Aplicacao de substituicoes de texto"

    LogStepStart "Remocao de marca d'agua e insercao de carimbo"
    RemoveWatermark doc
    InsertHeaderstamp doc
    LogStepComplete "Remocao de marca d'agua e insercao de carimbo"

    LogSection "LIMPEZA FINAL"

    LogStepStart "Limpeza de espacos multiplos"
    CleanMultipleSpaces doc
    LogStepComplete "Limpeza de espacos multiplos"

    LogStepStart "Controle de linhas em branco"
    LimitSequentialEmptyLines doc
    LogStepComplete "Controle de linhas em branco"

    LogStepStart "Substituicao de datas do plenario"
    ReplacePlenarioDateParagraph doc
    LogStepComplete "Substituicao de datas do plenario"

    LogSection "FINALIZACAO"

    LogStepStart "Configuracao de visualizacao"
    ConfigureDocumentView doc
    LogStepComplete "Configuracao de visualizacao"

    LogStepStart "Insercao de rodape"
    If Not InsertFooterStamp(doc) Then
        LogMessage "Falha na insercao do rodape", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    LogStepComplete "Insercao de rodape"

    LogStepStart "Ajustes finais de negrito e formatacao"
    ApplyBoldToSpecialParagraphs doc
    FormatVereadorParagraphs doc
    LogStepComplete "Ajustes finais de negrito e formatacao"

    LogStepStart "Formatacoes especiais (diante do exposto, requeiro)"
    FormatDianteDoExposto doc
    FormatRequeiroParagraphs doc
    FormatPorTodasRazoesParagraphs doc
    LogStepComplete "Formatacoes especiais (diante do exposto, requeiro)"

    LogStepStart "Remocao de realces e bordas"
    RemoveAllHighlightsAndBorders doc
    LogStepComplete "Remocao de realces e bordas"

    LogStepStart "Remocao de paginas vazias no final"
    RemoveEmptyPagesAtEnd doc
    LogStepComplete "Remocao de paginas vazias no final"

    LogStepStart "Aplicacao de formatacao final universal"
    ApplyUniversalFinalFormatting doc
    LogStepComplete "Aplicacao de formatacao final universal"

    LogStepStart "Adicao de espacamento especial (ementa, justificativa, data)"
    AddSpecialElementsSpacing doc
    LogStepComplete "Adicao de espacamento especial (ementa, justificativa, data)"

    LogStepStart "Ajuste final de recuos para Vereador (travessoes)"
    FixHyphenatedVereadorParagraphIndents doc
    LogStepComplete "Ajuste final de recuos para Vereador (travessoes)"

    LogMessage "Formatacao completa aplicada com sucesso", LOG_LEVEL_INFO
    LogMetric "Total de paragrafos", doc.Paragraphs.count
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatacao: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' AJUSTE FINAL - Zera recuo de paragrafos com marcador de Vereador/Vereadora (travessoes)
' Ao final do processamento, se existirem paragrafos contendo exatamente essas
' strings, garante recuo a esquerda = 0.
'================================================================================
Public Sub FixHyphenatedVereadorParagraphIndents(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim counter As Long
    Dim fixedCount As Long

    counter = 0
    fixedCount = 0

    For Each para In doc.Paragraphs
        counter = counter + 1
        If counter Mod 30 = 0 Then DoEvents

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
            On Error Resume Next
            para.Range.ListFormat.RemoveNumbers
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

            fixedCount = fixedCount + 1
        End If
    Next para

    If fixedCount > 0 Then
        LogMessage "Recuos ajustados para " & fixedCount & " paragrafo(s) de Vereador/Vereadora", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao ajustar recuos de Vereador/Vereadora: " & Err.Description, LOG_LEVEL_WARNING
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

    ' Construcao ASCII-safe de ",neste municipio" (com acento via ChrW)
    Dim municipio As String
    municipio = "mun" & ChrW(237) & "cipio" ' municipio
    Dim suffix1 As String
    Dim suffix2 As String
    suffix1 = ",neste " & municipio
    suffix2 = ", neste " & municipio

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
' EMENTA - Remove aspas envolventes ("...", “...”, ‘...’, '...')
' Regras:
' - Remove aspas duplas ou simples do inicio e fim do paragrafo da ementa
' - Trata aspas ASCII (" ') e tipograficas (“ ” ‘ ’)
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
    If firstCh = ChrW(8220) And lastCh = ChrW(8221) Then isMatch = True   ' “ ... ”
    If firstCh = Chr(39) And lastCh = Chr(39) Then isMatch = True          ' ' ... '
    If firstCh = ChrW(8216) And lastCh = ChrW(8217) Then isMatch = True   ' ‘ ... ’

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
            para.Range.text = cleanText
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

    ' Identifica apenas o 2 paragrafo (considerando apenas paragrafos com texto)
    actualParaIndex = 0
    secondParaIndex = 0

    ' Cache do count para performance
    Dim paraCount As Long
    paraCount = doc.Paragraphs.count

    ' Encontra o 2 paragrafo com conteudo (pula vazios)
    For i = 1 To paraCount
        If i > paraCount Then Exit For ' Protecao dinamica

        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o paragrafo tem texto ou conteudo visual, conta como paragrafo valido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1

            ' Registra o indice do 2 paragrafo
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For ' Ja encontramos o 2 paragrafo
            End If
        End If

        ' Protecao expandida: processa ate 20 paragrafos para encontrar o 2
        If i > 20 Then Exit For
    Next i

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
                para.Range.text = "Requer" & Mid(paraFullText, 9) & vbCr
                LogMessage "Palavra inicial 'Solicita' substituida por 'Requer' no 2 paragrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Verifica se inicia com "Pede" (case insensitive)
        If Not wasReplaced And Len(paraFullText) >= 4 Then
            lowerStart = LCase(Left(paraFullText, 4))
            If lowerStart = "pede" Then
                para.Range.text = "Requer" & Mid(paraFullText, 5) & vbCr
                LogMessage "Palavra inicial 'Pede' substituida por 'Requer' no 2 paragrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Verifica se inicia com "Sugere" (case insensitive)
        If Not wasReplaced And Len(paraFullText) >= 6 Then
            lowerStart = LCase(Left(paraFullText, 6))
            If lowerStart = "sugere" Then
                para.Range.text = "Indica" & Mid(paraFullText, 7) & vbCr
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
                para.Range.text = Left(paraFullText, Len(paraFullText) - 17) & vbCr
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
' HELPER FUNCTIONS FOR BLANK LINES - Funcoes auxiliares para linhas em branco
'================================================================================
' Nota: CountBlankLinesBefore ja esta definida nas linhas 918-958
' (secao de identificacao de estrutura do documento)

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

'================================================================================
' ENSURE SECOND PARAGRAPH BLANK LINES - Garante 2 linhas em branco no 2 paragrafo
'================================================================================
Public Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim secondParaIndex As Long
    Dim linesToAdd As Long
    Dim linesToAddAfter As Long

    secondParaIndex = GetSecondParagraphIndex(doc)
    linesToAdd = 0
    linesToAddAfter = 0

    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Dim para As Paragraph
        Set para = doc.Paragraphs(secondParaIndex)

        ' Verifica e corrige linhas em branco ANTES
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)

        If blankLinesBefore < 2 Then
            Dim insertionPoint As Range
            Set insertionPoint = para.Range
            insertionPoint.Collapse wdCollapseStart

            linesToAdd = 2 - blankLinesBefore

            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines

            ' Atualiza o indice (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If

        ' Verifica e corrige linhas em branco DEPOIS
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)

        If blankLinesAfter < 2 Then
            Dim insertionPointAfter As Range
            Set insertionPointAfter = para.Range
            insertionPointAfter.Collapse wdCollapseEnd

            linesToAddAfter = 2 - blankLinesAfter

            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If

        LogMessage "Linhas em branco do 2 paragrafo reforcadas (antes: " & (blankLinesBefore + linesToAdd) & ", depois: " & (blankLinesAfter + linesToAddAfter) & ")", LOG_LEVEL_INFO
    End If

    EnsureSecondParagraphBlankLines = True
    Exit Function

ErrorHandler:
    EnsureSecondParagraphBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do 2 paragrafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ENSURE PLENARIO BLANK LINES - Garante 2 linhas em branco antes e depois do Plenario
'================================================================================
Public Function EnsurePlenarioBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim paraTextCmp As String
    Dim i As Long
    Dim plenarioIndex As Long

    plenarioIndex = 0

    ' Localiza o paragrafo "Plenario Dr. Tancredo Neves"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            paraTextCmp = NormalizeForComparison(paraText)

            ' Procura por "Plenario" e "Tancredo Neves"
                If InStr(paraTextCmp, "plenario") > 0 And _
                    InStr(paraTextCmp, "tancredo") > 0 And _
                    InStr(paraTextCmp, "neves") > 0 Then
                plenarioIndex = i
                Exit For
            End If
        End If
    Next i

    If plenarioIndex > 0 Then
        ' Remove linhas vazias ANTES
        i = plenarioIndex - 1
        Do While i >= 1
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            If paraText = "" And Not HasVisualContent(para) Then
                para.Range.Delete
                plenarioIndex = plenarioIndex - 1
                i = i - 1
            Else
                Exit Do
            End If
        Loop

        ' Remove linhas vazias DEPOIS
        i = plenarioIndex + 1
        Do While i <= doc.Paragraphs.count
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            If paraText = "" And Not HasVisualContent(para) Then
                para.Range.Delete
            Else
                Exit Do
            End If
        Loop

        ' Insere EXATAMENTE 2 linhas em branco ANTES
        Set para = doc.Paragraphs(plenarioIndex)
        para.Range.InsertParagraphBefore
        para.Range.InsertParagraphBefore

        ' Insere EXATAMENTE 2 linhas em branco DEPOIS
        Set para = doc.Paragraphs(plenarioIndex + 2) ' +2 porque inserimos 2 antes
        para.Range.InsertParagraphAfter
        para.Range.InsertParagraphAfter

        LogMessage "Linhas em branco do Plenario reforcadas: 2 antes e 2 depois", LOG_LEVEL_INFO
    End If

    EnsurePlenarioBlankLines = True
    Exit Function

ErrorHandler:
    EnsurePlenarioBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do Plenario: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ENSURE SINGLE BLANK LINE BETWEEN PARAGRAPHS - Garante pelo menos 1 linha em branco entre paragrafos
'================================================================================
Public Function EnsureSingleBlankLineBetweenParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim para As Paragraph
    Dim NextPara As Paragraph
    Dim paraText As String
    Dim nextParaText As String
    Dim insertionPoint As Range
    Dim addedCount As Long

    addedCount = 0

    ' Percorre todos os paragrafos de tras para frente para nao afetar os indices
    For i = doc.Paragraphs.count - 1 To 1 Step -1
        Set para = doc.Paragraphs(i)
        Set NextPara = doc.Paragraphs(i + 1)

        ' Obtem texto limpo dos paragrafos
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

        ' Se ambos os paragrafos tem conteudo (texto ou imagem)
        If (paraText <> "" Or HasVisualContent(para)) And _
           (nextParaText <> "" Or HasVisualContent(NextPara)) Then

            ' Verifica se ha pelo menos uma linha em branco entre eles
            Dim hasBlankBetween As Boolean
            hasBlankBetween = False

            ' Verifica se o proximo paragrafo e imediatamente adjacente
            ' Isso seria indicado se nao ha paragrafo vazio entre eles
            If i + 1 <= doc.Paragraphs.count Then
                ' Se o indice do proximo paragrafo e i+1, eles sao adjacentes
                ' e precisamos verificar se ha linha em branco
                Dim checkIndex As Long
                For checkIndex = i + 1 To i + 1
                    If checkIndex <= doc.Paragraphs.count Then
                        Dim checkPara As Paragraph
                        Set checkPara = doc.Paragraphs(checkIndex)
                        Dim checkText As String
                        checkText = Trim(Replace(Replace(checkPara.Range.text, vbCr, ""), vbLf, ""))

                        ' Se o paragrafo entre eles esta vazio, ha linha em branco
                        If checkText = "" And Not HasVisualContent(checkPara) Then
                            hasBlankBetween = True
                        End If
                    End If
                Next checkIndex
            End If

            ' Se nao ha linha em branco, adiciona uma
            If Not hasBlankBetween Then
                Set insertionPoint = NextPara.Range
                insertionPoint.Collapse wdCollapseStart
                insertionPoint.InsertBefore vbCrLf
                addedCount = addedCount + 1
            End If
        End If
    Next i

    If addedCount > 0 Then
        LogMessage "Linhas em branco adicionadas entre paragrafos: " & addedCount, LOG_LEVEL_INFO
    End If

    EnsureSingleBlankLineBetweenParagraphs = True
    Exit Function

ErrorHandler:
    EnsureSingleBlankLineBetweenParagraphs = False
    LogMessage "Erro ao garantir linhas em branco entre paragrafos: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ENSURE BLANK LINES BETWEEN LONG PARAGRAPHS - Garante linha em branco entre paragrafos com mais de 10 palavras
'================================================================================
Public Function EnsureBlankLinesBetweenLongParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim para As Paragraph
    Dim NextPara As Paragraph
    Dim paraText As String
    Dim nextParaText As String
    Dim paraWordCount As Long
    Dim nextParaWordCount As Long
    Dim insertionPoint As Range
    Dim addedCount As Long

    addedCount = 0

    ' Percorre todos os paragrafos de tras para frente para nao afetar os indices
    For i = doc.Paragraphs.count - 1 To 1 Step -1
        If i >= doc.Paragraphs.count Then Exit For ' Protecao dinamica

        Set para = doc.Paragraphs(i)

        ' Verifica se ha proximo paragrafo
        If i + 1 <= doc.Paragraphs.count Then
            Set NextPara = doc.Paragraphs(i + 1)
        Else
            Exit For
        End If

        ' Obtem texto limpo dos paragrafos
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

        ' Conta palavras (divide por espacos)
        paraWordCount = 0
        nextParaWordCount = 0

        If paraText <> "" Then
            paraWordCount = UBound(Split(paraText, " ")) + 1
        End If

        If nextParaText <> "" Then
            nextParaWordCount = UBound(Split(nextParaText, " ")) + 1
        End If

        ' Se ambos os paragrafos tem mais de 10 palavras
        If paraWordCount > 10 And nextParaWordCount > 10 Then
            ' Verifica se ha linha em branco entre eles
            Dim hasBlankBetween As Boolean
            hasBlankBetween = False

            ' Verifica se eles sao adjacentes (sem linha em branco entre)
            ' Se i+1 e o proximo paragrafo e nao esta vazio, sao adjacentes
            If nextParaText <> "" Then
                hasBlankBetween = False
            Else
                hasBlankBetween = True
            End If

            ' Se nao ha linha em branco, adiciona uma
            If Not hasBlankBetween Then
                Set insertionPoint = NextPara.Range
                insertionPoint.Collapse wdCollapseStart
                insertionPoint.InsertBefore vbCrLf
                addedCount = addedCount + 1
            End If
        End If
    Next i

    If addedCount > 0 Then
        LogMessage "Linhas em branco adicionadas entre paragrafos longos (>10 palavras): " & addedCount, LOG_LEVEL_INFO
    End If

    EnsureBlankLinesBetweenLongParagraphs = True
    Exit Function

ErrorHandler:
    EnsureBlankLinesBetweenLongParagraphs = False
    LogMessage "Erro ao garantir linhas em branco entre paragrafos longos: " & Err.Description, LOG_LEVEL_WARNING
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

            ' Insere "Pag. X de Y" centralizado (numero da pagina e total de paginas)
            Set rngPage = footer.Range.Paragraphs.Last.Range
            rngPage.Collapse Direction:=wdCollapseStart
            rngPage.text = "P" & Chr(225) & "g. "
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
' LIMPA TEXTO PARA COMPARACAO
'================================================================================
Public Function CleanTextForComparison(text As String) As String
    On Error Resume Next
    CleanTextForComparison = text

    Dim result As String
    result = text

    ' Remove quebras de linha
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")

    ' Normaliza variacoes de caracteres
    result = Replace(result, Chr(160), " ")  ' Non-breaking space

    ' Remove multiplos espacos
    Dim counter As Long
    counter = 0
    Do While InStr(result, "  ") > 0 And counter < 100
        result = Replace(result, "  ", " ")
        counter = counter + 1
    Loop

    CleanTextForComparison = Trim(result)
End Function

'================================================================================
' VERIFICACAO DE DADOS SENSIVEIS (LGPD) - MODO ESTRITO
'================================================================================
' Objetivo:
' - Reduzir a checagem a achados realmente graves
' - Maximizar precisao usando validadores deterministicos quando possivel
'   (ex.: CPF/CNPJ com digitos verificadores, cartao com Luhn, CID com padrao estrito)
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

        ' Verifica se a linha termina com o padrao desejado
        If IsPageNumberLine(cleanText) Then
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
            ' Se a ultima palavra e numero, contem barra, ou e "N", "N.", "N" & Chr(186), remove.
            If IsNumeric(lastW) Or InStr(lastW, "/") > 0 Or UCase(lastW) = "N" Or UCase(lastW) = "N." Or UCase(lastW) = "N" & Chr(186) Or UCase(lastW) = "NO" Then
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

    ' SEMPRE aplica formatacao de titulo: caixa alta, negrito, sublinhado
    firstPara.Range.text = UCase(newText)

    ' Formatacao completa do titulo (primeira linha)
    With firstPara.Range.Font
        .Bold = True
        .Underline = wdUnderlineSingle
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

    ' Funcionalidade 16: "Área Pública" e "Roçagem" sempre em minusculas
    Dim areaPublicaCount As Long
    Dim rocagemCount As Long
    areaPublicaCount = ExecuteFindReplace(doc, "Área Pública", "área pública", False)
    areaPublicaCount = areaPublicaCount + ExecuteFindReplace(doc, "Area Publica", "área pública", False)
    rocagemCount = ExecuteFindReplace(doc, "Roçagem", "roçagem", False)
    rocagemCount = rocagemCount + ExecuteFindReplace(doc, "Rocagem", "roçagem", False)
    If areaPublicaCount > 0 Or rocagemCount > 0 Then
        LogMessage "Substituicao aplicada: 'Área Pública' e 'Roçagem' em minusculas (" & (areaPublicaCount + rocagemCount) & "x)", LOG_LEVEL_INFO
    End If

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
        replaceRng.text = newPhrase

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
        replaceRng.text = newText

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

            ' Adiciona apenas Justificativa e Anexo (Vereador nao recebe negrito)
            If cleanText = JUSTIFICATIVA_TEXT Or _
               IsAnexoPattern(cleanText) Then
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
        If pCleanText = JUSTIFICATIVA_TEXT Then
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
'================================================================================
Public Sub FormatVereadorParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim NextPara As Paragraph
    Dim i As Long
    Dim formattedCount As Long

    formattedCount = 0

    ' Procura por paragrafos com "vereador"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        ' OBS: O paragrafo pode conter pontuacao/travessoes/hifens.
        ' A deteccao abaixo ignora tudo que nao for letra e valida se sobrou apenas "vereador".
        If IsVereadorPattern(para.Range.text) Then
            ApplyVereadorParagraphFormatting para

            ' Formata linha ACIMA (se existir): centraliza, zera recuo, aplica caixa alta e negrito (somente se nao houver conteudo visual)
            If i > 1 Then
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

    If formattedCount > 0 Then
        LogMessage "Formatacao 'Vereador': " & formattedCount & " ocorrencias formatadas", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar paragrafos 'Vereador': " & Err.Description, LOG_LEVEL_ERROR
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

    ' Remove formatacao de lista (causa comum de recuo 1,25cm)
    On Error Resume Next
    para.Range.ListFormat.RemoveNumbers
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

    ' FASE 7: Processa "Excelentissimo Senhor Prefeito Municipal,"
    Dim prefeitoIndex As Long

    prefeitoIndex = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            paraTextLower = LCase(paraText)

            ' Procura por "Excelentissimo Senhor Prefeito Municipal" (case insensitive)
            If InStr(paraTextLower, "excelentissimo") > 0 And _
               InStr(paraTextLower, "senhor") > 0 And _
               InStr(paraTextLower, "prefeito") > 0 And _
               InStr(paraTextLower, "municipal") > 0 Then
                prefeitoIndex = i
                Exit For
            End If
        End If
    Next i

    If prefeitoIndex > 0 Then
        ' Remove linhas vazias depois e insere exatamente 2
        RemoveBlankLinesAfter doc, prefeitoIndex
        InsertBlankLinesAfter doc, prefeitoIndex, 2

        LogMessage "2 linhas em branco inseridas apos 'Excelentissimo Senhor Prefeito Municipal,'", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas em branco: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FUNCOES AUXILIARES PARA DETECCAO DE PADROES
'================================================================================
Public Function IsVereadorPattern(text As String) As Boolean
    IsVereadorPattern = (GetVereadorNormalizedWord(text) <> "")
End Function

Public Function GetVereadorNormalizedWord(text As String) As String
    Dim cleanText As String

    cleanText = Replace(Replace(text, vbCr, ""), vbLf, "")
    cleanText = Trim$(cleanText)
    cleanText = NormalizeLettersOnly(cleanText)

    ' Retorna com travessoes (em dash) e sem espacos antes/depois da expressao
    If cleanText = "vereador" Then
        GetVereadorNormalizedWord = ChrW(8212) & " Vereador " & ChrW(8212)
    ElseIf cleanText = "vereadora" Then
        GetVereadorNormalizedWord = ChrW(8212) & " Vereadora " & ChrW(8212)
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
Public Sub RestaurarBackup()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    If doc Is Nothing Then
        MsgBox "Nenhum documento ativo para restaurar.", vbExclamation, "Z7_STDPROPOSERS - Restaurar Backup"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Localiza todos os backups do documento na pasta de backups
    Dim backupFolder As String
    backupFolder = GetZ7StdProposersBackupsPath()

    If Not fso.FolderExists(backupFolder) Then
        MsgBox "Pasta de backups nao encontrada." & vbCrLf & vbCrLf & _
               "[i] O backup e criado apenas apos a primeira execucao de PadronizarDocumentoMain.", _
               vbExclamation, "Z7_STDPROPOSERS - Restaurar Backup"
        Exit Sub
    End If

    Dim docBaseName As String
    Dim docExtension As String
    docBaseName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)

    ' Coleta nomes de todos os backups do documento (nome = yyyy-mm-dd_HHmmss → ordem cronologica)
    Dim folder As Object
    Set folder = fso.GetFolder(backupFolder)

    Dim prefix As String
    prefix = LCase(docBaseName & "_backup_")

    Dim fileItem As Object
    Dim items() As String
    Dim itemCount As Long
    itemCount = 0

    For Each fileItem In folder.Files
        If Left(LCase(fileItem.Name), Len(prefix)) = prefix Then
            ReDim Preserve items(itemCount)
            items(itemCount) = fileItem.Name & "|" & fileItem.Path
            itemCount = itemCount + 1
        End If
    Next fileItem

    If itemCount = 0 Then
        MsgBox "Nenhum backup disponivel para o documento '" & doc.Name & "'." & vbCrLf & vbCrLf & _
               "[i] O backup e criado apenas apos a primeira execucao de PadronizarDocumentoMain.", _
               vbExclamation, "Z7_STDPROPOSERS - Restaurar Backup"
        Exit Sub
    End If

    ' Ordena alfanumericamente por nome: timestamp ISO -> o item[0] e o mais antigo
    Dim i As Long, j As Long, temp As String
    For i = 0 To itemCount - 2
        For j = i + 1 To itemCount - 1
            If items(i) > items(j) Then
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next j
    Next i

    Dim oldestParts() As String
    oldestParts = Split(items(0), "|")

    Dim targetBackupPath As String
    Dim targetBackupName As String
    targetBackupPath = oldestParts(1)
    targetBackupName = oldestParts(0)

    ' Confirma com usuario
    Dim confirmMsg As String
    confirmMsg = "[?] Deseja restaurar o backup mais antigo do documento?" & vbCrLf & vbCrLf & _
                 "[!] ATENCAO: O documento atual sera descartado!" & vbCrLf & vbCrLf & _
                 "[DIR] Documento atual: " & doc.Name & vbCrLf & _
                 "[DIR] Backup (mais antigo): " & targetBackupName

    If itemCount > 1 Then
        confirmMsg = confirmMsg & vbCrLf & _
                     "[i] Total de backups disponiveis: " & itemCount
    End If

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Z7_STDPROPOSERS - Confirmar Restauracao") <> vbYes Then
        Exit Sub
    End If

    Dim originalPath As String
    Dim originalName As String
    Dim discardedPath As String
    Dim timeStamp As String

    originalPath = doc.FullName
    originalName = doc.Name

    ' Cria timestamp para arquivo descartado
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")

    ' Nome do arquivo descartado: nome_discarded_timestamp.ext
    Dim baseName As String
    Dim extension As String
    baseName = fso.GetBaseName(originalName)
    extension = fso.GetExtensionName(originalName)

    discardedPath = fso.GetParentFolderName(originalPath) & "\" & _
                    baseName & "_discarded_" & timeStamp & "." & extension

    ' Protege contra conflito: exclui arquivo pre-existente
    If fso.FileExists(discardedPath) Then
        fso.DeleteFile discardedPath, True
    End If

    ' Salva documento atual como _discarded
    Application.StatusBar = "Salvando documento descartado..."
    doc.SaveAs2 discardedPath

    ' Fecha o documento descartado
    doc.Close SaveChanges:=False

    ' Protege contra conflito no caminho original
    If fso.FileExists(originalPath) Then
        fso.DeleteFile originalPath, True
    End If

    ' Copia o backup mais antigo para o local original
    Application.StatusBar = "Restaurando backup..."
    fso.CopyFile targetBackupPath, originalPath, True

    ' Abre o backup restaurado
    Application.Documents.Open originalPath

    Application.StatusBar = "Backup restaurado com sucesso! (z7_stdproposers)"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao restaurar backup"
    MsgBox "[ERRO] Falha ao restaurar backup:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "[i] O documento pode estar em estado inconsistente." & vbCrLf & _
           "   Verifique manualmente a pasta de backups.", _
           vbCritical, "Z7_STDPROPOSERS - Erro na Restauracao"
End Sub

'================================================================================
' LIMPEZA DE BACKUPS ANTIGOS
'================================================================================
Public Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error GoTo CleanExit

    If MAX_BACKUP_FILES < 1 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(backupFolder) Then GoTo CleanExit

    Dim folder As Object
    Set folder = fso.GetFolder(backupFolder)

    Dim fileItem As Object
    Dim prefix As String
    prefix = LCase(docBaseName & "_backup_")

    Dim items() As String
    Dim itemCount As Long
    itemCount = 0

    For Each fileItem In folder.Files
        If Left(LCase(fileItem.Name), Len(prefix)) = prefix Then
            ReDim Preserve items(itemCount)
            items(itemCount) = Format(fileItem.DateLastModified, "yyyymmddHHMMSS") & "|" & fileItem.Path
            itemCount = itemCount + 1
        End If
    Next fileItem

    If itemCount <= MAX_BACKUP_FILES Then GoTo CleanExit

    Dim i As Long, j As Long, temp As String
    For i = 0 To itemCount - 2
        For j = i + 1 To itemCount - 1
            If items(i) < items(j) Then
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next j
    Next i

    Dim idx As Long
    For idx = MAX_BACKUP_FILES To itemCount - 1
        Dim parts() As String
        parts = Split(items(idx), "|")
        On Error Resume Next
        fso.DeleteFile parts(1), True
        If Err.Number <> 0 Then
            If loggingEnabled Then
                LogMessage "Failed to delete old backup: " & parts(1) & " - " & Err.Description, LOG_LEVEL_WARNING
            End If
            Err.Clear
        Else
            If loggingEnabled Then
                LogMessage "Old backup removed: " & parts(1), LOG_LEVEL_INFO
            End If
        End If
        On Error GoTo CleanExit
    Next idx

CleanExit:
    On Error Resume Next
    Set items = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

'================================================================================
' LIMPEZA DE ESPACOS MULTIPLOS
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
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim trimmed As String
    Dim digitLen As Long
    Dim sepChar As String
    Dim formattedCount As Long
    Dim prevWasFormatted As Boolean
    Dim paraCounter As Long
    Dim prefixLen As Long
    Dim prefixRange As Range

    formattedCount = 0
    prevWasFormatted = False
    paraCounter = 0

    For Each para In doc.Paragraphs
        paraCounter = paraCounter + 1
        If paraCounter Mod 30 = 0 Then DoEvents

        ' Remove o marcador de paragrafo (Chr(13)) do final para analise
        paraText = Left(para.Range.Text, Len(para.Range.Text) - 1)
        trimmed = LTrim(paraText)

        ' Tamanho minimo: digito + separador + espaco + um char ("1. x")
        If Len(trimmed) < 4 Or Not IsNumeric(Left(trimmed, 1)) Then
            prevWasFormatted = False
        ElseIf para.Range.ListFormat.ListType <> wdListNoNumbering Then
            ' Ja tem formatacao de lista; mantém continuidade para proximos paragrafos
            prevWasFormatted = True
        Else
            ' Conta os digitos iniciais (suporta numeros multi-digito: "10. ", "123. ")
            digitLen = 0
            Do While digitLen < Len(trimmed) And IsNumeric(Mid(trimmed, digitLen + 1, 1))
                digitLen = digitLen + 1
            Loop

            ' Verifica padrao: digitos seguidos de "." ou ")" e espaco (ex: "1. ", "10) ")
            If Len(trimmed) >= digitLen + 2 Then
                sepChar = Mid(trimmed, digitLen + 1, 1)
                If (sepChar = "." Or sepChar = ")") And Mid(trimmed, digitLen + 2, 1) = " " Then
                    ' Calcula o tamanho do prefixo a remover (espacos iniciais + digitos + sep + espaco)
                    prefixLen = (Len(paraText) - Len(trimmed)) + digitLen + 2
                    Set prefixRange = para.Range
                    prefixRange.SetRange prefixRange.Start, prefixRange.Start + prefixLen
                    prefixRange.Delete

                    ' Aplica lista numerada automatica do Word (galeria 1 = simples)
                    para.Range.ListFormat.ApplyListTemplate _
                        ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(1), _
                        ContinuePreviousList:=prevWasFormatted, _
                        ApplyTo:=wdListApplyToThisPointForward

                    prevWasFormatted = True
                    formattedCount = formattedCount + 1
                Else
                    prevWasFormatted = False
                End If
            Else
                prevWasFormatted = False
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Paragrafos formatados como lista numerada: " & formattedCount, LOG_LEVEL_INFO
    End If

    FormatNumberedParagraphsIndent = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar lista numerada: " & Err.Description, LOG_LEVEL_WARNING
    FormatNumberedParagraphsIndent = False
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

    ' --- Substituicao gigante manual (maior que 255 chars) ---
    Dim repRange As Range
    Set repRange = doc.Range
    With repRange.Find
        .ClearFormatting
        .text = "Que cabe ao Poder Legislativo, dispor sobre as"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        Do While .Execute
            repRange.Expand wdParagraph
            Dim foundText As String
            foundText = Replace(repRange.text, vbCr, "")
            If InStr(foundText, "financeira e or") > 0 And InStr(foundText, "para atender tal solicita") > 0 Then
                repRange.text = "Cabe ao Poder Legislativo dispor sobre as matérias de competência do Município, especialmente assuntos de interesse local. Compete-lhe também a função de fiscalização dos atos do Poder Executivo, abrangendo os atos administrativos, de gestão e fiscalização financeira e orçamentária do município." & vbCr & "Desta forma, faço esta indicação para o prefeito determinar ao setor competente realize os atos administrativos para atender tal solicitação." & vbCr
                replacedCount = replacedCount + 1
            End If
            repRange.Collapse wdCollapseEnd
        Loop
    End With

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

        .text = "Indica ao Poder Executivo Municipal e aos órgãos competentes"
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
    ' agora apontam para o limbo (desalinhados). Forçamos a reconstrução do cache antes de prosseguir.
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

        ' ZOOM e mantido em 120% - unica configuracao que permanece alterada
        .Zoom.Percentage = 120
    End With

    ' Configuracoes especificas do Window (para reguas)
    docWindow.DisplayRulers = originalViewSettings.ShowHorizontalRuler
    docWindow.DisplayVerticalRuler = originalViewSettings.ShowVerticalRuler

    LogMessage "Configuracoes de visualizacao originais restauradas (zoom mantido em 120%)"
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
        Exit Sub
    End If

    ' Caminho do instalador
    installerPath = Environ("USERPROFILE") & "\AppData\Local\Z7\Apps\Z7_StdProposers\installer.cmd"

    ' Verifica se o instalador existe
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(installerPath) Then
        MsgBox "Instalador nao encontrado em:" & vbCrLf & installerPath & vbCrLf & vbCrLf & _
               "Baixe manualmente de: https://github.com/chrmsantos/Z7_StdProposers/raw/main/z7_stdproposers_installer.cmd", _
               vbExclamation, "Z7_STDPROPOSERS - Instalador Nao Encontrado"
        Exit Sub
    End If

    ' Salva todos os documentos abertos antes de executar o instalador
    Dim doc As Object
    For Each doc In Application.Documents
        If doc.Saved = False Then
            On Error Resume Next
            doc.Save
            On Error GoTo ErrorHandler
        End If
    Next doc

    ' Executa o instalador em uma nova janela de comando
    shellCmd = "cmd.exe /c """ & installerPath & """"
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

    LogMessage "Adicionando espacamento especial para ementa, justificativa e data...", LOG_LEVEL_INFO

    ' Garante sem espaco antes e depois da Ementa
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(ementaParaIndex).Format
            .SpaceBefore = 0   ' Sem espaco antes
            .SpaceAfter = 0    ' Sem espaco depois
        End With
        If Err.Number = 0 Then elementsProcessed = elementsProcessed + 1
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' Garante sem espaco antes e depois do Titulo Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(tituloJustificativaIndex).Format
            .SpaceBefore = 0   ' Sem espaco antes
            .SpaceAfter = 0    ' Sem espaco depois
        End With
        If Err.Number = 0 Then elementsProcessed = elementsProcessed + 1
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' Garante sem espaco antes e depois da Data
    If dataParaIndex > 0 And dataParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(dataParaIndex).Format
            .SpaceBefore = 0   ' Sem espaco antes
            .SpaceAfter = 0    ' Sem espaco depois
        End With
        If Err.Number = 0 Then elementsProcessed = elementsProcessed + 1
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    LogMessage "Espacamento especial aplicado a " & elementsProcessed & " elementos", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao adicionar espacamento especial: " & Err.Description, LOG_LEVEL_WARNING
End Sub

