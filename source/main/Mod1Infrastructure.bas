Attribute VB_Name = "Mod1Infrastructure"
Option Explicit

' Mod1Infrastructure.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Versao: 7.7.7-rc2
' Data: 2026-05-19
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
'================================================================================
' CONSTANTES DO WORD
'================================================================================
Public Const wdNoProtection As Long = -1
Public Const wdTypeDocument As Long = 0
Public Const wdHeaderFooterPrimary As Long = 1
Public Const wdAlignParagraphLeft As Long = 0
Public Const wdAlignParagraphCenter As Long = 1
Public Const wdAlignParagraphJustify As Long = 3
Public Const wdLineSpaceSingle As Long = 0
Public Const wdLineSpace1pt5 As Long = 1
Public Const wdLineSpacingMultiple As Long = 5
Public Const wdStatisticPages As Long = 2
Public Const msoTrue As Long = -1
Public Const msoFalse As Long = 0
Public Const msoPicture As Long = 13
Public Const msoTextEffect As Long = 15
Public Const wdCollapseEnd As Long = 0
Public Const wdCollapseStart As Long = 1
Public Const wdFieldPage As Long = 33
Public Const wdFieldNumPages As Long = 26
Public Const wdFieldEmpty As Long = -1
Public Const wdRelativeHorizontalPositionPage As Long = 1
Public Const wdRelativeVerticalPositionPage As Long = 1
Public Const wdWrapTopBottom As Long = 3
Public Const wdAlertsAll As Long = -1
Public Const wdAlertsNone As Long = 0
Public Const wdColorAutomatic As Long = -16777216
Public Const wdOrientPortrait As Long = 0
Public Const wdUnderlineNone As Long = 0
Public Const wdUnderlineSingle As Long = 1
Public Const wdTextureNone As Long = 0
Public Const wdPrintView As Long = 3

'================================================================================
' CONSTANTES DE FORMATACAO
'================================================================================
Public Const STANDARD_FONT As String = "Arial"
Public Const STANDARD_FONT_SIZE As Long = 12
Public Const FOOTER_FONT_SIZE As Long = 10
Public Const LINE_SPACING As Single = 14

Public Const TOP_MARGIN_CM As Double = 4.85
Public Const BOTTOM_MARGIN_CM As Double = 2
Public Const LEFT_MARGIN_CM As Double = 3
Public Const RIGHT_MARGIN_CM As Double = 3
Public Const HEADER_DISTANCE_CM As Double = 0.3
Public Const FOOTER_DISTANCE_CM As Double = 0.9

Public Const HEADER_IMAGE_RELATIVE_PATH As String = "\AppData\Local\Z7\Apps\Z7_StdProposers\assets\stamp.png"
Public Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Public Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Public Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19



' Caminho para o executavel de configuracao do prompt do Gemini
Public Const PROMPT_CONFIG_SCRIPT_RELATIVE_PATH As String = "\AppData\Local\Z7\Apps\Z7_StdProposers\ai\config_prompt\config_prompt.exe"

' Caminho para o executavel de chat em tempo real com a IA
Public Const CHAT_IA_SCRIPT_RELATIVE_PATH As String = "\AppData\Local\Z7\Apps\Z7_StdProposers\ai\chat_ia\chat_ia.exe"



'================================================================================
' CONSTANTES DE SISTEMA
'================================================================================
Public Const Z7_STDPROPOSERS_VERSION As String = "7.9.3"
Public Const MIN_SUPPORTED_VERSION As Long = 14
Public Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"
Public Const MAX_BACKUP_FILES As Long = 10
Public Const DEBUG_MODE As Boolean = False

Public Const LOG_LEVEL_INFO As Long = 1
Public Const LOG_LEVEL_WARNING As Long = 2
Public Const LOG_LEVEL_ERROR As Long = 3
Public Const LOG_BUFFER_FLUSH_SECONDS As Long = 5

Public Const MAX_RETRY_ATTEMPTS As Long = 3
Public Const RETRY_DELAY_MS As Long = 1000
Public Const MAX_LOOP_ITERATIONS As Long = 1000
Public Const MAX_INITIAL_PARAGRAPHS_TO_SCAN As Long = 50
Public Const MAX_OPERATION_TIMEOUT_SECONDS As Long = 300

' Verificacao de atualizacao
Public Const UPDATE_CHECK_COOLDOWN_MINUTES As Long = 60

Public Const CONSIDERANDO_PREFIX As String = "considerando"
Public Const CONSIDERANDO_MIN_LENGTH As Long = 12
Public Const JUSTIFICATIVA_TEXT As String = "justificativa"

'================================================================================
' CONSTANTES DE IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS
'================================================================================
' VARIAVEIS GLOBAIS
'================================================================================
Public undoGroupEnabled As Boolean
Public loggingEnabled As Boolean
Public logFilePath As String
Public formattingCancelled As Boolean
Public executionStartTime As Date
Public backupFilePath As String
Public errorCount As Long
Public warningCount As Long
Public infoCount As Long
Public logFileHandle As Integer
Public logBufferEnabled As Boolean
Public logBuffer As String
Public lastFlushTime As Date
Public currentLogSessionId As String
Public currentOperationId As String

' Cache de verificacao de atualizacao (evita chamadas repetidas e travamentos)
Public lastUpdateCheckAttempt As Date
Public lastUpdateCheckSucceeded As Boolean
Public cachedUpdateAvailable As Boolean
Public cachedLocalVersion As String
Public cachedRemoteVersion As String

' Cache de paragrafos para otimizacao
Public Type paragraphCache
    index As Long
    text As String
    cleanText As String
    hasImages As Boolean
    isSpecial As Boolean
    specialType As String
    needsFormatting As Boolean
    ' Identificadores de elementos estruturais da propositura
    isTitulo As Boolean
    isEmenta As Boolean
    isVocativo As Boolean
    isProposicaoContent As Boolean
    isTituloJustificativa As Boolean
    isJustificativaContent As Boolean
    isData As Boolean
    isAssinatura As Boolean
    isTituloAnexo As Boolean
    isAnexoContent As Boolean
End Type

Public paragraphCache() As paragraphCache
Public cacheSize As Long
Public cacheEnabled As Boolean
Public documentDirty As Boolean  ' Flag para otimizar pipeline de 2 passagens

' Barra de progresso
Public totalSteps As Long
Public currentStep As Long

Public Type ImageInfo
    paraIndex As Long
    ImageIndex As Long
    ImageType As String
    ImageData As Variant
    Position As Long
    WrapType As Long
    Width As Single
    Height As Single
    LeftPosition As Single
    TopPosition As Single
    AnchorRange As Range
End Type

Public savedImages() As ImageInfo
Public imageCount As Long

Public Type ViewSettings
    ViewType As Long
    ShowVerticalRuler As Boolean
    ShowHorizontalRuler As Boolean
    ShowFieldCodes As Boolean
    ShowBookmarks As Boolean
    ShowParagraphMarks As Boolean
    ShowSpaces As Boolean
    ShowTabs As Boolean
    ShowHiddenText As Boolean
    ShowOptionalHyphens As Boolean
    ShowAll As Boolean
    ShowDrawings As Boolean
    ShowObjectAnchors As Boolean
    ShowTextBoundaries As Boolean
    ShowHighlight As Boolean
    DraftFont As Boolean
    WrapToWindow As Boolean
    ShowPicturePlaceHolders As Boolean
    ShowFieldShading As Long
    TableGridlines As Boolean
End Type

Public originalViewSettings As ViewSettings

Public Type ListFormatInfo
    paraIndex As Long
    HasList As Boolean
    ListType As Long
    ListLevelNumber As Long
    ListString As String
End Type

Public savedListFormats() As ListFormatInfo
Public listFormatCount As Long

' Backup de paragrafos centralizados antes do processamento
Public Type CenteredParaInfo
    paraIndex As Long
    originalText As String  ' Usado como chave de identificacao (primeiros 50 chars)
End Type

Public savedCenteredParas() As CenteredParaInfo
Public centeredParaCount As Long

'================================================================================
' VARIAVEIS DE IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS
'================================================================================
' Indices dos elementos identificados no documento (0 = nao encontrado)
Public tituloParaIndex As Long
Public ementaParaIndex As Long
Public vocativoStartIndex As Long
Public vocativoEndIndex As Long
Public proposicaoStartIndex As Long
Public proposicaoEndIndex As Long
Public tituloJustificativaIndex As Long
Public justificativaStartIndex As Long
Public justificativaEndIndex As Long
Public dataParaIndex As Long
Public assinaturaStartIndex As Long
Public assinaturaEndIndex As Long
Public tituloAnexoIndex As Long
Public anexoStartIndex As Long
Public anexoEndIndex As Long

'================================================================================
' GERENCIAMENTO DE ESTADO DA APLICACAO
'================================================================================
Public Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "", Optional ByVal preserveStatusBar As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    Dim success As Boolean
    success = True

    With Application
        On Error Resume Next
        .ScreenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler

        On Error Resume Next
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler

        ' Nao modifica StatusBar se preserveStatusBar = True
        If Not preserveStatusBar Then
            If statusMsg <> "" Then
                On Error Resume Next
                .StatusBar = statusMsg
                If Err.Number <> 0 Then success = False
                On Error GoTo ErrorHandler
            ElseIf enabled Then
                On Error Resume Next
                .StatusBar = False
                If Err.Number <> 0 Then success = False
                On Error GoTo ErrorHandler
            End If
        End If

        On Error Resume Next
        .EnableCancelKey = 0
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
    End With

    SetAppState = success
    Exit Function

ErrorHandler:
    SetAppState = False
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW - CONFIGURACAO DE VISUALIZACAO
'================================================================================
Public Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Configurando visualizacao..."

    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow

    ' Configura APENAS o zoom para 130% - todas as outras configuracoes sao preservadas
    With docWindow.View
        .Zoom.Percentage = 130
        ' NAO altera mais o tipo de visualizacao - preserva o original
    End With

    ' Remove configuracoes que alteravam configuracoes globais do Word
    ' Estas configuracoes sao agora preservadas do estado original

    LogMessage "Visualizacao configurada: zoom definido para 130%, demais configuracoes preservadas"
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao configurar visualizacao: " & Err.Description, LOG_LEVEL_WARNING
    ConfigureDocumentView = False ' Nao falha o processo por isso
End Function

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM - SISTEMA DE PROTECAO DAS CONFIGURACOES DE VISUALIZACAO
'================================================================================


'================================================================================
' TRATAMENTO AMIGAVEL DE ERROS
'================================================================================
Public Sub ShowUserFriendlyError(errNum As Long, errDesc As String)
    Dim msg As String

    Select Case errNum
        Case 91 ' Object variable not set
            msg = "Erro: Objeto nao inicializado." & vbCrLf & vbCrLf & _
                  "Reinicie o Word."

        Case 5 ' Invalid procedure call
            msg = "Erro de configuracao." & vbCrLf & vbCrLf & _
                  "Formato valido: .docx"

        Case 70 ' Permission denied
            msg = "Permissao negada." & vbCrLf & vbCrLf & _
                  "Documento protegido ou somente leitura." & vbCrLf & _
                  "Salve uma copia."

        Case 53 ' File not found
            msg = "Arquivo nao encontrado." & vbCrLf & vbCrLf & _
                  "Verifique se foi salvo."

        Case Else
            msg = "Erro #" & errNum & ":" & vbCrLf & vbCrLf & _
                  errDesc & vbCrLf & vbCrLf & _
                  "Verifique o log."
    End Select

    MsgBox msg, vbCritical, "Z7_StdProposers v" & Z7_STDPROPOSERS_VERSION
End Sub

'================================================================================
' RECUPERACAO DE EMERGENCIA
'================================================================================
Public Sub EmergencyRecovery()
    On Error Resume Next

    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0

    ' Fecha UndoRecord se ainda estiver aberto
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "UndoRecord fechado durante recuperacao de emergencia", LOG_LEVEL_WARNING
    End If

    ' Limpa variaveis de protecao de imagens em caso de erro
    CleanupImageProtection

    ' Limpa variaveis de configuracoes de visualizacao em caso de erro
    CleanupViewSettings

    ' Limpa cache de paragrafos
    ClearParagraphCache

    LogMessage "Recuperacao de emergencia executada", LOG_LEVEL_ERROR

    CloseAllOpenFiles
End Sub

'================================================================================
' ATUALIZACAO DA BARRA DE PROGRESSO
'================================================================================
Public Sub UpdateProgress(message As String, percentComplete As Long)
    If message <> "" Then
        Application.StatusBar = "Padronizando... " & message
    Else
        Application.StatusBar = "Padronizando..."
    End If
    DoEvents
End Sub

'================================================================================
' SALVAMENTO INICIAL DO DOCUMENTO
'================================================================================
Public Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Salvando documento..."
    ' Log de inicio removido para performance

    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        LogMessage "Operacao de salvamento cancelada pelo usuario", LOG_LEVEL_INFO
        Application.StatusBar = "Cancelado"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Aguarda confirmacao do salvamento com timeout de seguranca
    Dim waitCount As Integer
    Dim maxWait As Integer
    maxWait = 10

    For waitCount = 1 To maxWait
        DoEvents
        If doc.Path <> "" Then Exit For
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
        Application.StatusBar = "Salvando... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.Path = "" Then
        LogMessage "Falha ao salvar documento apos " & maxWait & " tentativas", LOG_LEVEL_ERROR
        Application.StatusBar = "Falha ao salvar"
        SaveDocumentFirst = False
    Else
        ' Log de sucesso removido para performance
        Application.StatusBar = "Salvo"
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro durante salvamento: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    Application.StatusBar = "Erro ao salvar"
    SaveDocumentFirst = False
End Function

'================================================================================
' SISTEMA DE BACKUP
'================================================================================
Public Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Nao faz backup se documento nao foi realmente salvo (nao existe no disco)
    If doc.Path = "" Or Not fso.FileExists(doc.FullName) Then
        LogMessage "Backup ignorado - documento nao salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If

    Dim backupFolder As String
    Dim docName As String
    Dim docExtension As String
    Dim timeStamp As String
    Dim backupFileName As String

    ' Usa a funcao que garante o diretorio de backup
    backupFolder = EnsureBackupDirectory(doc)

    ' Extrai nome e extensao do documento
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)

    ' Cria timestamp para o backup
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")

    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timeStamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName

    ' Protege contra conflito: exclui arquivo pre-existente com mesmo nome
    If fso.FileExists(backupFilePath) Then
        fso.DeleteFile backupFilePath, True
        LogMessage "Backup anterior com mesmo nome excluido: " & backupFileName, LOG_LEVEL_INFO
    End If

    ' Salva uma copia do documento como backup
    Application.StatusBar = "Criando backup..."

    ' Salva o documento atual primeiro para garantir que esta atualizado
    doc.Save

    ' Cria uma copia do arquivo usando FileSystemObject
    fso.CopyFile doc.FullName, backupFilePath, True

    ' Limpa backups antigos se necessario
    CleanOldBackups backupFolder, docName

    LogMessage "Backup criado com sucesso: " & backupFileName, LOG_LEVEL_INFO
    Application.StatusBar = "Backup criado"

    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar backup: " & Err.Description, LOG_LEVEL_ERROR
    CreateDocumentBackup = False
End Function

'================================================================================
' FUNCOES DE CAMINHO - Estrutura do projeto
'================================================================================

Public Function GetProjectRootPath() As String
    GetProjectRootPath = Environ("USERPROFILE") & "\AppData\Local\Z7\Apps\Z7_StdProposers"
End Function

Public Function GetZ7StdProposersBackupsPath() As String
    GetZ7StdProposersBackupsPath = Environ("USERPROFILE") & "\AppData\Local\Temp"
End Function

Public Function GetZ7StdProposersRecoveryPath() As String
    GetZ7StdProposersRecoveryPath = GetProjectRootPath() & "\props\recovery_tmp"
End Function

Public Function GetZ7StdProposersLogsPath() As String
    GetZ7StdProposersLogsPath = GetProjectRootPath() & "\source\logs"
End Function

Public Sub EnsureZ7StdProposersFolders()
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim propsPath As String
    propsPath = GetProjectRootPath() & "\props"

    Dim sourcePath As String
    sourcePath = GetProjectRootPath() & "\source"

    If Not fso.FolderExists(propsPath) Then fso.CreateFolder propsPath
    If Not fso.FolderExists(sourcePath) Then fso.CreateFolder sourcePath

    If Not fso.FolderExists(GetZ7StdProposersBackupsPath()) Then fso.CreateFolder GetZ7StdProposersBackupsPath()
    If Not fso.FolderExists(GetZ7StdProposersRecoveryPath()) Then fso.CreateFolder GetZ7StdProposersRecoveryPath()
    If Not fso.FolderExists(GetZ7StdProposersLogsPath()) Then fso.CreateFolder GetZ7StdProposersLogsPath()

    Set fso = Nothing
End Sub

'================================================================================
' GERENCIAMENTO DE DIRETORIO DE BACKUP
'================================================================================
Public Function EnsureBackupDirectory(doc As Document) As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim backupPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Garante que a estrutura de pastas do projeto existe
    EnsureZ7StdProposersFolders

    ' SEMPRE USA %TEMP%\.z7_stdproposers\props\backups para todos os documentos
    backupPath = GetZ7StdProposersBackupsPath()

    ' Cria o diretorio se nao existir
    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
        LogMessage "Pasta de backup criada: " & backupPath, LOG_LEVEL_INFO
    End If

    EnsureBackupDirectory = backupPath
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar pasta de backup: " & Err.Description, LOG_LEVEL_ERROR
    ' Retorna pasta do documento ou TEMP como fallback
    If doc.Path <> "" Then
        EnsureBackupDirectory = doc.Path
    Else
        EnsureBackupDirectory = Environ("TEMP")
    End If
End Function

'================================================================================
' VERIFICACAO DE VERSAO E ATUALIZACAO
'================================================================================

' Funcao: CheckForUpdates
' Descricao: Verifica se ha uma nova versao disponivel no GitHub
' Retorna: True se houver atualizacao disponivel, False caso contrario
'================================================================================
Public Function CheckForUpdates(Optional forceCheck As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    Dim localVersion As String
    Dim remoteVersion As String
    Dim updateAvailable As Boolean

    CheckForUpdates = False
    
    LogMessage "Iniciando verificacao de atualizacao...", LOG_LEVEL_INFO

    ' Nao executar verificacao durante operacao critica (ex.: padronizacao em andamento)
    If undoGroupEnabled Then
        LogMessage "Verificacao de atualizacao ignorada: operacao de padronizacao em andamento", LOG_LEVEL_INFO
        Exit Function
    End If

    ' Cache: se ja checou com sucesso nesta sessao, reusa o resultado
    If Not forceCheck Then
        If lastUpdateCheckAttempt <> 0 Then
            If lastUpdateCheckSucceeded Then
                LogMessage "Verificacao de atualizacao ignorada: resultado em cache utilizado (v" & cachedRemoteVersion & ")", LOG_LEVEL_INFO
                CheckForUpdates = cachedUpdateAvailable
                Exit Function
            End If

            ' Se a ultima tentativa falhou recentemente, evita repetir (reduz chance de travamentos)
            If DateDiff("n", lastUpdateCheckAttempt, Now) < UPDATE_CHECK_COOLDOWN_MINUTES Then
                LogMessage "Verificacao de atualizacao ignorada devido ao cooldown de " & UPDATE_CHECK_COOLDOWN_MINUTES & " minutos", LOG_LEVEL_INFO
                CheckForUpdates = cachedUpdateAvailable
                Exit Function
            End If
        End If
    End If

    lastUpdateCheckAttempt = Now

    ' Obtem versao local
    localVersion = GetLocalVersion()
    If localVersion = "" Then
        LogMessage "Nao foi possivel obter a versao local do arquivo VERSION", LOG_LEVEL_WARNING
        lastUpdateCheckSucceeded = False
        Exit Function
    End If

    cachedLocalVersion = localVersion
    LogMessage "Versao local detectada: v" & localVersion, LOG_LEVEL_INFO

    ' Obtem versao remota do GitHub
    LogMessage "Buscando versao remota no GitHub...", LOG_LEVEL_INFO
    remoteVersion = GetRemoteVersion()
    If remoteVersion = "" Then
        LogMessage "Nao foi possivel obter a versao remota a partir do GitHub", LOG_LEVEL_WARNING
        lastUpdateCheckSucceeded = False
        cachedUpdateAvailable = False
        Exit Function
    End If

    cachedRemoteVersion = remoteVersion
    lastUpdateCheckSucceeded = True
    LogMessage "Versao remota detectada: v" & remoteVersion, LOG_LEVEL_INFO

    ' Compara versoes
    updateAvailable = CompareVersions(remoteVersion, localVersion) > 0
    cachedUpdateAvailable = updateAvailable

    If updateAvailable Then
        LogMessage "Nova atualizacao disponivel: v" & localVersion & " -> v" & remoteVersion, LOG_LEVEL_INFO
        CheckForUpdates = True
    Else
        LogMessage "Nenhuma atualizacao necessaria. O sistema esta atualizado (v" & localVersion & ")", LOG_LEVEL_INFO
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar atualizacoes: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    lastUpdateCheckSucceeded = False
    CheckForUpdates = False
End Function

' Funcao: GetLocalVersion
' Descricao: Le a versao instalada do arquivo VERSION local
' Retorna: String com a versao local ou "" em caso de erro
'================================================================================
Public Function GetLocalVersion() As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim versionFile As String
    Dim fileContent As String
    Dim version As String

    GetLocalVersion = ""

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Caminho do arquivo de versao local
    versionFile = GetProjectRootPath() & "\VERSION"

    If Not fso.FileExists(versionFile) Then
        LogMessage "Arquivo de versao local nao encontrado: " & versionFile, LOG_LEVEL_WARNING
        Exit Function
    End If

    ' Le o arquivo
    fileContent = ReadTextFile(versionFile)

    ' Extrai versao (X.Y.Z)
    version = ExtractVersionFromText(fileContent)

    GetLocalVersion = version

    Exit Function

ErrorHandler:
    LogMessage "Erro ao obter versao local: " & Err.Description, LOG_LEVEL_ERROR
    GetLocalVersion = ""
End Function

' Funcao: GetRemoteVersion
' Descricao: Baixa e le a versao disponivel no GitHub
' Retorna: String com a versao remota ou "" em caso de erro
'================================================================================
Public Function GetRemoteVersion() As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim response As String
    Dim version As String
    Dim statusCode As Long
    Dim usedServerHttp As Boolean

    GetRemoteVersion = ""

    ' URL do arquivo VERSION no GitHub
    url = "https://raw.githubusercontent.com/chrmsantos/Z7_StdProposers/main/VERSION"

    ' Cria objeto HTTP com timeout quando possivel (evita travamentos em rede lenta/bloqueada)
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number <> 0 Or http Is Nothing Then
        Err.Clear
        Set http = CreateObject("MSXML2.XMLHTTP")
    Else
        usedServerHttp = True
    End If
    On Error GoTo ErrorHandler

    ' Faz requisicao GET
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"

    ' Alguns MSXML podem falhar no header User-Agent; nao e critico
    On Error Resume Next
    http.setRequestHeader "User-Agent", "Z7_STDPROPOSERS/" & Z7_STDPROPOSERS_VERSION
    ' Aplica timeouts condicionalmente se ServerXMLHTTP estiver em uso
    If usedServerHttp Then
        http.setTimeouts 5000, 5000, 10000, 10000
    End If
    On Error GoTo ErrorHandler

    http.send

    statusCode = 0
    On Error Resume Next
    statusCode = CLng(http.Status)
    On Error GoTo ErrorHandler

    ' Verifica resposta
    If statusCode = 200 Then
        response = CStr(http.responseText)
        version = ExtractVersionFromText(response)
        If version <> "" Then
            GetRemoteVersion = version
        Else
            LogMessage "Resposta remota sem versao valida", LOG_LEVEL_WARNING
        End If
    Else
        If statusCode = 0 Then
            LogMessage "Falha ao buscar versao remota (sem status HTTP)", LOG_LEVEL_WARNING
        Else
            LogMessage "Erro HTTP ao buscar versao remota: " & CStr(statusCode), LOG_LEVEL_WARNING
        End If
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao obter versao remota: " & Err.Description, LOG_LEVEL_ERROR
    GetRemoteVersion = ""
End Function

' Funcao: ExtractVersionFromText
' Descricao: Extrai uma versao (X.Y.Z) de um texto usando regex
' Parametros:
'   - textValue: String contendo texto com versao
' Retorna: String com a versao extraida ou "" se nao encontrado
'================================================================================
Public Function ExtractVersionFromText(ByVal textValue As String) As String
    On Error GoTo ErrorHandler

    Dim regex As Object
    Dim matches As Object
    Dim pattern As String

    ExtractVersionFromText = ""

    Set regex = CreateObject("VBScript.RegExp")

    ' Pattern para extrair versao no formato X.Y.Z
    pattern = "([0-9]+)\.([0-9]+)\.([0-9]+)"

    regex.pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False

    Set matches = regex.Execute(textValue)

    If matches.count > 0 Then
        ExtractVersionFromText = matches(0).Value
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao extrair versao: " & Err.Description, LOG_LEVEL_ERROR
    ExtractVersionFromText = ""
End Function

' Funcao: CompareVersions
' Descricao: Compara duas versoes no formato X.Y.Z
' Parametros:
'   - version1: Primeira versao
'   - version2: Segunda versao
' Retorna: 1 se version1 > version2, -1 se version1 < version2, 0 se iguais
'================================================================================
Public Function CompareVersions(ByVal version1 As String, ByVal version2 As String) As Integer
    On Error GoTo ErrorHandler

    Dim v1Parts() As String
    Dim v2Parts() As String
    Dim i As Integer
    Dim v1Num As Long, v2Num As Long

    CompareVersions = 0

    ' Remove espacos
    version1 = Trim(version1)
    version2 = Trim(version2)

    ' Divide versoes em partes
    v1Parts = Split(version1, ".")
    v2Parts = Split(version2, ".")

    ' Compara cada parte
    For i = 0 To 2
        v1Num = 0
        v2Num = 0

        If i <= UBound(v1Parts) Then v1Num = CLng(v1Parts(i))
        If i <= UBound(v2Parts) Then v2Num = CLng(v2Parts(i))

        If v1Num > v2Num Then
            CompareVersions = 1
            Exit Function
        ElseIf v1Num < v2Num Then
            CompareVersions = -1
            Exit Function
        End If
    Next i

    Exit Function

ErrorHandler:
    LogMessage "Erro ao comparar versoes: " & Err.Description, LOG_LEVEL_ERROR
    CompareVersions = 0
End Function

' Funcao: ReadTextFile
' Descricao: Le o conteudo completo de um arquivo de texto
' Parametros:
'   - filePath: Caminho completo do arquivo
' Retorna: Conteudo do arquivo como String
'================================================================================
Public Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim file As Object
    Dim content As String

    ReadTextFile = ""

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1, False, -2) ' -2 = SystemDefault
        content = file.ReadAll
        file.Close
        ReadTextFile = content
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao ler arquivo: " & Err.Description, LOG_LEVEL_ERROR
    ReadTextFile = ""
End Function

' Sub: PromptForUpdate
' Descricao: Verifica se ha atualizacao e pergunta ao usuario se deseja atualizar
'================================================================================
Public Sub PromptForUpdate()
    On Error GoTo ErrorHandler

    Dim updateAvailable As Boolean
    Dim response As VbMsgBoxResult
    Dim installerPath As String
    Dim shellCmd As String

    LogMessage "PromptForUpdate acionado interativamente pelo usuario", LOG_LEVEL_INFO

    If undoGroupEnabled Then
        LogMessage "Atualizacao nao permitida: padronizacao em andamento", LOG_LEVEL_WARNING
        MsgBox "A verificacao de atualizacao nao pode ser executada durante a padronizacao." & vbCrLf & _
               "Aguarde a conclusao e tente novamente.", vbExclamation, "Z7_STDPROPOSERS - Atualizacao"
        Exit Sub
    End If

    ' Verifica se ha atualizacoes (forcando checagem remota por ser interativo)
    updateAvailable = CheckForUpdates(forceCheck:=True)

    If Not updateAvailable Then
        LogMessage "PromptForUpdate concluido: sistema ja esta atualizado", LOG_LEVEL_INFO
        MsgBox "Seu sistema Z7_STDPROPOSERS esta atualizado!", vbInformation, "Z7_STDPROPOSERS - Verificacao de Versao"
        Exit Sub
    End If

    LogMessage "Prompt de atualizacao exibido ao usuario para versao v" & cachedRemoteVersion, LOG_LEVEL_INFO

    ' Pergunta ao usuario se deseja atualizar
    Dim msgUpdate As String
    msgUpdate = "Uma nova versao do Z7_STDPROPOSERS esta disponivel!" & vbCrLf & vbCrLf & _
                "Deseja atualizar agora?" & vbCrLf & vbCrLf & _
                "O instalador sera executado e o Word sera fechado."
    response = MsgBox(msgUpdate, vbYesNo + vbQuestion, "Z7_STDPROPOSERS - Atualizacao Disponivel")

    If response = vbYes Then
        LogMessage "Usuario aceitou a atualizacao para v" & cachedRemoteVersion, LOG_LEVEL_INFO
        ' Caminho do instalador
        installerPath = Environ("USERPROFILE") & "\AppData\Local\Z7\Apps\Z7_StdProposers\installer.cmd"

        ' Verifica se o instalador existe
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(installerPath) Then
            ' Executa o instalador
            shellCmd = "cmd.exe /c """ & installerPath & """"
            LogMessage "Instalador localizado em: " & installerPath & ". Preparando salvamento de documentos...", LOG_LEVEL_INFO

            ' Salva todos os documentos abertos
            Dim doc As Object
            For Each doc In Application.Documents
                If doc.Saved = False Then
                    LogMessage "Salvando documento pendente: " & doc.Name, LOG_LEVEL_INFO
                    doc.Save
                End If
            Next doc

            ' Executa instalador e fecha o Word
            LogMessage "Disparando shell cmd para iniciar atualizacao: " & shellCmd, LOG_LEVEL_INFO
            CreateObject("WScript.Shell").Run shellCmd, 1, False

            LogMessage "Fechando Microsoft Word para atualizacao segura", LOG_LEVEL_INFO
            MsgBox "O instalador sera executado. O Word sera fechado agora.", vbInformation, "Z7_STDPROPOSERS - Atualizacao"
            Application.Quit SaveChanges:=wdSaveChanges
        Else
            LogMessage "ERRO CRITICO: Arquivo do instalador nao encontrado em: " & installerPath, LOG_LEVEL_ERROR
            MsgBox "Instalador nao encontrado em:" & vbCrLf & installerPath & vbCrLf & vbCrLf & _
                   "Baixe manualmente de: https://github.com/chrmsantos/Z7_StdProposers", _
                   vbExclamation, "Z7_STDPROPOSERS - Erro"
        End If
    Else
        LogMessage "Usuario recusou a atualizacao para v" & cachedRemoteVersion, LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao processar atualizacao: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    MsgBox "Erro ao processar atualizacao: " & Err.Description, vbCritical, "Z7_STDPROPOSERS - Erro"
End Sub


' Mod7Utils.bas
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



