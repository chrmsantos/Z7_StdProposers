' Mod6System.bas
Option Explicit

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

    MsgBox msg, vbCritical, "Z7_StdProposers v1.0-beta1"
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
    ' Mostra apenas "Padronizando..." durante a execucao
    Application.StatusBar = "Padronizando..."

    ' Forca atualizacao da tela
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
    GetZ7StdProposersBackupsPath = Environ("TEMP") & "\.z7_stdproposers\props\backups"
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

    Dim z7TempRoot As String
    z7TempRoot = Environ("TEMP") & "\.z7_stdproposers"

    Dim z7TempProps As String
    z7TempProps = z7TempRoot & "\props"

    If Not fso.FolderExists(z7TempRoot) Then fso.CreateFolder z7TempRoot
    If Not fso.FolderExists(z7TempProps) Then fso.CreateFolder z7TempProps
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
Public Function CheckForUpdates() As Boolean
    On Error GoTo ErrorHandler

    Dim localVersion As String
    Dim remoteVersion As String
    Dim updateAvailable As Boolean

    CheckForUpdates = False

    ' Nao executar verificacao durante operacao critica (ex.: padronizacao em andamento)
    If undoGroupEnabled Then
        LogMessage "Verificacao de atualizacao ignorada: operacao em andamento", LOG_LEVEL_INFO
        Exit Function
    End If

    ' Cache: se ja checou com sucesso nesta sessao, reusa o resultado
    If lastUpdateCheckAttempt <> 0 Then
        If lastUpdateCheckSucceeded Then
            CheckForUpdates = cachedUpdateAvailable
            Exit Function
        End If

        ' Se a ultima tentativa falhou recentemente, evita repetir (reduz chance de travamentos)
        If DateDiff("n", lastUpdateCheckAttempt, Now) < UPDATE_CHECK_COOLDOWN_MINUTES Then
            CheckForUpdates = cachedUpdateAvailable
            Exit Function
        End If
    End If

    lastUpdateCheckAttempt = Now

    ' Obtem versao local
    localVersion = GetLocalVersion()
    If localVersion = "" Then
        LogMessage "Nao foi possivel obter versao local", LOG_LEVEL_WARNING
        lastUpdateCheckSucceeded = False
        Exit Function
    End If

    cachedLocalVersion = localVersion

    ' Obtem versao remota do GitHub
    remoteVersion = GetRemoteVersion()
    If remoteVersion = "" Then
        LogMessage "Nao foi possivel obter versao remota", LOG_LEVEL_WARNING
        lastUpdateCheckSucceeded = False
        cachedUpdateAvailable = False
        Exit Function
    End If

    cachedRemoteVersion = remoteVersion
    lastUpdateCheckSucceeded = True

    ' Compara versoes
    updateAvailable = CompareVersions(remoteVersion, localVersion) > 0
    cachedUpdateAvailable = updateAvailable

    If updateAvailable Then
        LogMessage "Atualizacao disponivel: " & localVersion & " -> " & remoteVersion, LOG_LEVEL_INFO
        CheckForUpdates = True
    Else
        LogMessage "Sistema esta atualizado (v" & localVersion & ")", LOG_LEVEL_INFO
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar atualizacoes: " & Err.Description, LOG_LEVEL_ERROR
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

    If undoGroupEnabled Then
        MsgBox "A verificacao de atualizacao nao pode ser executada durante a padronizacao." & vbCrLf & _
               "Aguarde a conclusao e tente novamente.", vbExclamation, "Z7_STDPROPOSERS - Atualizacao"
        Exit Sub
    End If

    ' Verifica se ha atualizacoes
    updateAvailable = CheckForUpdates()

    If Not updateAvailable Then
        MsgBox "Seu sistema Z7_STDPROPOSERS esta atualizado!", vbInformation, "Z7_STDPROPOSERS - Verificacao de Versao"
        Exit Sub
    End If

    ' Pergunta ao usuario se deseja atualizar
    Dim msgUpdate As String
    msgUpdate = "Uma nova versao do Z7_STDPROPOSERS esta disponivel!" & vbCrLf & vbCrLf & _
                "Deseja atualizar agora?" & vbCrLf & vbCrLf & _
                "O instalador sera executado e o Word sera fechado."
    response = MsgBox(msgUpdate, vbYesNo + vbQuestion, "Z7_STDPROPOSERS - Atualizacao Disponivel")

    If response = vbYes Then
        ' Caminho do instalador
        installerPath = Environ("USERPROFILE") & "\AppData\Local\Z7\Apps\Z7_StdProposers\installer.cmd"

        ' Verifica se o instalador existe
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(installerPath) Then
            ' Executa o instalador
            shellCmd = "cmd.exe /c """ & installerPath & """"

            ' Salva todos os documentos abertos
            Dim doc As Object
            For Each doc In Application.Documents
                If doc.Saved = False Then
                    doc.Save
                End If
            Next doc

            ' Executa instalador e fecha o Word
            CreateObject("WScript.Shell").Run shellCmd, 1, False

            MsgBox "O instalador sera executado. O Word sera fechado agora.", vbInformation, "Z7_STDPROPOSERS - Atualizacao"
            Application.Quit SaveChanges:=wdSaveChanges
        Else
            MsgBox "Instalador nao encontrado em:" & vbCrLf & installerPath & vbCrLf & vbCrLf & _
                   "Baixe manualmente de: https://github.com/chrmsantos/Z7_StdProposers", _
                   vbExclamation, "Z7_STDPROPOSERS - Erro"
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao processar atualizacao: " & Err.Description, vbCritical, "Z7_STDPROPOSERS - Erro"
End Sub

