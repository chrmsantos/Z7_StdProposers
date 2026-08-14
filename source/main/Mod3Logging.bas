Attribute VB_Name = "Mod3Logging"
Option Explicit

' Mod3Logging
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)

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

Public Sub WriteTextUTF8(filePath As String, textContent As String, Optional appendMode As Boolean = False)
    On Error GoTo ErrorHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText textContent, 1 ' adWriteLine

    If appendMode And Dir(filePath) <> "" Then
        ' Converte apenas o novo conteudo para bytes, pulando o BOM UTF-8 (3 bytes),
        ' e faz append binario no arquivo existente sem re-ler o conteudo como texto.
        stream.Position = 0
        stream.Type = 1 ' adTypeBinary
        stream.Position = 3 ' salta BOM UTF-8

        Dim newBytes() As Byte
        newBytes = stream.Read
        stream.Close
        Set stream = Nothing

        Dim binStream As Object
        Set binStream = CreateObject("ADODB.Stream")
        binStream.Type = 1 ' adTypeBinary
        binStream.Open
        binStream.LoadFromFile filePath
        binStream.Position = binStream.Size
        binStream.Write newBytes
        binStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
        binStream.Close
        Set binStream = Nothing
    Else
        stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
        stream.Close
        Set stream = Nothing
    End If

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
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

    ' Salva logs no mesmo diretorio dos logs de IA (%LOCALAPPDATA%\Z7\Tmp\StdProposers\logs)
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
    timeStamp = Format(Now, "HH:mm:ss.") & Format(CLng(Timer * 1000) Mod 1000, "000")
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
