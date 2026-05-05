Option Explicit

' Mod4Main.bas
'================================================================================
' PONTO DE ENTRADA PRINCIPAL
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler

    executionStartTime = Now
    formattingCancelled = False
    undoGroupEnabled = False ' Reset inicial

    ' Verificacoes iniciais ANTES de iniciar UndoRecord
    If Not CheckWordVersion() Then
        Application.StatusBar = "Erro: Word 2010 ou superior necessario"
        LogMessage "Versao do Word " & Application.version & " nao suportada. Minimo: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
        MsgBox "Requer Word 2010 ou superior." & vbCrLf & _
               "Versao atual: " & Application.version, vbCritical, "Versao Incompativel"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = Nothing

    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento aberto"
        MsgBox "Nenhum documento esta aberto para processamento.", vbCritical, "Erro"
        Exit Sub
    End If
    Err.Clear
    On Error GoTo CriticalErrorHandler
    ' ---------------------------------------------------------------------------

    ' Inicializa sistema de logging ANTES de qualquer LogMessage
    If Not InitializeLogging(doc) Then
        Application.StatusBar = "Aviso: Log desabilitado"
    Else
        LogContextSnapshot doc, "INICIO"
    End If

    ' Inicializa sistema de progresso (18 etapas do pipeline - 2 passagens)
    InitializeProgress 18

    If Not SetAppState(False, "Iniciando...") Then
        LogMessage "Falha ao configurar estado da aplicacao", LOG_LEVEL_WARNING
    End If

    IncrementProgress "Verificando documento"
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If

    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Cancelado: documento nao salvo"
            LogMessage "Operacao cancelada - documento nao foi salvo", LOG_LEVEL_INFO
            GoTo CleanUp
        End If
    End If

    ' Cria backup do documento antes de qualquer modificacao
    IncrementProgress "Criando backup"
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
    End If

    ' Backup das configuracoes de visualizacao originais
    IncrementProgress "Salvando configuracoes"
    If Not BackupViewSettings(doc) Then
        LogMessage "Aviso: Falha no backup das configuracoes de visualizacao", LOG_LEVEL_WARNING
    End If

    ' Backup de imagens antes das formatacoes
    IncrementProgress "Protegendo imagens"
    If Not BackupAllImages(doc) Then
        LogMessage "Aviso: Falha no backup de imagens - continuando com protecao basica", LOG_LEVEL_WARNING
    End If

    ' Backup de formatacoes de lista antes das formatacoes
    IncrementProgress "Protegendo listas"
    If Not BackupListFormats(doc) Then
        LogMessage "Aviso: Falha no backup de listas - formatacoes de lista podem ser perdidas", LOG_LEVEL_WARNING
    End If

    ' Backup de paragrafos centralizados antes das formatacoes
    IncrementProgress "Protegendo paragrafos centralizados"
    If Not BackupCenteredParagraphs(doc) Then
        LogMessage "Aviso: Falha no backup de paragrafos centralizados", LOG_LEVEL_WARNING
    End If

    ' ---------------------------------------------------------------------------
    ' INICIO DO GRUPO DE DESFAZER (UndoRecord) - melhor esforco
    ' ---------------------------------------------------------------------------
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Z7_STDPROPOSERS - Padronizacao"
    If Err.Number = 0 Then
        undoGroupEnabled = True
        LogMessage "UndoRecord iniciado", LOG_LEVEL_INFO
    Else
        undoGroupEnabled = False
        Err.Clear
    End If
    On Error GoTo CriticalErrorHandler
    ' ---------------------------------------------------------------------------

    ' ==========================================================================
    ' PIPELINE DE FORMATACAO (DUPLA PASSAGEM OTIMIZADA)
    ' ==========================================================================

    LogMessage "=== PIPELINE DE FORMATACAO (2 PASSAGENS) ===", LOG_LEVEL_INFO

    ' Constroi cache de paragrafos (inclui identificacao de estrutura)
    IncrementProgress "Indexando paragrafos"
    BuildParagraphCache doc

    ' Executa formatacao em 2 passagens para garantir estabilidade
    ' Segunda passagem so executa se primeira fez alteracoes (flag dirty)
    Dim pipelinePass As Integer
    documentDirty = True  ' Primeira passagem sempre executa

    For pipelinePass = 1 To 2
        ' Pula segunda passagem se documento nao foi modificado
        If pipelinePass = 2 And Not documentDirty Then
            LogMessage "=== PASSAGEM 2 IGNORADA (sem alteracoes na passagem 1) ===", LOG_LEVEL_INFO
            Exit For
        End If

        documentDirty = False  ' Reset flag antes de cada passagem
        LogMessage "=== PASSAGEM " & pipelinePass & " DE 2 ===", LOG_LEVEL_INFO

        ' Formata documento
        IncrementProgress "Formatando documento (" & pipelinePass & " passagem)"
        If Not PreviousFormatting(doc) Then
            GoTo CleanUp
        End If

        ' Restaura imagens apos formatacoes
        IncrementProgress "Restaurando imagens (" & pipelinePass & " passagem)"
        If Not RestoreAllImages(doc) Then
            LogMessage "Aviso: Algumas imagens podem ter sido afetadas durante o processamento", LOG_LEVEL_WARNING
        End If
    Next pipelinePass

    ' Remove linhas em branco extras e aplica ajustes finais
    IncrementProgress "Removendo linhas em branco extras"
    RemoverLinhasEmBrancoExtras doc

    ' Restaura formatacoes de lista apos formatacoes
    IncrementProgress "Restaurando listas"
    If Not RestoreListFormats(doc) Then
        LogMessage "Aviso: Algumas formatacoes de lista podem nao ter sido restauradas", LOG_LEVEL_WARNING
    End If

    ' Formata paragrafos iniciados com numero (aplica recuo de lista numerada)
    IncrementProgress "Ajustando numeracao"
    If Not FormatNumberedParagraphsIndent(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de paragrafos numerados", LOG_LEVEL_WARNING
    End If

    ' Formata paragrafos iniciados com marcador (aplica recuo de lista com marcadores)
    IncrementProgress "Ajustando marcadores"
    If Not FormatBulletedParagraphsIndent(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de paragrafos com marcadores", LOG_LEVEL_WARNING
    End If

    ' Formata recuos de paragrafos com imagens (zera recuo a esquerda)
    IncrementProgress "Ajustando layout"
    If Not FormatImageParagraphsIndents(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de imagens", LOG_LEVEL_WARNING
    End If

    ' Centraliza imagem entre 5a e 7a linha apos Plenario
    IncrementProgress "Centralizando elementos"
    If Not CenterImageAfterPlenario(doc) Then
        LogMessage "Aviso: Falha ao centralizar imagem apos Plenario", LOG_LEVEL_WARNING
    End If

    ' Restaura centralizacao dos paragrafos que estavam centralizados antes do processamento
    IncrementProgress "Restaurando paragrafos centralizados"
    If Not RestoreCenteredParagraphs(doc) Then
        LogMessage "Aviso: Falha ao restaurar paragrafos centralizados", LOG_LEVEL_WARNING
    End If

    ' Garantia final de fonte: reaplica Arial 12 em todo o documento apos todos os
    ' ajustes pos-pipeline (substituicoes de texto, listas, imagens), pois operacoes
    ' como Find/Replace com Replacement.ClearFormatting podem deixar trechos com
    ' a fonte do estilo Normal (ex: Calibri) em vez de Arial 12.
    IncrementProgress "Garantindo fonte final"
    On Error Resume Next
    With doc.Range.Font
        .Name = STANDARD_FONT
        .size = STANDARD_FONT_SIZE
    End With
    On Error GoTo 0
    LogMessage "Fonte final garantida: " & STANDARD_FONT & " " & STANDARD_FONT_SIZE & "pt em todo o documento", LOG_LEVEL_INFO

    ' Restaura configuracoes de visualizacao originais (exceto zoom)
    IncrementProgress "Restaurando visualizacao"
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configuracoes de visualizacao podem nao ter sido restauradas", LOG_LEVEL_WARNING
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    IncrementProgress "Finalizando"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO
    LogContextSnapshot doc, "FIM"

    ' Calcula tempo de execucao em segundos
    Dim execSeconds As Long
    execSeconds = CLng((Now - executionStartTime) * 86400)

    ' Mostra mensagem final na barra de status
    Application.StatusBar = "Padronizacao concluida em " & execSeconds & "s, com " & errorCount & " erros e " & warningCount & " avisos! (z7_stdproposers)"

CleanUp:
    ' ---------------------------------------------------------------------------
    ' FIM DO GRUPO DE DESFAZER - SEMPRE fecha o UndoRecord
    ' ---------------------------------------------------------------------------
    On Error Resume Next
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "UndoRecord finalizado com sucesso", LOG_LEVEL_INFO
    End If
    Err.Clear
    On Error GoTo 0
    ' ---------------------------------------------------------------------------

    ClearParagraphCache ' Limpa cache de paragrafos
    SafeCleanup
    CleanupImageProtection       ' Limpa variaveis de protecao de imagens
    CleanupViewSettings          ' Limpa variaveis de configuracoes de visualizacao
    CleanupCenteredParaBackup    ' Limpa variaveis de backup de paragrafos centralizados

    ' Restaura estado da aplicacao preservando a StatusBar (mantem mensagem final)
    If Not SetAppState(True, "", True) Then
        LogMessage "Falha ao restaurar estado da aplicacao", LOG_LEVEL_WARNING
    End If

    SafeFinalizeLogging

    ' Mensagem de conclusao desativada - informacoes exibidas apenas na StatusBar
    ' If Not formattingCancelled Then
    '     Dim executionTimeText As String
    '     Dim duration As Double
    '     duration = (Now - executionStartTime) * 86400
    '     If duration < 60 Then
    '         executionTimeText = Format(duration, "0.0") & " segundos"
    '     ElseIf duration < 3600 Then
    '         executionTimeText = Format(Int(duration / 60), "0") & " minuto(s) e " & Format(duration Mod 60, "00") & " segundo(s)"
    '     Else
    '         executionTimeText = Format(Int(duration / 3600), "0") & " hora(s) e " & Format(Int((duration Mod 3600) / 60), "00") & " minuto(s)"
    '     End If
    '     Dim statusMsg As String
    '     If errorCount > 0 Then
    '         statusMsg = vbCrLf & vbCrLf & "[!] ATENCAO: " & errorCount & " erro(s) detectado(s) durante a execucao." & vbCrLf & _
    '                    "   Verifique o log para mais detalhes."
    '     ElseIf warningCount > 0 Then
    '         statusMsg = vbCrLf & vbCrLf & "[i] INFORMACAO: " & warningCount & " aviso(s) registrado(s) durante a execucao." & vbCrLf & _
    '                    "   Verifique o log para mais detalhes."
    '     Else
    '         statusMsg = vbCrLf & vbCrLf & "[OK] Nenhum erro ou aviso detectado durante a execucao."
    '     End If
    '     MsgBox "[OK] Processamento concluido com sucesso em " & executionTimeText & "!" & vbCrLf & vbCrLf & _
    '            "[DIR] Backup criado em:" & vbCrLf & _
    '            "   " & IIf(backupFilePath <> "", backupFilePath, GetZ7StdProposersBackupsPath()) & vbCrLf & vbCrLf & _
    '            "[LOG] Log salvo em:" & vbCrLf & _
    '            "   " & logFilePath & statusMsg, _
    '            vbInformation, "Z7_STDPROPOSERS - Padronizacao Concluida"
    ' End If

    ' Posiciona cursor no inicio do documento
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.Range(0, 0).Select
    End If
    On Error GoTo 0

    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRITICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"

    LogMessage errDesc, LOG_LEVEL_ERROR
    If Not doc Is Nothing Then
        LogContextSnapshot doc, "ERRO_CRITICO"
    End If
    Application.StatusBar = "Erro - verificar logs"

    ShowUserFriendlyError Err.Number, Err.Description
    EmergencyRecovery

    ' CRITICO: Garante fechamento do UndoRecord mesmo em erro
    GoTo CleanUp
End Sub

'================================================================================
' FUNCOES PUBLICAS DE ACESSO AOS ELEMENTOS ESTRUTURAIS
'================================================================================

'--------------------------------------------------------------------------------
' GetTituloRange - Retorna o Range do titulo
'--------------------------------------------------------------------------------
Public Function GetTituloRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloRange = Nothing

    If tituloParaIndex <= 0 Or tituloParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloRange = doc.Paragraphs(tituloParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetEmentaRange - Retorna o Range da ementa
'--------------------------------------------------------------------------------
Public Function GetEmentaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetEmentaRange = Nothing

    If ementaParaIndex <= 0 Or ementaParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetEmentaRange = doc.Paragraphs(ementaParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetEmentaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetProposicaoRange - Retorna o Range da proposicao (conjunto de paragrafos)
'--------------------------------------------------------------------------------
Public Function GetProposicaoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetProposicaoRange = Nothing

    If proposicaoStartIndex <= 0 Or proposicaoEndIndex <= 0 Then Exit Function
    If proposicaoStartIndex > doc.Paragraphs.count Then Exit Function
    If proposicaoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(proposicaoStartIndex).Range.Start
    endPos = doc.Paragraphs(proposicaoEndIndex).Range.End

    Set GetProposicaoRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetProposicaoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetTituloJustificativaRange - Retorna o Range do titulo "Justificativa"
'--------------------------------------------------------------------------------
Public Function GetTituloJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloJustificativaRange = Nothing

    If tituloJustificativaIndex <= 0 Or tituloJustificativaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloJustificativaRange = doc.Paragraphs(tituloJustificativaIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloJustificativaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetJustificativaRange - Retorna o Range da justificativa (conjunto de paragrafos)
'--------------------------------------------------------------------------------
Public Function GetJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetJustificativaRange = Nothing

    If justificativaStartIndex <= 0 Or justificativaEndIndex <= 0 Then Exit Function
    If justificativaStartIndex > doc.Paragraphs.count Then Exit Function
    If justificativaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(justificativaStartIndex).Range.Start
    endPos = doc.Paragraphs(justificativaEndIndex).Range.End

    Set GetJustificativaRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetJustificativaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetDataRange - Retorna o Range da data (Plenario)
'--------------------------------------------------------------------------------
Public Function GetDataRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetDataRange = Nothing

    If dataParaIndex <= 0 Or dataParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetDataRange = doc.Paragraphs(dataParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetDataRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetAssinaturaRange - Retorna o Range da assinatura (3 paragrafos + imagens)
'--------------------------------------------------------------------------------
Public Function GetAssinaturaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetAssinaturaRange = Nothing

    If assinaturaStartIndex <= 0 Or assinaturaEndIndex <= 0 Then Exit Function
    If assinaturaStartIndex > doc.Paragraphs.count Then Exit Function
    If assinaturaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(assinaturaStartIndex).Range.Start
    endPos = doc.Paragraphs(assinaturaEndIndex).Range.End

    Set GetAssinaturaRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetAssinaturaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetTituloAnexoRange - Retorna o Range do titulo "Anexo" ou "Anexos"
'--------------------------------------------------------------------------------
Public Function GetTituloAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloAnexoRange = Nothing

    If tituloAnexoIndex <= 0 Or tituloAnexoIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloAnexoRange = doc.Paragraphs(tituloAnexoIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloAnexoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetAnexoRange - Retorna o Range do anexo (todo conteudo abaixo do titulo)
'--------------------------------------------------------------------------------
Public Function GetAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetAnexoRange = Nothing

    If anexoStartIndex <= 0 Or anexoEndIndex <= 0 Then Exit Function
    If anexoStartIndex > doc.Paragraphs.count Then Exit Function
    If anexoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(anexoStartIndex).Range.Start
    endPos = doc.Paragraphs(anexoEndIndex).Range.End

    Set GetAnexoRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetAnexoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetProposituraRange - Retorna o Range de toda a propositura (documento completo)
'--------------------------------------------------------------------------------
Public Function GetProposituraRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetProposituraRange = Nothing

    If doc Is Nothing Then Exit Function
    Set GetProposituraRange = doc.Range
    Exit Function

ErrorHandler:
    Set GetProposituraRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetElementInfo - Retorna informacoes sobre todos os elementos identificados
' REFATORADO: Usa funcoes identificadoras ao inves de acesso direto as variaveis
'--------------------------------------------------------------------------------
Public Function GetElementInfo(doc As Document) As String
    On Error Resume Next

    Dim info As String
    Dim rng As Range

    info = "=== INFORMACOES DOS ELEMENTOS ESTRUTURAIS ===" & vbCrLf

    ' Titulo - usa GetTituloRange
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        info = info & "Titulo: Paragrafo " & tituloParaIndex & vbCrLf
    Else
        info = info & "Titulo: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Ementa - usa GetEmentaRange
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Ementa: Paragrafo " & ementaParaIndex & vbCrLf
    Else
        info = info & "Ementa: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Proposicao - usa GetProposicaoRange
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        info = info & "Proposicao: Paragrafos " & proposicaoStartIndex & " a " & proposicaoEndIndex & _
                      " (" & (proposicaoEndIndex - proposicaoStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Proposicao: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Titulo Justificativa - ainda usa variavel direta (nao tem funcao Get especifica)
    If tituloJustificativaIndex > 0 Then
        info = info & "Titulo Justificativa: Paragrafo " & tituloJustificativaIndex & vbCrLf
    Else
        info = info & "Titulo Justificativa: Nao identificado" & vbCrLf
    End If

    ' Justificativa - usa GetJustificativaRange
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Justificativa: Paragrafos " & justificativaStartIndex & " a " & justificativaEndIndex & _
                      " (" & (justificativaEndIndex - justificativaStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Justificativa: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Data - usa GetDataRange
    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        info = info & "Data (Plenario): Paragrafo " & dataParaIndex & vbCrLf
    Else
        info = info & "Data (Plenario): Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Assinatura - usa GetAssinaturaRange
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Assinatura: Paragrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex & _
                      " (" & (assinaturaEndIndex - assinaturaStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Assinatura: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    If tituloAnexoIndex > 0 Then
        info = info & "Titulo Anexo: Paragrafo " & tituloAnexoIndex & vbCrLf
        If anexoStartIndex > 0 And anexoEndIndex > 0 Then
            info = info & "Anexo: Paragrafos " & anexoStartIndex & " a " & anexoEndIndex & _
                          " (" & (anexoEndIndex - anexoStartIndex + 1) & " paragrafos)" & vbCrLf
        End If
    Else
        info = info & "Anexo: Nao presente" & vbCrLf
    End If

    info = info & "============================================="

    GetElementInfo = info
End Function

'================================================================================
' SUBROTINA PUBLICA: ABRIR REPOSITORIO DO GITHUB
'================================================================================
Public Sub AbrirReadme()
    On Error GoTo ErrorHandler

    Const GITHUB_REPO_URL As String = "https://github.com/chrmsantos/Z7_StdProposers"

    ' Abre o repositorio do GitHub no navegador padrao
    Application.StatusBar = "Abrindo repositorio do GitHub..."

    ' Usa o comando Shell com o protocolo http:// para abrir no navegador padrao
    CreateObject("WScript.Shell").Run GITHUB_REPO_URL, 1, False

    ' Log da operacao se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Repositorio do GitHub aberto pelo usuario: " & GITHUB_REPO_URL, LOG_LEVEL_INFO
    End If

    Application.StatusBar = "Repositorio aberto no navegador"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositorio"
    LogMessage "Erro ao abrir repositorio do GitHub: " & Err.Description, LOG_LEVEL_ERROR

    ' Tenta metodo alternativo
    On Error Resume Next
    shell "explorer.exe """ & GITHUB_REPO_URL & """", vbNormalFocus
End Sub

'================================================================================
' SUBROTINA PUBLICA: CONFIRMAR DESFAZIMENTO DA PADRONIZACAO
'================================================================================
Public Sub ConfirmarDesfazerPadronizacao()
    On Error GoTo ErrorHandler

    ' Verifica se ha um documento ativo
    Dim doc As Document
    Set doc = Nothing

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Exit Sub
    End If

    ' Verifica o numero de acoes disponiveis para desfazer
    Dim canUndo As Boolean
    canUndo = False

    On Error Resume Next
    canUndo = Application.CommandBars.ActionControl.enabled
    If Err.Number <> 0 Then canUndo = False
    On Error GoTo ErrorHandler

    ' Armazena informacoes antes do desfazer
    Dim beforeUndoCount As Long
    Dim docName As String
    Dim docPath As String

    beforeUndoCount = doc.Paragraphs.count
    docName = doc.Name
    docPath = doc.Path

    ' Executa o comando Desfazer (Undo)
    Application.StatusBar = "Desfazendo padronizacao..."
    On Error Resume Next
    doc.Undo
    On Error GoTo ErrorHandler

    ' Aguarda o Word processar o desfazer
    DoEvents

    ' Verifica se houve mudanca no documento
    Dim afterUndoCount As Long
    afterUndoCount = doc.Paragraphs.count

    ' Calcula a diferenca
    Dim changeCount As Long
    changeCount = Abs(beforeUndoCount - afterUndoCount)

    ' Cria mensagem informativa
    Dim undoMsg As String

    If changeCount > 0 Then
        undoMsg = "[<<] Padronizacao desfeita com sucesso!" & vbCrLf & vbCrLf & _
                  "[CHART] Alteracoes revertidas:" & vbCrLf & _
                  "    Paragrafos afetados: " & changeCount & vbCrLf & vbCrLf & _
                  "[DIR] Documento:" & vbCrLf & _
                  "   " & docName & vbCrLf & vbCrLf & _
                  "[i] DICA: O backup da padronizacao permanece disponivel." & vbCrLf & _
                  "   Use 'Abrir Pasta de Logs e Backups' para acessa-lo."
    Else
        undoMsg = "[<<] Desfazer executado!" & vbCrLf & vbCrLf & _
                  "[i] O documento foi revertido para o estado anterior." & vbCrLf & vbCrLf & _
                  "[DIR] Documento:" & vbCrLf & _
                  "   " & docName & vbCrLf & vbCrLf & _
                  "[i] DICA: O backup da padronizacao permanece disponivel." & vbCrLf & _
                  "   Use 'Abrir Pasta de Logs e Backups' para acessa-lo."
    End If

    ' Exibe mensagem de confirmacao
    MsgBox undoMsg, vbInformation, "Z7_STDPROPOSERS - Desfazer Padronizacao"

    ' Registra no log se estiver ativo
    If loggingEnabled Then
        LogMessage "Padronizacao desfeita pelo usuario - documento: " & docName, LOG_LEVEL_INFO
    End If

    Application.StatusBar = "Padronizacao desfeita"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao desfazer"

    ' Mensagem de erro generica
    MsgBox "Nao foi possivel desfazer a operacao." & vbCrLf & vbCrLf & _
           "[!] Possiveis causas:" & vbCrLf & _
           "    Nao ha operacoes para desfazer" & vbCrLf & _
           "    O documento foi fechado e reaberto" & vbCrLf & _
           "    Limite de desfazer atingido" & vbCrLf & vbCrLf & _
           "[i] SOLUCAO: Restaure manualmente a partir do backup." & vbCrLf & _
           "   Use 'Abrir Pasta de Logs e Backups' para acessar os backups.", _
           vbExclamation, "Z7_STDPROPOSERS - Erro ao Desfazer"

    If loggingEnabled Then
        LogMessage "Erro ao desfazer padronizacao: " & Err.Description, LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PUBLICA: DESFAZER COM CONFIRMACAO AUTOMATICA
' Esta sub pode ser chamada diretamente ou apos o usuario usar Ctrl+Z
'================================================================================
Public Sub NotificarDesfazerPadronizacao()
    On Error Resume Next

    ' Verifica se ha um documento ativo
    Dim doc As Document
    Set doc = ActiveDocument

    If doc Is Nothing Then Exit Sub

    ' Cria mensagem de confirmacao simplificada
    Dim msg As String
    msg = "[<<] Padronizacao desfeita!" & vbCrLf & vbCrLf & _
          "[OK] Todas as alteracoes da ultima padronizacao foram revertidas." & vbCrLf & vbCrLf & _
          "[DIR] Documento: " & doc.Name & vbCrLf & vbCrLf & _
          "[SAVE] O backup continua disponivel na pasta de backups." & vbCrLf & _
          "   Use 'Abrir Pasta de Logs e Backups' para acessa-lo."

    ' Exibe notificacao
    MsgBox msg, vbInformation, "Z7_STDPROPOSERS - Operacao Desfeita"

    ' Log se disponivel
    If loggingEnabled Then
        LogMessage "Notificacao de desfazer exibida para: " & doc.Name, LOG_LEVEL_INFO
    End If
End Sub

'================================================================================
' SUBROTINA PUBLICA: CORRIGIR GRAMATICA COM GEMINI
'================================================================================
Public Sub CorrigirGramaticaComGemini()
    ' Macro para enviar o texto selecionado para correcao via script Python
    ' usando a API do Gemini.
    
    Dim objShell As Object
    Dim comandoExecucao As String
    Dim caminhoScript As String
    Dim caminhoPython As String
    
    On Error GoTo ErrorHandler
    
    ' Obtem o caminho do executável usando o caminho relativo configurado em Mod1Config
    caminhoScript = Environ("USERPROFILE") & GRAMMAR_SCRIPT_RELATIVE_PATH
    
    If Dir(caminhoScript) = "" Then
        MsgBox "Executável do Corretor Gramatical não encontrado em:" & vbCrLf & caminhoScript & vbCrLf & vbCrLf & "Por favor, recompile o projeto ou verifique a instalação.", vbCritical, "Erro de Arquivo"
        Exit Sub
    End If
    
    ' Monta o comando completo com aspas em volta do caminho do executável
    comandoExecucao = """" & caminhoScript & """"
    
    ' Cria o objeto WScript.Shell
    Set objShell = CreateObject("WScript.Shell")
    
    Application.StatusBar = "Carregando o assistente de IA... Isso pode levar alguns segundos."
    DoEvents
    
    ' Executa o comando de forma ASSÍNCRONA para não travar o Word
    objShell.Run comandoExecucao, 0, False
    
    ' Opcionalmente exibe na status bar ou no log
    Application.StatusBar = "Revisao Gemini iniciada em segundo plano..."
    If loggingEnabled Then LogMessage "Revisao Gemini iniciada de forma assincrona.", LOG_LEVEL_INFO
    
    Exit Sub
    
ErrorHandler:
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Erro na revisao Gemini"
    If loggingEnabled Then LogMessage "Erro na revisao Gemini: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro ao tentar executar a revisao pelo Gemini: " & Err.Description, vbCritical, "Z7_StdProposers"
End Sub

'================================================================================
' SUBROTINA PUBLICA: CONFIGURAR PROMPT GEMINI
'================================================================================
Public Sub ConfigurarPromptGemini()
    ' Macro para abrir a interface em Python e editar o prompt da IA
    Dim objShell As Object
    Dim comandoExecucao As String
    Dim caminhoScript As String
    Dim caminhoPython As String
    
    On Error GoTo ErrorHandler
    
    ' Obtem o caminho do executável usando o caminho relativo configurado em Mod1Infrastructure
    caminhoScript = Environ("USERPROFILE") & PROMPT_CONFIG_SCRIPT_RELATIVE_PATH
    
    ' Monta o comando completo com aspas em volta do caminho do executável
    comandoExecucao = """" & caminhoScript & """"
    
    ' Cria o objeto WScript.Shell
    Set objShell = CreateObject("WScript.Shell")
    
    ' Muda o ponteiro do mouse para indicar carregamento
    System.Cursor = wdCursorWait
    
    ' Executa o comando SEM aguardar a conclusao, pois e uma janela interativa
    objShell.Run comandoExecucao, 0, False
    
    ' Retorna o ponteiro do mouse ao normal
    System.Cursor = wdCursorNormal
    
    Exit Sub
    
ErrorHandler:
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Erro ao abrir config Gemini"
    If loggingEnabled Then LogMessage "Erro ao abrir config Gemini: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro ao tentar abrir configuracoes do prompt Gemini: " & Err.Description, vbCritical, "Z7_StdProposers"
End Sub

'================================================================================
' SUBROTINA PUBLICA: CHAT COM GEMINI
'================================================================================
Public Sub ChatComGemini()
    ' Macro para abrir a interface em Python do Chat Interativo com a IA
    Dim objShell As Object
    Dim comandoExecucao As String
    Dim caminhoScript As String
    
    On Error GoTo ErrorHandler
    
    ' Obtem o caminho do executável usando o caminho relativo configurado em Mod1Infrastructure
    caminhoScript = Environ("USERPROFILE") & CHAT_IA_SCRIPT_RELATIVE_PATH
    
    If Dir(caminhoScript) = "" Then
        MsgBox "Executável do Chat IA não encontrado em:" & vbCrLf & caminhoScript & vbCrLf & vbCrLf & "Por favor, recompile o projeto ou verifique a instalação.", vbCritical, "Erro de Arquivo"
        Exit Sub
    End If
    
    ' Monta o comando completo com aspas em volta do caminho do executável
    comandoExecucao = """" & caminhoScript & """"
    
    ' Cria o objeto WScript.Shell
    Set objShell = CreateObject("WScript.Shell")
    
    ' Muda o ponteiro do mouse para indicar carregamento
    System.Cursor = wdCursorWait
    
    Application.StatusBar = "Carregando o chat interativo da IA... Isso pode levar alguns segundos."
    DoEvents
    
    ' Executa o comando SEM aguardar a conclusao, pois e uma janela interativa
    objShell.Run comandoExecucao, 0, False
    
    ' Retorna o ponteiro do mouse ao normal
    System.Cursor = wdCursorNormal
    
    Exit Sub
    
ErrorHandler:
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Erro ao abrir Chat Gemini"
    If loggingEnabled Then LogMessage "Erro ao abrir Chat Gemini: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro ao tentar abrir o Chat da IA Gemini: " & Err.Description, vbCritical, "Z7_StdProposers"
End Sub


