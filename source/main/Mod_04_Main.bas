Attribute VB_Name = "Mod_04_Main"
Option Explicit

' Mod_04_Main.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
'================================================================================
' PONTO DE ENTRADA PRINCIPAL
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler

    executionStartTime = Now
    formattingCancelled = False

    ' Verificacoes iniciais
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
        GoTo Cleanup
    End If

    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Cancelado: documento nao salvo"
            LogMessage "Operacao cancelada - documento nao foi salvo", LOG_LEVEL_INFO
            GoTo Cleanup
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

    ' ==========================================================================
    ' INICIO DO GRUPO DE DESFAZER (UndoRecord)
    ' Late binding via CallByName: evita erro de compilacao em versoes
    ' do Word onde UndoRecord nao resolve como early-bound.
    ' Agrupa todas as edicoes da padronizacao em uma unica acao Ctrl+Z.
    ' A flag global undoRecordActive bloqueia DoEvents parasitas durante
    ' todo o pipeline, prevenindo entradas fantasmas na pilha de undo.
    ' ==========================================================================
    On Error Resume Next
    Dim objUndoStart As Object
    Set objUndoStart = CallByName(Application, "UndoRecord", VbGet)
    If Err.Number = 0 Then
        If Not objUndoStart Is Nothing Then
            CallByName objUndoStart, "StartCustomRecord", VbMethod, _
                "Z7_STDPROPOSERS - Padronizacao"
            If Err.Number = 0 Then
                undoRecordActive = True
                LogMessage "UndoRecord iniciado para padronizacao", LOG_LEVEL_INFO
            Else
                undoRecordActive = False
                Err.Clear
            End If
        Else
            undoRecordActive = False
        End If
    Else
        undoRecordActive = False
        Err.Clear
    End If
    Set objUndoStart = Nothing
    On Error GoTo CriticalErrorHandler
    ' ==========================================================================

    ' ==========================================================================
    ' PIPELINE DE FORMATACAO (DUPLA PASSAGEM OTIMIZADA)
    ' ==========================================================================

    LogMessage "=== PIPELINE DE FORMATACAO (2 PASSAGENS) ===", LOG_LEVEL_INFO

    ' Constroi cache de paragrafos (inclui identificacao de estrutura)
    IncrementProgress "Indexando paragrafos"
    BuildParagraphCache doc

    ' Executa verificacao de coerencia da Ementa x Corpo (Modulo 13)
    IncrementProgress "Verificando coerencia ementa x corpo"
    CheckEmentaCoherence doc

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
        LogMessage "=== PASSAGEM " & pipelinePass & " DE 2 === ", LOG_LEVEL_INFO

        ' Reconstroi cache a partir da segunda passagem para evitar indices obsoletos
        ' (paragrafos podem ter sido removidos na passagem anterior)
        If pipelinePass > 1 Then
            IncrementProgress "Reindexando paragrafos (passagem " & pipelinePass & ")"
            BuildParagraphCache doc
        End If

        ' Formata documento
        ' Passagem 1: pipeline completo. Passagem 2: apenas etapas index-dependent
        ' (otimizacao: reduz ~60-70% do tempo da segunda passagem)
        IncrementProgress "Formatando documento (" & pipelinePass & " passagem)"
        If pipelinePass = 1 Then
            If Not PreviousFormatting(doc) Then
                GoTo Cleanup
            End If
        Else
            LogMessage "=== PASSAGEM 2: FORMATACAO SELETIVA (ETAPAS INDEX-DEPENDENT) ===", LOG_LEVEL_INFO
            If Not PreviousFormattingPass2(doc) Then
                GoTo Cleanup
            End If
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
    EnsureConsideringBlankLines doc

    ' Normalizacao generalizada de linhas puladas (maximo 1 linha em branco
    ' seguida; maximo 2 em volta da Data) e garantia de 2 linhas abaixo da Data.
    ' Executadas ANTES das garantias zonais abaixo, para nao desfazer os
    ' espacamentos das zonas especiais.
    NormalizarLinhasEmBranco doc
    GarantirEspacoAbaixoDaData doc

    ' Garantia FINAL de linhas em branco nas zonas especiais (Data: 2 acima;
    ' Titulo da Justificativa: 1 acima e abaixo; Ementa: 1 acima e abaixo).
    ' Executada DEPOIS de toda a padronizacao generalizada de linhas puladas,
    ' para nao ser desfeita por ela. Ordem de baixo para cima (Data -> Titulo
    ' Justificativa -> Ementa) para que os deslocamentos de indice nao afetem
    ' os elementos ja ajustados.
    ForceDataSpacing doc
    ForceJustificativaTitleSpacing doc
    ForceEmentaSpacing doc
    IdentifyDocumentStructure doc   ' atualiza os indices, que mudaram com as remocoes
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

    ' Remove formatacao de numero de paragrafos vazios ao final do processamento
    IncrementProgress "Ajustando paragrafos em branco"
    If Not RemoveNumberingFromBlankParagraphs(doc) Then
        LogMessage "Aviso: Falha ao remover formatacao de numero de paragrafos em branco", LOG_LEVEL_WARNING
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
    On Error GoTo CriticalErrorHandler
    LogMessage "Fonte final garantida: " & STANDARD_FONT & " " & STANDARD_FONT_SIZE & "pt em todo o documento", LOG_LEVEL_INFO

    ' Restaura configuracoes de visualizacao originais (exceto zoom)
    IncrementProgress "Restaurando visualizacao"
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configuracoes de visualizacao podem nao ter sido restauradas", LOG_LEVEL_WARNING
    End If

    ' Exibe eventuais avisos de divergencia entre Ementa e Corpo (Modulo 13)
    ShowEmentaCoherenceWarning

    If formattingCancelled Then
        GoTo Cleanup
    End If

    IncrementProgress "Finalizando"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO
    LogContextSnapshot doc, "FIM"

    ' Calcula tempo de execucao em segundos
    Dim execSeconds As Long
    execSeconds = CLng((Now - executionStartTime) * 86400)

    ' Mostra mensagem final na barra de status
    Application.StatusBar = RenderProgressBar(100, "Padronizacao concluida em " & execSeconds & "s, " & errorCount & " erros, " & warningCount & " avisos")

Cleanup:

    InitializeProgress 0 ' Zera contadores de progresso (evita estado obsoleto)
    ClearParagraphCache ' Limpa cache de paragrafos
    SafeCleanup
    CleanupImageProtection       ' Limpa variaveis de protecao de imagens
    CleanupViewSettings          ' Limpa variaveis de configuracoes de visualizacao

    ' Restaura estado da aplicacao preservando a StatusBar (mantem mensagem final).
    ' undoRecordActive=True: DoEvents em ReleaseObjects ficam bloqueados.
    ' ScreenUpdating=True e ScreenRefresh executam DENTRO do grupo de undo,
    ' capturando quaisquer operacoes internas do Word dentro do mesmo grupo.
    If Not SetAppState(True, "", True) Then
        LogMessage "Falha ao restaurar estado da aplicacao", LOG_LEVEL_WARNING
    End If

    ' Atualiza tela DENTRO do grupo de undo (undoRecordActive=True).
    ' Se executassem ScreenRefresh apos EndCustomRecord, o recalculo de
    ' layout do Word criaria uma entrada fantasma na pilha de undo,
    ' resultando em Access Violation no segundo Ctrl+Z.
    '
    ' CRITICO: Consolidamos On Error Resume Next para cobrir tanto
    ' ScreenRefresh quanto EndCustomRecord em um unico bloco protegido,
    ' eliminando riscos de trocas indevidas de handler no meio do processo.
    On Error Resume Next
    Application.ScreenRefresh

    ' CRITICO: EndCustomRecord e a ULTIMA operacao que pode afetar o estado
    ' do documento. Nenhuma operacao de tela/layout/doc acontece apos aqui.
    ' Isso garante que a pilha de undo termine com exatamente UMA entrada.
    If undoRecordActive Then
        Dim objUndoEnd As Object
        Set objUndoEnd = CallByName(Application, "UndoRecord", VbGet)
        If Err.Number = 0 Then
            If Not objUndoEnd Is Nothing Then
                CallByName objUndoEnd, "EndCustomRecord", VbMethod
            End If
        End If
        Err.Clear
        Set objUndoEnd = Nothing
        undoRecordActive = False  ' so reseta DEPOIS do fechamento
        LogMessage "UndoRecord finalizado com sucesso", LOG_LEVEL_INFO
    End If
    On Error GoTo 0

    SafeFinalizeLogging

    ' NOTA: Suporte a "Repetir" (F4) nao implementado intencionalmente.
    ' O risco de instabilidade da pilha de undo e alto (ver v9.0.0 do projeto),
    ' e o ganho e baixo. O foco e garantir a integridade da pilha de Desfazer,
    ' nao implementar Repetir.
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

    ' CRITICO: Fluxo para Cleanup garante fechamento do UndoRecord mesmo em erro
    GoTo Cleanup
End Sub

'================================================================================
' FUNCOES PUBLICAS DE ACESSO AOS ELEMENTOS ESTRUTURAIS
'================================================================================

Public Function GetTituloRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetTituloRange = Nothing
    If tituloParaIndex <= 0 Or tituloParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloRange = doc.Paragraphs(tituloParaIndex).Range
    Exit Function
ErrorHandler:
    Set GetTituloRange = Nothing
End Function

Public Function GetEmentaRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetEmentaRange = Nothing
    If ementaParaIndex <= 0 Or ementaParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetEmentaRange = doc.Paragraphs(ementaParaIndex).Range
    Exit Function
ErrorHandler:
    Set GetEmentaRange = Nothing
End Function

Public Function GetVocativoRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetVocativoRange = Nothing
    If vocativoStartIndex <= 0 Or vocativoEndIndex <= 0 Then Exit Function
    If vocativoStartIndex > vocativoEndIndex Then Exit Function
    If vocativoStartIndex > doc.Paragraphs.count Then Exit Function
    If vocativoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long, endPos As Long
    startPos = doc.Paragraphs(vocativoStartIndex).Range.Start
    endPos = doc.Paragraphs(vocativoEndIndex).Range.End

    Set GetVocativoRange = doc.Range(startPos, endPos)
    Exit Function
ErrorHandler:
    Set GetVocativoRange = Nothing
End Function

Public Function GetCorpoRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetCorpoRange = Nothing
    If corpoStartIndex <= 0 Or corpoEndIndex <= 0 Then Exit Function
    If corpoStartIndex > corpoEndIndex Then Exit Function
    If corpoStartIndex > doc.Paragraphs.count Then Exit Function
    If corpoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long, endPos As Long
    startPos = doc.Paragraphs(corpoStartIndex).Range.Start
    endPos = doc.Paragraphs(corpoEndIndex).Range.End

    Set GetCorpoRange = doc.Range(startPos, endPos)
    Exit Function
ErrorHandler:
    Set GetCorpoRange = Nothing
End Function

Public Function GetTituloJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetTituloJustificativaRange = Nothing
    If tituloJustificativaIndex <= 0 Or tituloJustificativaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloJustificativaRange = doc.Paragraphs(tituloJustificativaIndex).Range
    Exit Function
ErrorHandler:
    Set GetTituloJustificativaRange = Nothing
End Function

Public Function GetJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetJustificativaRange = Nothing
    If justificativaStartIndex <= 0 Or justificativaEndIndex <= 0 Then Exit Function
    If justificativaStartIndex > justificativaEndIndex Then Exit Function
    If justificativaStartIndex > doc.Paragraphs.count Then Exit Function
    If justificativaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long, endPos As Long
    startPos = doc.Paragraphs(justificativaStartIndex).Range.Start
    endPos = doc.Paragraphs(justificativaEndIndex).Range.End

    Set GetJustificativaRange = doc.Range(startPos, endPos)
    Exit Function
ErrorHandler:
    Set GetJustificativaRange = Nothing
End Function

Public Function GetDataRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetDataRange = Nothing
    If dataParaIndex <= 0 Or dataParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetDataRange = doc.Paragraphs(dataParaIndex).Range
    Exit Function
ErrorHandler:
    Set GetDataRange = Nothing
End Function

Public Function GetAssinaturaRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetAssinaturaRange = Nothing
    If assinaturaStartIndex <= 0 Or assinaturaEndIndex <= 0 Then Exit Function
    If assinaturaStartIndex > assinaturaEndIndex Then Exit Function
    If assinaturaStartIndex > doc.Paragraphs.count Then Exit Function
    If assinaturaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long, endPos As Long
    startPos = doc.Paragraphs(assinaturaStartIndex).Range.Start
    endPos = doc.Paragraphs(assinaturaEndIndex).Range.End

    Set GetAssinaturaRange = doc.Range(startPos, endPos)
    Exit Function
ErrorHandler:
    Set GetAssinaturaRange = Nothing
End Function

Public Function GetTituloAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetTituloAnexoRange = Nothing
    If tituloAnexoIndex <= 0 Or tituloAnexoIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloAnexoRange = doc.Paragraphs(tituloAnexoIndex).Range
    Exit Function
ErrorHandler:
    Set GetTituloAnexoRange = Nothing
End Function

Public Function GetAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetAnexoRange = Nothing
    If anexoStartIndex <= 0 Or anexoEndIndex <= 0 Then Exit Function
    If anexoStartIndex > anexoEndIndex Then Exit Function
    If anexoStartIndex > doc.Paragraphs.count Then Exit Function
    If anexoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long, endPos As Long
    startPos = doc.Paragraphs(anexoStartIndex).Range.Start
    endPos = doc.Paragraphs(anexoEndIndex).Range.End

    Set GetAnexoRange = doc.Range(startPos, endPos)
    Exit Function
ErrorHandler:
    Set GetAnexoRange = Nothing
End Function

Public Function GetProposituraRange(doc As Document) As Range
    On Error GoTo ErrorHandler
    Set GetProposituraRange = Nothing
    If doc Is Nothing Then Exit Function
    Set GetProposituraRange = doc.Range
    Exit Function
ErrorHandler:
    Set GetProposituraRange = Nothing
End Function

Public Function GetElementInfo(doc As Document) As String
    On Error Resume Next
    Dim info As String
    Dim rng As Range

    info = "=== INFORMACOES DOS ELEMENTOS ESTRUTURAIS ===" & vbCrLf

    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        info = info & "Titulo: Paragrafo " & tituloParaIndex & vbCrLf
    Else
        info = info & "Titulo: Nao identificado" & vbCrLf
    End If

    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Ementa: Paragrafo " & ementaParaIndex & vbCrLf
    Else
        info = info & "Ementa: Nao identificado" & vbCrLf
    End If

    Set rng = GetVocativoRange(doc)
    If Not rng Is Nothing Then
        info = info & "Vocativo: Paragrafos " & vocativoStartIndex & " a " & vocativoEndIndex & _
                      " (" & (vocativoEndIndex - vocativoStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Vocativo: Nao identificado" & vbCrLf
    End If

    Set rng = GetCorpoRange(doc)
    If Not rng Is Nothing Then
        info = info & "Proposicao: Paragrafos " & corpoStartIndex & " a " & corpoEndIndex & _
                      " (" & (corpoEndIndex - corpoStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Proposicao: Nao identificado" & vbCrLf
    End If

    If tituloJustificativaIndex > 0 Then
        info = info & "Titulo Justificativa: Paragrafo " & tituloJustificativaIndex & vbCrLf
    Else
        info = info & "Titulo Justificativa: Nao identificado" & vbCrLf
    End If

    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Justificativa: Paragrafos " & justificativaStartIndex & " a " & justificativaEndIndex & _
                      " (" & (justificativaEndIndex - justificativaStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Justificativa: Nao identificado" & vbCrLf
    End If

    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        info = info & "Data (Plenario): Paragrafo " & dataParaIndex & vbCrLf
    Else
        info = info & "Data (Plenario): Nao identificado" & vbCrLf
    End If

    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Assinatura: Paragrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex & _
                      " (" & (assinaturaEndIndex - assinaturaStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Assinatura: Nao identificado" & vbCrLf
    End If

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
' SUBROTINAS PUBLICAS E AUXILIARES
'================================================================================

Public Sub AbrirReadme()
    On Error GoTo ErrorHandler
    Const GITHUB_REPO_URL As String = "https://github.com/chrmsantos/Z7_StdProposers"
    Application.StatusBar = RenderProgressBar(50, "Abrindo repositorio do GitHub")
    CreateObject("WScript.Shell").Run GITHUB_REPO_URL, 1, False
    If loggingEnabled Then LogMessage "Repositorio do GitHub aberto pelo usuario: " & GITHUB_REPO_URL, LOG_LEVEL_INFO
    Application.StatusBar = "Repositorio aberto no navegador"
    Exit Sub
ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositorio"
    LogMessage "Erro ao abrir repositorio do GitHub: " & Err.Description, LOG_LEVEL_ERROR
    On Error Resume Next
    Shell "explorer.exe """ & GITHUB_REPO_URL & """", vbNormalFocus
End Sub

Public Sub ConfirmarDesfazerPadronizacao()
    On Error GoTo ErrorHandler
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    Dim beforeUndoCount As Long, docName As String
    beforeUndoCount = doc.Paragraphs.count
    docName = doc.Name

    Application.StatusBar = RenderProgressBar(40, "Desfazendo padronizacao")
    On Error Resume Next
    doc.Undo
    On Error GoTo ErrorHandler
    DoEvents

    Dim afterUndoCount As Long, changeCount As Long
    afterUndoCount = doc.Paragraphs.count
    changeCount = Abs(beforeUndoCount - afterUndoCount)

    Dim undoMsg As String
    If changeCount > 0 Then
        undoMsg = "[<<] Padronizacao desfeita com sucesso!" & vbCrLf & vbCrLf & _
                  "[CHART] Alteracoes revertidas:" & vbCrLf & _
                  "    Paragrafos afetados: " & changeCount & vbCrLf & vbCrLf & _
                  "[DIR] Documento: " & docName
    Else
        undoMsg = "[<<] Desfazer executado!" & vbCrLf & vbCrLf & _
                  "[i] O documento foi revertido para o estado anterior." & vbCrLf & vbCrLf & _
                  "[DIR] Documento: " & docName
    End If

    MsgBox undoMsg, vbInformation, "Z7_STDPROPOSERS - Desfazer Padronizacao"
    If loggingEnabled Then LogMessage "Padronizacao desfeita pelo usuario - documento: " & docName, LOG_LEVEL_INFO
    Application.StatusBar = "Padronizacao desfeita"
    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao desfazer"
    MsgBox "Nao foi possivel desfazer a operacao.", vbExclamation, "Z7_STDPROPOSERS - Erro ao Desfazer"
    If loggingEnabled Then LogMessage "Erro ao desfazer padronizacao: " & Err.Description, LOG_LEVEL_WARNING
End Sub

Public Sub NotificarDesfazerPadronizacao()
    On Error Resume Next
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    Dim msg As String
    msg = "[<<] Padronizacao desfeita!" & vbCrLf & vbCrLf & _
          "[OK] Todas as alteracoes da ultima padronizacao foram revertidas." & vbCrLf & vbCrLf & _
          "[DIR] Documento: " & doc.Name
    MsgBox msg, vbInformation, "Z7_STDPROPOSERS - Operacao Desfeita"
    If loggingEnabled Then LogMessage "Notificacao de desfazer exibida para: " & doc.Name, LOG_LEVEL_INFO
End Sub

Public Sub ConfigurarPromptGemini()
    Dim objShell As Object
    Dim comandoExecucao As String, caminhoScript As String
    On Error GoTo ErrorHandler
    
    caminhoScript = Environ("USERPROFILE") & PROMPT_CONFIG_SCRIPT_RELATIVE_PATH
    comandoExecucao = """" & caminhoScript & """"
    
    Set objShell = CreateObject("WScript.Shell")
    System.Cursor = wdCursorWait
    objShell.Run comandoExecucao, 0, False
    System.Cursor = wdCursorNormal
    Exit Sub
    
ErrorHandler:
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Erro ao abrir config Gemini"
    If loggingEnabled Then LogMessage "Erro ao abrir config Gemini: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro ao tentar abrir configuracoes do prompt Gemini: " & Err.Description, vbCritical, "Z7_StdProposers"
End Sub

Public Sub ChatComGemini()
    Dim objShell As Object
    Dim comandoExecucao As String, caminhoScript As String
    On Error GoTo ErrorHandler
    
    caminhoScript = Environ("USERPROFILE") & CHAT_IA_SCRIPT_RELATIVE_PATH
    If Dir(caminhoScript) = "" Then
        MsgBox "Executavel do Chat IA nao encontrado.", vbCritical, "Erro de Arquivo"
        Exit Sub
    End If
    
    comandoExecucao = """" & caminhoScript & """"
    Set objShell = CreateObject("WScript.Shell")
    System.Cursor = wdCursorWait
    Application.StatusBar = RenderProgressBar(15, "Carregando chat IA")
    DoEvents
    
    objShell.Run comandoExecucao, 0, False
    System.Cursor = wdCursorNormal
    Exit Sub
    
ErrorHandler:
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Erro ao abrir Chat Gemini"
    If loggingEnabled Then LogMessage "Erro ao abrir Chat Gemini: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro ao tentar abrir o Chat da IA Gemini: " & Err.Description, vbCritical, "Z7_StdProposers"
End Sub

Public Sub ComentarElementosPropositura()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Nenhum documento ativo.", vbExclamation, "Z7"
        Exit Sub
    End If

    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
    End If

    BuildParagraphCache doc
    Dim rng As Range
    Dim commentAddedCount As Long
    commentAddedCount = 0

    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] T" & ChrW(237) & "tulo"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Ementa"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetVocativoRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Vocativo"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetCorpoRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Corpo"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetTituloJustificativaRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] T" & ChrW(237) & "tulo da Justificativa"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Justificativa"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Data (Plen" & ChrW(225) & "rio)"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Assinatura"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetTituloAnexoRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] T" & ChrW(237) & "tulo do Anexo"
        commentAddedCount = commentAddedCount + 1
    End If

    Set rng = GetAnexoRange(doc)
    If Not rng Is Nothing Then
        doc.Comments.Add Range:=rng, Text:="[Z7] Anexo"
        commentAddedCount = commentAddedCount + 1
    End If

    ClearParagraphCache

    If commentAddedCount > 0 Then
        Application.StatusBar = RenderProgressBar(100, commentAddedCount & " partes comentadas com sucesso")
        MsgBox "Identificacao concluida! " & commentAddedCount & " partes estruturais foram marcadas.", vbInformation, "Z7 - Comentar Propositura"
    Else
        MsgBox "Nenhuma parte estrutural da propositura foi identificada.", vbExclamation, "Z7 - Comentar Propositura"
    End If

    Exit Sub

ErrorHandler:
    ClearParagraphCache
    MsgBox "Erro ao comentar elementos: " & Err.Description, vbCritical, "Erro de Execucao"
End Sub

'================================================================================
' NORMALIZACAO E REGRAS DE ESPACAMENTO DAS LINHAS EM BRANCO
'================================================================================

Public Sub NormalizarLinhasEmBranco(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim i As Long
    Dim runStart As Long
    Dim runEnd As Long
    Dim prevIdx As Long
    Dim nextIdx As Long
    Dim maxBlank As Long
    Dim removedCount As Long
    Dim excess As Long
    Dim deleteErr As Long

    removedCount = 0
    i = doc.Paragraphs.count
    Do While i >= 1
        If EhLinhaVaziaZ7(doc.Paragraphs(i)) Then
            ' Localiza o bloco de linhas vazias consecutivas
            runEnd = i
            runStart = i
            Do While runStart > 1
                If EhLinhaVaziaZ7(doc.Paragraphs(runStart - 1)) Then
                    runStart = runStart - 1
                Else
                    Exit Do
                End If
            Loop

            ' Padrao: 1 linha. Ao lado da Data (acima ou abaixo): 2 linhas
            maxBlank = 1
            prevIdx = runStart - 1
            nextIdx = runEnd + 1
            If prevIdx >= 1 Then
                If IsDataElement(doc.Paragraphs(prevIdx)) Then maxBlank = 2
            End If
            If nextIdx <= doc.Paragraphs.count Then
                If IsDataElement(doc.Paragraphs(nextIdx)) Then maxBlank = 2
            End If

            ' Apaga o excedente (sempre as primeiras do bloco) em uma unica
            ' delecao de range - uma unica repaginacao do Word, em vez de uma
            ' por paragrafo. Fallback: loop original em caso de erro.
            excess = (runEnd - runStart + 1) - maxBlank
            If excess > 0 Then
                On Error Resume Next
                doc.Range(doc.Paragraphs(runStart).Range.Start, _
                          doc.Paragraphs(runStart + excess - 1).Range.End).Delete
                deleteErr = Err.Number
                Err.Clear
                On Error GoTo ErrorHandler
                If deleteErr = 0 Then
                    removedCount = removedCount + excess
                Else
                    Do While (runEnd - runStart + 1) > maxBlank
                        doc.Paragraphs(runStart).Range.Delete
                        removedCount = removedCount + 1
                        runEnd = runEnd - 1
                    Loop
                End If
            End If

            i = runStart - 1
        Else
            i = i - 1
        End If
    Loop

    ' Regra de seguranca: indices estruturais ficam invalidos apos delecoes
    If removedCount > 0 Then IdentifyDocumentStructure doc

    LogMessage "NormalizarLinhasEmBranco: 1 linha em branco em todo o documento, 2 em volta da Data", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro em NormalizarLinhasEmBranco: " & Err.Description, LOG_LEVEL_WARNING
End Sub

Private Function EhLinhaVaziaZ7(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    Dim t As String
    t = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
    ' Evita HasVisualContent (COM) quando o paragrafo tem texto: o resultado
    ' so depende dele quando o texto e vazio (And do VBA nao faz curto-circuito)
    If Len(t) = 0 Then
        EhLinhaVaziaZ7 = Not HasVisualContent(para)
    Else
        EhLinhaVaziaZ7 = False
    End If
    Exit Function
ErrorHandler:
    EhLinhaVaziaZ7 = False
End Function

Public Sub GarantirEspacoAbaixoDaData(doc As Document)
    On Error GoTo ErrorHandler

    Const LINHAS_ABAIXO As Long = 2

    Dim i As Long
    Dim firstIdx As Long
    Dim dataIdx As Long
    Dim belowIdx As Long
    Dim blankCount As Long
    Dim faltam As Long
    Dim n As Long

    If doc Is Nothing Then Exit Sub

    ' 1. Usa o indice da Data se ele ainda estiver valido
    dataIdx = 0
    If dataParaIndex > 0 And dataParaIndex <= doc.Paragraphs.count Then
        If IsDataElement(doc.Paragraphs(dataParaIndex)) Then dataIdx = dataParaIndex
    End If

    ' 2. Senao, procura a Data de baixo para cima nos ultimos 15 paragrafos
    If dataIdx = 0 Then
        firstIdx = doc.Paragraphs.count - 15
        If firstIdx < 1 Then firstIdx = 1
        For i = doc.Paragraphs.count To firstIdx Step -1
            If IsDataElement(doc.Paragraphs(i)) Then
                dataIdx = i
                Exit For
            End If
        Next i
    End If

    If dataIdx = 0 Then
        LogMessage "GarantirEspacoAbaixoDaData: Data nao localizada", LOG_LEVEL_WARNING
        Exit Sub
    End If

    ' 3. Conta as linhas em branco que ja existem abaixo da Data
    blankCount = 0
    belowIdx = dataIdx + 1
    Do While belowIdx <= doc.Paragraphs.count
        If EhLinhaVaziaZ7(doc.Paragraphs(belowIdx)) Then
            blankCount = blankCount + 1
            belowIdx = belowIdx + 1
        Else
            Exit Do
        End If
    Loop

    ' 4. Insere se houver conteudo depois da Data e se faltarem linhas
    If belowIdx <= doc.Paragraphs.count And blankCount < LINHAS_ABAIXO Then
        faltam = LINHAS_ABAIXO - blankCount
        For n = 1 To faltam
            doc.Paragraphs(dataIdx).Range.InsertParagraphAfter
        Next n
        LogMessage "GarantirEspacoAbaixoDaData: " & faltam & " linha(s) inserida(s) abaixo da Data", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro em GarantirEspacoAbaixoDaData: " & Err.Description, LOG_LEVEL_WARNING
End Sub

Public Sub TesteEspacoData()
    Dim doc As Document
    Dim i As Long, idx As Long
    Dim proximo As String

    Set doc = ActiveDocument
    idx = 0

    For i = doc.Paragraphs.count To 1 Step -1
        If IsDataElement(doc.Paragraphs(i)) Then
            idx = i
            Exit For
        End If
    Next i

    If idx = 0 Then
        MsgBox "Data NAO localizada pelo IsDataElement.", vbExclamation
        Exit Sub
    End If

    If idx < doc.Paragraphs.count Then
        proximo = Left(doc.Paragraphs(idx + 1).Range.Text, 60)
    Else
        proximo = "(nao ha paragrafo depois)"
    End If

    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Data no paragrafo " & idx & " de " & doc.Paragraphs.count & ":" & vbCrLf & _
           Left(doc.Paragraphs(idx).Range.Text, 80) & vbCrLf & vbCrLf & _
           "Paragrafo seguinte: [" & proximo & "]" & vbCrLf & vbCrLf & _
           "Inserir uma linha em branco de teste apos a Data?", _
           vbYesNo + vbQuestion, "Z7 - Teste Espaco da Data")

    If resposta = vbYes Then
        doc.Paragraphs(idx).Range.InsertParagraphAfter
    End If
End Sub
