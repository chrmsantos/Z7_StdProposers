Attribute VB_Name = "Mod_03_Pipeline"
Option Explicit

' Mod_03_Pipeline
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@gmail.com)
'================================================================================



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
        If Not undoRecordActive Then
            DoEvents
        End If
    Next memoryCounter
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
    Do While InStr(result, "  ") > 0 And loopGuard < 10
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

    ' Excecao: "Vereadora," (vocativo) nao deve ser classificada como paragrafo especial de vereador
    If cleanText = "vereadora," Then
        DetectSpecialParagraph = False
        Exit Function
    End If

    ' Remove pontuacao final para analise
    Dim textForAnalysis As String
    textForAnalysis = cleanText

    Do While Len(textForAnalysis) > 0 And InStr(".,;:", Right(textForAnalysis, 1)) > 0
        textForAnalysis = Left(textForAnalysis, Len(textForAnalysis) - 1)
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

'--------------------------------------------------------------------------------
' SafeReplaceText - Substitui o texto de um range mantendo a formata��o original
'--------------------------------------------------------------------------------
Public Sub SafeReplaceText(ByVal rng As Range, ByVal newText As String)
    If rng Is Nothing Then Exit Sub
    
    Dim origFont As Font
    Dim origParaFormat As ParagraphFormat
    
    ' Salva a formata��o original com seguran�a
    On Error Resume Next
    Set origFont = rng.Font.Duplicate
    Set origParaFormat = rng.ParagraphFormat.Duplicate
    On Error GoTo 0
    
    ' Realiza a substitui��o
    rng.text = newText
    
    ' Restaura a formata��o original no novo range (que agora cont�m o novo texto)
    On Error Resume Next
    If Not origFont Is Nothing Then rng.Font = origFont
    If Not origParaFormat Is Nothing Then rng.ParagraphFormat = origParaFormat
    On Error GoTo 0
End Sub


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
    LogStepComplete "Normalizacao de quebras"

    LogStepStart "Limpeza estrutural"
    RemovePageNumberLines doc
    RemoveUnderscoreOnlyParagraphs doc
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

    LogStepStart "Substituicao de DAE por Poder Executivo Municipal em Indicacoes"
    ProcessEmentaIndicacao doc
    LogStepComplete "Substituicao de DAE por Poder Executivo Municipal em Indicacoes"

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

    LogStepStart "Formatacao do paragrafo 2 (ementa)"
    FormatSecondParagraph doc
    LogStepComplete "Formatacao do paragrafo 2 (ementa)"

    LogStepStart "Formatacao dos paragrafos 2-4 apos Ementa (justificado, recuo 2,5 cm)"
    FormatPostEmentaBodyParagraphs doc
    LogStepComplete "Formatacao dos paragrafos 2-4 apos Ementa (justificado, recuo 2,5 cm)"

    LogStepStart "Formatacao de considerandos"
    FormatConsiderandoParagraphs doc
    LogStepComplete "Formatacao de considerandos"

    LogStepStart "Insercao de linhas em branco (justificativa, plenario, prefeito)"
    InsertJustificativaBlankLines doc
    LogStepComplete "Insercao de linhas em branco (justificativa, plenario, prefeito)"

    LogStepStart "Aplicacao de substituicoes de texto"
    ApplyTextReplacements doc
    LogStepComplete "Aplicacao de substituicoes de texto"

    LogStepStart "Capitalizacao do inicio de paragrafos"
    CapitalizeFirstLetterOfParagraphs doc
    LogStepComplete "Capitalizacao do inicio de paragrafos"

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
    SubstituiVereadoraPorSenhoraVereadora doc
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

    LogStepStart "Remocao de dois pontos da justificativa"
    RemoveJustificativaColon doc
    LogStepComplete "Remocao de dois pontos da justificativa"

    LogStepStart "Adicao de espacamento especial (ementa, justificativa, data)"
    AddSpecialElementsSpacing doc
    LogStepComplete "Adicao de espacamento especial (ementa, justificativa, data)"

    LogStepStart "Ajuste final de recuos para Vereador (travessoes)"
    FixHyphenatedVereadorParagraphIndents doc
    LogStepComplete "Ajuste final de recuos para Vereador (travessoes)"

    LogStepStart "Insercao final de paragrafo em branco na Ementa (acima e abaixo)"
    ForceEmentaSpacing doc
    LogStepComplete "Insercao final de paragrafo em branco na Ementa (acima e abaixo)"

    LogStepStart "Insercao final de paragrafo em branco na Data (acima)"
    ForceDataSpacing doc
    LogStepComplete "Insercao final de paragrafo em branco na Data (acima)"

    LogMessage "Formatacao completa aplicada com sucesso", LOG_LEVEL_INFO
    LogMetric "Total de paragrafos", doc.Paragraphs.count
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatacao: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' FORMATACAO SELETIVA - SEGUNDA PASSAGEM (OTIMIZADA)
'================================================================================
' Executa apenas as etapas que dependem de indices estruturais,
' que podem ter sido invalidados por remocoes/insercoes de paragrafos
' na primeira passagem. Etapas idempotentes (ClearAllFormatting,
' ApplyStdFont, ApplyStdParagraphs, CapitalizeFirstLetter, etc.)
' sao omitidas pois ja foram aplicadas e nao precisam ser reexecutadas.
'
' Esta otimizacao reduz ~60-70% do tempo da segunda passagem
' em documentos com muitos paragrafos.
'================================================================================

Public Function PreviousFormattingPass2(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' =========================================================================
    ' ETAPAS INDEX-DEPENDENT (indices podem ter mudado apos insercoes/remocoes)
    ' =========================================================================
    LogSection "LIMPEZA E FORMATACAO (PASSAGEM 2 - SELETIVA)"

    LogStepStart "Limpeza de prefixo da ementa (P2)"
    RemoveEmentaLeadingLabelPrefix doc
    LogStepComplete "Limpeza de prefixo da ementa (P2)"

    LogStepStart "Limpeza de sufixo da ementa (P2)"
    RemoveEmentaTrailingMunicipioSuffix doc
    LogStepComplete "Limpeza de sufixo da ementa (P2)"

    LogStepStart "Remocao de aspas da ementa (P2)"
    RemoveEmentaQuotes doc
    LogStepComplete "Remocao de aspas da ementa (P2)"

    LogStepStart "Substituicao de DAE por Poder Executivo Municipal em Indicacoes (P2)"
    ProcessEmentaIndicacao doc
    LogStepComplete "Substituicao de DAE por Poder Executivo Municipal em Indicacoes (P2)"

    LogStepStart "Formatacao de titulo (P2)"
    FormatDocumentTitle doc
    LogStepComplete "Formatacao de titulo (P2)"

    LogSection "FORMATACOES ESPECIFICAS (PASSAGEM 2 - SELETIVA)"

    LogStepStart "Formatacao do paragrafo 2 (ementa) (P2)"
    FormatSecondParagraph doc
    LogStepComplete "Formatacao do paragrafo 2 (ementa) (P2)"

    LogStepStart "Formatacao dos paragrafos 2-4 apos Ementa (P2)"
    FormatPostEmentaBodyParagraphs doc
    LogStepComplete "Formatacao dos paragrafos 2-4 apos Ementa (P2)"

    LogStepStart "Formatacao de considerandos (P2)"
    FormatConsiderandoParagraphs doc
    LogStepComplete "Formatacao de considerandos (P2)"

    LogStepStart "Insercao de linhas em branco (P2)"
    InsertJustificativaBlankLines doc
    LogStepComplete "Insercao de linhas em branco (P2)"

    ' Substituicoes de texto (Find/Replace leves, podem precisar reexecutar)
    LogStepStart "Aplicacao de substituicoes de texto (P2)"
    ApplyTextReplacements doc
    LogStepComplete "Aplicacao de substituicoes de texto (P2)"

    LogSection "LIMPEZA FINAL (PASSAGEM 2 - SELETIVA)"

    LogStepStart "Substituicao de datas do plenario (P2)"
    ReplacePlenarioDateParagraph doc
    LogStepComplete "Substituicao de datas do plenario (P2)"

    LogSection "FINALIZACAO (PASSAGEM 2 - SELETIVA)"

    LogStepStart "Insercao de rodape (P2)"
    If Not InsertFooterStamp(doc) Then
        LogMessage "Falha na insercao do rodape (P2)", LOG_LEVEL_ERROR
        PreviousFormattingPass2 = False
        Exit Function
    End If
    LogStepComplete "Insercao de rodape (P2)"

    LogStepStart "Ajustes finais de negrito e formatacao (P2)"
    ApplyBoldToSpecialParagraphs doc
    SubstituiVereadoraPorSenhoraVereadora doc
    FormatVereadorParagraphs doc
    LogStepComplete "Ajustes finais de negrito e formatacao (P2)"

    LogStepStart "Formatacoes especiais (P2)"
    FormatDianteDoExposto doc
    FormatRequeiroParagraphs doc
    FormatPorTodasRazoesParagraphs doc
    LogStepComplete "Formatacoes especiais (P2)"

    LogStepStart "Remocao de dois pontos da justificativa (P2)"
    RemoveJustificativaColon doc
    LogStepComplete "Remocao de dois pontos da justificativa (P2)"

    LogStepStart "Adicao de espacamento especial (P2)"
    AddSpecialElementsSpacing doc
    LogStepComplete "Adicao de espacamento especial (P2)"

    LogStepStart "Ajuste final de recuos para Vereador (P2)"
    FixHyphenatedVereadorParagraphIndents doc
    LogStepComplete "Ajuste final de recuos para Vereador (P2)"

    LogStepStart "Insercao final de paragrafo em branco na Ementa (P2)"
    ForceEmentaSpacing doc
    LogStepComplete "Insercao final de paragrafo em branco na Ementa (P2)"

    LogStepStart "Insercao final de paragrafo em branco na Data (P2)"
    ForceDataSpacing doc
    LogStepComplete "Insercao final de paragrafo em branco na Data (P2)"

    LogMessage "Formatacao seletiva (Passagem 2) aplicada com sucesso", LOG_LEVEL_INFO
    PreviousFormattingPass2 = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatacao seletiva (Passagem 2): " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormattingPass2 = False
End Function

'================================================================================
' AJUSTE FINAL - Zera recuo de paragrafos com marcador de Vereador/Vereadora (travessoes)
' Ao final do processamento, se existirem paragrafos contendo exatamente essas
' strings, garante recuo a esquerda = 0.
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

    ' Coleta nomes de todos os backups do documento (nome = yyyy-mm-dd_HHmmss -> ordem cronologica)
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
    confirmMsg = "[�] Deseja restaurar o backup mais antigo do documento�" & vbCrLf & vbCrLf & _
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
    Application.StatusBar = RenderProgressBar(30, "Salvando documento descartado")
    doc.SaveAs2 discardedPath

    ' Fecha o documento descartado
    doc.Close SaveChanges:=False

    ' Protege contra conflito no caminho original
    If fso.FileExists(originalPath) Then
        fso.DeleteFile originalPath, True
    End If

    ' Copia o backup mais antigo para o local original
    Application.StatusBar = RenderProgressBar(35, "Restaurando backup")
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
' LIMPEZA DE ESPACOS MULTIPLOS
'================================================================================
