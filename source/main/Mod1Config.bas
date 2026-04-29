' Mod1Config.bas
Option Explicit

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
Public Const wdAlertsAll As Long = 0
Public Const wdAlertsNone As Long = -1
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

'================================================================================
' CONSTANTES DE SISTEMA
'================================================================================
Public Const Z7_STDPROPOSERS_VERSION As String = "3.0.0"
Public Const MIN_SUPPORTED_VERSION As Long = 14
Public Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"
Public Const MAX_BACKUP_FILES As Long = 10
Public Const DEBUG_MODE As Boolean = False

Public Const LOG_LEVEL_INFO As Long = 1
Public Const LOG_LEVEL_WARNING As Long = 2
Public Const LOG_LEVEL_ERROR As Long = 3

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

'================================================================================
' VARIAVEIS DE IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS
'================================================================================
' Indices dos elementos identificados no documento (0 = nao encontrado)
Public tituloParaIndex As Long
Public ementaParaIndex As Long
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

    LogMessage "Visualizacao configurada: zoom definido para 120%, demais configuracoes preservadas"
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao configurar visualizacao: " & Err.Description, LOG_LEVEL_WARNING
    ConfigureDocumentView = False ' Nao falha o processo por isso
End Function

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM - SISTEMA DE PROTECAO DAS CONFIGURACOES DE VISUALIZACAO
'================================================================================

