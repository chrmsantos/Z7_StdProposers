Attribute VB_Name = "Mod_06_WordMacro"
Option Explicit

' Mod_06_WordMacro.bas
' =============================================================================
' Z7_STDPROPOSERS - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
' Macros para integrar o Word com os apps da suite Z7 StdProposers AI
' ============================================================
' ResolverComando: monta o comando para um app da suite Z7.
'   appName  - nome do executavel sem extensao (ex: "correct_grammar")
'   Retorna o comando pronto para WScript.Shell.Run, ou "" se nao encontrado.
' ============================================================
Private Function ResolverComando(appName As String) As String
    Dim base As String
    base = Environ$("LOCALAPPDATA") & "\Z7\Apps\Z7_StdProposers\ai\"

    ' 1. Executavel compilado (instalado via Install.ps1) -- preferencial.
    Dim caminhoExe As String
    caminhoExe = base & appName & "\" & appName & ".exe"
    If Dir(caminhoExe) <> "" Then
        ResolverComando = """" & caminhoExe & """"
        Exit Function
    End If

    ' 2. Fallback: script .py com pyw -3 (ambiente de desenvolvimento).
    Dim caminhoScript As String
    caminhoScript = base & appName & ".py"
    If Dir(caminhoScript) <> "" Then
        ResolverComando = "pyw -3 """ & caminhoScript & """"
        Exit Function
    End If

    ResolverComando = ""
End Function

' ============================================================
' AbrirChatIA
'   Abre a interface de chat com contexto do documento ativo.
' ============================================================
Sub AbrirChatIA()
    On Error GoTo ErrorHandler

    Dim objShell As Object
    Dim comandoExecucao As String

    comandoExecucao = ResolverComando("chat_ia")

    If comandoExecucao = "" Then
        MsgBox "Executavel 'chat_ia' nao encontrado." & vbCrLf & _
               "Execute 'Install.ps1' para instalar os executaveis.", _
               vbExclamation, "Z7 StdProposers"
        Exit Sub
    End If

    Set objShell = CreateObject("WScript.Shell")
    ' Chat nao bloqueia o Word (False): janela fica aberta enquanto o usuario trabalha.
    objShell.Run comandoExecucao, 0, False

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao abrir Chat IA: " & Err.Description, vbCritical, "Erro de Execucao"
End Sub

' ============================================================
' ConfigurarPrompt
'   Abre a interface de configuracao do prompt Gemini.
' ============================================================
Sub ConfigurarPrompt()
    On Error GoTo ErrorHandler

    Dim objShell As Object
    Dim comandoExecucao As String

    comandoExecucao = ResolverComando("config_prompt")

    If comandoExecucao = "" Then
        MsgBox "Executavel 'config_prompt' nao encontrado." & vbCrLf & _
               "Execute 'Install.ps1' para instalar os executaveis.", _
               vbExclamation, "Z7 StdProposers"
        Exit Sub
    End If

    Set objShell = CreateObject("WScript.Shell")
    objShell.Run comandoExecucao, 0, False

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao abrir Configurar Prompt: " & Err.Description, vbCritical, "Erro de Execucao"
End Sub

