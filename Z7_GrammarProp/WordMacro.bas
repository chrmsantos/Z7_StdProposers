Attribute VB_Name = "MacroGeminiGrammar"
Sub CorrigirGramaticaComGemini()
    ' Macro para enviar o texto selecionado para correção via script Python
    ' usando a API do Gemini.
    
    Dim objShell As Object
    Dim comandoExecucao As String
    Dim caminhoScript As String
    Dim caminhoPython As String
    
    ' DEFINE O CAMINHO DO SCRIPT PYTHON AQUI:
    ' Se a pasta for diferente, atualize o caminho abaixo.
    caminhoScript = "C:\Users\csantos\AppData\Local\Z7\Apps\Z7_GrammarProp\correct_grammar.py"
    
    ' Comando base. Usamos 'pythonw' para executar sem exibir a janela preta do console (prompt de comando).
    ' Se quiser ver o console para buscar erros, troque "pythonw" por "python".
    caminhoPython = "pythonw"
    
    ' Monta o comando completo com aspas em volta do caminho do script para evitar problemas com espaços
    comandoExecucao = caminhoPython & " """ & caminhoScript & """"
    
    ' Cria o objeto WScript.Shell
    Set objShell = CreateObject("WScript.Shell")
    
    ' Muda o ponteiro do mouse para indicar carregamento
    Application.Cursor = wdCursorWait
    
    ' Executa o comando.
    ' O primeiro parâmetro (0) oculta a janela.
    ' O segundo parâmetro (True) faz o VBA aguardar até que o Python termine a execução.
    objShell.Run comandoExecucao, 0, True
    
    ' Retorna o ponteiro do mouse ao normal
    Application.Cursor = wdCursorNormal
    
    ' Como o script Python já substitui o texto diretamente via COM/win32com, 
    ' não precisamos manipular o Word a partir daqui, apenas avisamos o fim.
    ' (Opcional: remova o comentário da linha abaixo para ter um popup de aviso).
    ' MsgBox "Correção finalizada!", vbInformation, "Revisor Gemini"

End Sub
