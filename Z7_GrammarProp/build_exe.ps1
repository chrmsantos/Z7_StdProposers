$ErrorActionPreference = "Stop"
$pythonPath = "C:\Users\csantos\AppData\Local\Programs\Python\Python314\python.exe"
$pyinstallerPath = "C:\Users\csantos\AppData\Local\Programs\Python\Python314\Scripts\pyinstaller.exe"

Write-Host "Compilando correct_grammar.py..."
& $pyinstallerPath --onefile --noconsole correct_grammar.py

Write-Host "Compilando config_prompt.py..."
& $pyinstallerPath --onefile --noconsole config_prompt.py

Write-Host "Compilando chat_ia.py..."
& $pyinstallerPath --onefile --noconsole chat_ia.py

Write-Host "Movendo os executáveis para a raiz de Z7_GrammarProp..."
Move-Item -Path "dist\correct_grammar.exe" -Destination ".\correct_grammar.exe" -Force
Move-Item -Path "dist\config_prompt.exe" -Destination ".\config_prompt.exe" -Force
Move-Item -Path "dist\chat_ia.exe" -Destination ".\chat_ia.exe" -Force

Write-Host "Limpando arquivos temporários..."
Remove-Item -Path "build" -Recurse -Force
Remove-Item -Path "dist" -Recurse -Force
Remove-Item -Path "*.spec" -Force

Write-Host "Build concluído com sucesso!"
