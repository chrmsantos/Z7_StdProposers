$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$pyinstallerPath = "C:\Users\csantos\AppData\Local\Programs\Python\Python314\Scripts\pyinstaller.exe"

function Invoke-PyInstaller {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ScriptName
	)

	$process = Start-Process -FilePath $pyinstallerPath -ArgumentList @("--onefile", "--noconsole", $ScriptName) -NoNewWindow -Wait -PassThru
	if ($process.ExitCode -ne 0) {
		throw "Falha ao compilar $ScriptName (exit code: $($process.ExitCode))."
	}
}

Write-Host "Compilando correct_grammar.py..."
Invoke-PyInstaller -ScriptName "correct_grammar.py"

Write-Host "Compilando config_prompt.py..."
Invoke-PyInstaller -ScriptName "config_prompt.py"

Write-Host "Compilando chat_ia.py..."
Invoke-PyInstaller -ScriptName "chat_ia.py"

Write-Host "Movendo os executáveis para a raiz de Z7_GrammarProp..."
Move-Item -Path "dist\correct_grammar.exe" -Destination ".\correct_grammar.exe" -Force
Move-Item -Path "dist\config_prompt.exe" -Destination ".\config_prompt.exe" -Force
Move-Item -Path "dist\chat_ia.exe" -Destination ".\chat_ia.exe" -Force

Write-Host "Limpando arquivos temporários..."
Remove-Item -Path "build" -Recurse -Force
Remove-Item -Path "dist" -Recurse -Force
Remove-Item -Path "*.spec" -Force

Write-Host "Build concluído com sucesso!"
