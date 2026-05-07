$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$pyinstallerPath = "C:\Users\csantos\AppData\Local\Programs\Python\Python314\Scripts\pyinstaller.exe"

function Invoke-PyInstaller {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ScriptName
	)

	$baseName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)
	# Pre-create build directory to avoid PyInstaller bug with Python 3.14
	New-Item -ItemType Directory -Force -Path "build\$baseName" | Out-Null
	# Remove stale spec to force a clean analysis
	if (Test-Path "$baseName.spec") { Remove-Item "$baseName.spec" -Force }

	# --onedir: DLLs ficam pre-extraidas na pasta, eliminando 1-3s de extração em cada execução
	# --noconfirm: sobrescreve dist sem pedir confirmacao interativa
	# Nota: nao usar --clean pois remove o diretorio pre-criado (workaround bug Python 3.14)
	$process = Start-Process -FilePath $pyinstallerPath -ArgumentList @("--onedir", "--noconsole", "--noconfirm", $ScriptName) -NoNewWindow -Wait -PassThru
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

Write-Host "Instalando pastas de executaveis..."
foreach ($name in @("correct_grammar", "config_prompt", "chat_ia")) {
	# Remove instalação anterior para evitar DLLs obsoletas
	if (Test-Path ".\$name") { Remove-Item -Path ".\$name" -Recurse -Force }
	Copy-Item -Path "dist\$name" -Destination ".\$name" -Recurse -Force
}

Write-Host "Limpando arquivos temporarios..."
Remove-Item -Path "build" -Recurse -Force
Remove-Item -Path "dist" -Recurse -Force
Remove-Item -Path "*.spec" -Force

Write-Host "Build concluido com sucesso!"

