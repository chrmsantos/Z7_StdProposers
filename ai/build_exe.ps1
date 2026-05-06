$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$pyinstallerPath = "C:\Users\csantos\AppData\Local\Programs\Python\Python314\Scripts\pyinstaller.exe"
$scriptDir = $PSScriptRoot

function Invoke-PyInstaller {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ScriptName
	)

	$baseName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)
	$scriptPath = Join-Path $scriptDir $ScriptName

	if (-not (Test-Path $scriptPath)) {
		throw "Script nao encontrado: $scriptPath"
	}

	# Pre-create build directory to avoid PyInstaller bug with Python 3.14
	New-Item -ItemType Directory -Force -Path (Join-Path $scriptDir "build\$baseName") | Out-Null
	# Remove stale spec to force a clean analysis
	$specPath = Join-Path $scriptDir "$baseName.spec"
	if (Test-Path $specPath) { Remove-Item $specPath -Force }

	# --onedir: DLLs ficam pre-extraidas na pasta, eliminando 1-3s de extração em cada execução
	$process = Start-Process -FilePath $pyinstallerPath -ArgumentList @("--onedir", "--noconsole", "--clean", $scriptPath) -WorkingDirectory $scriptDir -NoNewWindow -Wait -PassThru
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
	$dest = Join-Path $scriptDir $name
	# Remove instalação anterior para evitar DLLs obsoletas
	if (Test-Path $dest) { Remove-Item -Path $dest -Recurse -Force }
	Copy-Item -Path (Join-Path $scriptDir "dist\$name") -Destination $dest -Recurse -Force
}

Write-Host "Limpando arquivos temporarios..."
Remove-Item -Path (Join-Path $scriptDir "build") -Recurse -Force
Remove-Item -Path (Join-Path $scriptDir "dist") -Recurse -Force
Get-Item (Join-Path $scriptDir "*.spec") -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "Build concluido com sucesso!"

