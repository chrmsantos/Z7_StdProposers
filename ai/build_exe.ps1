$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
# Garante que o script roda a partir do proprio diretorio (ai/), independente do cwd do chamador.
$scriptDir = $PSScriptRoot
Push-Location $scriptDir

try {

# Procura pyinstaller.exe: primeiro no PATH, depois nas instalacoes Python do sistema.
# Nao usa o Python do venv ativo para evitar o caso em que pyinstaller nao esta no venv.
$pyinstallerPath = $null
$pyinstallerCmd = Get-Command pyinstaller -ErrorAction SilentlyContinue
if ($pyinstallerCmd) {
    $pyinstallerPath = $pyinstallerCmd.Source
}

if (-not $pyinstallerPath) {
    $searchRoots = @(
        "$env:LOCALAPPDATA\Programs\Python",
        "$env:APPDATA\Python"
    )
    foreach ($root in $searchRoots) {
        $found = Get-ChildItem "$root\*\Scripts\pyinstaller.exe" -ErrorAction SilentlyContinue |
                 Sort-Object FullName -Descending |
                 Select-Object -First 1
        if ($found) { $pyinstallerPath = $found.FullName; break }
    }
}

if (-not $pyinstallerPath) {
    throw "PyInstaller nao encontrado. Execute (fora do venv): pip install pyinstaller"
}

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
	# IMPORTANTE: nao remover o spec nem usar --clean; o Analysis-00.toc em cache
	# evita o bug de Python 3.14 que remove o build dir durante analise completa.
	New-Item -ItemType Directory -Force -Path (Join-Path $scriptDir "build\$baseName") | Out-Null

	# --onedir: DLLs ficam pre-extraidas na pasta, eliminando 1-3s de extração em cada execução
	# --noconfirm: sobrescreve dist sem pedir confirmacao interativa
	# Nota: nao usar --clean pois remove o diretorio pre-criado (workaround bug Python 3.14)
	$pyiArgs = @("--onedir", "--noconsole", "--noconfirm", $scriptPath)
	$process = Start-Process -FilePath $pyinstallerPath -ArgumentList $pyiArgs -WorkingDirectory $scriptDir -NoNewWindow -Wait -PassThru
	if ($process.ExitCode -ne 0) {
		throw "Falha ao compilar $ScriptName (exit code: $($process.ExitCode))."
	}
}

function Install-Executable {
	param([Parameter(Mandatory = $true)][string]$Name)
	$src  = Join-Path $scriptDir "dist\$Name"
	$dest = Join-Path $scriptDir $Name
	if (-not (Test-Path $src)) { throw "dist\$Name nao encontrado apos compilacao." }
	if (Test-Path $dest) { Remove-Item -Path $dest -Recurse -Force }
	Copy-Item -Path $src -Destination $dest -Recurse -Force
	Write-Host "[$Name] instalado."
}

Write-Host "Compilando correct_grammar.py..."
Invoke-PyInstaller -ScriptName "correct_grammar.py"
Install-Executable -Name "correct_grammar"

Write-Host "Compilando config_prompt.py..."
Invoke-PyInstaller -ScriptName "config_prompt.py"
Install-Executable -Name "config_prompt"

Write-Host "Compilando chat_ia.py..."
Invoke-PyInstaller -ScriptName "chat_ia.py"
Install-Executable -Name "chat_ia"

Write-Host "Limpando arquivos temporarios..."
# Nao remover build/ - o cache Analysis-00.toc evita bug Python 3.14 na proxima execucao
Remove-Item -Path (Join-Path $scriptDir "dist") -Recurse -Force -ErrorAction SilentlyContinue
Get-Item (Join-Path $scriptDir "*.spec") -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "Build concluido com sucesso!"

} finally {
    Pop-Location
}
