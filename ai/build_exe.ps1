$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
# Garante que o script roda a partir do proprio diretorio (ai/), independente do cwd do chamador.
$scriptDir   = $PSScriptRoot
$projectRoot = Split-Path $scriptDir -Parent
$distDir     = Join-Path $projectRoot "dist"
$version     = (Get-Content (Join-Path $projectRoot "VERSION") -Raw).Trim()
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
		[string]$ScriptName,
		[switch]$OneFile
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
	$mode = if ($OneFile) { "--onefile" } else { "--onedir" }
	$pyiArgs = @($mode, "--noconsole", "--noconfirm", "--hidden-import=unicodedata", "--hidden-import=openai", "--hidden-import=jiter", "--hidden-import=certifi", "--hidden-import=pythoncom", "--hidden-import=win32com.client", "--hidden-import=win32com", $scriptPath)
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

function Package-Artifact {
	param([Parameter(Mandatory = $true)][string]$Name)
	$src  = Join-Path $scriptDir $Name
	$dest = Join-Path $distDir "$Name-v$version.zip"
	if (-not (Test-Path $src)) { throw "Pasta $Name nao encontrada para empacotar." }
	New-Item -ItemType Directory -Force -Path $distDir | Out-Null
	Remove-Item $dest -ErrorAction SilentlyContinue
	# Compress-Archive falha com arquivos .zip aninhados (ex: base_library.zip) no PS 5.1.
	# Usar a API .NET diretamente contorna esse problema.
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	[System.IO.Compression.ZipFile]::CreateFromDirectory($src, $dest)
	$sizeMB = [math]::Round((Get-Item $dest).Length / 1MB, 2)
	Write-Host "[$Name] empacotado -> dist\$Name-v$version.zip ($sizeMB MB)"
}



# ── Tcl/Tk environment for PyInstaller + tkinter ──────────────────────
# PyInstaller's _tkinter hook needs TCL_LIBRARY / TK_LIBRARY to bundle the
# Tcl/Tk data directories (_tcl_data / _tk_data).  Detect them from the
# active Python and export before invoking PyInstaller.
function Set-TkinterEnvironment {
	Write-Host "Detectando Tcl/Tk para empacotamento..."
	$pyLines = & python -c @"
import os, sys
prefix = sys.prefix
tcl_base = os.path.join(prefix, 'tcl')
if os.path.isdir(tcl_base):
    # Prefer the most specific versioned directory (e.g. tcl8.6 over tcl8)
    tcl_best = ''
    tk_best  = ''
    for name in os.listdir(tcl_base):
        full = os.path.join(tcl_base, name)
        if not os.path.isdir(full):
            continue
        low = name.lower()
        if low.startswith('tcl') and '.' in low:
            if low > tcl_best:
                tcl_best = low
                print('TCL_LIBRARY=' + full)
        elif low.startswith('tk') and '.' in low:
            if low > tk_best:
                tk_best = low
                print('TK_LIBRARY=' + full)
"@ 2>&1

	foreach ($line in $pyLines) {
		$s = "$line".Trim()
		if ($s -match '^TCL_LIBRARY=(.+)$') {
			$env:TCL_LIBRARY = $Matches[1]
			Write-Host "  TCL_LIBRARY = $env:TCL_LIBRARY"
		}
		elseif ($s -match '^TK_LIBRARY=(.+)$') {
			$env:TK_LIBRARY = $Matches[1]
			Write-Host "  TK_LIBRARY  = $env:TK_LIBRARY"
		}
	}

	if (-not $env:TCL_LIBRARY -or -not $env:TK_LIBRARY) {
		Write-Warning "Nao foi possivel detectar TCL_LIBRARY/TK_LIBRARY. O executavel pode falhar ao iniciar."
	}
}
Set-TkinterEnvironment

# ── Importa modulos .bas para Normal.dotm antes de empacotar ────────────
$importBasScript = Join-Path $projectRoot "scripts\import_bas_to_normal.py"
if (Test-Path $importBasScript) {
    Write-Host "Importando modulos .bas para Normal.dotm..."
    Push-Location $projectRoot
    & python $importBasScript
    Pop-Location
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Falha ao importar .bas para Normal.dotm. Continuando build..."
    }
} else {
    Write-Warning "Script import_bas_to_normal.py nao encontrado em: $importBasScript"
}

Write-Host "Compilando config_prompt.py..."
Invoke-PyInstaller -ScriptName "config_prompt.py"
Install-Executable -Name "config_prompt"
Package-Artifact   -Name "config_prompt"

Write-Host "Compilando chat_ia.py..."
Invoke-PyInstaller -ScriptName "chat_ia.py"
Install-Executable -Name "chat_ia"
Package-Artifact   -Name "chat_ia"

# ── Compile scripts/ utilities ───────────────────────────────────────────
$scriptsDir = Join-Path $projectRoot "scripts"
$importBasScript = Join-Path $scriptsDir "import_bas_to_normal.py"
if (Test-Path $importBasScript) {
    Write-Host "Compilando import_bas_to_normal.py..."
    $baseName = "import_bas_to_normal"
    New-Item -ItemType Directory -Force -Path (Join-Path $scriptsDir "build\$baseName") | Out-Null
    $pyiArgs = @("--onefile", "--noconsole", "--noconfirm",
                 "--hidden-import=unicodedata",
                 "--hidden-import=pythoncom",
                 "--hidden-import=win32com.client",
                 "--hidden-import=win32com",
                 $importBasScript)
    $process = Start-Process -FilePath $pyinstallerPath -ArgumentList $pyiArgs `
        -WorkingDirectory $scriptsDir -NoNewWindow -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        Write-Warning "Falha ao compilar import_bas_to_normal.py (exit code: $($process.ExitCode))."
    } else {
        $exeSrc = Join-Path $scriptsDir "dist\$baseName.exe"
        $exeDest = Join-Path $distDir "$baseName.exe"
        New-Item -ItemType Directory -Force -Path $distDir | Out-Null
        if (Test-Path $exeSrc) {
            Copy-Item $exeSrc $exeDest -Force
            Write-Host "[$baseName] copiado para dist\."
        }
        # Clean up scripts/ build artifacts
        Remove-Item -Path (Join-Path $scriptsDir "dist") -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path (Join-Path $scriptsDir "build") -Recurse -Force -ErrorAction SilentlyContinue
        Get-Item (Join-Path $scriptsDir "*.spec") -ErrorAction SilentlyContinue | Remove-Item -Force
    }
}

Write-Host "Limpando arquivos temporarios..."
# Nao remover build/ - o cache Analysis-00.toc evita bug Python 3.14 na proxima execucao
Remove-Item -Path (Join-Path $scriptDir "dist") -Recurse -Force -ErrorAction SilentlyContinue
Get-Item (Join-Path $scriptDir "*.spec") -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "Build concluido com sucesso!"

} finally {
    Pop-Location
}
