$ErrorActionPreference = "Stop"
$ai = "C:\Users\csantos\AppData\Local\Z7\Apps\Z7_StdProposers\ai"

$pyi = (Get-Command pyinstaller -ErrorAction SilentlyContinue)
if ($pyi) { $pyiPath = $pyi.Source }
else { $pyiPath = (Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\*\Scripts\pyinstaller.exe" | Sort-Object FullName -Descending | Select-Object -First 1).FullName }

# Compila APENAS correct_grammar.py (script pequeno, sem bug Python 3.14)
Write-Host "Compilando correct_grammar.py..."
New-Item -ItemType Directory -Force -Path "$ai\build\correct_grammar" | Out-Null
$proc = Start-Process -FilePath $pyiPath -ArgumentList @("--onedir","--noconsole","--noconfirm","$ai\correct_grammar.py") -WorkingDirectory $ai -NoNewWindow -Wait -PassThru
Write-Host "correct_grammar exit: $($proc.ExitCode)"
if ($proc.ExitCode -eq 0) {
    if (Test-Path "$ai\correct_grammar") { Remove-Item -Path "$ai\correct_grammar" -Recurse -Force }
    Copy-Item -Path "$ai\dist\correct_grammar" -Destination "$ai\correct_grammar" -Recurse -Force
    Write-Host "correct_grammar instalado com fix do status bar."
} else {
    Write-Host "ERRO: correct_grammar nao compilou."
}
