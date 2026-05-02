$inputFile = "source\main\Proposers.bas"
$outputDir = "source\main"

$mappings = @{
    "Mod1Infrastructure" = @("CONSTANTES", "VARIAVEIS GLOBAIS", "VARIAVEIS DE IDENTIFICACAO", "Indices dos elementos", "GERENCIAMENTO DE ESTADO", "CONFIGURE DOCUMENT VIEW", "VIEW SETTINGS PROTECTION SYSTEM", "SISTEMA DE LOGS", "BARRA DE PROGRESSO", "TRATAMENTO.*ERROS", "RECUPERACAO DE EMERGENCIA", "SISTEMA DE BACKUP", "ATUALIZACAO DA BARRA", "GERENCIAMENTO DE DIRETORIO DE BACKUP", "VERIFICACAO DE VERSAO E ATUALIZACAO", "SALVAMENTO INICIAL", "FUNCOES AUXILIARES DE LIMPEZA DE TEXTO", "FUNCOES DE CAMINHO", "UTILITARIO", "ACESSO SEGURO", "RETORNA O MINIMO", "CONTA DIGITOS", "NORMALIZA TEXTO PARA COMPARACAO", "CALCULA A DISTANCIA", "GET.*")
    "Mod4Main" = @("PONTO DE ENTRADA PRINCIPAL", "SUBROTINAS PUBLICAS", "API PUBLICA", "FUNCOES PUBLICAS", "CONCLUIR", "ABRIR REPOSITORIO", "CONFIRMAR DESFAZIMENTO", "DESFAZER COM CONFIRMACAO")
    "Mod2Engine" = @("CACHE", "IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS", "FUNCOES DE VALIDACAO", "IS DOCUMENT HEALTHY", "IS OPERATION TIMEOUT", "VERIFICACAO DE VERSAO DO WORD", "VERIFICA DADOS", "VALIDACAO", "DETECTA", "VERIFICA CONSISTENCIA", "EXTRAI", "OBTEM TEXTO DA", "LOCALIZA O PARAGRAFO", "VERIFICA SE", "IMAGENS", "IMAGEM", "LISTAS", "LIST FORMATS", "INSERCAO DE IMAGEM")
}

$lines = Get-Content $inputFile

$outputFiles = @{}
$currentFile = "Mod1Infrastructure.bas"
$outputFiles[$currentFile] = New-Object System.Collections.Generic.List[String]
$outputFiles[$currentFile].Add("Option Explicit`r`n") 

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    
    # Check if section header
    if ($line.Trim().StartsWith("'") -and $line.Contains("=====")) {
        if ($i + 1 -lt $lines.Count -and -not $lines[$i+1].Trim().StartsWith("'=====") -and $lines[$i+1].Trim().StartsWith("'")) {
            $secName = $lines[$i+1].Trim().Substring(1).Trim()
            
            # Check mappings
            $matched = $false
            foreach ($key in $mappings.Keys) {
                foreach ($pattern in $mappings[$key]) {
                    if ($secName -match "(?i)$pattern") {
                        $currentFile = "$key.bas"
                        if (-not $outputFiles.ContainsKey($currentFile)) {
                            $outputFiles[$currentFile] = New-Object System.Collections.Generic.List[String]
                            $outputFiles[$currentFile].Add("Option Explicit`r`n")
                        }
                        $matched = $true
                        break
                    }
                }
                if ($matched) { break }
            }
            if (-not $matched) {
                $currentFile = "Mod3Pipeline.bas"
                if (-not $outputFiles.ContainsKey($currentFile)) {
                    $outputFiles[$currentFile] = New-Object System.Collections.Generic.List[String]
                    $outputFiles[$currentFile].Add("Option Explicit`r`n")
                }
            }
        }
    }

    if ($line.Trim() -eq "Option Explicit") {
        continue
    }

    $processedLine = $line -replace '(?i)^(\s*)Private\s+', '$1Public '
    
    $outputFiles[$currentFile].Add($processedLine)
}

foreach ($file in $outputFiles.Keys) {
    if ($file -ne "Mod1Infrastructure.bas" -and $file -ne "Mod4Main.bas" -and $file -ne "Mod2Engine.bas" -and $file -ne "Mod3Pipeline.bas") {
        continue
    }
    $path = Join-Path $outputDir $file
    [System.IO.File]::WriteAllLines($path, $outputFiles[$file], [System.Text.Encoding]::UTF8)
    Write-Host "Created $path with $($outputFiles[$file].Count) lines"
}

Write-Host "Split complete."
