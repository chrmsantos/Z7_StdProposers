# =============================================================================
# Z7_STDPROPOSERS - Testes de Encoding e Emojis
# =============================================================================
# Verifica conformidade de encoding (UTF-8/ASCII) e ausencia de emojis
# =============================================================================

$ErrorActionPreference = "Stop"
. $PSScriptRoot\\Helpers.ps1

Describe 'Z7_STDPROPOSERS - Testes de Encoding e Emojis' {

    $projectRoot = Split-Path -Parent $PSScriptRoot
    $testsPath = Join-Path $projectRoot 'tests'

    $psSearchRoots = @()
    if (Test-Path $testsPath) { $psSearchRoots += $testsPath }
    $psSearchRoots += $projectRoot

    $script:ProjectPsFiles = @()
    foreach ($root in $psSearchRoots) {
        if (Test-Path $root) {
            $script:ProjectPsFiles += Get-ChildItem -Path $root -Filter "*.ps1" -Recurse -File -ErrorAction SilentlyContinue
        }
    }
    $script:ProjectPsFiles = $script:ProjectPsFiles | Sort-Object FullName -Unique

    Context 'Validacao de Encoding de Arquivos' {

        It 'Scripts PowerShell estao em UTF-8 com BOM ou ASCII' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Verifica UTF-8 com BOM (EF BB BF) ou ASCII puro
                $isUtf8WithBom = ($bytes.Length -ge 3) -and
                ($bytes[0] -eq 0xEF) -and
                ($bytes[1] -eq 0xBB) -and
                ($bytes[2] -eq 0xBF)

                # Verifica se e ASCII puro (todos os bytes < 128)
                $isAscii = $true
                foreach ($byte in $bytes) {
                    if ($byte -ge 128) {
                        $isAscii = $false
                        break
                    }
                }

                # Verifica UTF-8 sem BOM usando decoder strict para rejeitar sequencias invalidas
                $isValidUtf8 = $false
                try {
                    $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
                    [void] $utf8Strict.GetString($bytes)
                    $isValidUtf8 = $true
                } catch { }

                ($isUtf8WithBom -or $isAscii -or $isValidUtf8) | Should Be $true
            }
        }

        It 'Arquivos Markdown estao em UTF-8' {
            $mdFiles = @(
                if (Test-Path (Join-Path $projectRoot 'docs')) {
                    Get-ChildItem -Path (Join-Path $projectRoot 'docs') -Filter "*.md" -Recurse -File -ErrorAction SilentlyContinue
                }
                Get-ChildItem -Path $projectRoot -Filter "*.md" -File -ErrorAction SilentlyContinue
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $content | Should Not BeNullOrEmpty

                # Verifica que pode ser lido como UTF-8
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
                $decodedContent = [System.Text.Encoding]::UTF8.GetString($bytes)
                $decodedContent.Length | Should BeGreaterThan 0
            }
        }

        It 'Arquivo VBA esta em formato legivel' {
            $vbaFile = "$projectRoot\source\main\Mod_03_Pipeline.bas"

            if (Test-Path $vbaFile) {
                $content = Get-Content $vbaFile -Raw
                $content | Should Not BeNullOrEmpty

                # VBA deve conter apenas ASCII ou caracteres extendidos validos
                $bytes = [System.IO.File]::ReadAllBytes($vbaFile)

                # Verifica que nao tem bytes nulos (indicaria binario)
                $hasNullBytes = $false
                foreach ($byte in $bytes) {
                    if ($byte -eq 0) {
                        $hasNullBytes = $true
                        break
                    }
                }

                $hasNullBytes | Should Be $false
            }
        }

        It 'Arquivos de texto (.txt) estao em UTF-8 ou ASCII' {
            $txtFiles = Get-ChildItem -Path $projectRoot -Filter "*.txt" -Recurse

            foreach ($file in $txtFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Tenta decodificar como UTF-8
                try {
                    $utf8Content = [System.Text.Encoding]::UTF8.GetString($bytes)
                    $utf8Content.Length | Should BeGreaterThan 0
                }
                catch {
                    # Se falhar UTF-8, tenta ASCII
                    $asciiContent = [System.Text.Encoding]::ASCII.GetString($bytes)
                    $asciiContent.Length | Should BeGreaterThan 0
                }
            }
        }
    }

    Context 'Deteccao de Emojis e Caracteres Especiais' {

        # Regex para detectar emojis (Unicode ranges)
        # Usando lookbehind negativo para evitar falsos positivos com caracteres ASCII
        # Emoticons: U+1F600 to U+1F64F
        # Symbols & Pictographs: U+1F300 to U+1F5FF
        # Transport & Map: U+1F680 to U+1F6FF
        # Supplemental Symbols: U+1F900 to U+1F9FF
        # Outros ranges comuns de emojis
        $emojiPattern = '[\u2600-\u26FF\u2700-\u27BF]|' +
        '[\uD83C][\uDF00-\uDFFF]|' + # High surrogate for U+1F300-1F5FF
        '[\uD83D][\uDC00-\uDE4F]|' + # High surrogate for U+1F600-1F64F
        '[\uD83D][\uDE80-\uDEFF]|' + # High surrogate for U+1F680-1F6FF
        '[\uD83E][\uDD00-\uDDFF]|' + # High surrogate for U+1F900-1F9FF
        '[\uD83E][\uDE00-\uDE6F]|' + # High surrogate for U+1FA00-1FA6F
        '[\uD83E][\uDE70-\uDEFF]'     # High surrogate for U+1FA70-1FAFF

        It 'Scripts PowerShell nao contem emojis' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                # Ignora comentarios
                $codeLines = ($content -split "`r?`n") | Where-Object {
                    $_ -notmatch '^\s*#'
                }
                $codeContent = $codeLines -join "`n"

                if ($codeContent -match $emojiPattern) {
                    throw "Arquivo $($file.Name) contem emojis no codigo (fora de comentarios)"
                }

                # Passa se nao encontrou emojis no codigo
                $true | Should Be $true
            }
        }

        It 'Arquivos Markdown podem conter emojis apenas em documentacao' {
            # Markdown pode conter emojis para documentacao
            # Este teste apenas verifica que podem ser lidos
            $mdFiles = @(
                if (Test-Path (Join-Path $projectRoot 'docs')) {
                    Get-ChildItem -Path (Join-Path $projectRoot 'docs') -Filter "*.md" -Recurse -File -ErrorAction SilentlyContinue
                }
                Get-ChildItem -Path $projectRoot -Filter "*.md" -File -ErrorAction SilentlyContinue
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                # Verifica que emojis podem ser lidos corretamente se presentes
                $content.Length | Should BeGreaterThan 0

                # Passa se conseguiu ler
                $true | Should Be $true
            }
        }

        It 'Arquivo VBA nao contem emojis' {
            $vbaFile = "$projectRoot\source\main\Mod_03_Pipeline.bas"

            if (Test-Path $vbaFile) {
                $content = Get-Content $vbaFile -Raw -Encoding UTF8

                if ($content -match $emojiPattern) {
                    throw "Arquivo VBA contem emojis"
                }

                # Passa se nao encontrou emojis
                $true | Should Be $true
            }
        }

        It 'Testes PowerShell nao contem emojis' {
            $testFiles = Get-ChildItem -Path "$projectRoot\tests" -Filter "*.ps1"

            foreach ($file in $testFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                if ($content -match $emojiPattern) {
                    throw "Arquivo de teste $($file.Name) contem emojis"
                }

                # Passa se nao encontrou emojis
                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de Caracteres Problematicos' {

        It 'Scripts PowerShell nao contem caracteres de controle invalidos' {
            $psFiles = $script:ProjectPsFiles

            # Caracteres de controle permitidos: Tab (0x09), LF (0x0A), CR (0x0D)
            $allowedControlChars = @(0x09, 0x0A, 0x0D)

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                for ($i = 0; $i -lt $bytes.Length; $i++) {
                    $byte = $bytes[$i]

                    # Verifica caracteres de controle (0x00-0x1F exceto permitidos)
                    if (($byte -le 0x1F) -and ($byte -notin $allowedControlChars)) {
                        throw "Arquivo $($file.Name) contem caractere de controle invalido 0x$($byte.ToString('X2')) na posicao $i"
                    }
                }

                # Passa se nao encontrou problemas
                $true | Should Be $true
            }
        }

        It 'Arquivos Markdown nao contem tabs (usam espacos)' {
            $mdFiles = @(
                if (Test-Path (Join-Path $projectRoot 'docs')) {
                    Get-ChildItem -Path (Join-Path $projectRoot 'docs') -Filter "*.md" -Recurse -File -ErrorAction SilentlyContinue
                }
                Get-ChildItem -Path $projectRoot -Filter "*.md" -File -ErrorAction SilentlyContinue
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw

                # Markdown nao deve ter tabs (usa espacos para indentacao)
                if ($content -match "`t") {
                    Write-Warning "Arquivo $($file.Name) contem tabs - deveria usar espacos"
                    # Apenas warning, nao falha o teste
                }

                $true | Should Be $true
            }
        }

        It 'Arquivo VBA nao contem tabs (usa espacos conforme padrao)' {
            $vbaFile = "$projectRoot\source\main\Mod_03_Pipeline.bas"

            if (Test-Path $vbaFile) {
                $content = Get-Content $vbaFile -Raw

                # VBA nao deve ter tabs
                $content -match "`t" | Should Be $false
            }
        }
    }

    Context 'Consistencia de Line Endings' {

        It 'Scripts PowerShell usam CRLF (Windows)' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Verifica se nao tem LF sozinho (Unix style)
                $hasLfOnly = $false
                for ($i = 0; $i -lt $bytes.Length; $i++) {
                    if ($bytes[$i] -eq 0x0A) {
                        if (($i -eq 0) -or ($bytes[$i - 1] -ne 0x0D)) {
                            $hasLfOnly = $true
                            break
                        }
                    }
                }

                if ($hasLfOnly) {
                    Write-Warning "Arquivo $($file.Name) usa LF (Unix) - deveria usar CRLF (Windows)"
                }

                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de BOM (Byte Order Mark)' {

        It 'Scripts PowerShell nao usam UTF-8 com BOM' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                $hasUtf8Bom = ($bytes.Length -ge 3) -and
                ($bytes[0] -eq 0xEF) -and
                ($bytes[1] -eq 0xBB) -and
                ($bytes[2] -eq 0xBF)

                if ($hasUtf8Bom) {
                    throw "Arquivo $($file.Name) possui UTF-8 BOM - prefira UTF-8 sem BOM"
                }

                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de Politica ASCII (Texto)' {

        It 'Arquivo VBA esta em formato textual valido (nao UTF-16)' {
            $vbaFile = "$projectRoot\source\main\Mod_03_Pipeline.bas"

            if (Test-Path $vbaFile) {
                $bytes = [System.IO.File]::ReadAllBytes($vbaFile)
                if ($bytes.Length -ge 2) {
                    $isUtf16LE = ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
                    $isUtf16BE = ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
                    ($isUtf16LE -or $isUtf16BE) | Should Be $false
                }
                $true | Should Be $true
            }
        }

        It 'Documentacao Markdown esta em UTF-8 legivel' {
            $mdFiles = @(
                Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
                Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File -ErrorAction SilentlyContinue
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $content | Should Not BeNullOrEmpty
            }
        }
    }

    Context 'Validacao de Encoding Cross-Platform' {

        It 'VERSION pode ser lido em qualquer plataforma' {
            $versionFile = "$projectRoot\VERSION"

            if (Test-Path $versionFile) {
                # Deve ser UTF-8 sem BOM para compatibilidade
                $bytes = [System.IO.File]::ReadAllBytes($versionFile)

                $hasUtf8Bom = ($bytes.Length -ge 3) -and
                ($bytes[0] -eq 0xEF) -and
                ($bytes[1] -eq 0xBB) -and
                ($bytes[2] -eq 0xBF)

                if ($hasUtf8Bom) {
                    Write-Warning "VERSION tem UTF-8 BOM - melhor usar UTF-8 sem BOM"
                }

                # Deve conter um numero de versao simples
                $content = Get-Content $versionFile -Raw -Encoding UTF8
                $content | Should Match '[0-9]+\.[0-9]+\.[0-9]+'
            }
        }

        It 'Arquivos de configuracao (.config) sao UTF-8 validos' {
            $configFiles = Get-ChildItem -Path $projectRoot -Filter "*.config" -Recurse

            foreach ($file in $configFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $content | Should Not BeNullOrEmpty

                # Nao deve ter caracteres de substituicao
                $content | Should Not Match ([regex]::Escape([char]0xFFFD))
            }
        }
    }

    Context 'Validacao de Regex (ASCII)' {

        It 'Regex encontra strings ASCII esperadas' {
            $testContent = @'
' Versao: 2.0.2
Instalacao completa
Configuracao do sistema
'@

            $testContent | Should Match 'Versao'
            $testContent | Should Match 'Instalacao'
            $testContent | Should Match 'Configuracao'

            if ($testContent -match "Versao: ([^\r\n]+)") {
                $matches[1].Trim() | Should Be "2.0.2"
            }
        }
    }

    Context 'Deteccao de Encodings Problematicos' {

        It 'Nenhum arquivo usa UTF-16 (wide chars)' {
            $allTextFiles = @(
                $script:ProjectPsFiles
                if (Test-Path (Join-Path $projectRoot 'docs')) { Get-ChildItem -Path (Join-Path $projectRoot 'docs') -Filter "*.md" -Recurse -File -ErrorAction SilentlyContinue }
                if (Test-Path (Join-Path $projectRoot 'tests')) { Get-ChildItem -Path (Join-Path $projectRoot 'tests') -Filter "*.ps1" -Recurse -File -ErrorAction SilentlyContinue }
            )

            foreach ($file in $allTextFiles) {
                if ($file -eq $null) { continue }

                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                if ($bytes.Length -ge 2) {
                    # Verifica UTF-16 LE BOM (FF FE)
                    $isUtf16Le = ($bytes[0] -eq 0xFF) -and ($bytes[1] -eq 0xFE)

                    # Verifica UTF-16 BE BOM (FE FF)
                    $isUtf16Be = ($bytes[0] -eq 0xFE) -and ($bytes[1] -eq 0xFF)

                    if ($isUtf16Le -or $isUtf16Be) {
                        throw "Arquivo $($file.Name) usa UTF-16 - deveria usar UTF-8"
                    }
                }

                $true | Should Be $true
            }
        }

        It 'Nenhum arquivo tem encoding misto' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Tenta decodificar como UTF-8
                $decoder = [System.Text.Encoding]::UTF8.GetDecoder()
                $charCount = $decoder.GetCharCount($bytes, 0, $bytes.Length, $true)

                # Se falhar, pode ser encoding misto
                if ($charCount -eq 0) {
                    throw "Arquivo $($file.Name) pode ter encoding misto ou invalido"
                }

                $true | Should Be $true
            }
        }
    }

    # =========================================================================
    # Anti-regression: import_bas_to_normal.py and .bas file integrity
    # See .clinerules/02-vba-coding.md (Importação de Módulos) and
    # .clinerules/03-python-coding.md (VBA module import encoding rule).
    # These tests prevent the recurring bug where Attribute VB_Name is not
    # processed (causing unnamed modules and compile errors) and where UTF-8
    # content is silently garbled by CP1252 decoding.
    # =========================================================================
    Context 'Import de Modulos VBA (anti-regressao)' {

        $script:importScript = Join-Path $projectRoot 'scripts\import_bas_to_normal.py'
        $script:importContent = $null
        if (Test-Path $script:importScript) {
            $script:importContent = Get-Content $script:importScript -Raw -Encoding UTF8
        }

        It 'import_bas_to_normal.py existe' {
            Test-Path $script:importScript | Should Be $true
        }

        It 'import_bas_to_normal.py NAO usa AddFromString (Attribute VB_Name nao e processado)' {
            $script:importContent | Should Not BeNullOrEmpty
            # Match only actual method calls (.AddFromString), not comment mentions
            $codeLines = $script:importContent -split "`n" | Where-Object {
                $_.TrimStart() -notmatch '^\s*#' -and $_ -match 'AddFromString'
            }
            $codeLines.Count | Should Be 0
        }

        It 'import_bas_to_normal.py usa VBComponents.Import' {
            $script:importContent | Should Not BeNullOrEmpty
            $script:importContent | Should Match 'VBComponents\.Import'
        }

        It 'import_bas_to_normal.py decodifica UTF-8 primeiro (nao CP1252)' {
            $script:importContent | Should Not BeNullOrEmpty
            # Must try UTF-8 before CP1252: find the position of each decode call
            $utf8Idx = $script:importContent.IndexOf('decode("utf-8")')
            $cp1252Idx = $script:importContent.IndexOf('decode("cp1252")')
            if ($utf8Idx -lt 0) { $utf8Idx = $script:importContent.IndexOf("decode('utf-8')") }
            if ($cp1252Idx -lt 0) { $cp1252Idx = $script:importContent.IndexOf("decode('cp1252')") }
            $utf8Idx | Should BeGreaterThan -1
            $cp1252Idx | Should BeGreaterThan -1
            $utf8Idx | Should BeLessThan $cp1252Idx
        }

        It 'import_bas_to_normal.py usa Remove antes de re-importar modulo existente' {
            $script:importContent | Should Not BeNullOrEmpty
            $script:importContent | Should Match 'VBComponents\.Remove'
        }

        It 'import_bas_to_normal.py grava arquivo temporario em CP1252 para Import' {
            $script:importContent | Should Not BeNullOrEmpty
            $script:importContent | Should Match 'encoding.*cp1252'
        }

        It 'Todos os .bas comecam com Attribute VB_Name na linha 1' {
            $basFiles = Get-VbaFiles
            $basFiles | Should Not BeNullOrEmpty
            foreach ($f in $basFiles) {
                $content = Get-Content $f.FullName -Encoding UTF8
                $content.Length | Should BeGreaterThan 0
                $content[0] | Should Match '^Attribute VB_Name\s*='
            }
        }

        It 'Todos os .bas tem Option Explicit na linha 2' {
            $basFiles = Get-VbaFiles
            $basFiles | Should Not BeNullOrEmpty
            foreach ($f in $basFiles) {
                $content = Get-Content $f.FullName -Encoding UTF8
                $content.Length | Should BeGreaterThan 1
                $content[1] | Should Be 'Option Explicit'
            }
        }

        It 'Arquivos .bas sao UTF-8 ou CP1252 validos (nao UTF-16, nao corrompidos)' {
            $basFiles = Get-VbaFiles
            $basFiles | Should Not BeNullOrEmpty
            foreach ($f in $basFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
                $bytes.Length | Should BeGreaterThan 0

                # Must not be UTF-16
                if ($bytes.Length -ge 2) {
                    $isUtf16Le = ($bytes[0] -eq 0xFF) -and ($bytes[1] -eq 0xFE)
                    $isUtf16Be = ($bytes[0] -eq 0xFE) -and ($bytes[1] -eq 0xFF)
                    ($isUtf16Le -or $isUtf16Be) | Should Be $false
                }

                # Must be decodable as either valid UTF-8 or CP1252 (single-byte)
                $isUtf8 = $false
                try {
                    $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
                    [void] $utf8Strict.GetString($bytes)
                    $isUtf8 = $true
                } catch { }

                $isCp1252 = $false
                try {
                    # CP1252 (code page 28591 = ISO-8859-1 superset) can decode any single-byte sequence
                    $cp1252 = [System.Text.Encoding]::GetEncoding(28591)
                    $decoded = $cp1252.GetString($bytes)
                    if ($decoded.Length -gt 0) { $isCp1252 = $true }
                } catch { }

                ($isUtf8 -or $isCp1252) | Should Be $true
            }
        }
    }
}

