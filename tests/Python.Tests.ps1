#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

# ---------------------------------------------------------------------------
# Módulos Python esperados
# ---------------------------------------------------------------------------
$ExpectedModules = @(
    'z7_logging.py',
    'z7_gemini_key.py',
    'z7_theme.py',
    'config_prompt.py',
    'chat_ia.py'
)

# Arquivos de teste Python esperados
$ExpectedTestFiles = @(
    'test_z7_logging.py',
    'test_z7_gemini_key.py',
    'test_z7_theme.py',
    'test_config_prompt.py',
    'test_chat_ia.py'
)

Describe 'Z7_STDPROPOSERS - Python Logging and Test Harness' {

    # -----------------------------------------------------------------------
    # Existência dos módulos
    # -----------------------------------------------------------------------
    Context 'Modulos Python existem' {
        foreach ($mod in $ExpectedModules) {
            It "Modulo $mod existe em ai/" {
                $repoRoot = Get-RepoRoot
                Test-Path "$repoRoot\ai\$mod" | Should Be $true
            }
        }
    }

    # -----------------------------------------------------------------------
    # Convenções de código
    # -----------------------------------------------------------------------
    Context 'Convencoes de codigo' {
        It 'Scripts Python principais usam o logger compartilhado' {
            $repoRoot = Get-RepoRoot
            foreach ($script in @('config_prompt.py', 'chat_ia.py')) {
                $content = Get-Content "$repoRoot\ai\$script" -Raw -Encoding UTF8
                $content | Should Match 'from z7_logging import configure_component_logger'
                $content | Should Match 'LOGGER = configure_component_logger'
            }
        }

        It 'Scripts que usam API Gemini importam de z7_gemini_key' {
            $repoRoot = Get-RepoRoot
            foreach ($script in @('chat_ia.py')) {
                $content = Get-Content "$repoRoot\ai\$script" -Raw -Encoding UTF8
                $content | Should Match 'from z7_gemini_key import'
            }
        }

        It 'Scripts que acessam Word usam GetActiveObject como conexao primaria' {
            $repoRoot = Get-RepoRoot
            foreach ($script in @('config_prompt.py', 'chat_ia.py')) {
                $content = Get-Content "$repoRoot\ai\$script" -Raw -Encoding UTF8
                $content | Should Match 'GetActiveObject'
            }
        }
    }

    # -----------------------------------------------------------------------
    # Existência dos arquivos de teste
    # -----------------------------------------------------------------------
    Context 'Arquivos de teste Python existem' {
        foreach ($testFile in $ExpectedTestFiles) {
            It "Arquivo de teste $testFile existe em tests/python/" {
                $repoRoot = Get-RepoRoot
                Test-Path "$repoRoot\tests\python\$testFile" | Should Be $true
            }
        }
    }

    # -----------------------------------------------------------------------
    # Execução dos testes unitários (auto-descoberta)
    # -----------------------------------------------------------------------
    Context 'Suites de testes unitarios Python executam com sucesso' {
        It 'Todos os testes Python passam via unittest discover' {
            $repoRoot = Get-RepoRoot
            $cmd = "Set-Location '$repoRoot'; py -3 -m unittest discover -s tests/python -p 'test_*.py' -v"
            $output = powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $cmd 2>&1 | Out-String
            $LASTEXITCODE | Should Be 0
            $output | Should Match 'OK'
        }

        # Testes individuais por módulo para relatório granular
        foreach ($testFile in $ExpectedTestFiles) {
            It "Testes de $testFile passam" {
                $repoRoot = Get-RepoRoot
                $module = $testFile -replace '\.py$', '' -replace '/', '.'
                $cmd = "Set-Location '$repoRoot'; py -3 -m unittest tests.python.$module -v"
                $output = powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $cmd 2>&1 | Out-String
                $LASTEXITCODE | Should Be 0
                $output | Should Match 'OK'
            }
        }
    }
}

