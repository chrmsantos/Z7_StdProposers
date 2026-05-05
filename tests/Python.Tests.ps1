#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'Z7_STDPROPOSERS - Python Logging and Test Harness' {
    It 'Tem modulo de logging compartilhado no Python' {
        $repoRoot = Get-RepoRoot
        Test-Path (Join-Path $repoRoot 'ai' 'z7_logging.py') | Should Be $true
    }

    It 'Scripts Python principais usam o logger compartilhado' {
        $repoRoot = Get-RepoRoot

        foreach ($script in @('correct_grammar.py', 'config_prompt.py', 'chat_ia.py')) {
            $content = Get-Content (Join-Path $repoRoot 'ai' $script) -Raw -Encoding UTF8
            $content | Should Match 'from z7_logging import configure_component_logger'
            $content | Should Match 'LOGGER = configure_component_logger'
        }
    }

    It 'Suite de testes unitarios Python de logging existe' {
        $repoRoot = Get-RepoRoot
        Test-Path (Join-Path $repoRoot 'tests\python\test_z7_logging.py') | Should Be $true
    }

    It 'Teste unitario Python do logger executa com sucesso' {
        $repoRoot = Get-RepoRoot
        $cmd = "Set-Location '$repoRoot'; py -3 -m unittest tests.python.test_z7_logging -v"
        $output = powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $cmd 2>&1 | Out-String
        $LASTEXITCODE | Should Be 0
        $output | Should Match 'OK'
    }
}

