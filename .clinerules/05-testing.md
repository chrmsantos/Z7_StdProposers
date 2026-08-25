---
paths:
  - "tests/**"
  - "tests/*.ps1"
  - "tests/python/*.py"
  - "**/*.Tests.ps1"
  - "**/test_*.py"
---

# Regras de Testes

## Estrutura de Testes

| Suite | Runner | Arquivos |
|-------|--------|---------|
| **All** | `Run-Tests.ps1 -TestSuite All` | Todos os `.Tests.ps1` + pytest |
| **VBA** | `Run-Tests.ps1 -TestSuite VBA` | `VBA.Tests.ps1` |
| **Encoding** | `Run-Tests.ps1 -TestSuite Encoding` | `Encoding.Tests.ps1` |
| **Python** | `Run-Tests.ps1 -TestSuite Python` | `Python.Tests.ps1` + pytest |
| **VBA-Logging** | `Run-Tests.ps1 -TestSuite VBA-Logging` | `VBA-Logging.Tests.ps1` |
| **VBA-AIStructure** | `Run-Tests.ps1 -TestSuite VBA-AIStructure` | `VBA-AIStructure.Tests.ps1` |

## Comando Padrão
```powershell
powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1 -TestSuite All -NoProgress
```

## Regras Fundamentais

### 1. NUNCA enfraqueça um teste
Se um teste falha, o código está errado — não o teste. Testes de encoding, import e estrutura são barreiras de segurança contra regressões conhecidas.

### 2. Testes de Encoding (críticos — NÃO ALTERE)
- `Tests.ps1` de `import_bas_to_normal.py`:
  - NÃO use `AddFromString()` (deve ser 0 ocorrências em linhas não-comentadas)
  - Use `VBComponents.Import`
  - Decodifique UTF-8 primeiro, CP1252 depois
  - Use `VBComponents.Remove` antes de re-importar
  - Grava temp file CP1252
  - `.bas` começam com `Attribute VB_Name` (linha 1) e `Option Explicit` (linha 2)

### 3. Testes Python
- Local: `tests/python/`
- Runner: `pytest`
- Baseline: 116+ testes
- Regra do tkinter: SEMPRE faça mock de `tkinter` antes de importar código que o use

### 4. Progress Preference
Use `-NoProgress` para evitar travamentos no VS Code. O runner já configura `$ProgressPreference = 'SilentlyContinue'`.

### 5. Cobertura Mínima
- Todo novo módulo ou função pública DEVE ter teste correspondente.
- Mudanças em `Mod_02_Engine.bas`, `Mod_03_Pipeline.bas` ou `Mod_04_Main.bas` exigem validação com `All`.
- Mudanças em scripts de import exigem `Encoding`.

### 6. Antes de Commit
```powershell
# 1. Corrigir encoding
python scripts/fix_bas_encoding.py

# 2. Rodar todos os testes
powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1 -TestSuite All -NoProgress

# 3. Só faça commit se TUDO passar
```