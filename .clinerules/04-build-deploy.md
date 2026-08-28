---
paths:
  - "ai/build_exe.ps1"
  - "ai/*.spec"
  - "dist/**"
  - "VERSION"
  - "ai/build/**"
---

# Regras de Build e Deploy

## Fluxo Completo de Build + Deploy (GitHub Release)

### Pré-requisitos
- PyInstaller instalado no sistema (fora de venv): `pip install pyinstaller`
- PowerShell 5.1+
- Git configurado com acesso ao repositório remoto

### Passo a Passo — Build

1. **Verificar versão**: Leia `VERSION` na raiz. Confirme que `Z7_STDPROPOSERS_VERSION` em `source/main/Mod_01_Infrastructure.bas` coincide.

2. **Verificar encoding dos .bas**: Execute `python scripts/fix_bas_encoding.py --check`. Se houver arquivos com "NEEDS FIX", execute sem `--check` primeiro.

3. **Executar testes**: `powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1 -TestSuite All -NoProgress`. SÓ PROSSIGA se todos passarem.

4. **Build dos executáveis**: `powershell -ExecutionPolicy Bypass -File ai/build_exe.ps1`
   - Compila `chat_ia.py`, `config_prompt.py`, `import_bas_to_normal.py`
   - Artefatos em `dist/`: `chat_ia-v<VERSION>.zip`, `config_prompt-v<VERSION>.zip`, `import_bas_to_normal.exe`
   - **NÃO delete o diretório `ai/build/`** — o cache `Analysis-00.toc` evita bug do Python 3.14

5. **Verificar artefatos**: Liste `dist/` e confirme que os 3 arquivos existem com a versão correta:
   ```
   dist/chat_ia-v8.10.1.zip
   dist/config_prompt-v8.10.1.zip
   dist/import_bas_to_normal.exe
   ```

### Passo a Passo — Deploy (GitHub Release)

6. **Commit e push**: Commit todas as alterações com mensagem descritiva no padrão: `v<VERSION>: <resumo>`. Push para o branch atual.

7. **Criar tag**: `git tag -a v<VERSION> -m "Release v<VERSION>"` e `git push origin v<VERSION>`

8. **Criar release no GitHub**:
   Use `gh release create` se o GitHub CLI estiver disponível:
   ```powershell
   gh release create v<VERSION> `
     dist/chat_ia-v<VERSION>.zip `
     dist/config_prompt-v<VERSION>.zip `
     dist/import_bas_to_normal.exe `
     --title "v<VERSION>" `
     --notes "Release notes aqui"
   ```
   Se `gh` não estiver disponível, crie a release manualmente via interface web.

### Troubleshooting Comum

| Problema | Causa Provável | Solução |
|----------|---------------|---------|
| `PyInstaller nao encontrado` | PyInstaller fora do PATH | Localizar em `$env:LOCALAPPDATA\Programs\Python\*\Scripts\pyinstaller.exe` |
| `UnauthorizedAccessError` no zip | `Compress-Archive` com `.zip` aninhados | O script já usa `ZipFile::CreateFromDirectory()` — não altere |
| Falha na análise PyInstaller | Bug Python 3.14 remove build dir | O script já pré-cria `build/<name>/` — não remova |
| Testes falham após mudança | Encoding ou regressão | Execute `fix_bas_encoding.py`, revise o diff |
| `AddFromString` detectado | Alguém usou `AddFromString` | Substitua por `VBComponents.Import` com arquivo temp CP1252 |
| `Attribute VB_Name` ausente | Salvou .bas como UTF-8 sem BOM? | Execute `fix_bas_encoding.py` |

### Regras de Versionamento

- **SEMPRE** incremente a versão em `VERSION` ANTES do build.
- **SEMPRE** atualize `Z7_STDPROPOSERS_VERSION` em `Mod_01_Infrastructure.bas` para coincidir.
- O build NUNCA deve ser feito sem que os testes `All` passem.
- Não faça release com `--no-backup` ou `--dry-run` ativos.