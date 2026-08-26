---
paths:
  - "ai/*.py"
  - "ai/**/*.py"
  - "scripts/*.py"
  - "tests/python/*.py"
---

# Regras de Codificação Python

## Arquitetura dos Módulos Python

| Arquivo | Responsabilidade |
|---------|-----------------|
| `ai/chat_ia.py` | Chat UI com contexto do documento Word. Lazy loading de imports pesados (~2.5s savings). |
| `ai/config_prompt.py` | UI de edição de prompts. Layout side-by-side: Corretor de Propositura + Chat IA. |
| `ai/z7_logging.py` | Logger estruturado com `RotatingFileHandler` (2MB/3 backups). |
| `ai/z7_api_key.py` | Gerenciamento de API key com DPAPI. |
| `ai/z7_theme.py` | Temas da UI (tkinter). |
| `scripts/import_bas_to_normal.py` | Importa `.bas` no Normal.dotm do Word. |
| `scripts/import_ui_to_word.py` | Importa UI customizada (QAT) no Word. |
| `scripts/fix_bas_encoding.py` | Corrige encoding de `.bas` para CP1252. |

## Regras Obrigatórias

### 1. Encoding
- Todos os `.py` devem ser **UTF-8** com line endings **LF**.
- Strings acentuadas são permitidas em Python.

### 2. Logging
- Use `z7_logging.py` para todo logging de componentes Python.
- API: `configure_component_logger(component, level)`, `log_exception(logger, context, exc)`, `build_log_path(component)`.
- Logs vão para `%LOCALAPPDATA%\Z7\Apps\Z7_StdProposers\setup\logs`.

### 3. Imports
- `chat_ia.py` e `config_prompt.py` usam **lazy loading** para imports pesados (tkinter, PIL, openai).
- NÃO faça imports no topo do arquivo se o módulo só é usado em um branch condicional.

### 4. Contexto Word
- `chat_ia.py` usa `_context_pending`: se a chamada Gemini inicial falhar (503), o texto do documento é inserido na primeira mensagem.
- `_find_word_with_documents` detecta Word via ROT por class moniker (CLSID `{000209FF-…}`) E item moniker legado.

### 5. Testes — Regra do tkinter
**NUNCA use `tkinter.Tk()` real em testes unitários.** Múltiplas instâncias causam `_tkinter.TclError`.
```python
# Sempre faça:
sys.modules['tkinter'] = mock.MagicMock()
```
antes de importar código que usa tkinter.

### 6. PyInstaller Build
- Script: `ai/build_exe.ps1`
- Modo: `--onedir` (DLLs pré-extraídas, elimina 1-3s de extração)
- NÃO usar `--clean` (bug Python 3.14)
- Pré-criar `build/<name>/` para evitar bug do PyInstaller
- Empacotamento usa `[System.IO.Compression.ZipFile]::CreateFromDirectory()` (NÃO `Compress-Archive` — bug com `.zip` aninhados no PS 5.1)
- Hidden imports obrigatórios: `unicodedata`, `openai`, `jiter`, `certifi`, `pythoncom`, `win32com.client`, `win32com`

### 7. Word.officeUI
- Office 2016+/365: `%APPDATA%\Microsoft\Office\Word.officeUI` (Roaming)
- Versões antigas: `%LOCALAPPDATA%\Microsoft\Office\Word.officeUI` (Local)

### 8. Prompt Revision
- Prompt de revisão em `revision_prompt.txt` (carregado por `CarregarPromptRevisao()`), fallback hardcoded.
- Preservação de formatação: `SubstituirTextoPreservandoFormatacao` salva/restaura Borders, Shading, KeepWithNext e protege marcas de parágrafo (¶).