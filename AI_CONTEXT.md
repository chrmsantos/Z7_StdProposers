# Z7_STDPROPOSERS - AI Assistant Developer Context

> Note to AI Agents: read this document before modifying the VBA pipeline or Python integration.
>
> Last updated: 2026-08-25 (v8.10.1 — robust format preservation in text revision, side-by-side prompt editor, CLSID ROT fix).

## 1. Project Overview

Z7_StdProposers is a Microsoft Word automation project for Brazilian legislative document standardization.

The solution has two coordinated parts:

- VBA formatting engine in `source/main/`.
- Python Gemini integration in `ai/`.

## 2. Current Codebase Architecture

### 2.1 VBA (12 modules)

The active VBA architecture is consolidated into 12 modules:

- `Mod_01_Infrastructure.bas`: constants, global state, paths (`GetZ7StdProposersDataPath`, `GetZ7StdProposersLogsPath`, etc.), safe wrappers, backup/system helpers.
- `Mod_02_Engine.bas`: structure detection (AI-first with heuristic fallback), cache system, image/list handling and restoration helpers.
- `Mod_03_Pipeline.bas`: core formatting pipeline (double-pass), normalization/cleanup routines (including blank paragraph numbering removal), logging primitives.
- `Mod_04_Main.bas`: macro entrypoints (`PadronizarDocumentoMain`, public API helpers, Gemini integration bridge calling the blank paragraph cleanup).
- `Mod_11_RevisionText.bas`: AI text revision via OpenRouter API. Public entrypoints: `TestarRevisaoTextoSelecionado` (selected text), `CorrigirProposituraComIA` (selected text with detailed metrics/status), `DiagnosticarOpenRouter` (connectivity test). Uses DPAPI-encrypted API key and configurable model via `config_prompt.py`. Revision prompt is now configurable via `revision_prompt.txt` (loaded by `CarregarPromptRevisao()`, falls back to hardcoded `MontarPromptRevisao()`). Integrates with project logging (`LogMessage`) and progress system (`InitializeProgress`/`IncrementProgress`). As of v8.10.1, `SubstituirTextoPreservandoFormatacao` now saves/restores full paragraph formatting including Borders, Shading, KeepWithNext, and protects paragraph marks (¶) from being destroyed during text replacement.
- `Mod_12_AIStructure.bas`: AI-based document structure identification via OpenRouter API. Public entrypoint: `IdentifyDocumentStructureWithAI`. Sends document text with paragraph markers to AI, parses JSON response with paragraph ranges for each structural element (Titulo, Ementa, Vocativo, Corpo, Justificativa, Data, Assinatura, Anexo). Used as primary identification method by `IdentifyDocumentStructure` in `Mod_02_Engine.bas`.

### 2.2 Python (Gemini integration)

Main files in `ai/`:

- `chat_ia.py`: chat UI with Word document context. `_context_pending` flag: if the initial Gemini call fails (e.g. 503), the full document text is prepended to the user's first message instead. Heavy imports deferred via lazy loading (opens UI instantly, ~2.5 s savings). Now includes grammar correction and consistency verification directly in the UI, including dynamic validation of question consistency and coherence in relation to the document context. As of v8.10.1, `_find_word_with_documents` also detects Word instances registered in the ROT by class moniker (CLSID `{000209FF-…}`) in addition to the legacy `!Word.Application.N` item moniker, fixing cases where modern Word registered via class moniker was missed.
- `config_prompt.py`: UI for prompt editing. Hosts `DEFAULT_PROMPT` (grammar and general consistency prompt, including question checking rules), `DEFAULT_CONSISTENCY_PROMPT` (controls consistency check output classification rules), `DEFAULT_CHAT_SYSTEM_PROMPT` (Chat IA system prompt), and `DEFAULT_REVISION_PROMPT` (text revision prompt used by `CorrigirProposituraComIA`). Both chat and revision prompts are editable in the same window. As of v8.10.1, UI layout changed to side-by-side columns: Corretor de Propositura (left) and Chat IA (right), with renaming from "Chat LÉIA" to "Chat IA" and "Revisão de Textos" to "Corretor de Propositura".
- `z7_logging.py`: shared structured logger; uses `RotatingFileHandler` (2 MB / 3 backups).
- `build_exe.ps1`: PyInstaller build workflow for `.exe` artifacts.

## 3. Operational Conventions

### A. Double-pass formatting pipeline

The document formatter uses a two-pass strategy:

1. Pass 1 normalizes and clears structural debris.
2. Pass 2 applies alignment-sensitive formatting.

Pass 2 must run only when Pass 1 changed content (`documentDirty = True`).

### B. Structural heuristics

Key anchors are inferred from text/format signatures:

- Titulo and Ementa.
- Justificativa header/body.
- Plenario/Data anchors.
- Assinatura block (includes authorship by "Presidente" and "Prefeito" — fixed in v6.2.1).
- Anexo header/body.

### C. Index invalidation rule (critical)

Any routine that deletes paragraphs must refresh structure indices before consuming global index pointers.

Required pattern:

```vba
If removedCount > 0 Then IdentifyDocumentStructure doc
```

Do not move this call above the deletion loop.

### D. UndoRecord safety rule

Main flow must keep `StartCustomRecord` and `EndCustomRecord` paired across all exit paths.

Current design in `Mod_04_Main.bas` routes failures through `CriticalErrorHandler -> GoTo CleanUp`, and `CleanUp` closes UndoRecord under guarded error handling.

### E. COM discipline rule

Avoid per-character COM object churn (`Range.Characters(n)`) in hot loops. Prefer operating directly on `Range`.

Example:

```vba
pRange.MoveEnd wdCharacter, -1
pRange.Delete
```

### F. Word StatusBar encoding rule

`word.StatusBar` is set via COM (ANSI/CP1252 on Windows). Do **not** use accented or non-ASCII characters in StatusBar strings — they render as garbage in the Word status bar. Use unaccented ASCII equivalents (e.g. `"verificacao"` not `"verificação"`).

### G. VBA module import encoding rule (critical — recurring regression)

When importing `.bas` files into Word's VBProject (via `import_bas_to_normal.py` or any similar script), two rules **must** be followed:

1. **Never use `CodeModule.AddFromString()` for `.bas` content that contains `Attribute` lines.** `AddFromString()` does NOT process VBA `Attribute` directives — it treats `Attribute VB_Name = "..."` as regular code, which (a) leaves the module unnamed and (b) causes a compile error. Always use `VBComponents.Import()` instead, which properly processes `Attribute VB_Name` and `Attribute VB_Base`. For existing modules, remove them first with `VBComponents.Remove()`, then re-import.

2. **Source `.bas` files are UTF-8.** Always decode as UTF-8 first, fall back to CP1252 only if UTF-8 decoding fails. `VBComponents.Import()` reads files using the system ANSI codepage (CP1252), so write a temporary `.bas` file in CP1252 encoding before calling `Import()` — this preserves accented characters through the round-trip.

**Anti-regression tests** enforcing these rules live in `tests/Encoding.Tests.ps1` (Context: `Import de Modulos VBA`). If any test fails, do NOT weaken the test — fix the import script instead.

### H. Word.officeUI path rule (critical)

Word reads QAT (Quick Access Toolbar) customizations from `%APPDATA%\Microsoft\Office\Word.officeUI` (Roaming) in Office 2016+/365. Older versions use `%LOCALAPPDATA%\Microsoft\Office\Word.officeUI` (Local).

### I. Tkinter in tests rule

Never use real `tkinter.Tk()` in unit tests — multiple `Tk()` instances in the same process cause `_tkinter.TclError`. Always add `"tkinter": mock.MagicMock()` to the `sys.modules` stubs dict before calling any code that imports tkinter.

## 4. Logging and Observability

### 4.1 VBA logging

VBA logging core is in `Mod_03_Pipeline.bas` (`InitializeLogging`, `LogMessage`, `SafeFinalizeLogging`).

Current behavior includes:

- Session/operation metadata (`currentLogSessionId`, `currentOperationId`).
- Structured elapsed-time line format with operation marker (`[op=...]`).
- Buffered flushing controlled by `LOG_BUFFER_FLUSH_SECONDS`.
- Context snapshots via `LogContextSnapshot doc, "..."`.

Current snapshots are executed in `Mod_04_Main.bas` at:

- Pipeline start (`INICIO`).
- Successful completion (`FIM`).
- Critical error path (`ERRO_CRITICO`).

### 4.2 Python logging

Python scripts share `z7_logging.py` and write UTF-8 logs to:

- `%LOCALAPPDATA%\Z7\Apps\StdProposers\LocalConfigs\logs`

Key API: `configure_component_logger(component, level)`, `log_exception(logger, context, exc)`, `build_log_path(component)`, `get_logs_dir()`.

Log files are managed by `RotatingFileHandler` (max 2 MB, 3 backups). Format: `YYYY-MM-DD HH:MM:SS.mmm | LEVEL | z7.<component> | mensagem`.

Log coverage includes:

- Startup and initialization milestones.
- Word integration attempts.
- OpenRouter request lifecycle.
- Exception stack traces.

## 5. Testing Model

### 5.1 Test runner

`tests/Run-Tests.ps1` supports:

- `All`
- `VBA`
- `Encoding`
- `Python`
- `VBA-Logging`

### 5.2 Current test suite layout

- `tests/All.Tests.ps1`: integrity checks aligned to current modular architecture.
- `tests/VBA.Tests.ps1`: VBA architecture and API contracts for Mod1..Mod4.
- `tests/VBA-IdentifierFunctions.Tests.ps1`: identifier-range safety and declarations in `Mod_04_Main.bas`.
- `tests/VBA-Logging.Tests.ps1`: observability assertions (session/op IDs, snapshots, logging primitives).
- `tests/Python.Tests.ps1`: Python integration checks and unittest invocation.
- `tests/python/test_z7_logging.py`: unit tests for `z7_logging.py`.
- `tests/python/test_z7_api_key.py`: unit tests for `z7_api_key.py`.
- `tests/python/test_chat_ia.py`: unit tests covering AI init, document reading, streaming, fallback, and edge cases.
- `tests/Encoding.Tests.ps1`: encoding/line-ending policy checks (UTF-8 safe, CRLF warnings, no UTF-16).

### 5.3 Current baseline

As of v8.7.0 (2026-08-14), `Run-Tests.ps1 -TestSuite All -NoProgress` passes. Python unit tests: 116 total via `pytest`.

## 6. Immediate Development Guidelines

1. Keep logic in the current 4-module VBA topology; do not reintroduce monoliths.
2. Preserve `Option Explicit` in every VBA module.
3. For formatting core changes, focus `Mod_02_Engine.bas`, `Mod_03_Pipeline.bas`, and `Mod_04_Main.bas` according to responsibility boundaries.
4. After any paragraph-deletion change, enforce a structure refresh before index usage (Section 3.C).
5. Preserve UndoRecord closure invariants (Section 3.D).
6. If editing Python integrations, keep logging instrumentation through `z7_logging.py`.
7. Validate with test suites before finalizing changes.

## 7. Build and Runtime Notes (Python)

- Dependency install: `install_requirements.bat`.
- Build executables: `ai/build_exe.ps1`.
  - `Package-Artifact` uses `[System.IO.Compression.ZipFile]::CreateFromDirectory()` (not `Compress-Archive`) — required because PowerShell 5.1's `Compress-Archive` raises `UnauthorizedAccessError` on nested `.zip` files such as `_internal/base_library.zip`.
  - Build artifacts are placed in `dist/` as `<name>-v<VERSION>.zip`; old PyInstaller output (`ai/dist/`) is cleaned up after packaging.
- Word macro launcher (`WordMacro.bas`) uses `pyw -3` and path fallback logic for script location.
- Footer page numbering format: `Página X de Y` (changed from `X-Y` in v6.2.1). Font size and color are applied to the full footer paragraph.

When changing runtime paths, update both VBA constants/macros and relevant docs together.
