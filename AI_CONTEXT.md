# Z7_STDPROPOSERS - AI Assistant Developer Context

> Note to AI Agents: read this document before modifying the VBA pipeline or Python integration.
>
> Last updated: 2026-05-06 (v6.5.0-beta1 — lazy loading, footer format, RotatingFileHandler, signature heuristics).

## 1. Project Overview

Z7_StdProposers is a Microsoft Word automation project for Brazilian legislative document standardization.

The solution has two coordinated parts:

- VBA formatting engine in `source/main/`.
- Python Gemini integration in `ai/`.

## 2. Current Codebase Architecture

### 2.1 VBA (4 modules)

The active VBA architecture is consolidated into four modules:

- `Mod1Infrastructure.bas`: constants, global state, paths, safe wrappers, backup/system helpers.
- `Mod2Engine.bas`: structure detection heuristics, cache system, image/list handling and restoration helpers.
- `Mod3Pipeline.bas`: core formatting pipeline (double-pass), normalization/cleanup routines, logging primitives.
- `Mod4Main.bas`: macro entrypoints (`PadronizarDocumentoMain`, public API helpers, Gemini integration bridge).

### 2.2 Python (Gemini integration)

Main files in `ai/`:

- `correct_grammar.py`: corrects selected text in Word. Heavy imports (`google.genai`, `win32com`, `win32crypt`) are deferred via lazy loading.
- `config_prompt.py`: UI for prompt editing.
- `chat_ia.py`: chat UI with Word document context. Heavy imports are deferred via lazy loading (opens UI instantly, ~2.5 s savings).
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

Current design in `Mod4Main.bas` routes failures through `CriticalErrorHandler -> GoTo CleanUp`, and `CleanUp` closes UndoRecord under guarded error handling.

### E. COM discipline rule

Avoid per-character COM object churn (`Range.Characters(n)`) in hot loops. Prefer operating directly on `Range`.

Example:

```vba
pRange.MoveEnd wdCharacter, -1
pRange.Delete
```

## 4. Logging and Observability

### 4.1 VBA logging

VBA logging core is in `Mod3Pipeline.bas` (`InitializeLogging`, `LogMessage`, `SafeFinalizeLogging`).

Current behavior includes:

- Session/operation metadata (`currentLogSessionId`, `currentOperationId`).
- Structured elapsed-time line format with operation marker (`[op=...]`).
- Buffered flushing controlled by `LOG_BUFFER_FLUSH_SECONDS`.
- Context snapshots via `LogContextSnapshot doc, "..."`.

Current snapshots are executed in `Mod4Main.bas` at:

- Pipeline start (`INICIO`).
- Successful completion (`FIM`).
- Critical error path (`ERRO_CRITICO`).

### 4.2 Python logging

Python scripts share `z7_logging.py` and write UTF-8 logs to:

- `%LOCALAPPDATA%\Z7\Tmp\StdProposers\logs`

Key API: `configure_component_logger(component, level)`, `log_exception(logger, context, exc)`, `build_log_path(component)`, `get_logs_dir()`.

Log files are managed by `RotatingFileHandler` (max 2 MB, 3 backups). Format: `YYYY-MM-DD HH:MM:SS.mmm | LEVEL | z7.<component> | mensagem`.

Log coverage includes:

- Startup and initialization milestones.
- Word integration attempts.
- Gemini request lifecycle.
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
- `tests/VBA-IdentifierFunctions.Tests.ps1`: identifier-range safety and declarations in `Mod4Main.bas`.
- `tests/VBA-Logging.Tests.ps1`: observability assertions (session/op IDs, snapshots, logging primitives).
- `tests/Python.Tests.ps1`: Python integration checks and unittest invocation.
- `tests/python/test_z7_logging.py`: unit tests for `z7_logging.py`.
- `tests/Encoding.Tests.ps1`: encoding/line-ending policy checks (UTF-8 safe, CRLF warnings, no UTF-16).

### 5.3 Current baseline

As of v6.5.0-beta1 (2026-05-06), `Run-Tests.ps1 -TestSuite All -NoProgress` passes.

## 6. Immediate Development Guidelines

1. Keep logic in the current 4-module VBA topology; do not reintroduce monoliths.
2. Preserve `Option Explicit` in every VBA module.
3. For formatting core changes, focus `Mod2Engine.bas`, `Mod3Pipeline.bas`, and `Mod4Main.bas` according to responsibility boundaries.
4. After any paragraph-deletion change, enforce a structure refresh before index usage (Section 3.C).
5. Preserve UndoRecord closure invariants (Section 3.D).
6. If editing Python integrations, keep logging instrumentation through `z7_logging.py`.
7. Validate with test suites before finalizing changes.

## 7. Build and Runtime Notes (Python)

- Dependency install: `install_requirements.bat`.
- Build executables: `ai/build_exe.ps1`.
- Word macro launcher (`WordMacro.bas`) uses `pyw -3` and path fallback logic for script location.
- Footer page numbering format: `Pág. X de Y` (changed from `X-Y` in v6.2.1). Font size and color are applied to the full footer paragraph.

When changing runtime paths, update both VBA constants/macros and relevant docs together.

