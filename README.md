# Z7_StdProposers

## Sistema de Padronização de Proposituras Legislativas

Z7_StdProposers is an advanced, robust VBA macro project designed exclusively for Microsoft Word. It automatically sanitizes, structures, and formats complex Brazilian legislative documents (Proposituras), turning raw, unformatted text into perfectly aligned, standardized legal documents.

## ✨ Features

- **Automated Element Identification:** Uses an advanced heuristic engine to detect structural parts of a document such as *Título*, *Ementa*, *Justificativa*, *Data (Plenário)*, and *Assinaturas*.
- **Two-Pass Formatting Pipeline:**
  - **Pass 1:** Normalizes syntax, limits blank lines, clears arbitrary text boundaries, and purges bad encodings.
  - **Pass 2:** Applies pixel-perfect layout alignments, margins, and custom indentation rules seamlessly.
- **Visual Content Protection:** Explicitly caches and protects images, bullet points, numbered lists, and header layouts from getting destroyed during the document formatting process.
- **Intelligent Clause Parsing:** Finds and formats specific legal clauses heavily used in Brazilian legislation like *“Considerando”*, *“Ante o Exposto”*, and *“In Loco”*.
- **Failsafe Executions:** Includes heavily engineered error recovery, undo-group wraps (`UndoRecord`), and pre-execution backups to ensure MS Word never critically fails during processing.

## 🏗️ Architecture

The VBA codebase is currently consolidated into 4 main modules in `source/main/`:

- `Mod1Infrastructure.bas`: Constants, global state, cross-cutting helpers, paths, backup/system integrations.
- `Mod2Engine.bas`: Structural detection heuristics, paragraph cache, image/list preservation routines.
- `Mod3Pipeline.bas`: Core formatting pipeline (double-pass), normalization, cleanup, and logging primitives.
- `Mod4Main.bas`: Public entrypoints/macros, orchestration, integration with engine/pipeline and Gemini helpers.

The repository also contains a Python integration package in `ai/` for Gemini-based grammar correction and chat utilities.

## 🚀 Installation & Usage

1. Open Microsoft Word.
2. Launch the **Visual Basic for Applications (VBA) Editor** (`ALT` + `F11`).
3. Import the `.bas` files found in the `source/main/` folder into your `Normal.dotm` or dedicated Document Template.
4. Go to `Debug -> Compile Project` to ensure your Word environment resolves the inter-module Public references.
5. Create a Ribbon Button or Quick Access Toolbar shortcut pointing to the `PadronizarDocumentoMain` macro.
6. Click the macro while editing a document to execute the Z7_StdProposers standardized pipeline!

## ✅ Automated Tests

Test suites are under `tests/` and can be executed via:

- `tests\Run-Tests.ps1 -TestSuite All -NoProgress`
- `tests\Run-Tests.ps1 -TestSuite VBA -NoProgress`
- `tests\Run-Tests.ps1 -TestSuite Encoding -NoProgress`
- `tests\Run-Tests.ps1 -TestSuite Python -NoProgress`
- `tests\Run-Tests.ps1 -TestSuite VBA-Logging -NoProgress`

`tests\run-tests.cmd` is available as a convenience wrapper for Windows environments.

## 📜 License

This project is licensed under the **GNU GPLv3** License. See the `LICENSE` file for more details.

