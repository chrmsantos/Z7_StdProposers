# Z7_STDPROPOSERS - AI Assistant Developer Context

> **Note to AI Agents**: Read this document before attempting to modify or refactor the Z7_StdProposers VBA codebase. This documentation will save you from re-analyzing the structural paradigms and pipeline mechanics of the project.
>
> **Last updated**: 2026-04-20 — Applied three architectural fixes (stale index pointers, UndoRecord audit, COM overhead). See Section 5.

## 1. Project Overview
**Z7_StdProposers** (Sistema de Padronização de Proposituras Legislativas) is an advanced document-formatting macro built for Microsoft Word. It is designed to automatically detect structural components of Brazilian legislative documents (Proposituras) and apply a highly tailored, fault-tolerant standard layout.

## 2. Codebase Architecture (VBA)
Historically a massive ~12,000-line monolithic `Proposers.bas` file, the codebase was initially refactored into 7 modules, and subsequently consolidated into **4 pilares lógicos** inside `source/main/` to balance maintenance and cohesion:

- **`Mod1Infrastructure.bas`**: Combina Config, Utils e System. Holds all initialization logic, default path declarations, UDTs, UI View setup states, system-level integrations (telemetry, progress bars, emergency recovery) and generic cross-cutting helpers (Safe Property wrappers).
- **`Mod2Engine.bas`**: Combina Core e Media. Houses the engine mechanics, heuristic parsers designed to detect text blocks (`Ementa`, `Plenário`, etc), the paragraph caching pipeline, and complex layout adjustments associated with bullet lists and image protections before/after formatting.
- **`Mod3Pipeline.bas`**: Mantém o antigo Process. The massive core formatting pipeline. Governs text normalization, the double-pass formatting workflow, special clauses, and blank line synchronization.
- **`Mod4Main.bas`**: Ponto de entrada das macros (ex: `PadronizarDocumentoMain`, `CorrigirGramaticaComGemini`). Liga as funções da Infraestrutura, as análises da Engine e a execução do Pipeline.

## 3. Structural Conventions and Nuances

### A. The Double-Pass Pipeline Concept
Z7_StdProposers formats documents using an optimized two-pass approach to ensure layout stability:
1. **Pass 1**: Normalizes characters and removes raw text debris. It establishes the baseline flow. If Pass 1 successfully flags `documentDirty = True`, it will trigger Pass 2.
2. **Pass 2**: Injects complex alignments (like the Ementa indents or Header images).

### B. Element Recognition (Heuristics)
Because legislative documents are rarely well-formed natively, the system parses strings looking for signatures:
- **Ementa**: Sits usually after the Title and ends right before the main proposition. Sometimes requires indent measurements (`Ementa_Min_Left_Indent`) to be identified robustly.
- **Plenario**: Detected by searching for `"plenario"`. Used as an anchor to align Data and Assinaturas that follow.
- **Justificativa**: Detected by the header `"justificativa"`. Requires aggressive insertion/synchronization of blank lines.

### C. Resource Protection & Safe Handlers
Never trust Microsoft Word's COM architecture to execute cleanly under pressure.
- **Image Protection System**: Found in `ModMedia`. Due to formatting resets blowing out images, images are essentially copied/cached and re-injected post-format.
- **UndoGroups (CustomRecord)**: Z7_StdProposers wraps executions in `Application.UndoRecord.StartCustomRecord`. All exit paths — including `CriticalErrorHandler → EmergencyRecovery → GoTo CleanUp` and every explicit `GoTo CleanUp` — converge at the single `CleanUp:` label in `ModProcess.bas`, which calls `EndCustomRecord` guarded by `On Error Resume Next`. This guarantees the native Undo stack is never permanently broken even during fatal crashes.
- **Error Recovery (`EmergencyRecovery`)**: If an unhandled exception surfaces, this subroutine triggers a safe shutdown (reverting View states, enabling screen updating, dropping COM objects) to ensure MS Word doesn't freeze or lock up.

### D. Scoping Strategy
Variables, Constants, Types, Subs, and Functions across the 7 modules use `Public` scope to interact with one another. If creating local memory for a temporary macro, strictly define variables with `Dim` locally inside the function.

### E. Structural Index Invalidation Rule
> **Critical pattern to preserve.** Any subroutine that **physically deletes paragraphs** from the document must call `IdentifyDocumentStructure doc` *after* the deletion loop and *before* any subsequent logic that reads global index variables (`tituloParaIndex`, `ementaParaIndex`, `tituloJustificativaIndex`, `dataParaIndex`, etc.).
>
> Background: global index variables store the ordinal position (1-based `Long`) of structural elements. Deleting a paragraph above any of these anchors silently shifts all indices below the deletion point, causing formatters to target wrong paragraphs (stale pointer / off-by-N bug).
>
> The established pattern (as implemented in `RemoverLinhasEmBrancoExtras`):
> ```vba
> If removedCount > 0 Then IdentifyDocumentStructure doc
> ```
> Do **not** remove or hoist this call above the deletion loop.

### F. COM Interface Discipline
Avoid interacting with individual characters via the `Characters` collection inside hot loops. Creating per-character COM proxy objects (`Range.Characters(n)`) is expensive and degrades stability under long documents. Instead:
- To delete a single trailing character from a `Range`, shrink the range with `pRange.MoveEnd wdCharacter, -1` then call `pRange.Delete`.
- For bulk font changes that must skip inline images, use `Range.Font` on the whole range with `On Error Resume Next` rather than iterating character by character. Reserve `FormatCharacterByCharacter` (in `ModProcess.bas`) only for paragraphs that are explicitly confirmed to contain inline images.

## 4. Immediate Development Guidelines
When updating this codebase:
1. Do not merge functional logic back into one monolithic script. Keep logic contained to its respective `Mod[Type].bas` file.
2. **Always include `Option Explicit`** at the top of any new `.bas` file.
3. If adjusting styles or font margins, strictly go to `ModConfig.bas` and `ModProcess.bas`.
4. Run thorough checks via `Debug -> Compile` mapped inside Word when adjusting cross-module parameters to avoid scope disconnection crashes.
5. After any operation that deletes paragraphs, re-run `IdentifyDocumentStructure doc` before using any global index variable (see §3.E).

## 5. Architectural Fixes Applied (2026-04-20)

Three bugs identified in `analise_projeto_vba.md` were corrected in `ModProcess.bas`:

### Fix 1 — Stale Index Pointers (`RemoverLinhasEmBrancoExtras`)
`RemoverLinhasEmBrancoExtras` deleted blank paragraphs in a reverse loop but then used `tituloJustificativaIndex` in the subsequent `For Each para` loop without refreshing the indices. This caused the Vereador/cargo centering logic to fire on wrong paragraphs.

**Resolution**: Added `If removedCount > 0 Then IdentifyDocumentStructure doc` immediately after the deletion loop and the text-replacement block, and immediately before the `adjustCounter` loop. The `IdentifyDocumentStructure` call re-scans the document and rewrites all global index variables to their correct post-deletion positions.

### Fix 2 — UndoRecord Leak Audit
The `analise_projeto_vba.md` diagnosed a potential `EndCustomRecord` leak when a fatal error bypassed the `CleanUp:` label. Audit confirmed all exit paths — including `CriticalErrorHandler` inside `PadronizarDocumentoMain` — unconditionally route through `GoTo CleanUp`, where `EndCustomRecord` is called under `On Error Resume Next`. No code change was required; the fix was already in place.

### Fix 3 — COM Overhead in Space-Deletion Loop
The line `pRange.Characters(1).Delete` allocated a COM proxy object on every loop iteration to delete a single space character. Replaced with:
```vba
pRange.MoveEnd wdCharacter, -1
pRange.Delete
```
This operates directly on the `Range` object already in scope, eliminating the intermediate `Characters` collection proxy and reducing COM round-trips.
