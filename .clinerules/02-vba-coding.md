---
paths:
  - "source/main/*.bas"
  - "source/**/*.bas"
---

# Regras de Codificação VBA

## Estrutura dos Módulos

O projeto usa 12 módulos VBA com responsabilidades BEM definidas:

| Módulo | Responsabilidade |
|--------|-----------------|
| `Mod_01_Infrastructure.bas` | Constantes, paths, safe wrappers, backup |
| `Mod_02_Engine.bas` | Detecção de estrutura (AI-first + fallback heurístico), cache, imagens/listas |
| `Mod_03_Pipeline.bas` | Pipeline de formatação (double-pass), normalização, logging |
| `Mod_04_Main.bas` | Entrypoints públicos, `PadronizarDocumentoMain`, bridge Gemini |
| `Mod_05_Logging.bas` | Infraestrutura auxiliar de logging |
| `Mod_06_WordMacro.bas` | Lançador de macros Python via `pyw -3` |
| `Mod_07_Formatting.bas` | Rotinas de formatação de parágrafos |
| `Mod_08_Ementa.bas` | Formatação específica de ementas |
| `Mod_09_SpecialParagraphs.bas` | Parágrafos especiais (justificativa, assinatura, etc.) |
| `Mod_10_Validation.bas` | Validações pós-formatação |
| `Mod_11_RevisionText.bas` | Revisão de texto via OpenRouter API |
| `Mod_12_AIStructure.bas` | Identificação de estrutura via AI (OpenRouter) |

**NÃO adicione novos módulos.** Modifique apenas os existentes, respeitando responsabilidades.

## Padrões de Código

### Remoção de Parágrafos com Segurança
```vba
Dim removedCount As Long
removedCount = 0
' ... loop de deleção, incrementando removedCount ...
If removedCount > 0 Then
    IdentifyDocumentStructure doc  ' refresh obrigatório!
End If
' Agora pode usar índices da estrutura
```

### Manipulação de Range (evitar Characters())
```vba
' ERRADO — lento, alto churn COM:
For i = 1 To rng.Characters.Count
    ' ...
Next

' CORRETO:
pRange.MoveEnd wdCharacter, -1
pRange.Delete
```

### UndoRecord com CleanUp
```vba
On Error GoTo CriticalErrorHandler
doc.UndoRecord.StartCustomRecord "Operação"

' ... corpo principal ...

CleanUp:
    On Error Resume Next
    doc.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Exit Sub

CriticalErrorHandler:
    LogMessage "ERRO_CRITICO", Err.Description
    Resume CleanUp
```

### StatusBar (apenas ASCII)
```vba
word.StatusBar = "Processando..."        ' OK
word.StatusBar = "Processando..."        ' ERRADO — acento!
```

## Encoding

- Arquivos `.bas` são **CP1252** (Latin1). NUNCA salve em UTF-8.
- Após editar, execute: `python scripts/fix_bas_encoding.py`
- Linha 1: `Attribute VB_Name = "Mod_XX_..."` (obrigatório para `VBComponents.Import`)
- Linha 2: `Option Explicit` (obrigatório)

## Importação de Módulos (Contexto Python)

Ao modificar `scripts/import_bas_to_normal.py`:
1. NUNCA use `AddFromString()` — use `VBComponents.Import()`
2. Decodifique `.bas` como UTF-8 primeiro, CP1252 como fallback
3. Escreva arquivo temporário CP1252 antes de `Import()`
4. Remova módulo existente com `VBComponents.Remove()` antes de re-importar