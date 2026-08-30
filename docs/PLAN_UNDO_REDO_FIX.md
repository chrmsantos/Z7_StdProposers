# Plano de Ação — Blindagem da Pilha de Desfazer/Refazer
## `PadronizarDocumentoMain` — Z7 StdProposers v9.5.0

---

## 1. Diagnóstico Técnico

### 1.1 Causa raiz do "Desfazer Fantasma" (crash no 2º Ctrl+Z)

O crash ocorre quando **qualquer operação que force o Word a recalcular layout, tela ou modelo de objeto acontece DEPOIS de `EndCustomRecord`**. O Word processa essas operações como se fossem edições novas, gerando uma segunda entrada de undo **inválida** (sem `StartCustomRecord` correspondente).

Casos históricos já corrigidos neste projeto:

| Versão | Operação parasita | Efeito |
|--------|-------------------|--------|
| v9.0.0 | `Application.OnRepeat` | Corrompia pilha de undo |
| v9.2.0 | `doc.Save` / `doc.Range(0,0).Select` | Criavam entradas órfãs |
| v9.2.1 | `DoEvents`/`ScreenRefresh` fora de ordem | Layout recálculo gerava entrada fantasma |
| v9.2.1 | `doc.UndoClear` | Criava entradas fantasmas |

**Observação:** o v9.3.0 removeu a integração `StartCustomRecord`/`EndCustomRecord` por segurança. O v9.5.0 (este PR) reintegra com blindagem completa.

### 1.2 Por que "Repetir" (F4) não é implementado

`Application.OnRepeat` foi removido no v9.0.0 por corromper a pilha de undo. Reimplementar reintroduziria exatamente esse risco. **Decisão: abortar "Repetir", focar 100% no "Desfazer".**

### 1.3 Regras de ouro do UndoRecord seguro

1. `StartCustomRecord` = primeira operação geradora de edição (após backups não-destrutivos).
2. `ScreenRefresh` + `ScreenUpdating=True` = executam **DENTRO** do grupo de undo.
3. `EndCustomRecord` = **ÚLTIMA** operação de qualquer natureza. Nenhuma operação de tela/layout/`DoEvents` apos ela.
4. `doc.UndoClear` é **PROIBIDO** em qualquer ponto.
5. Todo caminho de saída (sucesso, `GoTo CleanUp`, erro crítico) passa por um único bloco `CleanUp`.
6. `emergencyRecovery` pode fechar o registro antecipadamente (idempotente), mas **NÃO** reseta `undoRecordActive` — `CleanUp` faz isso centralizadamente.

---

## 2. O que foi alterado (plano executado)

### 2.1 `source/main/Mod_04_Main.bas` — CleanUp consolidado

**Antes (fragilizado):** dois blocos `On Error Resume Next` separados — um para `ScreenRefresh`, outro para `EndCustomRecord`, com `On Error GoTo 0` entre eles.

**Depois (robusto):** um único `On Error Resume Next` cobre `ScreenRefresh` + fechamento do `EndCustomRecord`, com cancelamento do handler via `On Error GoTo 0` apenas no final do bloco.

```vba
    On Error Resume Next
    Application.ScreenRefresh

    If undoRecordActive Then
        Dim objUndoEnd As Object
        Set objUndoEnd = CallByName(Application, "UndoRecord", VbGet)
        If Err.Number = 0 Then
            If Not objUndoEnd Is Nothing Then
                CallByName objUndoEnd, "EndCustomRecord", VbMethod
            End If
        End If
        Err.Clear
        Set objUndoEnd = Nothing
        undoRecordActive = False  ' so reseta DEPOIS do fechamento
        LogMessage "UndoRecord finalizado com sucesso", LOG_LEVEL_INFO
    End If
    On Error GoTo 0
```

### 2.2 `tests/VBA.Tests.ps1` — 5 novos testes de regressão

| Teste | O que valida |
|-------|-------------|
| `PadronizarDocumentoMain USA StartCustomRecord` | Integração undo ativa |
| `PadronizarDocumentoMain USA EndCustomRecord` | Fechamento do grupo |
| `undoRecordActive apos StartCustomRecord` | Flag setada corretamente |
| `EmergencyRecovery fecha UndoRecord quando undoRecordActive` | Recuperação idempotente |
| `CleanUp fecha EndCustomRecord DEPOIS de ScreenRefresh` | Ordem obrigatória no bloco |
| `Nenhuma operacao perigosa apos EndCustomRecord` | Ve-se que nada perigoso está após o fechamento |
| `PadronizarDocumentoMain nao usa Selection em lugar algum` | Anti-regressão de Selection |
| `Mod_04_Main nao contem Application.OnRepeat` | Proibição de OnRepeat |

### 2.3 Arquivos tocados

| Arquivo | Natureza da alteração |
|---------|----------------------|
| `source/main/Mod_04_Main.bas` | Consolidou bloco `On Error` + comentário anti-repetir |
| `source/main/Mod_01_Infrastructure.bas` | Nenhuma (já estava correto) |
| `tests/VBA.Tests.ps1` | 5 testes de regressão (undo) + fix de regex |
| `tests/VBA-AIStructure.Tests.ps1` | Fix de CRLF (pre-existing) |
| `VERSION` | `9.4.0` → `9.5.0` |
| `source/main/Mod_01_Infrastructure.bas` | `Z7_STDPROPOSERS_VERSION` → `"9.5.0"` |
| `RELEASE_NOTES.md` | Nova entrada v9.5.0 |

---

## 3. Instruções de Validação

### 3.1 Validação automática (já executada)

```
python scripts/fix_bas_encoding.py    → OK (todos .bas)
Testes VBA (Pester)                   → OK (todos passaram)
Testes Encoding                       → OK (todos passaram)
Testes Python (pytest)                → OK (todos passaram)
```

### 3.2 Validação manual com Word (checklist pós-fix)

1. Eleje `PadronizarDocumentoMain` (Alt+P) em um documento de teste.
2. **Ctrl+Z** uma vez → documento deve reverter integralmente em uma única ação.
3. **Ctrl+Z** uma segunda vez → NÃO deve haver crash; deve desfazer ação anterior ou nada.
4. Fechar e reabrir o documento, repetir os passos acima.
5. Rodar `CorrigirProposituraComIA` (Alt+C) e verificar que o undo funciona independentemente.