# Z7 StdProposers — Convenções de Projeto

> **alwaysApply: true**
> Regras fundamentais que se aplicam a TODAS as interações neste projeto.

## 1. Arquitetura do Projeto

Este é um projeto de automação Microsoft Word para padronização de documentos legislativos brasileiros.

Duas partes coordenadas:
- **VBA** (`source/main/`): 12 módulos (`Mod_01` a `Mod_12`) — engine de formatação no Word.
- **Python** (`ai/`): Integração Gemini/OpenRouter via `chat_ia.py`, `config_prompt.py`, `z7_logging.py`, `z7_api_key.py`, `z7_theme.py`.

## 2. Encoding — REGRA CRÍTICA

| Tipo de Arquivo | Encoding |
|----------------|----------|
| `.bas` (VBA) | **CP1252 (Latin1)** — NUNCA salvar como UTF-8 |
| `.py` (Python) | **UTF-8**, LF |
| `.ps1` (PowerShell) | **UTF-8**, CRLF |
| `.md`, `.txt` | **UTF-8**, CRLF |

**Após editar qualquer `.bas`, execute:** `python scripts/fix_bas_encoding.py`

## 3. Limites de Módulo

- NÃO crie novos módulos VBA. Os 12 existentes são o limite.
- NÃO reintroduza monólitos. Cada módulo tem responsabilidade clara.
- Preserve `Option Explicit` em todo `.bas` (linha 2, após `Attribute VB_Name`).
- Preserve `Attribute VB_Name = "..."` na linha 1 de todo `.bas`.

## 4. Regras de Segurança VBA

### 4.1 Invalidação de Índices (CRÍTICO)
Toda rotina que deleta parágrafos DEVE recarregar a estrutura:
```vba
If removedCount > 0 Then IdentifyDocumentStructure doc
```
NUNCA mova esta chamada para antes do loop de deleção.

### 4.2 UndoRecord
`StartCustomRecord` e `EndCustomRecord` devem estar pareados em TODOS os caminhos de saída.
Falhas roteiam via `CriticalErrorHandler -> GoTo CleanUp`.

**NUNCA use `doc.UndoClear`** — nem antes de `StartCustomRecord`, nem após `EndCustomRecord`. `doc.UndoClear` cria entradas fantasmas na pilha de undo do Word que causam Access Violation (crash) quando o usuário tenta desfazer pela segunda vez. Use APENAS `StartCustomRecord`/`EndCustomRecord` para agrupar operações.

### 4.3 Disciplina COM
Evite `Range.Characters(n)` em loops quentes. Prefira operar no `Range` diretamente.

### 4.4 StatusBar Encoding
`word.StatusBar` NÃO aceita caracteres acentuados. Use ASCII: `"verificacao"`, não `"verificação"`.

## 5. Pipeline de Formatação (Double-Pass)

1. Pass 1: normaliza e limpa debris estruturais.
2. Pass 2: formatação sensível a alinhamento.
   - Pass 2 só executa se Pass 1 alterou conteúdo (`documentDirty = True`).

### 5.1 Heurísticas Estruturais (Elementos Detectados)

Âncoras-chave inferidas por assinaturas de texto/formatacão:

- **Título** e **Ementa**.
- Cabeçalho/corpo da **Justificativa**.
- Âncoras de **Plenário/Data**.
- Bloco de **Assinatura** (inclui autoria por "Presidente" e "Prefeito").
- Cabeçalho/corpo de **Anexo**.

## 6. Logging

- **VBA**: `Mod_03_Pipeline.bas` (`InitializeLogging`, `LogMessage`, `SafeFinalizeLogging`). Snapshots em INICIO, FIM, ERRO_CRITICO. Metadados de sessão/operação (`currentLogSessionId`, `currentOperationId`) com formato de elapsed-time estruturado e marcador `[op=...]`. Buffered flushing controlado por `LOG_BUFFER_FLUSH_SECONDS`.
- **Python**: `z7_logging.py` com `RotatingFileHandler` (2MB / 3 backups). Logs em `%LOCALAPPDATA%\Z7\Apps\Z7_StdProposers\setup\logs`. Formato: `YYYY-MM-DD HH:MM:SS.mmm | LEVEL | z7.<component> | mensagem`.

## 7. Testes

- Runner: `tests/Run-Tests.ps1 -TestSuite All -NoProgress`
- Suites: `All`, `VBA`, `Encoding`, `Python`, `VBA-Logging`, `VBA-AIStructure`
- Pester para testes PowerShell, pytest para Python.
- **SEMPRE execute os testes após qualquer mudança** — não faça commit sem testes passando.

## 8. Versionamento

- Versão no arquivo `VERSION` na raiz.
- `Mod_01_Infrastructure.bas` contém `Z7_STDPROPOSERS_VERSION` que DEVE coincidir com `VERSION`.
- Build artifacts em `dist/` seguem o padrão `<name>-v<VERSION>.zip`.

## 9. Regra de Ouro

Ao editar qualquer código:
1. **Leia** as regras em `.clinerules/` antes de começar.
2. **Execute** `python scripts/fix_bas_encoding.py` após editar `.bas`.
3. **Execute** `tests/Run-Tests.ps1 -TestSuite All -NoProgress` antes de finalizar.
4. **NUNCA** enfraqueça um teste para fazê-lo passar — corrija o código.

## 10. Critérios para Agentes de IA

### 10.1 Comportamento de Undo/Redo
- `PadronizarDocumentoMain` usa `StartCustomRecord`/`EndCustomRecord` via late-binding (`CallByName`) para agrupar todas as edições em um único comando de desfazer. A flag global `undoRecordActive` bloqueia `DoEvents` parasitas durante o pipeline.
- `CorrigirProposituraComIA` funciona como um único comando de desfazer/refazer no Word (via `StartCustomRecord`/`EndCustomRecord`).
- **NUNCA use `doc.UndoClear`** — cria entradas fantasmas na pilha de undo que causam Access Violation (crash) no Word.
- NÃO chame `DoEvents` entre `EndCustomRecord` e `SetAppState` no CleanUp — pode criar entradas parasitas na pilha de undo.

### 10.2 Segurança e Estabilidade
- Use estrutura de tratamento de erros adequada com `GoTo CleanUp`.
- Evite chamadas que possam corromper a pilha de desfazer, como `Application.OnRepeat` (já removida do código).