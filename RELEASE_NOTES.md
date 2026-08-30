## v9.5.0 — Z7 StdProposers

### Correcoes Criticas

- **Blindagem final da pilha de desfazer (UndoRecord)**: consolidado o bloco `On Error Resume Next` no `CleanUp` de `PadronizarDocumentoMain`, eliminando a troca indevida de handler entre `ScreenRefresh` e `EndCustomRecord` que podia gerar entradas fantasmas na pilha de undo e causar Access Violation no segundo Ctrl+Z
- **Integracao undo restaurada**: `StartCustomRecord`/`EndCustomRecord` reintegrados com protecoes robustas (flag `undoRecordActive`, bloqueio de `DoEvents` parasitas, ordem obrigatoria no CleanUp, idempotencia em `EmergencyRecovery`)

### Testes

- **Novos testes Pester de regressao (undo)**: 5 testes adicionados validam presenca de `StartCustomRecord`/`EndCustomRecord`, flag `undoRecordActive`, ordem `ScreenRefresh < EndCustomRecord < undoRecordActive=False`, ausencia de `Selection.`, e proibicao de `Application.OnRepeat`
- **Teste de limite de operacoes**: valida que nenhuma operacao perigosa ocorre apos `EndCustomRecord` ate o `Exit Sub` da macro

### Decisoes de Design

- **Suporte a "Repetir" (F4) nao implementado**: `Application.OnRepeat` ja foi removido no v9.0.0 por corromper a pilha de undo; reintroduzir seria regressao. Documentado em comentario no codigo

---

## v9.3.0 — Z7 StdProposers

### Mudancas Estruturais
- **Remocao de integracao undo em PadronizarDocumentoMain**: a macro nao registra mais operacoes no historico de desfazer do Word. As alteracoes sao permanentes e nao podem ser desfeitas com Ctrl+Z. Isso elimina definitivamente o crash no segundo desfazer.
- **Removidos StartCustomRecord/EndCustomRecord** de PadronizarDocumentoMain
- **Removido undoGroupEnabled** — flag ja nao e necessaria

### Correcoes de Encoding
- **Sanitizacao de texto da IA**: nova funcao `SanitizarTextoIA` que remove caracteres de controle, BOM markers e normaliza quebras de linha antes de inserir texto corrigido no documento
- **DesescaparJSON corrigido**: `\n` agora produz `vbCr` (separador de paragrafo do Word) em vez de `vbCrLf` que causava soft-line-breaks fantasmas
- **Validacao de codepoints Unicode**: rejeita NUL, BOM markers (U+FFFE/U+FFFF) e surrogates isolados na resposta da IA
- **LimparRespostaIA corrigido**: deteccao de blocos markdown usa `vbCr` em vez de `vbLf`

### Testes
- **7 novos testes de encoding**: validam presenca de SanitizarTextoIA, uso correto de vbCr em DesescaparJSON, sanitizacao em SubstituirTextoPreservandoFormatacao e ProcessarTextoComIA, validacao de codepoints Unicode
- **Testes de undo atualizados**: validam ausencia de StartCustomRecord/EndCustomRecord em PadronizarDocumentoMain

### Assets
- chat_ia-v9.3.0.zip — Chat IA com contexto do documento
- config_prompt-v9.3.0.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

## v9.2.1 — Z7 StdProposers

### Correcoes Criticas
- **Crash no 2o Desfazer corrigido**: removido doc.UndoClear que causava entradas fantasmas na pilha de undo do Word, resultando em Access Violation ao desfazer pela segunda vez apos PadronizarDocumentoMain
- **Controle de DoEvents durante UndoRecord**: adicionada flag global undoRecordActive que impede chamadas de DoEvents enquanto o grupo de undo esta ativo, prevenindo criacao de entradas parasitas
- **ScreenRefresh reposicionado**: movido para apos SetAppState no CleanUp para evitar interferencia com a pilha de undo

### Melhorias
- **42 chamadas de DoEvents atualizadas**: todas as funcoes de processamento de paragrafos agora verificam undoRecordActive antes de chamar DoEvents
- **6 novos testes de regressao**: testes adicionados para prevenir reintroducao do bug de undo
- **Documentacao atualizada**: .clinerules e PROCESSAMENTO_PADRONIZAR_DOCUMENTO.md atualizados com regras de seguranca de undo

### Assets
- chat_ia-v9.2.1.zip — Chat IA com contexto do documento
- config_prompt-v9.2.1.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

## v9.2.0 — Z7 StdProposers

### Correcoes Criticas
- **Crash no 2o Desfazer corrigido**: removidas operacoes `doc.Save` e `doc.Range(0,0).Select` do `CleanUp` que criavam entradas orfas na pilha de undo apos `EndCustomRecord`, causando Access Violation no Word ao desfazer pela segunda vez
- **doc.Save removido de CreateDocumentBackup**: operacao de salvamento antes do `StartCustomRecord` removida para evitar entradas de undo indesejadas; backup continua sendo feito via `fso.CopyFile`

### Refatoracao
- **Remocao de funcionalidades descontinuadas**: removidas chamadas `BackupListFormats`, `BackupCenteredParagraphs`, `RestoreListFormats`, `FormatBulletedParagraphsIndent`, `RestoreCenteredParagraphs`, `CleanupCenteredParaBackup`, `RemovePageBreaks` e verificacao de dados sensiveis (`CheckSensitiveData`) do pipeline
- **Documentacao atualizada**: PROCESSAMENTO_PADRONIZAR_DOCUMENTO.md e PASSO_A_PASSO_PADRONIZAR_DOCUMENTO.md sincronizados com o estado atual do codigo

### Assets
- chat_ia-v9.2.0.zip — Chat IA com contexto do documento
- config_prompt-v9.2.0.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

## v9.1.0 — Z7 StdProposers

### Correcoes
- **Ementa formatada como titulo**: corrigido bug onde paragrafos em branco acima do titulo causavam desalinhamento de indices estruturais, fazendo a ementa receber a formatacao do titulo (negrito, sublinhado, centralizado)

### Melhorias
- Zoom de visualizacao padronizado para **130%** (antes inconsistente: 120% na configuracao inicial e 140% na restauracao)

### Documentacao
- Atualizadas referencias de zoom em PASSO_A_PASSO_PADRONIZAR_DOCUMENTO.md e PROCESSAMENTO_PADRONIZAR_DOCUMENTO.md

### Assets
- chat_ia-v9.1.0.zip — Chat IA com contexto do documento
- config_prompt-v9.1.0.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

## v9.0.0 — Z7 StdProposers

### Correcoes Criticas
- **Crash no 2o Desfazer corrigido**: removida chamada incompativel Application.OnRepeat que corrompia a pilha de undo e causava Access Violation no Word
- **_remove_z7_modules refatorado**: substituida enumeracao COM fragil por iteracao baseada em indice (VBComponents.Count + VBComponents.Item(i)), garantindo remocao confiavel de modulos Z7 existentes

### Melhorias
- Zoom de visualizacao ajustado para **120%** (antes 140%)
- AI_CONTEXT.md excluido — conteudo distribuido nos arquivos .clinerules/ e .cline/custom_modes.json
- import_bas_to_normal.py: removidos artefatos PLACEHOLDER_PART; _remove_z7_modules robusto

### Documentacao
- .clinerules/01-project-conventions.md: +heuristicas estruturais, +detalhes logging VBA/Python, atualizada regra de ouro
- .clinerules/02-vba-coding.md: +detalhes Mod_11/Mod_12, +rodape (formato Pagina X de Y)
- .clinerules/03-python-coding.md: +deteccao de Word multi-estrategia
- .clinerules/05-testing.md: +tabela completa de 14 arquivos de teste

### Assets
- chat_ia-v9.0.0.zip — Chat IA com contexto do documento
- config_prompt-v9.0.0.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA