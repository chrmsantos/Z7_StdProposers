## v9.10.0 — Z7 StdProposers

### Correcoes / Padronizacao de Espacamento

- **2 linhas em branco nas zonas especiais ao final da formatação**: ficam **exatamente 2 linhas em branco** acima da **Data**, acima do **Título da Justificativa** e acima e abaixo da **Ementa** (antes ficava apenas 1 em todos os casos)
- **`RemoverLinhasEmBrancoExtras` com zonas protegidas**: a padronização generalizada de linhas puladas mantém no máximo 1 linha vazia consecutiva, mas no máximo **2** nas zonas protegidas (acima/abaixo da ementa, acima do título da justificativa, acima da data)
- **Garantia final pós-padronização** (`PadronizarDocumentoMain`): `ForceDataSpacing`, `ForceJustificativaTitleSpacing` e `ForceEmentaSpacing` rodam **depois** de `RemoverLinhasEmBrancoExtras`/`EnsureConsideringBlankLines` (ordem de baixo para cima), para a regra não ser desfeita por formatação generalizada posterior
- **`ForceEmentaSpacing` corrigido**: agora garante **exatamente 2** linhas acima (a regra antiga de 3 acima foi removida) e também remove excesso; `ForceDataSpacing` garante exatamente 2 acima da Data (com validação/fallback de localização via `IsDataElement`)
- **Nova rotina `ForceJustificativaTitleSpacing`** (`Mod_09_SpecialParagraphs.bas`): garante exatamente 2 linhas em branco acima do título "Justificativa"

### Testes

- **6 novos testes de contrato (Pester)** em `VBA.Tests.ps1`: exatamente 2 linhas na ementa/data/título da justificativa, preservação das zonas protegidas em `RemoverLinhasEmBrancoExtras` e ordem da garantia final após a padronização de linhas puladas

### Sincronizacao de Versao

- Versao 9.10.0 alinhada em `VERSION`, `Z7_STDPROPOSERS_VERSION` (`Mod_01_Infrastructure.bas`) e `_APP_VERSION` (`config_prompt.py` e `chat_ia.py`)

---


## v9.9.0 — Z7 StdProposers

### Novidades

- **Quebra de parágrafo antes de sufixo de Vereador** (`BreakParagraphBeforeVereadorSuffix`, `Mod_09_SpecialParagraphs.bas`): logo no início das formatações (Passagem 1 de `PreviousFormatting`, logo após a normalização de quebras `^l → ^p`), parágrafos **de uma única linha** terminados por `" - vereador"` / `" - vereadora"` (hífen ASCII ou en-dash `–`) são quebrados imediatamente antes do sufixo, colocando o sufixo em parágrafo próprio
- Comparação *case-insensitive*; pontuação final opcional (`.` ou `,`) após a palavra é aceita
- Guard multi-linha (`ParagraphHasMultipleLines`, *fail-closed*): parágrafos que ultrapassam uma linha não são quebrados; prefixo vazio nunca gera parágrafo vazio
- Após as quebras, a estrutura do documento é recarregada (`IdentifyDocumentStructure doc`) — índices estruturais invalidados pelas novas marcas de parágrafo

### Testes

- **5 novos testes de contrato (Pester)** em `VBA.Tests.ps1`: presença da rotina, cobertura de hífen/en-dash + vereador/vereadora + case-insensitive + pontuação opcional, guard de linha única, proteção de prefixo vazio + refresh de estrutura, e posição da etapa no início do pipeline

### Sincronizacao de Versao

- Versao 9.9.0 alinhada em `VERSION`, `Z7_STDPROPOSERS_VERSION` (`Mod_01_Infrastructure.bas`) e `_APP_VERSION` (`config_prompt.py` e `chat_ia.py` — atualizados de 9.7.0)

---


## v9.7.0 — Z7 StdProposers

### Seguranca (anti prompt-injection)

- **Texto do documento e SEMPRE DADO, nunca prompt**: os tres fluxos que enviam texto do documento a IA (`CorrigirProposituraComIA`/`TestarRevisaoTextoSelecionado`, `IdentifyDocumentStructureWithAI` e Chat IA) agora garantem estruturalmente que o conteudo extraido do Word seja processado como texto a revisar/segmentar/contextualizar — jamais como prompt, instrucao ou comando
- **Envelope de dados sem marcador de fechamento** (`MontarMensagemDados`, `AI_MontarMensagemDados`, `_wrap_doc_data`): a regiao de dados vai do marcador `<<<INICIO_...>>>` ate o FINAL da mensagem, impedindo breakout por injecao de marcador dentro do proprio documento
- **Guard anti-injecao em codigo** (`MontarGuardAntiInjecao`, `AI_MontarGuardAntiInjecao`, `_with_doc_data_guard`): anexado ao system prompt depois do prompt configuravel — vale para qualquer `revision_prompt.txt`/`chat_system_prompt.txt` e nao pode ser removido por eles
- **`RemoverEnvelopeResposta`** (Mod_11): remove eventual eco do envelope apenas nas bordas da resposta da IA
- **Chat IA**: marcadores `---INICIO/FIM DO DOCUMENTO---` (com fechamento, passiveis de breakout) substituidos pelo envelope de dados; no fallback `_context_pending`, a mensagem do usuario vai antes e os dados por ultimo

### Testes

- **17 novos testes de contrato**: 7 Pester (`VBA-AIStructure.Tests.ps1`) + 10 pytest (`test_chat_ia.py::TestDocDataPromptIsolation`), incluindo guarda de regressao contra marcadores com fechamento e cenario de injecao que repete o marcador

### Sincronizacao de Versao

- Versao 9.7.0 alinhada em `VERSION`, `Z7_STDPROPOSERS_VERSION` (`Mod_01_Infrastructure.bas`) e `_APP_VERSION` (`config_prompt.py` e `chat_ia.py`)

### Assets

- chat_ia-v9.7.0.zip — Chat IA com contexto do documento
- config_prompt-v9.7.0.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

## v9.6.1 — Z7 StdProposers

### Melhorias

- **Corretor IA limitado a um paragrafo por vez**: `CorrigirProposituraComIA` recusa selecoes que abrangem mais de um paragrafo e exibe um aviso gentil (MsgBox) sem alterar o documento; cursor solto ou selecao dentro de um paragrafo continuam funcionando normalmente
- **Nova funcao `SelecaoAbrangeMultiplosParagrafos`**: conta os paragrafos tocados pela selecao descontando um eventual paragrafo final nao tocado, evitando falso positivo em selecao de paragrafo unico que inclui a propria marca de paragrafo (vbCr)

### Testes

- **Novo teste de contrato (Pester)**: valida que `CorrigirProposituraComIA` contem a validacao de paragrafo unico e o aviso ao usuario dentro do proprio entrypoint

### Sincronizacao de Versao

- Versao 9.6.1 alinhada em `VERSION`, `Z7_STDPROPOSERS_VERSION` (`Mod_01_Infrastructure.bas`) e `_APP_VERSION` (`config_prompt.py` e `chat_ia.py` — este ultimo estava desatualizado em 8.12.4)

### Assets

- chat_ia-v9.6.1.zip — Chat IA com contexto do documento
- config_prompt-v9.6.1.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

## v9.6.0 — Z7 StdProposers

### Correcoes

- **VBProject.Name estavel apos importacao**: corrigido em `import_bas_to_normal.py` o bug que fazia o `VBProject.Name` mudar para `TemplateProject` ao importar os modulos Z7
- **Prompts de revisao e Chat IA aprimorados** (`config_prompt`): prompt do Chat IA reformulado com paragrafos separados, instrucoes para ignorar erros de formatacao/datas e indicar localizacao por paragrafo/linha; Corretor de Propositura passa a preservar a palavra 'Indica' (nao substitui por 'Indico')

### Mudancas Estruturais

- **Atalhos de teclado removidos**: `AutoOpen` e `RegistrarAtalhosTeclado` removidos de `Mod_04_Main` (atalhos Alt+P e Alt+C descontinuados)
- **`NormalTemplate.Save` removido do VBA**: evita gravacao do template durante a padronizacao
- **`_APP_VERSION` sincronizado** em `config_prompt.py` (8.12.4 → 9.6.0, alinhado com `VERSION`)

### Testes

- **Testes de anti-regressao atualizados**: validam a ausencia dos atalhos removidos e a estabilidade de `VBProject.Name` apos a importacao

### Assets

- chat_ia-v9.6.0.zip — Chat IA com contexto do documento
- config_prompt-v9.6.0.zip — Editor de prompts side-by-side
- import_bas_to_normal.exe — Importador de modulos VBA

---

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