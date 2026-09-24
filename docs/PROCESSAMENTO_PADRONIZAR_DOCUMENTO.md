# Processamento completo da macro `PadronizarDocumentoMain`

> **Descrição geral:** Este documento detalha, em lista numerada e na ordem exata de execução, todo o processamento realizado pela macro **`PadronizarDocumentoMain`** (módulo `Mod_04_Main.bas`), desde as verificações iniciais até a finalização e o salvamento do documento.
>
> **Observação sobre o nome:** não existe uma macro chamada `PadronizarProposituraMain` no projeto. O entrypoint público de padronização é **`PadronizarDocumentoMain`** (`Public Sub PadronizarDocumentoMain()` — `Mod_04_Main.bas`, linha 14). Este documento descreve essa macro.

---

## Visão geral das fases

A macro executa em sete fases encadeadas:

| Fase | Conteúdo |
|------|----------|
| 1 | Inicialização e verificações iniciais |
| 2 | Backups (documento, visualização, imagens) |
| 3 | Início do grupo de desfazer (`UndoRecord`) |
| 4 | Pipeline de formatação (até duas passagens) |
| 5 | Ajustes pós-pipeline |
| 6 | Finalização e registro de sucesso |
| 7 | `CleanUp` / tratamento de erros |

---

## Fase 1 — Inicialização e verificações iniciais

1. Define as variáveis globais de execução:
   - `executionStartTime = Now` (marca o início para cálculo de tempo).
   - `formattingCancelled = False` (flag de cancelamento).

2. Chama `CheckWordVersion()`:
   - Se a versão do Word for inferior ao mínimo suportado (`MIN_SUPPORTED_VERSION`, Word 2010+), exibe `MsgBox` crítico, registra `LOG_LEVEL_ERROR` e encerra (`Exit Sub`) sem modificar nada.

3. Obtém o documento ativo (`ActiveDocument`):
   - Se não houver nenhum documento aberto, exibe `MsgBox` crítico e encerra.

4. Inicializa o sistema de logging via `InitializeLogging(doc)`:
   - Se a inicialização falhar, avisa na barra de status que o log está desabilitado.
   - Se tiver sucesso, registra o snapshot de contexto `"INICIO"` via `LogContextSnapshot`.

5. Inicializa o sistema de progresso com `InitializeProgress 18` (18 etapas do pipeline em duas passagens).

6. Configura o estado da aplicação com `SetAppState(False, "Iniciando...")`:
   - Desabilita interações do usuário durante o processamento.
   - Em caso de falha, registra `LOG_LEVEL_WARNING`.

7. Avança o progresso (`IncrementProgress "Verificando documento"`) e executa `PreviousChecking(doc)`:
   - Valida tipo de documento, proteção, modo somente-leitura e espaço em disco.
   - Se a verificação falhar, desvia para `CleanUp`.

8. Se o documento ainda não foi salvo (`doc.Path = ""`):
   - Chama `SaveDocumentFirst(doc)` para solicitar o salvamento prévio.
   - Se o usuário cancelar/falhar, registra `"Operacao cancelada - documento nao foi salvo"` e desvia para `CleanUp`.

---

## Fase 2 — Backups (sem alteração do documento)

9. Cria backup do documento: `CreateDocumentBackup(doc)` — salva uma cópia `.docm` com *timestamp* na pasta de backups. Em falha, registra aviso e prossegue.

10. Faz backup das configurações de visualização: `BackupViewSettings(doc)` — salva tipo de vista, régua, marcadores, zoom, etc.

11. Faz backup de imagens: `BackupAllImages(doc)` — cataloga posição, dimensões e tipo (inline/flutuante) de todas as imagens.

---

## Fase 3 — Início do Grupo de Desfazer (UndoRecord)

12. **Inicia o `UndoRecord`** via late-binding (`CallByName`): `StartCustomRecord "Z7_STDPROPOSERS - Padronizacao"`. Define `undoRecordActive = True`, que bloqueia **todas** as 42+ chamadas de `DoEvents` durante o pipeline, prevenindo entradas parasitas na pilha de undo. Em caso de falha no late-binding, `undoRecordActive` permanece `False` e a macro prossegue sem UndoRecord (modo degradado).

---

## Fase 4 — Pipeline de formatação (até duas passagens)

15. Registra `"=== PIPELINE DE FORMATACAO (2 PASSAGENS) ==="` no log.

16. Constrói o cache de parágrafos: `BuildParagraphCache doc`:
    - Para cada parágrafo, captura texto bruto, texto normalizado (sem acentos, caixa baixa), presença de imagens e se é parágrafo especial (CONSIDERANDO, Justificativa, Vereador, Diante do exposto, Requeiro, Anexo).
    - Internamente dispara a identificação da estrutura do documento (`IdentifyDocumentStructure`), que identifica por IA ou, como fallback, por heurística (posição/texto/formatação) os elementos: **Título**, **Ementa**, **Vocativo**, **Corpo**, **Título da Justificativa**, **Justificativa**, **Data (Plenário)**, **Assinatura**, **Título do Anexo** e **Anexo**.

17. Força a primeira passagem: `documentDirty = True`.

18. Entra no loop de passagens `For pipelinePass = 1 To 2`:
    1. **Controle de passagem 2:** se for a passagem 2 e `documentDirty = False` (nada foi alterado na passagem 1), registra `"PASSAGEM 2 IGNORADA"` e sai do loop.
    2. **Reset da flag:** `documentDirty = False` antes de cada passagem.
    3. **Reindexação:** se `pipelinePass > 1`, reconstrói o cache (`BuildParagraphCache doc`) para evitar índices obsoletos (parágrafos podem ter sido removidos/inseridos na passagem anterior).
    4. **Formatação:**
       - Passagem 1 → `PreviousFormatting(doc)` (pipeline completo).
       - Passagem 2 → `PreviousFormattingPass2(doc)` (apenas etapas index-dependentes, reduzindo ~60–70% do tempo).
    5. **Restauração de imagens:** após cada passagem, `RestoreAllImages(doc)`.

### 4.1 — `PreviousFormatting` (Passagem 1 — pipeline completo)

19. `ApplyPageSetup(doc)` — configuração de página:
    - Margens superior/inferior (`TOP_MARGIN_CM`, `BOTTOM_MARGIN_CM`).
    - Margens esquerda/direita (`LEFT_MARGIN_CM`, `RIGHT_MARGIN_CM`).
    - Distâncias de cabeçalho e rodapé.
    - Gutter = 0; orientação retrato (`wdOrientPortrait`).

20. `ClearAllFormatting doc` — limpeza total de formatação de cada parágrafo (preservando imagens).

21. Normalização de quebras:
    - `ReplaceLineBreaksWithParagraphBreaks doc` — converte quebras de linha em quebras de parágrafo.

22. Quebra de parágrafo antes de sufixo de Vereador: `BreakParagraphBeforeVereadorSuffix doc` — parágrafos terminados por `" - vereador"` / `" - vereadora"` (hífen ASCII ou en-dash `–`) que ocupam **no máximo uma linha** são quebrados imediatamente antes do sufixo, colocando o sufixo em parágrafo próprio:
    - Comparação *case-insensitive*; pontuação final opcional (`.` ou `,`) após a palavra é aceita.
    - Parágrafos multi-linha não são quebrados (guard `ParagraphHasMultipleLines`, *fail-closed*) e prefixo vazio não gera parágrafo vazio.
    - Se houver quebras, recarrega a estrutura (`IdentifyDocumentStructure doc`) — índices estruturais invalidados pelas novas marcas de parágrafo.

23. Limpeza estrutural:
    - `RemovePageNumberLines doc` — remove linhas de número de página.
    - `RemoveUnderscoreOnlyParagraphs doc` — remove parágrafos compostos apenas por *underscores*.
    - `CleanDocumentStructure doc` — limpa *debris* estruturais.
    - `RemoveAllTabMarks doc` — remove todas as marcas de tabulação.

24. Limpeza de prefixo da ementa: `RemoveEmentaLeadingLabelPrefix doc`.

25. Limpeza de sufixo da ementa: `RemoveEmentaTrailingMunicipioSuffix doc`.

26. Remoção de aspas da ementa: `RemoveEmentaQuotes doc`.

27. Substituição de "DAE" por "Poder Executivo Municipal" na ementa de Indicações: `ProcessEmentaIndicacao doc`.

28. Formatação do título: `FormatDocumentTitle doc`.

29. Aplicação da fonte padrão:
    - Se cache ativo → `ApplyStdFontOptimized(doc)`; em falha, *fallback* para `ApplyStdFont(doc)`.
    - Senão → `ApplyStdFont(doc)` (Arial 12).

30. Formatação de parágrafos: `ApplyStdParagraphs(doc)`.

31. Formatação do parágrafo 2 (ementa): `FormatSecondParagraph doc`.

32. Formatação dos parágrafos 2–4 após a ementa (justificado, recuo 2,5 cm): `FormatPostEmentaBodyParagraphs doc`.

33. Formatação de considerandos: `FormatConsiderandoParagraphs doc`.

34. Inserção de linhas em branco (justificativa, plenário, prefeito): `InsertJustificativaBlankLines doc`.

35. Aplicação de substituições de texto: `ApplyTextReplacements doc`.

36. Capitalização do início dos parágrafos: `CapitalizeFirstLetterOfParagraphs doc`.

37. Marca d'água e carimbo:
    - `RemoveWatermark doc` — remove marca d'água.
    - `InsertHeaderstamp doc` — insere carimbo no cabeçalho.

38. Limpeza de espaços múltiplos: `CleanMultipleSpaces doc`.

39. Controle de linhas em branco: `LimitSequentialEmptyLines doc`.

40. Substituição de datas do plenário: `ReplacePlenarioDateParagraph doc`.

41. Configuração de visualização: `ConfigureDocumentView doc`.

42. Inserção de rodapé: `InsertFooterStamp(doc)` (em falha, retorna `False` e aborta a passagem).

43. Ajustes finais de negrito e formatação:
    - `ApplyBoldToSpecialParagraphs doc`.
    - `SubstituiVereadoraPorSenhoraVereadora doc`.
    - `FormatVereadorParagraphs doc`.

44. Formatações especiais:
    - `FormatDianteDoExposto doc`.
    - `FormatRequeiroParagraphs doc`.
    - `FormatPorTodasRazoesParagraphs doc`.

45. Remoção de realces e bordas: `RemoveAllHighlightsAndBorders doc`.

46. Remoção de páginas vazias no final: `RemoveEmptyPagesAtEnd doc`.

47. Aplicação de formatação final universal: `ApplyUniversalFinalFormatting doc`.

48. Remoção de dois-pontos da justificativa: `RemoveJustificativaColon doc`.

49. Adição de espaçamento especial (ementa, justificativa, data): `AddSpecialElementsSpacing doc`.

50. Ajuste final de recuos para Vereador (travessões): `FixHyphenatedVereadorParagraphIndents doc`.

51. Garantia final de **2 linhas em branco** na ementa (acima e abaixo): `ForceEmentaSpacing doc`.

52. Garantia final de **2 linhas em branco** na data (acima): `ForceDataSpacing doc`.

53. Registra métrica `"Total de paragrafos"` e retorna `True`.

### 4.2 — `PreviousFormattingPass2` (Passagem 2 — seletiva)

> Executa somente as etapas **index-dependentes**, que podem ter sido invalidadas por inserções/remoções de parágrafos da passagem 1. Etapas idempotentes (limpeza total, fonte padrão, formatação de parágrafos, capitalização, etc.) são omitidas.

54. `RemoveEmentaLeadingLabelPrefix doc` (P2).

55. `RemoveEmentaTrailingMunicipioSuffix doc` (P2).

56. `RemoveEmentaQuotes doc` (P2).

57. `ProcessEmentaIndicacao doc` (P2).

58. `FormatDocumentTitle doc` (P2).

59. `FormatSecondParagraph doc` (P2).

60. `FormatPostEmentaBodyParagraphs doc` (P2).

61. `FormatConsiderandoParagraphs doc` (P2).

62. `InsertJustificativaBlankLines doc` (P2).

63. `ApplyTextReplacements doc` (P2).

64. `ReplacePlenarioDateParagraph doc` (P2).

65. `InsertFooterStamp(doc)` (P2).

66. Ajustes finais de negrito e formatação (P2): `ApplyBoldToSpecialParagraphs`, `SubstituiVereadoraPorSenhoraVereadora`, `FormatVereadorParagraphs`.

67. Formatações especiais (P2): `FormatDianteDoExposto`, `FormatRequeiroParagraphs`, `FormatPorTodasRazoesParagraphs`.

68. Remoção de dois-pontos da justificativa (P2): `RemoveJustificativaColon doc`.

69. Adição de espaçamento especial (P2): `AddSpecialElementsSpacing doc`.

70. Ajuste final de recuos para Vereador (P2): `FixHyphenatedVereadorParagraphIndents doc`.

71. Garantia final de **2 linhas em branco** na ementa (P2): `ForceEmentaSpacing doc`.

72. Garantia final de **2 linhas em branco** na data (P2): `ForceDataSpacing doc`.

---

## Fase 5 — Ajustes pós-pipeline

73. Remoção de linhas em branco extras: `RemoverLinhasEmBrancoExtras doc` — máximo 1 linha vazia consecutiva (máximo **2** nas zonas protegidas: acima/abaixo da ementa, acima do título da justificativa e acima da data).

74. Garantia de linhas em branco após "CONSIDERANDO": `EnsureConsideringBlankLines doc`.

75. Garantia final de **2 linhas em branco** nas zonas especiais (acima da data, acima do título da justificativa, acima e abaixo da ementa), executada **depois** de toda a padronização generalizada de linhas puladas para não ser desfeita: `ForceDataSpacing doc`, `ForceJustificativaTitleSpacing doc` e `ForceEmentaSpacing doc` (ordem de baixo para cima).

78. Formatação de recuos de imagens: `FormatImageParagraphsIndents(doc)` — zera recuo esquerdo/primeira linha e centraliza parágrafos com imagens inline.

79. Centralização de imagem após o Plenário: `CenterImageAfterPlenario(doc)` — centraliza imagem localizada entre a 5ª e a 7ª linha após o parágrafo "Plenário".

81. Remoção de numeração de parágrafos em branco: `RemoveNumberingFromBlankParagraphs(doc)`.

82. Garantia final de fonte: reaplica **Arial 12** em todo o `doc.Range.Font` (corrige trechos que Find/Replace possam ter deixado com a fonte do estilo Normal, ex.: Calibri).

83. Restauração das configurações de visualização: `RestoreViewSettings(doc)` — restaura tudo, **exceto o zoom**, que é mantido em 130%.

---

## Fase 6 — Finalização e registro de sucesso

84. Se `formattingCancelled = True`, desvia para `CleanUp`.

85. Avança o progresso (`IncrementProgress "Finalizando"`).

86. Registra sucesso: `LogMessage "Documento padronizado com sucesso"` e `LogContextSnapshot doc, "FIM"`.

87. Calcula o tempo de execução: `execSeconds = CLng((Now - executionStartTime) * 86400)`.

88. Exibe na barra de status: `"Padronizacao concluida em Xs, com Y erros e Z avisos! (z7_stdproposers)"`.

---

## Fase 7 — `CleanUp` (sempre executado) e tratamento de erros

89. **Fecha o `UndoRecord`** se `undoRecordActive = True`: `CallByName(objUndoEnd, "EndCustomRecord", VbMethod)` via late-binding. Esta operação deve vir **antes** de qualquer outra operação no `CleanUp` para garantir que as edições sejam agrupadas como uma única ação de desfazer.

90. Limpa o cache de parágrafos: `ClearParagraphCache`.

91. Executa `SafeCleanup` (limpeza geral de objetos/temp).

92. Limpa as variáveis de proteção:
    - `CleanupImageProtection` (imagens).
    - `CleanupViewSettings` (configurações de visualização).

93. Restaura o estado da aplicação: `SetAppState(True, "", True)` — preservando a barra de status (mantém a mensagem final).

94. Atualiza a tela: `Application.ScreenRefresh` (chamado APENAS após `SetAppState` restaurar `ScreenUpdating`).

95. Finaliza o logging: `SafeFinalizeLogging`.

> **`CriticalErrorHandler`:** em caso de erro crítico, registra `"ERRO CRITICO #<número>: <descrição> em <fonte> (Linha: <linha>)"` no log, registra snapshot `"ERRO_CRITICO"` e retoma a execução em `CleanUp`, garantindo que o estado da aplicação seja restaurado.

---

## Observações importantes

1. **Duas passagens:** o pipeline roda até duas vezes. A passagem 2 só executa se a passagem 1 alterou o documento (`documentDirty = True`) e, antes de rodar, reconstrói o cache para manter os índices válidos.

2. **Desfazer (Ctrl+Z):** todas as edições são agrupadas em um único `UndoRecord`; é possível desfazer toda a padronização com um único Ctrl+Z. **`doc.UndoClear` NÃO é usado** — cria entradas fantasmas que causam crash no Word.

3. **Proteção de imagens:** formatações de fonte e parágrafo preservam imagens inline e *shapes* flutuantes; em parágrafos com imagem, a formatação é aplicada caractere a caractere para evitar danos.

4. **Backup automático:** antes de qualquer modificação, o documento original é salvo como backup (`.docm` com *timestamp*).

5. **Parágrafos especiais:** "Justificativa", "Anexo(s)", "Vereador/Vereadora", "CONSIDERANDO", "Diante do exposto" e "Requeiro" recebem tratamento diferenciado em várias etapas (negrito, alinhamento, recuos, espaçamento).

6. **Regra de parágrafos multi-linha:** parágrafos com mais de 1 linha (quebra por largura ou Shift+Enter) **não** recebem alinhamento centralizado nem recuo à esquerda/primeira linha igual a zero. O guard `ParagraphHasMultipleLines` (`Mod_01_Infrastructure.bas`, *fail-closed*) é aplicado nas formatações de elementos especiais e centralizações (título, assinatura, imagens, plenário, vereador e rodapé). A normalização/reset (`ClearAllFormatting`), a formatação de corpo/ementa (`ApplyStdParagraphs`, `FormatPostEmentaBodyParagraphs`, `FormatSecondParagraph`) e a normalização de listas com marcadores (`FormatBulletedParagraphsIndent`) continuam zerando recuos normalmente.

