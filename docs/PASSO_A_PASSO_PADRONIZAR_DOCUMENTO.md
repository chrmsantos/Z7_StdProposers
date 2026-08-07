# Passo-a-Passo: Edições realizadas por `PadronizarDocumentoMain`

> **Descrição geral:** Este documento lista, na ordem exata de execução, todas as formatações e substituições de texto aplicadas ao documento ao rodar a subrotina `PadronizarDocumentoMain` (módulo `Mod4Main.bas`).

---

## Resumo do fluxo principal

A subrotina opera em duas fases:
1. **Pré-pipeline:** verificações, backups e proteções.
2. **Pipeline de formatação (duas passagens):** aplica todas as edições ao documento.
3. **Pós-pipeline:** ajustes finais, restaurações e garantias.

---

## 1. Verificações Iniciais (sem edição)

| Etapa | Descrição |
|---|---|
| 1.1 | Verifica se o Word é versão 2010 ou superior (`CheckWordVersion`). |
| 1.2 | Verifica se existe um documento ativo. |
| 1.3 | Inicializa o sistema de logging e registra snapshot de contexto. |
| 1.4 | Executa `PreviousChecking`: valida tipo de documento, proteção, modo somente-leitura, espaço em disco, dados sensíveis (CPF/CNPJ/cartão/CID). |
| 1.5 | Se o documento não estiver salvo, solicita o salvamento prévio. |

---

## 2. Backups (sem edição no documento)

| Etapa | Descrição |
|---|---|
| 2.1 | **Backup do documento:** salva cópia `.docm` com timestamp na pasta de backups. |
| 2.2 | **Backup das configurações de visualização:** salva tipo de vista, régua, marcadores, etc. |
| 2.3 | **Backup de imagens:** cataloga posição, dimensões e tipo (inline/flutuante) de todas as imagens. |
| 2.4 | **Backup de listas:** salva tipo, nível e string de cada parágrafo com formatação de lista. |
| 2.5 | **Backup de parágrafos centralizados:** salva índices e texto dos parágrafos centralizados (excluindo Heading/Título). |

---

## 3. Início do Grupo de Desfazer (UndoRecord)

| Etapa | Descrição |
|---|---|
| 3.1 | Inicia o `UndoRecord` customizado "Z7\_STDPROPOSERS - Padronizacao". Todas as edições subsequentes são agrupadas em uma única ação de desfazer. |

---

## 4. Construção do Cache de Parágrafos (`BuildParagraphCache`)

| Etapa | Descrição |
|---|---|
| 4.1 | Para cada parágrafo: captura texto bruto, texto normalizado (sem acentos, caixa baixa), se tem imagens, se é parágrafo especial (CONSIDERANDO, Justificativa, Vereador, Diante do exposto, Requeiro, Anexo). |
| 4.2 | **Identificação da estrutura do documento** (`IdentifyDocumentStructure`): identifica por heurísticas posição/texto/formatação os seguintes elementos: **Título**, **Ementa**, **Vocativo**, **Proposição**, **Título da Justificativa**, **Justificativa**, **Data (Plenário)**, **Assinatura**, **Título do Anexo**, **Anexo**. |

---

## 5. Pipeline de Formatação — Passagem 1 (e possivelmente Passagem 2)

> Se a Passagem 1 alterar o documento (`documentDirty = True`), uma Passagem 2 é executada. A Passagem 2 reconstrói o cache antes de rodar.

Todas as operações abaixo ocorrem **dentro** de `PreviousFormatting(doc)`:

### 5.1 Configuração de Página (`ApplyPageSetup`)

- **Margens** superior e inferior: conforme constantes (`TOP_MARGIN_CM`, `BOTTOM_MARGIN_CM`).
- **Margens** esquerda e direita: conforme constantes (`LEFT_MARGIN_CM`, `RIGHT_MARGIN_CM`).
- Distância de cabeçalho e rodapé: conforme constantes.
- Gutter: 0.
- Orientação: retrato (`wdOrientPortrait`).

### 5.2 Limpeza Total de Formatação (`ClearAllFormatting`)

Para **cada parágrafo** do documento (preservando imagens):

- **Fonte:** reset completo → Arial 12, cor automática, sem negrito, sem itálico, sem sublinhado.
- **Parágrafo:** reset completo → alinhamento à esquerda, espaçamento entre linhas 12pt, espaço antes/depois = 0, recuo esquerda/direita/primeira linha = 0.
- **Bordas e sombreamento:** desabilitados.
- **Estilo:** redefinido para "Normal".

### 5.3 Normalização de Quebras de Linha (`ReplaceLineBreaksWithParagraphBreaks`)

- Substitui **todas** as quebras de linha manuais (`^l` / Shift+Enter) por quebras de parágrafo (`^p` / Enter).

### 5.4 Remoção de Quebras de Página (`RemovePageBreaks`)

- Remove **todas** as quebras de página manuais (`^m`).

### 5.5 Remoção de Linhas de Paginação (`RemovePageNumberLines`)

- Remove parágrafos que seguem o padrão `$NUMERO$/$ANO$ – Página N` (linhas de paginação de requerimento, indicação e moção).
- Remove também a linha em branco subsequente, se existir.

### 5.6 Remoção de Parágrafos com Apenas Underlines (`RemoveUnderscoreOnlyParagraphs`)

- Remove parágrafos compostos exclusivamente por caracteres `_` (linhas decorativas).

### 5.7 Limpeza da Estrutura do Documento (`CleanDocumentStructure`)

- Remove **linhas vazias** antes do primeiro parágrafo com texto.
- Remove **espaços em branco** no início das linhas (após quebra de parágrafo e no início absoluto do documento).
- Remove **tabs** no início das linhas.

### 5.8 Remoção de Todas as Tabulações (`RemoveAllTabMarks`)

- Substitui **todas** as tabulações (`^t`) por espaço simples.

### 5.9 Limpeza do Prefixo da Ementa (`RemoveEmentaLeadingLabelPrefix`)

- Remove do parágrafo da ementa os prefixos `"EMENTA:"` ou `"ASSUNTO:"` (case-insensitive), incluindo espaços e o caractere `:` após a palavra.

### 5.10 Limpeza do Sufixo da Ementa (`RemoveEmentaTrailingMunicipioSuffix`)

- Remove do final da ementa a string `", neste município"` (com ou sem acento, com ou sem espaço após a vírgula).
- Se a ementa não terminava com ponto final antes da remoção, insere `.` ao final.

### 5.11 Remoção de Aspas da Ementa (`RemoveEmentaQuotes`)

- Remove aspas envolventes na ementa: aspas duplas retas (`"..."`), aspas duplas tipográficas (`"…"`), aspas simples retas (`'...'`) e aspas simples tipográficas (`'…'`).
- Trata o caso em que o ponto final está após a aspa de fechamento (`"texto".`).

### 5.12 Substituição de DAE por Poder Executivo Municipal em Indicações (`ProcessEmentaIndicacao`)

- **Apenas se** o título começar com "INDICAÇÃO": se a ementa iniciar com `"Indica ao DAE"` ou `"Sugere ao DAE"`, substitui por `"Indica ao Poder Executivo Municipal"`.

### 5.13 Formatação do Título (`FormatDocumentTitle`)

- Encontra o primeiro parágrafo com texto.
- Remove ponto final do título, se existir.
- **Se for proposição** (Indicação, Requerimento ou Moção): normaliza o número/ano no formato `Nº $NUMERO$/$ANO$`.
- Aplica formatação ao título:
  - **Caixa alta** (`AllCaps = True`).
  - **Negrito** (`Bold = True`).
  - **Sublinhado** (`Underline = wdUnderlineSingle`).
  - **Alinhamento:** centralizado.
  - **Recuos:** todos zero.
  - **Espaço depois:** 6pt.

> **Nota:** Esta etapa consolida a formatação do 1º parágrafo (anteriormente etapa separada `FormatFirstParagraph`), que aplicava as mesmas formatações (AllCaps, Bold, Underline, centralizado, recuos zero). A chamada a `FormatFirstParagraph` foi removida do pipeline.

### 5.14 Aplicação de Fonte Padrão (`ApplyStdFont` ou `ApplyStdFontOptimized`)

Para **cada parágrafo** com texto (preservando imagens):

- **Fonte:** Arial 12pt, cor automática.
- **Sublinhado:** removido (exceto nos 3 primeiros parágrafos / título).
- **Negrito:** removido (exceto para parágrafos especiais: "Justificativa" e "Anexo/Anexos").

### 5.15 Formatação de Parágrafos (`ApplyStdParagraphs`)

Para **cada parágrafo** (preservando imagens e parágrafos especiais):

- **Limpeza de espaços múltiplos:** remove espaços duplos, tabs extras, espaços antes/depois de quebras de linha.
- **Espaçamento entre linhas:** múltiplo, conforme `LINE_SPACING`.
- **Recuo direito:** 0.
- **Espaço antes/depois:** 0.
- **Recuos condicionais:**
  - Parágrafos centralizados: recuo esquerda e primeira linha = 0.
  - Parágrafos com recuo esquerda ≥ 5cm: recuo esquerda = 9cm (ementa).
  - Demais parágrafos com recuo < 5cm: recuo esquerda = 0, primeira linha = 2,5cm.
- **Alinhamento:** parágrafos alinhados à esquerda passam para justificado.

### 5.16 Formatação do 2º Parágrafo — Ementa (`FormatSecondParagraph`)

> **Nota:** Marcadores e numeração existentes nos parágrafos **não são editados** em nenhum momento durante o processamento.

- Normaliza quebras de linha antes de processar.
- **Substituições iniciais:**
  - `"Solicita"` → `"Requer"` (mantendo o restante).
  - `"Pede"` → `"Requer"` (mantendo o restante).
  - `"Sugere"` → `"Indica"` (mantendo o restante).
- Remove `", neste município"` do final, se presente.
- **Insere 2 linhas em branco ANTES** do parágrafo.
- **Formatação:** recuo esquerda = 9cm, primeira linha = 0, recuo direito = 0, justificado.
- **Insere 2 linhas em branco DEPOIS** do parágrafo.

### 5.17 Formatação de "CONSIDERANDO" e "ANTE O EXPOSTO" (`FormatConsiderandoParagraphs`)

- Parágrafos que começam com `"considerando"` (seguido de espaço, vírgula, ponto-e-vírgula ou dois-pontos):
  - `"considerando"` → `"CONSIDERANDO"` em **caixa alta** e **negrito**.
- Parágrafos que começam com `"ante o exposto"`:
  - `"ante o exposto"` → `"ANTE O EXPOSTO"` em **caixa alta** e **negrito**.

### 5.18 Substituições de Texto (`ApplyTextReplacements`)

Executa as seguintes substituições globais (Find/Replace):

| # | Busca | Substituição | Sensível a maiúsculas |
|---|---|---|---|
| 1 | 16 variantes de `"d'Oeste"` (aspas retas, acento agudo, grave, tipográficas, maiúsculas/minúsculas) | `"d'Oeste"` | Não |
| 2 | `" ao Setor, "` | `" ao setor competente"` | Sim |
| 3 | `" Setor Competente "` | `" setor competente "` | Não |
| 4 | No 1º parágrafo após a ementa: `"para sugerir"` | `"para indicar"` | Não |
| 5 | No 1º parágrafo após a ementa: normaliza abertura Art. 108 → `"setor competente,"` | Conforme regra | Não |
| 6 | Variantes de `"tapa-buracos"` (com/sem aspas, com/sem hífen) | `"tapa-buracos"` | Não |
| 7 | `"in loco"` com aspas → sem aspas; aplica **itálico** em todas as ocorrências | — | Não |
| 8 | `"Área Pública"` (maiúscula) → `"área pública"` (minúscula) | `"área pública"` | Sim |
| 9 | `"Roçagem"` (maiúscula) → `"roçagem"` (minúscula) | `"roçagem"` | Sim |
| 10 | `"retorne à esta Casa de Leis com as seguintes respostas"` | `"retorne à esta Casa de Leis com as seguintes informações"` | Não |
| 11 | `" Jd "` | `" Jd. "` | Sim |
| 12 | `" aos nº "` / `" aos n° "` | `" ao nº "` / `" ao n° "` | Sim |
| 13 | `" nos nº "` / `" nos n° "` | `" no nº "` / `" no n° "` | Sim |
| 14 | `"Nº"` / `"N°"` / `"nº"` → `"n°"` (exceto no título) | `"n°"` | Sim |
| 15 | Parágrafos com `"tikinho tk"` (variantes com aspas) | `"TIKINHO TK"` | Não |
| 16 | Garante espaço não separável após `nº`/`n°` antes de algarismos | — | — |
| 17 | Substitui **todos** os espaços não separáveis por espaços comuns (exceto após `nº`/`n°` antes de algarismos) | — | — |

### 5.19 Capitalização do Início dos Parágrafos (`CapitalizeFirstLetterOfParagraphs`)

- Para cada parágrafo: se a primeira letra do texto for minúscula, converte para maiúscula.
- Exemplo: `"nos termos..."` → `"Nos termos..."`.

### 5.20 Remoção de Marca d'Água (`RemoveWatermark`)

- Percorre cabeçalhos e rodapés de todas as seções.
- Remove shapes (imagens ou efeitos de texto) cujo nome ou texto alternativo contenha `"Watermark"`.

### 5.21 Inserção de Carimbo no Cabeçalho (`InsertHeaderstamp`)

- Para cada seção do documento:
  - Limpa o cabeçalho existente.
  - Insere a imagem `stamp.png` (logo da Câmara) centralizada horizontalmente, com largura e altura conforme constantes.
  - Fonte do cabeçalho: Arial 12.

### 5.22 Limpeza de Espaços Múltiplos (`CleanMultipleSpaces`)

- Remove espaços duplos (iterativamente até sobrar um).
- Remove espaços antes/depois de quebras de parágrafo.
- Converte tabs múltiplos em um único tab, depois tabs em espaços.
- Garante espaço após a palavra `"CONSIDERANDO"` (corrige caso a próxima palavra tenha ficado grudada).

### 5.23 Controle de Linhas em Branco Sequenciais (`LimitSequentialEmptyLines`)

- Normaliza quebras de linha (`^l` → `^p`).
- Remove excesso de linhas em branco consecutivas: máximo permitido = **1 linha vazia** entre parágrafos.
- Usa múltiplas passadas de Find/Replace (`^p^p^p` → `^p^p`, etc.) com fallback por loop.

### 5.24 Substituição do Parágrafo de Data (Plenário) (`ReplacePlenarioDateParagraph`)

- Localiza parágrafos com até 80 caracteres que contenham 2+ termos como `"Plenário"`, `"Dr. Tancredo Neves"`, `"Palacio 15 de Junho"`, nomes de meses, etc.
- **Substitui** o texto por: `Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$.`
- **Formata:** centralizado, sem recuos, sem espaço antes/depois.

### 5.25 Configuração de Visualização (`ConfigureDocumentView`)

- Define o zoom em **140%**.
- Mantém as demais configurações de visualização preservadas.

### 5.26 Inserção de Rodapé (`InsertFooterStamp`)

- Para cada seção:
  - Limpa o rodapé existente.
  - **À esquerda:** iniciais do usuário (Arial 6pt, cinza).
  - **Centro:** `"Página X de Y"` (Arial 9pt), usando campos `PAGE` e `NUMPAGES`.

### 5.27 Negrito em Parágrafos Especiais (`ApplyBoldToSpecialParagraphs`)

- **"Justificativa":** aplica negrito, Arial 12, centralizado, sem recuos.
- **"Anexo"/"Anexos":** aplica negrito, Arial 12, alinhado à esquerda, sem recuos.

### 5.28 Substituição de "Vereadora," por "Senhora Vereadora," (`SubstituiVereadoraPorSenhoraVereadora`)

- Se um parágrafo contém apenas `"Vereadora,"` e o anterior contém apenas `"Senhores Vereadores,"`, substitui por `"Senhora Vereadora,"`.

### 5.29 Formatação de Parágrafos "Vereador" (`FormatVereadorParagraphs`)

Para cada parágrafo que contenha apenas "Vereador" (com ou sem hífens/travessões):

- **Remove linhas em branco** imediatamente acima.
- **Texto normalizado:** `"Vereador"` (sem hífens, sem travessões).
- **Fonte:** normal (sem negrito, sem itálico, sem sublinhado, sem caixa alta), Arial 12.
- **Alinhamento:** centralizado.
- **Recuos:** todos zero.
- **Formata a 1ª linha acima:** caixa alta, negrito, centralizada, sem recuos.
- **Formata a 2ª linha acima:** negrito, centralizada, recuo esquerda = 0.
- **Formata a linha abaixo:** centralizada, sem recuos.

### 5.30 Formatação de "Diante do Exposto" (`FormatDianteDoExposto`)

- Parágrafos que começam com `"Diante do exposto"`: aplica **negrito** e **caixa alta** nos primeiros 17 caracteres.

### 5.31 Formatação de "Requeiro" (`FormatRequeiroParagraphs`)

- Parágrafos que começam com `"Requeiro"`: aplica **negrito** e **caixa alta** nas primeiras 8 caracteres (a palavra "REQUEIRO").

### 5.32 Formatação de "Por Todas as Razões" (`FormatPorTodasRazoesParagraphs`)

- Parágrafos que começam com `"Por todas as razões aqui expostas"` (33 caracteres) ou `"Pelas razões aqui expostas"` (28 caracteres): aplica **negrito** na frase.

### 5.33 Remoção de Realces e Bordas (`RemoveAllHighlightsAndBorders`)

- Remove **realce** (highlight) de todo o documento.
- Remove **bordas** de todos os parágrafos.

### 5.34 Remoção de Páginas Vazias no Final (`RemoveEmptyPagesAtEnd`)

- Verifica se a(s) última(s) página(s) do documento está(ão) vazia(s).
- Remove parágrafos vazios do final até que a última página tenha conteúdo.

### 5.35 Formatação Final Universal (`ApplyUniversalFinalFormatting`)

- Para **todo** o documento (em lote):
  - **Fonte:** Arial 12.
  - **Espaçamento entre linhas:** simples (`wdLineSpaceSingle`).
  - **Espaço antes/depois:** 0.
  - **Hifenação automática:** desativada.
- Reaplica formatação de Vereador (se necessário).

### 5.36 Remoção de Dois-Pontos da Justificativa (`RemoveJustificativaColon`)

- `"Justificativa:"` ou `"Justificação:"` → `"Justificativa"` (remove dois-pontos).
- Normaliza caixa: `"JUSTIFICATIVA"` ou `"justificativa"` → `"Justificativa"`.

### 5.37 Espaçamento Especial (`AddSpecialElementsSpacing`)

- **Garante 1 linha em branco** entre o Título e a Ementa (se não existir).
- **Zera espaço antes/depois** da Ementa.
- **Zera espaço antes/depois** do Título da Justificativa.
- **Zera espaço antes/depois** da Data.

### 5.38 Ajuste Final de Recuos para Vereador (`FixHyphenatedVereadorParagraphIndents`)

- Parágrafos com `"- Vereador -"`, `"- Vereadora -"` ou variantes: zera todos os recuos (esquerda, direita, primeira linha).

---

## 6. Pós-Pipeline (após as duas passagens)

### 6.1 Restauração de Imagens (`RestoreAllImages`)

- Verifica se as imagens (inline e flutuantes) ainda estão nas posições e dimensões originais.
- Corrige dimensões e posições de shapes flutuantes que tenham sido alteradas.

### 6.2 Remoção de Linhas em Branco Extras (`RemoverLinhasEmBrancoExtras`)

- **Espaçamento simples** em todos os parágrafos (entre linhas = 12pt, espaço antes/depois = 0).
- Remove parágrafos vazios consecutivos (mantém no máximo 1).
- Remove parágrafos que contenham apenas um espaço (`" "`).
- **Substituições de texto adicionais:**
  - `"por intermedio do Setor,"` → `"por intermédio do Setor competente,"`.
  - `"Indica ao Poder Executivo Municipal efetue"` → `"Indica ao Poder Executivo Municipal que efetue"`.
  - `"Indica ao Poder Executivo Municipal e aos órgãos competentes"` → `"Indica ao Poder Executivo Municipal"`.
  - `"Fomos procurados por municipes, solicitando..."` → normaliza pontuação e acentos.
- **Reatualiza índices estruturais** após deleções físicas (reconstrói cache se necessário).
- **Ajustes por parágrafo:**
  - Parágrafo do Plenário: sem espaço antes/depois.
  - Parágrafos de assinatura (Vereador, Vereadora, Presidente, Prefeito) após a Justificativa: centraliza, zera recuos; parágrafo anterior recebe negrito e centralização; parágrafo seguinte centralizado.

### 6.3 Garantia de Linha em Branco após CONSIDERANDO (`EnsureConsideringBlankLines`)

- Para cada parágrafo que começa com `"CONSIDERANDO"`: se não houver parágrafo vazio logo abaixo, **insere um**.

### 6.4 Restauração de Listas (`RestoreListFormats`)

- Para cada parágrafo que possuía formatação de lista antes do processamento:
  - Remove formatação de lista existente.
  - Reaplica conforme o tipo original (marcadores, numeração simples, numeração de tópicos, etc.).
  - Restaura o nível da lista.

### 6.5 Formatação de Recuos com Marcadores (`FormatBulletedParagraphsIndent`)

- Parágrafos que começam com marcadores (`*`, `-`, `>`, `+`, `~`):
  - Aplica recuo esquerda de **36 pontos** (≈1,27cm), primeira linha = 0.

### 6.6 Formatação de Recuos de Imagens (`FormatImageParagraphsIndents`)

- Parágrafos com imagens inline:
  - **Recuo esquerda:** 0.
  - **Primeira linha:** 0.
  - **Alinhamento:** centralizado.

### 6.7 Centralização de Imagem após Plenário (`CenterImageAfterPlenario`)

- Localiza o parágrafo `"Plenário Dr. Tancredo Neves"`.
- Nas **linhas 5 a 7** após o Plenário: se houver imagem, **centraliza**.

### 6.8 Restauração de Parágrafos Centralizados (`RestoreCenteredParagraphs`)

- Para cada parágrafo que estava centralizado antes do processamento (registrado no backup):
  - Reaplica **centralização**.
  - Zera **recuo esquerda** e **primeira linha**.

### 6.9 Remoção de Numeração de Parágrafos em Branco (`RemoveNumberingFromBlankParagraphs`)

- Parágrafos vazios que tenham formatação de lista: **remove a numeração/marcador**.

### 6.10 Garantia Final de Fonte

- Reaplica **Arial 12** em **todo** o documento (range completo) como garantia final, pois operações Find/Replace podem ter deixado trechos com fonte do estilo Normal (Calibri).

### 6.11 Restauração de Configurações de Visualização (`RestoreViewSettings`)

- Restaura **todas** as configurações de visualização originais (tipo de vista, régua, marcadores, etc.).
- **Exceção:** o zoom é mantido em **140%**.

---

## 7. Finalização

| Etapa | Descrição |
|---|---|
| 7.1 | Registra log de sucesso e snapshot de contexto final. |
| 7.2 | Calcula tempo de execução e exibe na barra de status. |
| 7.3 | **Fecha o UndoRecord** (todas as edições ficam disponíveis para desfazer em uma única ação Ctrl+Z). |
| 7.4 | Limpa cache de parágrafos e variáveis de proteção de imagens. |
| 7.5 | Restaura estado da aplicação. |
| 7.6 | Finaliza o sistema de logging. |
| 7.7 | Posiciona cursor no início do documento. |
| 7.8 | **Salva o documento** automaticamente. |

---

## Observações Importantes

1. **Proteção de imagens:** todas as formatações de fonte e parágrafo são aplicadas preservando imagens inline e shapes flutuantes. Quando um parágrafo contém imagens, a formatação é aplicada caractere por caractere para evitar danos.

2. **Duas passagens:** o pipeline é executado duas vezes se a primeira passagem fizer alterações. A segunda passagem reconstrói o cache para garantir que os índices estão atualizados após possíveis remoções/adições de parágrafos.

3. **Parágrafos especiais:** os parágrafos identificados como "Justificativa", "Anexo"/"Anexos" e "Vereador"/"Vereadora" recebem tratamento diferenciado em diversas etapas (negrito preservado, alinhamento específico, recuos zero, etc.).

4. **Desfazer (Ctrl+Z):** todas as edições são agrupadas em um único `UndoRecord`, permitindo desfazer toda a padronização com um único Ctrl+Z.

5. **Backup automático:** antes de qualquer modificação, o documento original é salvo como backup. Os backups são mantidos com limite configurável.