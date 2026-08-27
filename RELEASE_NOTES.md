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
- import_ui_to_word.exe — Importador de UI customizada

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