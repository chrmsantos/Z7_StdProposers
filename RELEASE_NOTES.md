# Release Notes

## v5.0.1-beta — 2026-05-04 — Observability & Testing Sprint

### Resumo

Esta versão introduz um sistema de observabilidade completo para os componentes Python e VBA,
novos conjuntos de testes automatizados e a correção de defeitos encontrados durante a sprint.

---

### Novidades

#### Python — Logging Estruturado (`z7_logging.py`)

- Novo módulo compartilhado `Z7_GrammarProp/z7_logging.py` com:
  - `configure_component_logger(component, level)` — cria um logger com arquivo timestamped em `%USERPROFILE%\AppData\Local\Z7\Tmp\StdProposers\logs\`
  - `log_exception(logger, context, exc)` — registra exceções com traceback completo
  - Formato de linha: `YYYY-MM-DD HH:MM:SS.mmm | LEVEL | component | mensagem`
- `correct_grammar.py`, `config_prompt.py` e `chat_ia.py` instrumentados com LOGGER em todos os pontos críticos de fluxo
- Erros de importação de dependências opcionais (`win32com`, `win32crypt`) capturados com diálogo GUI amigável

#### VBA — Logging Aprimorado (`Mod1Infrastructure.bas` / `Mod3Pipeline.bas`)

- `currentLogSessionId` e `currentOperationId` — identificadores únicos por sessão e operação
- Formato de linha atualizado inclui `[op=ID]` para rastreabilidade cruzada
- `LOG_BUFFER_FLUSH_SECONDS = 5` — constante configurável para intervalo de flush do buffer
- Nova `Public Sub LogContextSnapshot(doc As Document, contextName As String)` — registra estado do documento (parágrafos, páginas, caracteres, proteção, somente-leitura, salvo)
- `LogContextSnapshot` chamado automaticamente em INICIO, FIM e ERRO_CRITICO em `Mod4Main.bas`

#### Testes Automatizados

- **`tests/VBA-Logging.Tests.ps1`** — valida declarações de funções de log, globals de sessão/operação, formato `[op=]`, chamadas de LogContextSnapshot e cobertura mínima de logs
- **`tests/Python.Tests.ps1`** — valida existência de `z7_logging.py`, imports em todos os scripts e executa suite unittest
- **`tests/python/test_z7_logging.py`** — testes unitários para `build_log_path`, `configure_component_logger` (criação de arquivo) e `get_logs_dir`
- **`tests/Run-Tests.ps1`** — adicionados suites `Python` e `VBA-Logging` ao seletor `-TestSuite`

---

### Correções

| Componente | Problema | Solução |
|---|---|---|
| `correct_grammar.py` | `ModuleNotFoundError: No module named 'win32com'` na inicialização | Try/except com diálogo GUI e fallback gracioso |
| `build_exe.ps1` | Falhas falsas por saída em stderr do PyInstaller | Migrado para `Start-Process ... -Wait -PassThru` com verificação de exit code |
| `tests/VBA.Tests.ps1` | Referências a `Modulo1.bas` (arquivo deletado) | Reescrito para arquitetura modular (Mod1–Mod4) |
| `tests/Encoding.Tests.ps1` | Política ASCII estrita falhava em arquivos UTF-8 | Relaxado para UTF-8 válido |
| `tests/VBA-IdentifierFunctions.Tests.ps1` | Buscava funções em arquivo único deletado | Reescrito para buscar em todos os `.bas` de `source/main/` |
| `tests/All.Tests.ps1` | Verificava arquivos de documentação inexistentes | Corrigida lista para arquivos reais do repositório |
| `tests/python/test_z7_logging.py` | Falha no cleanup de `TemporaryDirectory` no Windows | FileHandler fechado explicitamente antes do `with` sair |

---

### Documentação Atualizada

- **`README.md`** — arquitetura de 4 módulos, comandos de teste por suite
- **`Z7_GrammarProp/README.md`** — fluxo DPAPI (sem `.env`), invocação `pyw -3`, `z7_logging.py`, instruções de build
- **`AI_CONTEXT.md`** — reescrito com topologia de módulos atual, arquitetura de logging/testes, regras de manutenção

---

### Executáveis Recompilados

| Arquivo | Tamanho |
|---|---|
| `Z7_GrammarProp/correct_grammar.exe` | ~11.5 MB |
| `Z7_GrammarProp/config_prompt.exe` | ~11.5 MB |
| `Z7_GrammarProp/chat_ia.exe` | ~11.5 MB |

Compilados com PyInstaller 6.20.0 / Python 3.14.4.

---

### Como Atualizar

1. Substitua os arquivos `.exe` em `Z7_GrammarProp/`
2. O VBA detecta a nova versão automaticamente via `GetLocalVersion()` e exibe aviso de atualização disponível

---

### Versão Anterior

`v5.0.0` — release estável anterior
