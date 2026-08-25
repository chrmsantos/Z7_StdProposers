---
paths:
  - "VERSION"
  - ".git/**"
  - ".gitignore"
  - ".github/**"
---

# Regras de Git e Versionamento

## Fluxo de Versionamento

1. **Leia** `VERSION` para saber a versão atual.
2. **Incremente** conforme SemVer (MAJOR.MINOR.PATCH):
   - **MAJOR**: Reestruturação de módulos, breaking changes na API VBA.
   - **MINOR**: Novos módulos VBA, novas features no Python, novos testes.
   - **PATCH**: Correções de bugs, melhorias de encoding, ajustes de logging.
3. **Atualize** `VERSION` E `Mod_01_Infrastructure.bas` (`Z7_STDPROPOSERS_VERSION`).
4. **Commit** com a mensagem: `v<VERSION>: <descrição curta>`

## Mensagens de Commit

Padrão:
```
v<MAJOR.MINOR.PATCH>: <resumo em português>

- <item 1>
- <item 2>
- <item 3>
```

Exemplo:
```
v8.10.1: preservacao robusta de formatacao na revisao de texto

- SubstituirTextoPreservandoFormatacao salva/restaura Borders, Shading, KeepWithNext
- Protecao de marcas de paragrafo durante substituicao de texto
- Side-by-side prompt editor (Corretor de Propositura + Chat IA)
- CLSID ROT fix para deteccao de Word moderno
```

## Branches

- `main`: branch principal. Código estável, releases são feitas daqui.
- `develop`: desenvolvimento ativo (se existir).
- Features branches: `feature/<nome>` — merge em `develop` ou `main`.

## Tags

- Tags seguem o padrão `v<VERSION>` (ex: `v8.10.1`).
- Toda release no GitHub DEVE ter tag correspondente.
- Criar tag: `git tag -a v<VERSION> -m "Release v<VERSION>"`

## Gitignore

- `ai/build/` — NÃO incluir no gitignore (cache Analysis-00.toc evita bug Python 3.14)
- `ai/dist/` — limpo após build, não commitado
- `dist/` — artefatos de release (zip + exe), commitados na release
- `*.pyc`, `__pycache__/` — ignorados

## Antes de Push

Checklist:
- [ ] `VERSION` atualizado
- [ ] `Mod_01_Infrastructure.bas` com `Z7_STDPROPOSERS_VERSION` coincidente
- [ ] `python scripts/fix_bas_encoding.py` executado
- [ ] `Run-Tests.ps1 -TestSuite All -NoProgress` — TODOS passando
- [ ] Mensagem de commit no padrão