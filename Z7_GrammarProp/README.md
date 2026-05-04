# Corretor Gramatical com Gemini IA para Word

Este projeto integra a API do Google Gemini ao Microsoft Word, permitindo a correção gramatical de textos selecionados no contexto de proposituras legislativas no Brasil. O modelo faz correções pontuais mantendo o jargão jurídico e a estrutura original, atuando com o mínimo possível de alterações.

## Arquivos do Projeto

- `correct_grammar.py`: fluxo principal de correção gramatical para seleção ativa no Word.
- `config_prompt.py`: interface para editar o prompt-base salvo localmente.
- `chat_ia.py`: interface de chat com contexto do documento ativo no Word.
- `z7_logging.py`: logger compartilhado entre os scripts Python.
- `WordMacro.bas`: macro VBA para disparar a correção via Python (`pyw -3`).
- `install_requirements.bat`: instalador de dependências com validação de `pywin32`.
- `build_exe.ps1`: build dos executáveis com PyInstaller.

## Passos para Instalação

### 1. Preparando o Ambiente Python

1. Certifique-se de ter o [Python](https://www.python.org/downloads/) instalado no seu computador e adicionado ao "PATH" do Windows.
2. Dê um duplo-clique no arquivo `install_requirements.bat` e aguarde a finalização para que ele instale as bibliotecas (`google-generativeai`, `pywin32` e `python-dotenv`).

### 2. Configurando a Chave da API (Gemini)

1. Acesse o [Google AI Studio](https://aistudio.google.com/app/apikey) e crie/obtenha uma Chave de API (API Key).
2. Execute a macro de correção pela primeira vez: o sistema solicitará a chave via interface gráfica.
3. A chave é criptografada com DPAPI do Windows e armazenada em:
    - `%LOCALAPPDATA%\Z7\Tmp\StdProposers\gemini.key`

Nenhum `.env` é necessário no fluxo atual.

### 3. Configurando a Macro no Microsoft Word

1. Abra o Microsoft Word e crie ou abra um documento qualquer.
2. Pressione as teclas `ALT + F11` para abrir o Editor do Visual Basic (VBA).
3. No menu superior, vá em **Arquivo > Importar Arquivo...** (ou *File > Import File...*).
4. Navegue até a pasta do projeto (`c:\Users\csantos\AppData\Local\Z7\Apps\Z7_StdProposers\Z7_GrammarProp`) e selecione o arquivo `WordMacro.bas`.
5. Feche o Editor do Visual Basic (pode fechar no X vermelho).

### 4. Adicionando um Botão à Interface (Faixa de Opções) do Word

1. No Word, clique na guia **Arquivo** (File) e depois em **Opções** (Options) lá embaixo.
2. Na janela que abrir, vá em **Personalizar Faixa de Opções** (Customize Ribbon).
3. Do lado direito, encontre a guia onde quer que o botão apareça (por exemplo, na guia *Revisão* ou na *Página Inicial*) e clique em **Novo Grupo** (New Group). Você pode renomear esse grupo para algo como "IA".
4. Do lado esquerdo, no menu suspenso *Escolher comandos em:*, selecione **Macros**.
5. Na lista, encontre a macro chamada `Normal.MacroGeminiGrammar.CorrigirGramaticaComGemini` (ou com um nome parecido, onde você a salvou).
6. Selecione a macro, certifique-se de que o "Novo Grupo" que você acabou de criar está selecionado do lado direito e clique em **Adicionar >>** (Add).
7. Se quiser, selecione a macro adicionada do lado direito e clique em **Renomear...** para escolher um ícone bonitinho (como uma varinha ou check) e dar um nome mais amigável, como "Corrigir com Gemini".
8. Clique em **OK**.

## Como Usar

1. Selecione um trecho de texto no seu documento do Word que deseja revisar.
2. Clique no botão "Corrigir com Gemini" que você acabou de adicionar na Faixa de Opções (ou rode a macro diretamente em *Exibir > Macros*).
3. O ponteiro do mouse virará um ícone de carregamento e, instantes depois, o texto que você selecionou será substituído pela versão gramaticalmente corrigida pelo Gemini!

> **Observação técnica:** A macro executa o Python em segundo plano usando `pythonw`. Portanto, não aparecerá nenhuma tela preta (console) enquanto a requisição estiver sendo feita.

## Logs e Diagnóstico

Os scripts Python escrevem logs estruturados (UTF-8) em:

- `%LOCALAPPDATA%\Z7\Tmp\StdProposers\logs`

Os logs incluem eventos de inicialização, integração com Word, chamadas ao Gemini e stack trace em caso de falha.

## Build dos Executáveis

Para recompilar os executáveis (`correct_grammar.exe`, `config_prompt.exe`, `chat_ia.exe`):

1. Abra PowerShell na pasta `Z7_GrammarProp`.
2. Execute: `./build_exe.ps1`
3. O script compila com PyInstaller, move os `.exe` para a raiz da pasta e limpa artefatos temporários (`build`, `dist`, `*.spec`).
