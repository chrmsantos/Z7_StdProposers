# Corretor Gramatical com Gemini IA para Word

Este projeto integra a API do Google Gemini ao Microsoft Word, permitindo a correção gramatical de textos selecionados no contexto de proposituras legislativas no Brasil. O modelo faz correções pontuais mantendo o jargão jurídico e a estrutura original, atuando com o mínimo possível de alterações.

## Arquivos do Projeto

- `correct_grammar.py`: Script Python responsável por conectar-se ao Word aberto, ler o texto selecionado, enviar para a API do Gemini e substituir a seleção pela resposta corrigida.
- `WordMacro.bas`: Módulo VBA contendo a macro que deverá ser importada no seu Microsoft Word.
- `.env.example`: Modelo do arquivo que armazenará a sua chave da API de forma segura.
- `install_requirements.bat`: Arquivo executável para instalar as dependências do Python com apenas um duplo-clique.

## Passos para Instalação

### 1. Preparando o Ambiente Python
1. Certifique-se de ter o [Python](https://www.python.org/downloads/) instalado no seu computador e adicionado ao "PATH" do Windows.
2. Dê um duplo-clique no arquivo `install_requirements.bat` e aguarde a finalização para que ele instale as bibliotecas (`google-generativeai`, `pywin32` e `python-dotenv`).

### 2. Configurando a Chave da API (Gemini)
1. Acesse o [Google AI Studio](https://aistudio.google.com/app/apikey) e crie/obtenha uma Chave de API (API Key).
2. Na pasta do projeto (`c:\Users\csantos\AppData\Local\Z7\Apps\Z7_GrammarProp`), renomeie o arquivo `.env.example` para `.env` (remova o `.example`).
3. Abra o arquivo `.env` (pode ser com o Bloco de Notas) e substitua `sua_chave_api_aqui` pela chave que você copiou do site do Google. Salve o arquivo.

### 3. Configurando a Macro no Microsoft Word
1. Abra o Microsoft Word e crie ou abra um documento qualquer.
2. Pressione as teclas `ALT + F11` para abrir o Editor do Visual Basic (VBA).
3. No menu superior, vá em **Arquivo > Importar Arquivo...** (ou *File > Import File...*).
4. Navegue até a pasta do projeto (`c:\Users\csantos\AppData\Local\Z7\Apps\Z7_GrammarProp`) e selecione o arquivo `WordMacro.bas`. 
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
