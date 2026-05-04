import os
import sys
import tkinter as tk
from tkinter import messagebox, simpledialog
from pathlib import Path
from z7_logging import configure_component_logger, log_exception

LOGGER = configure_component_logger("correct_grammar")

try:
    import win32com.client
    import win32crypt
except ModuleNotFoundError as e:
    LOGGER.error("Missing Python dependency: %s", e.name)
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    messagebox.showerror(
        "Z7 StdProposers - Dependencia ausente",
        "Nao foi possivel iniciar o corretor porque falta a dependencia Python: "
        f"{e.name}.\n\n"
        "Execute o arquivo install_requirements.bat da pasta Z7_GrammarProp "
        "e tente novamente."
    )
    root.destroy()
    sys.exit(1)

import google.generativeai as genai

def get_api_key():
    LOGGER.info("Loading Gemini API key")
    # Define o caminho do arquivo de chave
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        print("Erro: Variavel USERPROFILE nao encontrada.")
        LOGGER.error("USERPROFILE env var not found")
        return None
        
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_file = key_dir / 'gemini.key'
    
    # Se o arquivo ja existir, tenta ler e descriptografar
    if key_file.exists():
        try:
            with open(key_file, 'rb') as f:
                encrypted_key = f.read()
            # CryptUnprotectData retorna uma tupla (descricao, dados_descriptografados)
            _, decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)
            LOGGER.info("API key loaded from encrypted key file")
            return decrypted_key.decode('utf-8')
        except Exception as e:
            print(f"Erro ao descriptografar a chave: {e}")
            log_exception(LOGGER, "Failed to decrypt API key", e)
            # Se falhar, vamos pedir novamente
            pass
            
    # Se nao existir ou falhar, pede a chave via Tkinter
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal
    
    # Mantem a janela de dialogo no topo
    root.attributes('-topmost', True)
    
    api_key = simpledialog.askstring(
        "Z7 StdProposers - Primeira Inicializacao", 
        "Insira sua chave da API do Google Gemini:\n(Ela sera criptografada e salva localmente)",
        parent=root
    )
    
    root.destroy()
    
    if not api_key or not api_key.strip():
        print("Chave nao fornecida pelo usuario.")
        LOGGER.warning("User did not provide API key")
        return None
        
    api_key = api_key.strip()
    
    # Criptografa a chave usando DPAPI (seguranca atrelada ao usuario atual do Windows)
    try:
        encrypted_key = win32crypt.CryptProtectData(api_key.encode('utf-8'), 'Z7_Gemini_Key', None, None, None, 0)
        
        # Cria os diretorios se nao existirem
        key_dir.mkdir(parents=True, exist_ok=True)
        
        # Salva o arquivo
        with open(key_file, 'wb') as f:
            f.write(encrypted_key)
        LOGGER.info("API key encrypted and persisted")
    except Exception as e:
        print(f"Erro ao criptografar ou salvar a chave: {e}")
        log_exception(LOGGER, "Failed to save encrypted API key", e)
        # Mesmo se falhar ao salvar, podemos retornar a chave em memoria para uso atual
        
    return api_key

def main():
    LOGGER.info("Starting grammar correction flow")
    api_key = get_api_key()
    if not api_key:
        LOGGER.warning("Aborting grammar flow because API key is unavailable")
        return

    # Configura a API do Gemini
    genai.configure(api_key=api_key)
    
    # Inicializa o modelo (gemini-1.5-pro e otimo para tarefas complexas de raciocinio, 
    # ou gemini-1.5-flash para respostas mais rapidas)
    model = genai.GenerativeModel('gemini-3.1-pro-preview')
    LOGGER.info("Gemini model configured")

    try:
        # Conecta ao aplicativo Word que já está em execução
        word = win32com.client.Dispatch("Word.Application")
        LOGGER.info("Connected to running Word instance")
    except Exception as e:
        log_exception(LOGGER, "Failed to connect to Word", e)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Z7 StdProposers - Erro", f"Erro ao conectar ao Word:\n{e}")
        return

    # Pega o texto selecionado pelo usuário
    selection = word.Selection
    text = selection.Text.strip()

    # Se nada estiver selecionado ou houver apenas espaços em branco, encerra
    if not text or len(text) < 2:
        LOGGER.info("Selection too short; nothing to process")
        return

    # Lê o prompt configurado do arquivo, se existir
    user_profile = os.environ.get('USERPROFILE')
    base_prompt = ""
    if user_profile:
        prompt_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'gemini_prompt.txt'
        if prompt_file.exists():
            try:
                with open(prompt_file, 'r', encoding='utf-8') as f:
                    base_prompt = f.read().strip()
                LOGGER.info("Custom prompt loaded from file")
            except Exception:
                LOGGER.warning("Unable to read custom prompt; falling back to default")
                pass
                
    if not base_prompt:
        base_prompt = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

    prompt = f"{base_prompt}\n\nTexto original:\n{text}\n"

    try:
        # Faz a requisição para a API do Gemini
        response = model.generate_content(prompt)
        corrected_text = response.text.strip()
        LOGGER.info("Gemini response received")
        
        if corrected_text:
            # Preserva a formatação substituindo o texto da seleção
            selection.Text = corrected_text
            LOGGER.info("Selection updated with corrected content")
        else:
            LOGGER.warning("Gemini returned empty response")
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning("Z7 StdProposers - Aviso", "A API do Gemini retornou um texto vazio.")
            
    except Exception as e:
        log_exception(LOGGER, "Gemini API request failed", e)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}")

if __name__ == "__main__":
    main()
