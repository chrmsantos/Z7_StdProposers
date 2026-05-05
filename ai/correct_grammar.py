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
        "Execute o arquivo install_requirements.bat da pasta ai "
        "e tente novamente."
    )
    root.destroy()
    sys.exit(1)

import google.genai as genai

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
    
    try:
        try:
            word = win32com.client.GetObject(Class="Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
        word.StatusBar = False # Limpa o status bar de carregamento do VBA
        LOGGER.info("Connected to running Word instance")
    except Exception as e:
        log_exception(LOGGER, "Failed to connect to Word", e)
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        messagebox.showerror("Z7 StdProposers - Erro", f"Erro ao conectar ao Word:\n{e}")
        return

    try:
        selection = word.Selection
        text = selection.Text.strip()
    except Exception as e:
        LOGGER.error("Erro ao obter selection: %s", str(e))
        return

    if not text or len(text) < 2:
        LOGGER.info("Selection too short; nothing to process")
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # 1. Privacy Warning
    if not messagebox.askokcancel(
        "Aviso de Privacidade - Z7 StdProposers",
        "O trecho selecionado será enviado para a API do Google Gemini para revisão gramatical.\n\n"
        "Certifique-se de que não há dados sigilosos e que o uso está de acordo com as diretrizes do seu órgão.\n\n"
        "Deseja continuar?",
        parent=root
    ):
        LOGGER.info("User cancelled at privacy warning")
        root.destroy()
        return

    api_key = get_api_key()
    if not api_key:
        LOGGER.warning("Aborting grammar flow because API key is unavailable")
        root.destroy()
        return

    client = genai.Client(api_key=api_key)
    LOGGER.info("Gemini client configured")

    # Mostrar carregamento
    loading = tk.Toplevel(root)
    loading.title("Z7 StdProposers")
    loading.geometry("300x100")
    loading.attributes('-topmost', True)
    tk.Label(loading, text="Consultando Google Gemini...\nPor favor, aguarde.", font=("Segoe UI", 10)).pack(expand=True)
    loading.update()

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
                pass
                
    if not base_prompt:
        base_prompt = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

    prompt = f"{base_prompt}\n\nTexto original:\n{text}\n"

    model_name = 'gemini-2.0-flash'
    if user_profile:
        try:
            model_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'selected_model.txt'
            if model_file.exists():
                with open(model_file, 'r', encoding='utf-8') as f:
                    saved_model = f.read().strip()
                    if saved_model:
                        model_name = saved_model
        except Exception:
            pass

    try:
        response = client.models.generate_content(model=model_name, contents=prompt)
        corrected_text = response.text.strip()
        LOGGER.info("Gemini response received with model: %s", model_name)
    except Exception as e:
        loading.destroy()
        log_exception(LOGGER, "Gemini API request failed", e)
        error_msg = str(e).lower()
        if "403" in error_msg or "401" in error_msg or "invalid api key" in error_msg or "api_key" in error_msg:
            try:
                key_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'gemini.key'
                if key_file.exists():
                    key_file.unlink()
            except Exception:
                pass
            messagebox.showerror("Z7 StdProposers - Erro na API", "Sua chave da API é inválida ou expirou. O arquivo de chave foi deletado.\nTente rodar a macro novamente para inserir uma nova chave.", parent=root)
        else:
            messagebox.showerror("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}", parent=root)
        root.destroy()
        return

    loading.destroy()

    if not corrected_text:
        messagebox.showwarning("Z7 StdProposers - Aviso", "A API do Gemini retornou um texto vazio.", parent=root)
        root.destroy()
        return

    def apply_text():
        selection.Text = corrected_text
        root.destroy()

    def copy_text():
        root.clipboard_clear()
        root.clipboard_append(corrected_text)
        root.update()
        messagebox.showinfo("Copiado", "Texto copiado para a Área de Transferência!", parent=result_window)
        root.destroy()

    def cancel_action():
        root.destroy()

    result_window = tk.Toplevel(root)
    result_window.title("Revisão Concluída")
    result_window.geometry("700x500")
    result_window.attributes('-topmost', True)
    result_window.protocol("WM_DELETE_WINDOW", cancel_action)
    
    tk.Label(result_window, text="Texto Sugerido:", font=("Segoe UI", 10, "bold")).pack(pady=(10, 5))
    text_area = tk.Text(result_window, wrap=tk.WORD, height=15, font=("Segoe UI", 11))
    text_area.pack(expand=True, fill=tk.BOTH, padx=15, pady=5)
    text_area.insert(tk.END, corrected_text)
    
    btn_frame = tk.Frame(result_window)
    btn_frame.pack(fill=tk.X, pady=15)
    
    tk.Button(btn_frame, text="Substituir no Word (Atenção: remove formatação rica)", 
              command=apply_text, font=("Segoe UI", 9), fg="#b91c1c").pack(side=tk.LEFT, padx=15)
              
    tk.Button(btn_frame, text="Copiar para Área de Transferência", 
              command=copy_text, font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=15)

    root.mainloop()

if __name__ == "__main__":
    main()

