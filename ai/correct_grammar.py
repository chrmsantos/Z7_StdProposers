import os
import sys
import tkinter as tk
from pathlib import Path
import z7_theme
from z7_logging import configure_component_logger, log_exception

LOGGER = configure_component_logger("correct_grammar")

def get_api_key() -> str | None:
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
            import win32crypt
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
    
    api_key = z7_theme.ask_string(
        "Z7 StdProposers - Primeira Inicializacao", 
        "Insira sua chave da API do Google Gemini:\n(Ela sera criptografada e salva localmente)",
        parent=root,
        show="*"
    )
    
    root.destroy()
    
    if not api_key or not api_key.strip():
        print("Chave nao fornecida pelo usuario.")
        LOGGER.warning("User did not provide API key")
        return None
        
    api_key = api_key.strip()
    
    # Criptografa a chave usando DPAPI (seguranca atrelada ao usuario atual do Windows)
    try:
        import win32crypt
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

def main() -> None:
    LOGGER.info("Starting grammar correction flow")
    
    try:
        import win32com.client
        try:
            word = win32com.client.GetObject(Class="Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
        
        try:
            word.Application.Run("CreateDocumentBackup", word.ActiveDocument)
            LOGGER.info("Document backup created successfully.")
        except Exception as backup_e:
            LOGGER.warning("Could not run CreateDocumentBackup macro: %s", str(backup_e))
            
        word.StatusBar = "Z7: Corrigindo gramática do trecho selecionado..."
        LOGGER.info("Connected to running Word instance")
    except Exception as e:
        log_exception(LOGGER, "Failed to connect to Word", e)
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        z7_theme.show_error("Z7 StdProposers - Erro", f"Erro ao conectar ao Word:\n{e}", parent=root)
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

    if not z7_theme.ask_privacy_warning(
        "Aviso de Privacidade - Z7 StdProposers",
        "O trecho selecionado será enviado para a API do Google Gemini para correção gramatical.\n\n"
        "Certifique-se de que não há dados sigilosos e que o uso está de acordo com as diretrizes do seu órgão.\n\n"
        "Deseja continuar?",
        key="correct_grammar",
        parent=root
    ):
        LOGGER.info("User cancelled grammar correction at privacy warning")
        root.destroy()
        return

    api_key = get_api_key()
    if not api_key:
        LOGGER.warning("Aborting grammar flow because API key is unavailable")
        root.destroy()
        return

    try:
        import google.genai as genai
    except ModuleNotFoundError as e:
        LOGGER.error("Missing Python dependency: %s", e.name)
        z7_theme.show_error(
            "Z7 StdProposers - Dependencia ausente",
            "Nao foi possivel iniciar a API porque falta a dependencia Python: "
            f"{e.name}.\n\n"
            "Execute o arquivo install_requirements.bat da pasta ai "
            "e tente novamente.",
            parent=root
        )
        root.destroy()
        sys.exit(1)

    client = genai.Client(api_key=api_key)
    LOGGER.info("Gemini client configured")

    # Mostrar carregamento
    colors = z7_theme.get_theme_colors()
    loading = tk.Toplevel(root)
    loading.title("Z7 StdProposers")
    loading.geometry("300x100")
    loading.configure(bg=colors["bg"])
    loading.attributes('-topmost', True)
    z7_theme._center_window(loading, root)
    tk.Label(loading, text="Consultando Google Gemini...\nPor favor, aguarde.", font=("Segoe UI", 10), bg=colors["bg"], fg=colors["fg"]).pack(expand=True)
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
            z7_theme.show_error("Z7 StdProposers - Erro na API", "Sua chave da API é inválida ou expirou. O arquivo de chave foi deletado.\nTente rodar a macro novamente para inserir uma nova chave.", parent=root)
        else:
            z7_theme.show_error("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}", parent=root)
        root.destroy()
        return

    loading.destroy()

    if not corrected_text:
        z7_theme.show_warning("Z7 StdProposers - Aviso", "A API do Gemini retornou um texto vazio.", parent=root)
        root.destroy()
        return

    try:
        # Tenta salvar os atributos básicos de formatação do texto selecionado
        try:
            font_name = selection.Font.Name
            font_size = selection.Font.Size
            font_bold = selection.Font.Bold
            font_italic = selection.Font.Italic
        except Exception:
            font_name, font_size, font_bold, font_italic = None, None, None, None

        selection.Text = corrected_text
        
        # Reaplica os atributos básicos, se possível
        if font_name is not None:
            try:
                selection.Font.Name = font_name
                selection.Font.Size = font_size
                selection.Font.Bold = font_bold
                selection.Font.Italic = font_italic
            except Exception:
                pass
                
        LOGGER.info("Text replaced in Word document directly, keeping formatting")
        word.StatusBar = "Z7: Correção finalizada."
    except Exception as e:
        log_exception(LOGGER, "Failed to apply text to Word", e)
        z7_theme.show_error("Z7 StdProposers - Erro", f"Não foi possível substituir o texto:\n{e}", parent=root)

    root.destroy()

if __name__ == "__main__":
    main()

