import os
import sys
import threading
import tkinter as tk
from pathlib import Path
import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_gemini_key import get_api_key, delete_api_key

LOGGER = configure_component_logger("correct_grammar")

def main() -> None:
    LOGGER.info("Starting grammar correction flow")

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.update()

    try:
        import win32com.client
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.GetObject(Class="Word.Application")
        
        try:
            word.Application.Run("CreateDocumentBackup", word.ActiveDocument)
            LOGGER.info("Document backup created successfully.")
        except Exception as backup_e:
            LOGGER.warning("Could not run CreateDocumentBackup macro: %s", str(backup_e))
            
        word.StatusBar = "Z7: Corrigindo gramática do trecho selecionado..."
        LOGGER.info("Connected to running Word instance")
    except Exception as e:
        log_exception(LOGGER, "Failed to connect to Word", e)
        z7_theme.show_error("Z7 StdProposers - Erro", f"Erro ao conectar ao Word:\n{e}", parent=root)
        root.destroy()
        return

    try:
        selection = word.Selection
        if selection is None:
            LOGGER.error("word.Selection is None: no active document or selection in Word")
            word.StatusBar = "Z7: Nenhum documento ou seleção disponível."
            root.destroy()
            return
        text = selection.Text.strip()
    except Exception as e:
        LOGGER.error("Erro ao obter selection: %s", str(e))
        word.StatusBar = "Z7: Nenhum documento ou seleção disponível."
        root.destroy()
        return

    if not text or len(text) < 2:
        LOGGER.info("Selection too short; nothing to process")
        word.StatusBar = "Z7: Selecione um trecho de texto antes de executar."
        root.destroy()
        return

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

    api_key = get_api_key(parent=root)
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

    client = genai.Client(api_key=api_key, http_options={'timeout': 60_000})
    LOGGER.info("Gemini client configured")

    from config_prompt import load_prompt, load_ai_model
    base_prompt = load_prompt()
    LOGGER.info("Prompt loaded via config_prompt")
    prompt = f"{base_prompt}\n\n---INICIO DO TEXTO SELECIONADO---\n{text}\n---FIM DO TEXTO SELECIONADO---\n"
    model_name = load_ai_model()
    LOGGER.info("Model loaded via config_prompt: %s", model_name)

    api_result: dict = {}

    def _call_api() -> None:
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            api_result["corrected_text"] = response.text.strip()
            LOGGER.info("Gemini response received with model: %s", model_name)
        except Exception as exc:
            api_result["error"] = exc
            log_exception(LOGGER, "Gemini API request failed", exc)

    loading_win = tk.Toplevel(root)
    loading_win.title("Z7 StdProposers")
    loading_win.resizable(False, False)
    loading_win.attributes('-topmost', True)
    loading_win.grab_set()
    tk.Label(
        loading_win,
        text="Corrigindo gramática com Gemini AI...\nAguarde...",
        padx=20,
        pady=20,
    ).pack()
    loading_win.update()

    api_thread = threading.Thread(target=_call_api, daemon=True)
    api_thread.start()

    def _poll() -> None:
        if api_thread.is_alive():
            root.after(200, _poll)
            return
        loading_win.destroy()
        root.quit()

    root.after(200, _poll)
    root.mainloop()

    if "error" in api_result:
        e = api_result["error"]
        error_msg = str(e).lower()
        if "403" in error_msg or "401" in error_msg or "invalid api key" in error_msg or "api_key" in error_msg:
            if z7_theme.ask_ok_cancel(
                "Z7 StdProposers - Chave Inválida",
                "Sua chave da API parece inválida ou expirou.\n\nDeseja removê-la para inserir uma nova na próxima execução?",
                parent=root
            ):
                delete_api_key()
                z7_theme.show_info("Z7 StdProposers", "Chave removida. Tente rodar a macro novamente para inserir uma nova.", parent=root)
            else:
                z7_theme.show_error("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}", parent=root)
        else:
            z7_theme.show_error("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}", parent=root)
        root.destroy()
        return

    corrected_text = api_result.get("corrected_text", "")

    if not corrected_text:
        z7_theme.show_warning("Z7 StdProposers - Aviso", "A API do Gemini retornou um texto vazio.", parent=root)
        root.destroy()
        return

    # Tenta salvar os atributos básicos de formatação do texto selecionado
    try:
        font_name = selection.Font.Name
        font_size = selection.Font.Size
        font_bold = selection.Font.Bold
        font_italic = selection.Font.Italic
    except Exception:
        font_name, font_size, font_bold, font_italic = None, None, None, None

    undo_started = False
    try:
        # Agrupa todas as alterações em um único passo de desfazer (Ctrl+Z = 1 clique)
        try:
            word.Application.UndoRecord.StartCustomRecord("Z7 - Correção Gramatical")
            undo_started = True
        except Exception:
            pass

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
    finally:
        if undo_started:
            try:
                word.Application.UndoRecord.EndCustomRecord()
            except Exception:
                pass

    root.destroy()

if __name__ == "__main__":
    main()

