import os
import sys
import threading
import tkinter as tk
from pathlib import Path
import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_gemini_key import get_api_key, delete_api_key

LOGGER = configure_component_logger("correct_grammar")


def _show_result_window(corrected_text: str, parent: tk.Tk) -> "str | None":
    """Exibe o texto corrigido para revisão antes de substituir no Word.

    Retorna o texto final (possivelmente editado pelo usuário) ou None se cancelado.
    """
    colors = z7_theme.get_theme_colors()
    bg = colors["bg"]
    fg = colors["fg"]
    text_bg = colors["text_bg"]
    border = colors["border"]
    btn_primary_bg = colors["btn_primary_bg"]
    btn_primary_fg = colors["btn_primary_fg"]
    btn_primary_hover = colors["btn_primary_hover"]

    result: dict = {}

    win = tk.Toplevel(parent)
    win.title("Z7 StdProposers - Correção Gramatical")
    win.geometry("680x440")
    win.minsize(500, 320)
    win.configure(bg=bg)
    win.attributes("-topmost", True)
    if parent and parent.winfo_viewable():
        win.transient(parent)
    win.grab_set()
    win.after(100, lambda: (win.lift(), win.focus_force()))

    tk.Label(
        win,
        text="\u2713  Texto corrigido \u2014 revise e confirme",
        font=("Segoe UI", 13, "bold"),
        bg=bg, fg=fg,
        anchor="w",
    ).pack(fill=tk.X, padx=20, pady=(18, 6))

    tk.Frame(win, bg=border, height=1).pack(fill=tk.X, padx=20, pady=(0, 10))

    frame = tk.Frame(win, bg=border, bd=0)
    frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

    text_widget = tk.Text(
        frame,
        wrap=tk.WORD,
        font=("Segoe UI", 10),
        bg=text_bg, fg=fg,
        insertbackground=fg,
        relief=tk.FLAT,
        padx=10, pady=10,
    )
    scrollbar = tk.Scrollbar(frame, command=text_widget.yview)
    text_widget.config(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=1, pady=1)

    display_text = corrected_text.replace('\r\n', '\n').replace('\r', '\n')
    text_widget.insert(tk.END, display_text)
    text_widget.focus_set()

    btn_frame = tk.Frame(win, bg=bg)
    btn_frame.pack(fill=tk.X, padx=20, pady=(0, 18))

    def on_apply() -> None:
        result["text"] = text_widget.get("1.0", tk.END).rstrip("\n").replace('\n', '\r')
        win.destroy()

    def on_cancel() -> None:
        win.destroy()

    tk.Button(
        btn_frame,
        text="Cancelar",
        font=("Segoe UI", 10),
        bg=border, fg=fg,
        activebackground=border, activeforeground=fg,
        relief=tk.FLAT, cursor="hand2",
        command=on_cancel, width=12,
    ).pack(side=tk.RIGHT, padx=(6, 0))

    tk.Button(
        btn_frame,
        text="Substituir no Word",
        font=("Segoe UI", 10, "bold"),
        bg=btn_primary_bg, fg=btn_primary_fg,
        activebackground=btn_primary_hover, activeforeground=btn_primary_fg,
        relief=tk.FLAT, cursor="hand2",
        command=on_apply, width=18,
    ).pack(side=tk.RIGHT)

    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.bind("<Escape>", lambda e: on_cancel())

    parent.wait_window(win)
    return result.get("text")


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
            
        word.StatusBar = "Z7: Corrigindo gramatica do documento..."
        LOGGER.info("Connected to running Word instance")
    except Exception as e:
        log_exception(LOGGER, "Failed to connect to Word", e)
        z7_theme.show_error("Z7 StdProposers - Erro", f"Erro ao conectar ao Word:\n{e}", parent=root)
        root.destroy()
        return

    # 'word.ActiveDocument' lança com_error (não retorna None) quando nenhum
    # documento está aberto — capturar separadamente do erro de leitura de conteúdo
    try:
        doc = word.ActiveDocument
    except Exception:
        doc = None
    if doc is None:
        LOGGER.error("No active document in Word")
        word.StatusBar = "Z7: Erro IA - nenhum documento aberto."
        z7_theme.show_error(
            "Z7 StdProposers",
            "Nenhum documento está aberto no Word.\nAbra um documento e tente novamente.",
            parent=root,
        )
        root.destroy()
        return
    try:
        full_text = doc.Content.Text.strip().replace('\r\n', '\n').replace('\r', '\n')
    except Exception as e:
        LOGGER.error("Erro ao obter texto do documento: %s", str(e))
        word.StatusBar = "Z7: Erro IA - falha ao ler o documento."
        z7_theme.show_error(
            "Z7 StdProposers - Erro",
            f"Não foi possível ler o texto do documento:\n{e}",
            parent=root,
        )
        root.destroy()
        return

    if not full_text or len(full_text) < 10:
        LOGGER.info("Document too short; nothing to process")
        word.StatusBar = "Z7: Documento vazio ou muito curto para correcao."
        root.destroy()
        return

    if not z7_theme.ask_privacy_warning(
        "Aviso de Privacidade - Z7 StdProposers",
        "O texto completo do documento será enviado para a API do Google Gemini para correção gramatical.\n\n"
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
        word.StatusBar = "Z7: Erro IA - chave da API indisponivel."
        root.destroy()
        return

    try:
        import google.genai as genai
    except ModuleNotFoundError as e:
        LOGGER.error("Missing Python dependency: %s", e.name)
        word.StatusBar = "Z7: Erro IA - dependencia Python ausente."
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

    client = genai.Client(api_key=api_key, http_options={'timeout': 120_000})
    LOGGER.info("Gemini client configured")

    from config_prompt import load_prompt, load_ai_model
    base_prompt = load_prompt()
    LOGGER.info("Prompt loaded via config_prompt")
    prompt = f"{base_prompt}\n\n---INICIO DO DOCUMENTO---\n{full_text}\n---FIM DO DOCUMENTO---\n"
    model_name = load_ai_model()
    LOGGER.info("Model loaded via config_prompt: %s", model_name)

    word.StatusBar = "Z7: Corrigindo gramatica com Gemini... Aguarde."
    LOGGER.info("Sending full document text to Gemini for grammar correction")

    api_result: dict = {}

    def _call_api() -> None:
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            api_result["corrected_text"] = response.text.strip()
            LOGGER.info("Gemini response received with model: %s", model_name)
        except Exception as exc:
            api_result["error"] = exc
            log_exception(LOGGER, "Gemini API request failed", exc)

    # Exibe janela de progresso enquanto a chamada à API é feita em background
    colors = z7_theme.get_theme_colors()
    loading_win = tk.Toplevel(root)
    loading_win.title("Z7 StdProposers")
    loading_win.resizable(False, False)
    loading_win.configure(bg=colors["bg"])
    loading_win.attributes("-topmost", True)
    tk.Label(
        loading_win,
        text="Corrigindo gramática com Gemini AI...\nAguarde, isso pode levar alguns instantes.",
        font=("Segoe UI", 10),
        bg=colors["bg"], fg=colors["fg"],
        justify=tk.CENTER,
        padx=30, pady=30,
    ).pack()
    loading_win.update_idletasks()
    loading_win.geometry(
        "+%d+%d" % (
            loading_win.winfo_screenwidth() // 2 - 200,
            loading_win.winfo_screenheight() // 2 - 60,
        )
    )
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
                word.StatusBar = "Z7: Erro IA - chave da API invalida."
                z7_theme.show_info("Z7 StdProposers", "Chave removida. Tente rodar a macro novamente para inserir uma nova.", parent=root)
            else:
                word.StatusBar = "Z7: Erro IA - falha na chamada ao Gemini."
                z7_theme.show_error("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}", parent=root)
        else:
            word.StatusBar = "Z7: Erro IA - falha na chamada ao Gemini."
            z7_theme.show_error("Z7 StdProposers - Erro na API", f"Erro ao chamar a API do Gemini:\n{e}", parent=root)
        root.destroy()
        return

    corrected_text = api_result.get("corrected_text", "")

    if not corrected_text:
        word.StatusBar = "Z7: Erro IA - resposta vazia do Gemini."
        z7_theme.show_warning("Z7 StdProposers - Aviso", "A API do Gemini retornou um texto vazio.", parent=root)
        root.destroy()
        return

    final_text = _show_result_window(corrected_text, parent=root)
    if not final_text:
        LOGGER.info("User cancelled grammar correction at review window")
        word.StatusBar = "Z7: Correcao cancelada."
        root.destroy()
        return

    undo_started = False
    try:
        # Agrupa todas as alterações em um único passo de desfazer (Ctrl+Z = 1 clique)
        try:
            word.Application.UndoRecord.StartCustomRecord("Z7 - Correção Gramatical")
            undo_started = True
        except Exception:
            pass

        doc.Content.Text = final_text

        LOGGER.info("Document text replaced with corrected version")
        word.StatusBar = "Z7: Correcao finalizada."
    except Exception as e:
        log_exception(LOGGER, "Failed to apply text to Word", e)
        word.StatusBar = "Z7: Erro IA - falha ao aplicar texto corrigido."
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

