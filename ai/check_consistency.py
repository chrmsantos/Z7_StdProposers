import sys
import threading
import tkinter as tk

import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_gemini_key import get_api_key, delete_api_key

LOGGER = configure_component_logger("check_consistency")

_NO_ISSUE_MARKER = "sem inconsistências graves detectadas"


def _show_issues_window(title: str, analysis: str, parent: tk.Tk) -> None:
    """Exibe janela rolável com as inconsistências graves identificadas pela IA."""
    colors = z7_theme.get_theme_colors()
    bg = colors["bg"]
    fg = colors["fg"]
    text_bg = colors["text_bg"]
    border = colors["border"]
    btn_primary = colors.get("btn_primary_bg", "#2563eb")
    btn_primary_hover = colors.get("btn_primary_hover", "#1d4ed8")

    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("700x500")
    win.minsize(540, 380)
    win.configure(bg=bg)
    win.attributes('-topmost', True)
    if parent and parent.winfo_viewable():
        win.transient(parent)
    win.grab_set()
    win.after(100, lambda: (win.lift(), win.focus_force()))

    tk.Label(
        win,
        text="\u26a0\ufe0f  Inconsistências graves detectadas",
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
        state=tk.DISABLED,
    )
    scrollbar = tk.Scrollbar(frame, command=text_widget.yview)
    text_widget.config(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=1, pady=1)

    text_widget.config(state=tk.NORMAL)
    text_widget.insert(tk.END, analysis)
    text_widget.config(state=tk.DISABLED)

    btn_frame = tk.Frame(win, bg=bg)
    btn_frame.pack(fill=tk.X, padx=20, pady=(0, 18))

    tk.Button(
        btn_frame,
        text="Fechar",
        font=("Segoe UI", 10, "bold"),
        bg=btn_primary, fg="white",
        activebackground=btn_primary_hover, activeforeground="white",
        relief=tk.FLAT, cursor="hand2",
        command=win.destroy, width=14,
    ).pack(side=tk.RIGHT)

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.bind("<Escape>", lambda e: win.destroy())

    parent.wait_window(win)


def main() -> None:
    LOGGER.info("Starting consistency check flow")

    # Inicializa o Tk ANTES das operações COM para evitar conflito
    # na fila de mensagens Win32 que causa travamento do dialog
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
        word.StatusBar = "Z7: Verificando consistencia da propositura..."
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
        full_text = doc.Range().Text.strip()
    except Exception as e:
        log_exception(LOGGER, "Failed to read document text", e)
        word.StatusBar = "Z7: Erro IA - falha ao ler o documento."
        z7_theme.show_error(
            "Z7 StdProposers - Erro",
            f"Não foi possível ler o texto do documento:\n{e}",
            parent=root,
        )
        root.destroy()
        return

    if not full_text or len(full_text) < 10:
        word.StatusBar = "Z7: Documento vazio ou muito curto para analise."
        LOGGER.info("Document too short; nothing to analyze")
        root.destroy()
        return

    if not z7_theme.ask_privacy_warning(
        "Aviso de Privacidade - Z7 StdProposers",
        "O texto completo da propositura será enviado para a API do Google Gemini "
        "para verificação de consistência lógica e semântica.\n\n"
        "Certifique-se de que não há dados sigilosos e que o uso está de acordo "
        "com as diretrizes do seu órgão.\n\n"
        "Deseja continuar?",
        key="check_consistency",
        parent=root,
    ):
        LOGGER.info("User cancelled consistency check at privacy warning")
        root.destroy()
        return

    api_key = get_api_key(parent=root)
    if not api_key:
        LOGGER.warning("Aborting consistency flow because API key is unavailable")
        root.destroy()
        return

    try:
        import google.genai as genai
    except ModuleNotFoundError as e:
        LOGGER.error("Missing Python dependency: %s", e.name)
        word.StatusBar = "Z7: Erro IA - dependencia Python ausente."
        z7_theme.show_error(
            "Z7 StdProposers - Dependência ausente",
            f"Não foi possível iniciar a API porque falta a dependência Python: "
            f"{e.name}.\n\n"
            "Execute o arquivo install_requirements.bat da pasta ai e tente novamente.",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    client = genai.Client(api_key=api_key, http_options={'timeout': 120_000})
    LOGGER.info("Gemini client configured")

    from config_prompt import load_consistency_prompt, load_ai_model
    base_prompt = load_consistency_prompt()
    model_name = load_ai_model()
    LOGGER.info("Consistency prompt and model loaded: %s", model_name)

    prompt = (
        f"{base_prompt}\n\n"
        "---INICIO DA PROPOSITURA---\n"
        f"{full_text}\n"
        "---FIM DA PROPOSITURA---\n"
    )

    word.StatusBar = "Z7: Analisando consistencia com Gemini... Aguarde."
    LOGGER.info("Sending document to Gemini for consistency analysis")

    # Exibe janela de progresso enquanto a chamada à API é feita em background
    colors = z7_theme.get_theme_colors()
    loading_win = tk.Toplevel(root)
    loading_win.title("Z7 StdProposers")
    loading_win.resizable(False, False)
    loading_win.configure(bg=colors["bg"])
    loading_win.attributes("-topmost", True)
    tk.Label(
        loading_win,
        text="Analisando a propositura com Gemini AI...\nAguarde, isso pode levar alguns instantes.",
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

    api_result: dict = {}

    def _call_api() -> None:
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            api_result["analysis"] = response.text.strip()
        except Exception as exc:
            api_result["error"] = exc
            log_exception(LOGGER, "Gemini API request failed", exc)

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
                "Sua chave da API parece inválida ou expirou.\n\n"
                "Deseja removê-la para inserir uma nova na próxima execução?",
                parent=root,
            ):
                delete_api_key()
                z7_theme.show_info(
                    "Z7 StdProposers",
                    "Chave removida. Tente rodar a macro novamente para inserir uma nova.",
                    parent=root,
                )
            else:
                z7_theme.show_error(
                    "Z7 StdProposers - Erro na API",
                    f"Erro ao chamar a API do Gemini:\n{e}",
                    parent=root,
                )
        else:
            z7_theme.show_error(
                "Z7 StdProposers - Erro na API",
                f"Erro ao chamar a API do Gemini:\n{e}",
                parent=root,
            )
        word.StatusBar = "Z7: Erro IA - falha na verificacao de consistencia."
        root.destroy()
        return

    analysis = api_result.get("analysis", "")
    LOGGER.info("Gemini consistency response received")

    if not analysis:
        z7_theme.show_warning(
            "Z7 StdProposers - Aviso",
            "A API do Gemini retornou uma resposta vazia.",
            parent=root,
        )
        word.StatusBar = "Z7: Erro IA - resposta vazia do Gemini."
        root.destroy()
        return

    has_issues = _NO_ISSUE_MARKER not in analysis.lower()
    word.StatusBar = "Z7: Verificacao de consistencia concluida."
    LOGGER.info("Consistency analysis complete. Issues detected: %s", has_issues)

    if has_issues:
        _show_issues_window(
            "Z7 StdProposers - Verificação de Consistência",
            analysis,
            parent=root,
        )
    else:
        z7_theme.show_info(
            "Z7 StdProposers - Verificação de Consistência",
            "Nenhuma inconsistência grave detectada na propositura.",
            parent=root,
        )

    root.destroy()


if __name__ == "__main__":
    main()
