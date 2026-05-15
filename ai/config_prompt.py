import json
import tkinter as tk
from pathlib import Path
import z7_theme
from z7_logging import configure_component_logger, get_data_dir, log_exception

LOGGER = configure_component_logger("config_prompt")

GITHUB_REPO_URL = "https://github.com/chrmsantos/Z7_StdProposers"

DEFAULT_PROMPT = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

DEFAULT_CONSISTENCY_PROMPT = """Você é um especialista em análise jurídica e legislativa no idioma Português do Brasil.
Analise criteriosamente a propositura legislativa abaixo em busca de inconsistências graves.

Considere como inconsistências graves:
1. Divergências entre o conteúdo da ementa e o restante do texto;
2. Contradições entre nomes de pessoas, lugares, logradouros ou endereços;
3. Contradições entre tipos de moção ou natureza do ato legislativo;
4. Contradições lógicas ou semânticas internas ao texto;
5. Qualquer incoerência que comprometa a validade ou o sentido jurídico do documento.

Não reporte como problemas:
- Pequenas divergências de grafia ou acentuação;
- Diferenças na ordem de palavras que não alterem o sentido;
- Pequenos erros formais ou gramaticais que não criem contradição lógica.

Se encontrar inconsistências graves, liste-as numericamente, uma por linha, de forma sucinta (máximo 2 linhas por item). Para cada item, indique apenas o tipo de inconsistência e os trechos conflitantes — sem introduções, sem conclusões, sem repetições.
Se NÃO encontrar inconsistências graves, responda APENAS com: "Sem inconsistências graves detectadas."

Seja direto e conciso. Responda em Português do Brasil."""

def get_prompt_file_path() -> Path:
    return get_data_dir() / "gemini_prompt.txt"


def get_consistency_prompt_file_path() -> Path:
    return get_data_dir() / "consistency_prompt.txt"


def load_api_key() -> str:
    from z7_gemini_key import read_stored_api_key
    return read_stored_api_key()

def save_api_key(api_key: str) -> None:
    from z7_gemini_key import write_api_key
    write_api_key(api_key)

def load_prompt() -> str:
    prompt_file = get_prompt_file_path()
    if prompt_file.exists():
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                LOGGER.info("Loaded custom prompt file")
                return f.read()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom prompt", e)
    return DEFAULT_PROMPT


def load_consistency_prompt() -> str:
    consistency_file = get_consistency_prompt_file_path()
    if consistency_file.exists():
        try:
            with open(consistency_file, 'r', encoding='utf-8') as f:
                LOGGER.info("Loaded custom consistency prompt file")
                return f.read()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom consistency prompt", e)
    return DEFAULT_CONSISTENCY_PROMPT

def get_model_file_path() -> Path:
    return get_data_dir() / "selected_model.txt"

def load_ai_model() -> str:
    model_file = get_model_file_path()
    if model_file.exists():
        try:
            with open(model_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom model", e)
    return "gemini-2.5-flash"

def save_ai_model(model_name: str) -> None:
    model_file = get_model_file_path()
    try:
        with open(model_file, 'w', encoding='utf-8') as f:
            f.write(model_name)
        LOGGER.info("Model saved successfully: %s", model_name)
    except Exception as e:
        log_exception(LOGGER, "Failed to save model", e)

def save_prompt(grammar_text: str, consistency_text: str, root: tk.Tk, model_var: tk.StringVar,
               privacy_chat_var: tk.BooleanVar | None = None,
               privacy_grammar_var: tk.BooleanVar | None = None,
               privacy_consistency_var: tk.BooleanVar | None = None) -> None:
    if not grammar_text.strip():
        LOGGER.warning("Grammar prompt save blocked because text is empty")
        z7_theme.show_warning("Aviso", "O prompt do Corretor Gramatical não pode estar vazio.", parent=root)
        return
    if not consistency_text.strip():
        LOGGER.warning("Consistency prompt save blocked because text is empty")
        z7_theme.show_warning("Aviso", "O prompt do Verificador de Consistência não pode estar vazio.", parent=root)
        return

    prompt_file = get_prompt_file_path()
    try:
        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(grammar_text.strip())
        LOGGER.info("Grammar prompt saved successfully")
    except Exception as e:
        log_exception(LOGGER, "Failed to save grammar prompt", e)
        z7_theme.show_error("Erro", f"Erro ao salvar prompt gramatical:\n{e}", parent=root)
        return

    consistency_file = get_consistency_prompt_file_path()
    try:
        with open(consistency_file, 'w', encoding='utf-8') as f:
            f.write(consistency_text.strip())
        LOGGER.info("Consistency prompt saved successfully")
    except Exception as e:
        log_exception(LOGGER, "Failed to save consistency prompt", e)
        z7_theme.show_error("Erro", f"Erro ao salvar prompt de consistência:\n{e}", parent=root)
        return

    save_ai_model(model_var.get())

    # Save privacy prefs (True = aviso desativado; ausente/False = aviso ativo)
    if privacy_chat_var is not None or privacy_grammar_var is not None or privacy_consistency_var is not None:
        prefs = z7_theme.load_privacy_prefs()
        if privacy_chat_var is not None:
            if privacy_chat_var.get():
                prefs.pop('chat_ia', None)       # reativar aviso
            else:
                prefs['chat_ia'] = True          # suprimir aviso
        if privacy_grammar_var is not None:
            if privacy_grammar_var.get():
                prefs.pop('correct_grammar', None)
            else:
                prefs['correct_grammar'] = True
        if privacy_consistency_var is not None:
            if privacy_consistency_var.get():
                prefs.pop('check_consistency', None)
            else:
                prefs['check_consistency'] = True
        z7_theme.save_privacy_prefs(prefs)
        LOGGER.info("Privacy prefs saved")

    z7_theme.show_info("Sucesso", "Configurações salvas com sucesso!", parent=root)
    root.destroy()

def restore_default(text_widget: tk.Text) -> None:
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, DEFAULT_PROMPT)


def restore_default_consistency(text_widget: tk.Text) -> None:
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, DEFAULT_CONSISTENCY_PROMPT)

class AppTheme:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.mode = z7_theme.load_theme()
        self.widgets = {}

    def toggle(self) -> None:
        self.mode = 'dark' if self.mode == 'light' else 'light'
        z7_theme.save_theme(self.mode)
        self.apply()

    def apply(self) -> None:
        colors = z7_theme.get_theme_colors(self.mode)
        bg = colors["bg"]
        fg = colors["fg"]
        fg_muted = colors["fg_muted"]
        text_bg = colors["text_bg"]
        border = colors["border"]
        btn_sec_bg = colors["btn_sec_bg"]
        btn_sec_fg = colors["btn_sec_fg"]
        btn_sec_hover = colors["btn_sec_hover"]
            
        self.root.configure(bg=bg)
        
        if 'title_lbl' in self.widgets:
            self.widgets['title_lbl'].configure(bg=bg, fg=fg)
        if 'info_lbl' in self.widgets:
            self.widgets['info_lbl'].configure(bg=bg, fg=fg_muted)
        if 'btn_frame' in self.widgets:
            self.widgets['btn_frame'].configure(bg=bg)
        if 'border_frame' in self.widgets:
            self.widgets['border_frame'].configure(bg=border)
        if 'text_area' in self.widgets:
            self.widgets['text_area'].configure(bg=text_bg, fg=fg, insertbackground=fg)
        if 'model_frame' in self.widgets:
            self.widgets['model_frame'].configure(bg=bg)
        if 'model_lbl' in self.widgets:
            self.widgets['model_lbl'].configure(bg=bg, fg=fg)
        if 'model_dropdown' in self.widgets:
            self.widgets['model_dropdown'].configure(bg=text_bg, fg=fg, insertbackground=fg)

        if 'api_frame' in self.widgets:
            self.widgets['api_frame'].configure(bg=bg)
        if 'api_btn' in self.widgets:
            self.widgets['api_btn'].configure(bg=btn_sec_bg, fg=btn_sec_fg,
                                              activebackground=btn_sec_hover, activeforeground=fg)

        if 'privacy_frame' in self.widgets:
            self.widgets['privacy_frame'].configure(bg=bg)
        if 'privacy_lbl' in self.widgets:
            self.widgets['privacy_lbl'].configure(bg=bg, fg=fg)
        for cb in self.widgets.get('privacy_checks', []):
            cb.configure(bg=bg, fg=fg_muted, activebackground=bg, activeforeground=fg,
                         selectcolor=text_bg)

        if 'tab_frame' in self.widgets:
            self.widgets['tab_frame'].configure(bg=bg)
        active_idx = self.widgets.get('active_tab_idx', 0)
        for i, btn in enumerate(self.widgets.get('tab_btns', [])):
            if i == active_idx:
                btn.configure(
                    bg=colors["btn_primary_bg"], fg=colors["btn_primary_fg"],
                    activebackground=colors["btn_primary_hover"],
                    activeforeground=colors["btn_primary_fg"],
                )
            else:
                btn.configure(
                    bg=btn_sec_bg, fg=btn_sec_fg,
                    activebackground=btn_sec_hover, activeforeground=fg,
                )

        for btn in self.widgets.get('sec_btns', []):
            btn.configure(bg=btn_sec_bg, fg=btn_sec_fg, activebackground=btn_sec_hover, activeforeground=fg)
        
        if 'toggle_btn' in self.widgets:
            icon = "🌙 Modo Escuro" if self.mode == 'light' else "☀️ Modo Claro"
            self.widgets['toggle_btn'].configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)


def open_api_key_dialog(parent: tk.Tk, theme_mode: str) -> None:
    colors = z7_theme.get_theme_colors(theme_mode)
    bg          = colors["bg"]
    fg          = colors["fg"]
    text_bg     = colors["text_bg"]
    btn_sec_bg  = colors["btn_sec_bg"]
    btn_sec_fg  = colors["btn_sec_fg"]
    btn_sec_hover = colors["btn_sec_hover"]
    btn_primary = colors.get("btn_primary_bg", "#2563eb")

    dialog = tk.Toplevel(parent)
    dialog.title("Chave de API Gemini")
    dialog.geometry("460x155")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes('-topmost', True)
    dialog.configure(bg=bg)

    tk.Label(dialog, text="Chave de API do Gemini:", font=("Segoe UI", 10, "bold"),
             bg=bg, fg=fg).pack(anchor="w", padx=25, pady=(20, 5))

    entry_frame = tk.Frame(dialog, bg=bg)
    entry_frame.pack(fill=tk.X, padx=25)

    entry_var = tk.StringVar(value=load_api_key())
    entry = tk.Entry(entry_frame, textvariable=entry_var, font=("Segoe UI", 11),
                     relief=tk.FLAT, bd=2, show="*", bg=text_bg, fg=fg,
                     insertbackground=fg)
    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def toggle_visibility() -> None:
        if entry.cget("show") == "*":
            entry.config(show="")
            show_btn.config(text="Ocultar")
        else:
            entry.config(show="*")
            show_btn.config(text="Mostrar")

    show_btn = tk.Button(entry_frame, text="Mostrar", font=("Segoe UI", 9),
                         relief=tk.FLAT, cursor="hand2",
                         bg=btn_sec_bg, fg=btn_sec_fg,
                         activebackground=btn_sec_hover, activeforeground=fg,
                         command=toggle_visibility)
    show_btn.pack(side=tk.LEFT, padx=(8, 0))

    def do_save() -> None:
        save_api_key(entry_var.get().strip())
        LOGGER.info("API key saved from dialog")
        dialog.destroy()

    btn_row = tk.Frame(dialog, bg=bg)
    btn_row.pack(fill=tk.X, padx=25, pady=(15, 0))

    tk.Button(btn_row, text="Cancelar", font=("Segoe UI", 10, "bold"),
              relief=tk.FLAT, cursor="hand2",
              bg=btn_sec_bg, fg=btn_sec_fg,
              activebackground=btn_sec_hover, activeforeground=fg,
              command=dialog.destroy).pack(side=tk.RIGHT, padx=(10, 0))

    tk.Button(btn_row, text="Salvar", font=("Segoe UI", 10, "bold"),
              relief=tk.FLAT, cursor="hand2",
              bg=btn_primary, fg="white",
              activebackground="#1d4ed8", activeforeground="white",
              command=do_save).pack(side=tk.RIGHT)


def main() -> None:
    LOGGER.info("Starting prompt configuration UI")
    
    try:
        import win32com.client
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.GetObject(Class="Word.Application")
        word.StatusBar = "Z7: Abrindo Configurações..."
    except Exception as e:
        LOGGER.warning("Could not connect to Word to update status bar: %s", str(e))
        
    root = tk.Tk()
    root.title("Configurar Prompt do Gemini")
    root.geometry("700x600")
    root.minsize(600, 500)
    
    theme = AppTheme(root)
    
    # Faz a janela aparecer na frente
    root.attributes('-topmost', True)
    
    # Botão de tema no canto superior
    toggle_btn = tk.Button(root, font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2", command=theme.toggle, bd=0)
    toggle_btn.place(relx=0.03, rely=0.03)
    theme.widgets['toggle_btn'] = toggle_btn
    
    lbl = tk.Label(root, text="Instruções para a Inteligência Artificial", font=("Segoe UI", 16, "bold"))
    lbl.pack(pady=(35, 5))  # Aumentado o pady superior para não sobrepor o botão de tema
    theme.widgets['title_lbl'] = lbl
    
    info_lbl = tk.Label(root, text="Personalize o comportamento do modelo ajustando o prompt abaixo.", font=("Segoe UI", 10))
    info_lbl.pack(pady=(0, 20))
    theme.widgets['info_lbl'] = info_lbl

    btn_frame = tk.Frame(root)
    theme.widgets['btn_frame'] = btn_frame

    # Seleção de Modelo
    model_frame = tk.Frame(root)
    model_frame.pack(fill=tk.X, padx=25, pady=(0, 10))
    theme.widgets['model_frame'] = model_frame

    model_lbl = tk.Label(model_frame, text="Modelo de IA:", font=("Segoe UI", 10, "bold"))
    model_lbl.pack(side=tk.LEFT)
    theme.widgets['model_lbl'] = model_lbl

    model_var = tk.StringVar(root)
    current_model = load_ai_model()
    model_var.set(current_model)

    model_entry = tk.Entry(model_frame, textvariable=model_var, font=("Segoe UI", 10), relief=tk.FLAT, bd=2)
    model_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
    theme.widgets['model_dropdown'] = model_entry

    # Chave de API
    api_frame = tk.Frame(root)
    api_frame.pack(fill=tk.X, padx=25, pady=(0, 10))
    theme.widgets['api_frame'] = api_frame

    api_btn = tk.Button(api_frame, text="🔑 Chave de API...", font=("Segoe UI", 10),
                        relief=tk.FLAT, cursor="hand2",
                        command=lambda: open_api_key_dialog(root, theme.mode))
    api_btn.pack(side=tk.LEFT)
    theme.widgets['api_btn'] = api_btn

    # Avisos de Privacidade
    privacy_prefs = z7_theme.load_privacy_prefs()

    privacy_frame = tk.Frame(root)
    privacy_frame.pack(fill=tk.X, padx=25, pady=(0, 10))
    theme.widgets['privacy_frame'] = privacy_frame

    privacy_lbl = tk.Label(privacy_frame, text="Avisos de Privacidade:", font=("Segoe UI", 10, "bold"))
    privacy_lbl.pack(side=tk.LEFT)
    theme.widgets['privacy_lbl'] = privacy_lbl

    # Checked = exibir aviso (pref ausente); Unchecked = não exibir (pref True)
    privacy_chat_var = tk.BooleanVar(value=not privacy_prefs.get('chat_ia', False))
    privacy_grammar_var = tk.BooleanVar(value=not privacy_prefs.get('correct_grammar', False))
    privacy_consistency_var = tk.BooleanVar(value=not privacy_prefs.get('check_consistency', False))

    cb_chat = tk.Checkbutton(privacy_frame, text="Chat IA",
                             variable=privacy_chat_var, font=("Segoe UI", 10),
                             relief=tk.FLAT, cursor="hand2")
    cb_chat.pack(side=tk.LEFT, padx=(15, 0))

    cb_grammar = tk.Checkbutton(privacy_frame, text="Corretor Gramatical",
                                variable=privacy_grammar_var, font=("Segoe UI", 10),
                                relief=tk.FLAT, cursor="hand2")
    cb_grammar.pack(side=tk.LEFT, padx=(10, 0))

    cb_consistency = tk.Checkbutton(privacy_frame, text="Verificador de Consistência",
                                    variable=privacy_consistency_var, font=("Segoe UI", 10),
                                    relief=tk.FLAT, cursor="hand2")
    cb_consistency.pack(side=tk.LEFT, padx=(10, 0))

    theme.widgets['privacy_checks'] = [cb_chat, cb_grammar, cb_consistency]

    # Seletor de abas de prompt
    prompt_buffers = {
        'grammar': load_prompt(),
        'consistency': load_consistency_prompt(),
    }
    current_tab = tk.StringVar(value='grammar')

    tab_frame = tk.Frame(root)
    tab_frame.pack(fill=tk.X, padx=25, pady=(0, 4))
    theme.widgets['tab_frame'] = tab_frame
    theme.widgets['active_tab_idx'] = 0

    def switch_tab(tab_name: str, tab_idx: int) -> None:
        if current_tab.get() == tab_name:
            return
        prompt_buffers[current_tab.get()] = text_area.get("1.0", tk.END)
        current_tab.set(tab_name)
        theme.widgets['active_tab_idx'] = tab_idx
        text_area.delete("1.0", tk.END)
        text_area.insert(tk.END, prompt_buffers[tab_name].strip())
        theme.apply()

    grammar_tab_btn = tk.Button(
        tab_frame, text="Corretor Gramatical",
        font=("Segoe UI", 10, "bold"), relief=tk.FLAT, cursor="hand2",
        command=lambda: switch_tab('grammar', 0), padx=12,
    )
    grammar_tab_btn.pack(side=tk.LEFT)

    consistency_tab_btn = tk.Button(
        tab_frame, text="Verificador de Consistência",
        font=("Segoe UI", 10, "bold"), relief=tk.FLAT, cursor="hand2",
        command=lambda: switch_tab('consistency', 1), padx=12,
    )
    consistency_tab_btn.pack(side=tk.LEFT, padx=(4, 0))

    theme.widgets['tab_btns'] = [grammar_tab_btn, consistency_tab_btn]

    frame = tk.Frame(root) # Borda sutil
    theme.widgets['border_frame'] = frame
    
    text_area = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 11), relief=tk.FLAT, padx=12, pady=12)
    text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=1, pady=1)
    theme.widgets['text_area'] = text_area

    scrollbar = tk.Scrollbar(frame, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

    # Carrega o prompt da aba ativa (gramática por padrão)
    text_area.insert(tk.END, prompt_buffers['grammar'])

    # Estilos de botão
    btn_font = ("Segoe UI", 10, "bold")

    def do_save() -> None:
        prompt_buffers[current_tab.get()] = text_area.get("1.0", tk.END)
        save_prompt(
            grammar_text=prompt_buffers['grammar'],
            consistency_text=prompt_buffers['consistency'],
            root=root,
            model_var=model_var,
            privacy_chat_var=privacy_chat_var,
            privacy_grammar_var=privacy_grammar_var,
            privacy_consistency_var=privacy_consistency_var,
        )

    def do_restore_default() -> None:
        if current_tab.get() == 'grammar':
            restore_default(text_area)
        else:
            restore_default_consistency(text_area)

    save_btn = tk.Button(btn_frame, text="Salvar Configuração", width=20, bg="#2563eb", fg="white", font=btn_font, relief=tk.FLAT, activebackground="#1d4ed8", activeforeground="white", cursor="hand2", command=do_save)
    save_btn.pack(side=tk.RIGHT, padx=25)
    
    cancel_btn = tk.Button(btn_frame, text="Cancelar", width=15, font=btn_font, relief=tk.FLAT, cursor="hand2", command=root.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=5)

    restore_btn = tk.Button(btn_frame, text="Restaurar Padrão", width=18, font=btn_font, relief=tk.FLAT, cursor="hand2", command=do_restore_default)
    restore_btn.pack(side=tk.LEFT, padx=25)

    import webbrowser
    github_btn = tk.Button(btn_frame, text="⧉ GitHub", font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
                           command=lambda: webbrowser.open(GITHUB_REPO_URL))
    github_btn.pack(side=tk.LEFT, padx=(0, 5))
    theme.widgets.setdefault('sec_btns', []).append(github_btn)

    theme.widgets['sec_btns'] = [cancel_btn, restore_btn]
    
    # Aplica o tema inicial
    theme.apply()

    # Pack in order so btn_frame is fixed at the bottom
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=25)
    frame.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=25, pady=(0, 10))

    root.mainloop()

if __name__ == "__main__":
    main()
