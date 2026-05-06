import os
import json
import tkinter as tk
from pathlib import Path
import win32crypt
import z7_theme
from z7_logging import configure_component_logger, log_exception

LOGGER = configure_component_logger("config_prompt")

DEFAULT_PROMPT = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

def get_prompt_file_path() -> Path | None:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'gemini_prompt.txt'

def load_api_key() -> str:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return ""
    key_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'gemini.key'
    if key_file.exists():
        try:
            with open(key_file, 'rb') as f:
                encrypted_key = f.read()
            _, decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)
            return decrypted_key.decode('utf-8')
        except Exception as e:
            log_exception(LOGGER, "Failed to decrypt API key", e)
    return ""

def save_api_key(api_key: str) -> None:
    api_key = api_key.strip()
    if not api_key:
        return
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_file = key_dir / 'gemini.key'
    try:
        encrypted_key = win32crypt.CryptProtectData(api_key.encode('utf-8'), 'Z7_Gemini_Key', None, None, None, 0)
        key_dir.mkdir(parents=True, exist_ok=True)
        with open(key_file, 'wb') as f:
            f.write(encrypted_key)
        LOGGER.info("API key encrypted and persisted")
    except Exception as e:
        log_exception(LOGGER, "Failed to persist API key", e)

def load_prompt() -> str:
    prompt_file = get_prompt_file_path()
    if prompt_file and prompt_file.exists():
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                LOGGER.info("Loaded custom prompt file")
                return f.read()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom prompt", e)
    return DEFAULT_PROMPT

def get_model_file_path() -> Path | None:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'selected_model.txt'

def load_ai_model() -> str:
    model_file = get_model_file_path()
    if model_file and model_file.exists():
        try:
            with open(model_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom model", e)
    return "gemini-2.5-flash"

def save_ai_model(model_name: str) -> None:
    model_file = get_model_file_path()
    if model_file:
        try:
            with open(model_file, 'w', encoding='utf-8') as f:
                f.write(model_name)
            LOGGER.info(f"Model saved successfully: {model_name}")
        except Exception as e:
            log_exception(LOGGER, "Failed to save model", e)

def save_prompt(text_widget: tk.Text, root: tk.Tk, model_var: tk.StringVar, api_var: tk.StringVar) -> None:
    new_prompt = text_widget.get("1.0", tk.END).strip()
    if not new_prompt:
        LOGGER.warning("Prompt save blocked because text is empty")
        z7_theme.show_warning("Aviso", "O prompt não pode estar vazio.", parent=root)
        return

    prompt_file = get_prompt_file_path()
    if prompt_file:
        try:
            with open(prompt_file, 'w', encoding='utf-8') as f:
                f.write(new_prompt)
            LOGGER.info("Prompt saved successfully")
            
            # Save the model
            save_ai_model(model_var.get())
            
            # Save the API key
            save_api_key(api_var.get())
            
            z7_theme.show_info("Sucesso", "Configurações salvas com sucesso!", parent=root)
            root.destroy()
        except Exception as e:
            log_exception(LOGGER, "Failed to save config", e)
            z7_theme.show_error("Erro", f"Erro ao salvar:\n{e}", parent=root)

def restore_default(text_widget: tk.Text) -> None:
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, DEFAULT_PROMPT)

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
        if 'api_lbl' in self.widgets:
            self.widgets['api_lbl'].configure(bg=bg, fg=fg)
        if 'api_entry' in self.widgets:
            self.widgets['api_entry'].configure(bg=text_bg, fg=fg, insertbackground=fg)
            
        for btn in self.widgets.get('sec_btns', []):
            btn.configure(bg=btn_sec_bg, fg=btn_sec_fg, activebackground=btn_sec_hover, activeforeground=fg)
        
        if 'toggle_btn' in self.widgets:
            icon = "🌙 Modo Escuro" if self.mode == 'light' else "☀️ Modo Claro"
            self.widgets['toggle_btn'].configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)

def main() -> None:
    LOGGER.info("Starting prompt configuration UI")
    
    try:
        import win32com.client
        try:
            word = win32com.client.GetObject(Class="Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
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

    api_lbl = tk.Label(api_frame, text="Chave de API:", font=("Segoe UI", 10, "bold"))
    api_lbl.pack(side=tk.LEFT)
    theme.widgets['api_lbl'] = api_lbl

    api_var = tk.StringVar(root)
    current_key = load_api_key()
    api_var.set(current_key)

    api_entry = tk.Entry(api_frame, textvariable=api_var, font=("Segoe UI", 10), relief=tk.FLAT, bd=2, show="*")
    api_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
    theme.widgets['api_entry'] = api_entry
    
    frame = tk.Frame(root) # Borda sutil
    theme.widgets['border_frame'] = frame
    
    text_area = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 11), relief=tk.FLAT, padx=12, pady=12)
    text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=1, pady=1)
    theme.widgets['text_area'] = text_area

    scrollbar = tk.Scrollbar(frame, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

    # Carrega o prompt atual
    current_prompt = load_prompt()
    text_area.insert(tk.END, current_prompt)

    # Estilos de botão
    btn_font = ("Segoe UI", 10, "bold")
    
    save_btn = tk.Button(btn_frame, text="Salvar Configuração", width=20, bg="#2563eb", fg="white", font=btn_font, relief=tk.FLAT, activebackground="#1d4ed8", activeforeground="white", cursor="hand2", command=lambda: save_prompt(text_area, root, model_var, api_var))
    save_btn.pack(side=tk.RIGHT, padx=25)
    
    cancel_btn = tk.Button(btn_frame, text="Cancelar", width=15, font=btn_font, relief=tk.FLAT, cursor="hand2", command=root.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=5)

    restore_btn = tk.Button(btn_frame, text="Restaurar Padrão", width=18, font=btn_font, relief=tk.FLAT, cursor="hand2", command=lambda: restore_default(text_area))
    restore_btn.pack(side=tk.LEFT, padx=25)

    theme.widgets['sec_btns'] = [cancel_btn, restore_btn]
    
    # Aplica o tema inicial
    theme.apply()

    # Pack in order so btn_frame is fixed at the bottom
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=25)
    frame.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=25, pady=(0, 10))

    root.mainloop()

if __name__ == "__main__":
    main()
