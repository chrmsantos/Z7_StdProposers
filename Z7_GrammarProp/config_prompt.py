import os
import json
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from z7_logging import configure_component_logger, log_exception

LOGGER = configure_component_logger("config_prompt")

DEFAULT_PROMPT = """Você é um especialista em revisão de textos legislativos no idioma Português do Brasil.
Abaixo está um trecho de uma propositura legislativa.
Sua tarefa é corrigir gramaticalmente o texto, realizando O MÍNIMO POSSÍVEL de alterações em relação ao original.
Mantenha o tom formal, o jargão jurídico/legislativo e a estrutura da frase intactos, corrigindo apenas erros de ortografia, concordância, regência, pontuação ou crase evidentes.
Não adicione ponto final se o texto original não possuir um.
Retorne APENAS o texto corrigido, sem adicionar nenhum comentário, explicação, formatação markdown ou aspas extras."""

def get_prompt_file_path():
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'gemini_prompt.txt'

def get_theme_file_path():
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'theme_config.json'

def load_prompt():
    prompt_file = get_prompt_file_path()
    if prompt_file and prompt_file.exists():
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                LOGGER.info("Loaded custom prompt file")
                return f.read()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom prompt", e)
    return DEFAULT_PROMPT

def get_model_file_path():
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'selected_model.txt'

def load_ai_model():
    model_file = get_model_file_path()
    if model_file and model_file.exists():
        try:
            with open(model_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom model", e)
    return "gemini-2.5-flash"

def save_ai_model(model_name):
    model_file = get_model_file_path()
    if model_file:
        try:
            with open(model_file, 'w', encoding='utf-8') as f:
                f.write(model_name)
            LOGGER.info(f"Model saved successfully: {model_name}")
        except Exception as e:
            log_exception(LOGGER, "Failed to save model", e)

def save_prompt(text_widget, root, model_var):
    new_prompt = text_widget.get("1.0", tk.END).strip()
    if not new_prompt:
        LOGGER.warning("Prompt save blocked because text is empty")
        messagebox.showwarning("Aviso", "O prompt não pode estar vazio.")
        return

    prompt_file = get_prompt_file_path()
    if prompt_file:
        try:
            with open(prompt_file, 'w', encoding='utf-8') as f:
                f.write(new_prompt)
            LOGGER.info("Prompt saved successfully")
            
            # Save the model
            save_ai_model(model_var.get())
            
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            root.destroy()
        except Exception as e:
            log_exception(LOGGER, "Failed to save config", e)
            messagebox.showerror("Erro", f"Erro ao salvar:\n{e}")

def restore_default(text_widget):
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, DEFAULT_PROMPT)

def load_theme():
    theme_file = get_theme_file_path()
    if theme_file and theme_file.exists():
        try:
            with open(theme_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get('theme', 'light')
        except Exception as e:
            log_exception(LOGGER, "Failed to load theme config", e)
    return 'light'

def save_theme(theme_mode):
    theme_file = get_theme_file_path()
    if theme_file:
        try:
            with open(theme_file, 'w', encoding='utf-8') as f:
                json.dump({'theme': theme_mode}, f)
            LOGGER.info("Theme saved: %s", theme_mode)
        except Exception as e:
            log_exception(LOGGER, "Failed to save theme config", e)

class AppTheme:
    def __init__(self, root):
        self.root = root
        self.mode = load_theme()
        self.widgets = {}

    def toggle(self):
        self.mode = 'dark' if self.mode == 'light' else 'light'
        save_theme(self.mode)
        self.apply()

    def apply(self):
        if self.mode == 'dark':
            bg = "#1e1e1e"
            fg = "#e4e4e4"
            fg_muted = "#a0a0a0"
            text_bg = "#252526"
            border = "#333333"
            btn_sec_bg = "#333333"
            btn_sec_fg = "#cccccc"
            btn_sec_hover = "#444444"
        else:
            bg = "#f3f4f6"
            fg = "#111827"
            fg_muted = "#4b5563"
            text_bg = "#ffffff"
            border = "#d1d5db"
            btn_sec_bg = "#e5e7eb"
            btn_sec_fg = "#374151"
            btn_sec_hover = "#d1d5db"
            
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
            self.widgets['model_dropdown'].configure(bg=text_bg, fg=fg)
            
        for btn in self.widgets.get('sec_btns', []):
            btn.configure(bg=btn_sec_bg, fg=btn_sec_fg, activebackground=btn_sec_hover, activeforeground=fg)
        
        if 'toggle_btn' in self.widgets:
            icon = "🌙 Modo Escuro" if self.mode == 'light' else "☀️ Modo Claro"
            self.widgets['toggle_btn'].configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)

def main():
    LOGGER.info("Starting prompt configuration UI")
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
    
    MODELS = [
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash",
        "gemini-2.0-pro-exp",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
    ]
    
    if current_model not in MODELS:
        MODELS.append(current_model)
        
    model_var.set(current_model)

    model_dropdown = tk.OptionMenu(model_frame, model_var, *MODELS)
    model_dropdown.config(font=("Segoe UI", 10), relief=tk.FLAT, bd=1, highlightthickness=1)
    model_dropdown.pack(side=tk.LEFT, padx=10)
    theme.widgets['model_dropdown'] = model_dropdown
    
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
    
    save_btn = tk.Button(btn_frame, text="Salvar Configuração", width=20, bg="#2563eb", fg="white", font=btn_font, relief=tk.FLAT, activebackground="#1d4ed8", activeforeground="white", cursor="hand2", command=lambda: save_prompt(text_area, root, model_var))
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
