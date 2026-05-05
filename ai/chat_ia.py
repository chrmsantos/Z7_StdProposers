import os
import json
import threading
import tkinter as tk
from tkinter import simpledialog, messagebox, scrolledtext
from pathlib import Path

from z7_logging import configure_component_logger, log_exception

LOGGER = configure_component_logger("chat_ia")

# ==========================================
# Funções de Configuração e Chave
# ==========================================
def get_api_key(root_for_dialog: tk.Tk | None = None) -> str | None:
    LOGGER.info("Loading API key for chat session")
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        LOGGER.error("USERPROFILE env var not found")
        return None
        
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_file = key_dir / 'gemini.key'
    
    if key_file.exists():
        try:
            import win32crypt
            with open(key_file, 'rb') as f:
                encrypted_key = f.read()
            _, decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)
            LOGGER.info("API key loaded from encrypted file")
            return decrypted_key.decode('utf-8')
        except Exception as e:
            log_exception(LOGGER, "Failed to decrypt API key", e)
            
    api_key = simpledialog.askstring(
        "Z7 StdProposers", 
        "Insira sua chave da API do Google Gemini:\n(Ela será criptografada e salva localmente)",
        parent=root_for_dialog
    )
    
    if not api_key or not api_key.strip():
        LOGGER.warning("User did not provide API key")
        return None
        
    api_key = api_key.strip()
    
    try:
        import win32crypt
        encrypted_key = win32crypt.CryptProtectData(api_key.encode('utf-8'), 'Z7_Gemini_Key', None, None, None, 0)
        key_dir.mkdir(parents=True, exist_ok=True)
        with open(key_file, 'wb') as f:
            f.write(encrypted_key)
        LOGGER.info("API key encrypted and persisted")
    except Exception as e:
        log_exception(LOGGER, "Failed to persist API key", e)
        
    return api_key

def get_theme_file_path() -> Path | None:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'theme_config.json'

def load_theme() -> str:
    theme_file = get_theme_file_path()
    if theme_file and theme_file.exists():
        try:
            with open(theme_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get('theme', 'light')
        except Exception as e:
            log_exception(LOGGER, "Failed to load theme config", e)
    return 'light'

def save_theme(theme_mode: str) -> None:
    theme_file = get_theme_file_path()
    if theme_file:
        try:
            with open(theme_file, 'w', encoding='utf-8') as f:
                json.dump({'theme': theme_mode}, f)
            LOGGER.info("Theme saved: %s", theme_mode)
        except Exception as e:
            log_exception(LOGGER, "Failed to save theme config", e)

# ==========================================
# Classe Principal do Chat
# ==========================================
class ChatApp:
    def __init__(self, root: tk.Tk) -> None:
        LOGGER.info("Initializing ChatApp UI")
        self.root = root
        self.root.title("Chat com a IA - Z7 StdProposers")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        chat_width = screen_width // 4
        chat_height = screen_height
        chat_left = screen_width - chat_width
        chat_top = 0
        
        self.root.geometry(f"{chat_width}x{chat_height}+{chat_left}+{chat_top}")
        self.root.minsize(300, 500)
        self.root.attributes('-topmost', True)
        
        self.resize_word_window(screen_width, screen_height)
        
        self.mode = load_theme()
        self.client = None
        self.chat_session = None
        self.is_generating = False
        
        self.build_ui()
        self.apply_theme()
        
        # Privacy Warning before starting AI
        if not messagebox.askokcancel(
            "Aviso de Privacidade - Z7 StdProposers",
            "O texto do seu documento atual será enviado para a API do Google Gemini para servir de contexto do chat.\n\n"
            "Certifique-se de que não há dados sigilosos e que o uso está de acordo com as diretrizes do seu órgão.\n\n"
            "Deseja iniciar o assistente?",
            parent=self.root
        ):
            LOGGER.info("User cancelled chat at privacy warning")
            self.root.destroy()
            return
            
        self.init_ai()

    def resize_word_window(self, screen_width: int, screen_height: int) -> None:
        try:
            import win32com.client
            try:
                word = win32com.client.GetObject(Class="Word.Application")
            except Exception:
                word = win32com.client.Dispatch("Word.Application")
                
            word.WindowState = 1  # wdWindowStateNormal
            word.Left = 0
            word.Top = 0
            word.StatusBar = False # Limpa o status de carregamento do VBA
            
            # A API COM do Word trabalha em 'Points', enquanto o Tkinter usa 'Pixels'.
            # Precisamos converter para que a janela não fique gigante ou pequena demais.
            target_width_px = screen_width * 3 // 4
            target_height_px = screen_height
            
            try:
                word.Width = word.PixelsToPoints(target_width_px)
                word.Height = word.PixelsToPoints(target_height_px, True)
            except Exception:
                # Fallback genérico de 96 DPI (1 px = 0.75 points)
                word.Width = target_width_px * 0.75
                word.Height = target_height_px * 0.75
                
            LOGGER.info("Word window resized to 3/4 of screen")
        except Exception as e:
            log_exception(LOGGER, "Failed to resize Word window", e)
        
    def toggle_theme(self) -> None:
        self.mode = 'dark' if self.mode == 'light' else 'light'
        save_theme(self.mode)
        self.apply_theme()

    def apply_theme(self) -> None:
        if self.mode == 'dark':
            bg = "#1e1e1e"
            fg = "#e4e4e4"
            fg_muted = "#a0a0a0"
            text_bg = "#252526"
            btn_sec_hover = "#444444"
            btn_primary_bg = "#005a9e"
            user_tag_color = "#60a5fa"
            ai_tag_color = "#34d399"
        else:
            bg = "#f3f4f6"
            fg = "#111827"
            fg_muted = "#4b5563"
            text_bg = "#ffffff"
            btn_sec_hover = "#d1d5db"
            btn_primary_bg = "#2563eb"
            user_tag_color = "#2563eb"
            ai_tag_color = "#10b981"
            
        self.root.configure(bg=bg)
        self.top_frame.configure(bg=bg)
        self.title_lbl.configure(bg=bg, fg=fg)
        self.status_lbl.configure(bg=bg, fg=fg_muted)
        
        icon = "🌙 Tema Escuro" if self.mode == 'light' else "☀️ Tema Claro"
        self.toggle_btn.configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)
        
        self.chat_area.configure(bg=text_bg, fg=fg, insertbackground=fg)
        self.input_frame.configure(bg=bg)
        self.input_text.configure(bg=text_bg, fg=fg, insertbackground=fg)
        self.send_btn.configure(bg=btn_primary_bg, fg="white", activebackground=btn_sec_hover, activeforeground="white")
        
        self.chat_area.tag_config("user_tag", font=("Segoe UI", 11, "bold"), foreground=user_tag_color)
        self.chat_area.tag_config("ai_tag", font=("Segoe UI", 11, "bold"), foreground=ai_tag_color)

    def build_ui(self) -> None:
        # Top Frame
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=(15, 10), padx=20)
        
        self.toggle_btn = tk.Button(self.top_frame, font=("Segoe UI", 10), relief=tk.FLAT, cursor="hand2", command=self.toggle_theme, bd=0)
        self.toggle_btn.pack(side=tk.LEFT)
        
        self.title_lbl = tk.Label(self.top_frame, text="Assistente de IA", font=("Segoe UI", 16, "bold"))
        self.title_lbl.pack(side=tk.LEFT, padx=20)
        
        self.status_lbl = tk.Label(self.top_frame, text="Carregando contexto...", font=("Segoe UI", 10, "italic"))
        self.status_lbl.pack(side=tk.RIGHT)
        
        # Chat Area
        self.chat_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, font=("Segoe UI", 11), relief=tk.FLAT, padx=15, pady=15, state=tk.DISABLED)
        self.chat_area.pack(expand=True, fill=tk.BOTH, padx=20, pady=5)
        
        # Input Area
        self.input_frame = tk.Frame(self.root)
        self.input_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(10, 20))
        
        self.input_text = tk.Text(self.input_frame, wrap=tk.WORD, height=3, font=("Segoe UI", 11), relief=tk.FLAT, padx=12, pady=12)
        self.input_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        self.input_text.bind("<Return>", self.on_enter)
        self.input_text.bind("<Shift-Return>", self.on_shift_enter)
        
        self.send_btn = tk.Button(self.input_frame, text="Enviar", width=12, font=("Segoe UI", 11, "bold"), relief=tk.FLAT, cursor="hand2", command=self.send_message)
        self.send_btn.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))

    def append_message(self, role: str, message: str) -> None:
        self.chat_area.config(state=tk.NORMAL)
        if role == "User":
            self.chat_area.insert(tk.END, "Você:\n", "user_tag")
            self.chat_area.insert(tk.END, f"{message}\n\n", "user_msg")
        else:
            self.chat_area.insert(tk.END, "Gemini:\n", "ai_tag")
            self.chat_area.insert(tk.END, f"{message}\n\n", "ai_msg")
            
        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def on_enter(self, event: tk.Event) -> str:
        self.send_message()
        return "break"
        
    def on_shift_enter(self, event: tk.Event) -> str:
        self.input_text.insert(tk.INSERT, "\n")
        return "break"

    def init_ai(self) -> None:
        # Inicialização em background para não travar a UI
        threading.Thread(target=self._init_ai_thread, daemon=True).start()
        
    def _init_ai_thread(self) -> None:
        try:
            import win32com.client
            import google.genai as genai
            from google.genai import types
            
            api_key = get_api_key(self.root)
            if not api_key:
                LOGGER.error("Chat initialization aborted: missing API key")
                self.root.after(0, lambda: self.status_lbl.config(text="Erro: Chave API ausente."))
                return
                
            self.client = genai.Client(api_key=api_key)
            
            try:
                try:
                    word = win32com.client.GetObject(Class="Word.Application")
                except Exception:
                    word = win32com.client.Dispatch("Word.Application")
                doc_text = word.ActiveDocument.Content.Text
                if len(doc_text) > 150000:
                    doc_text = doc_text[:150000] # Limita tamanho pra segurança
                LOGGER.info("Loaded Word active document context for chat")
            except Exception as e:
                log_exception(LOGGER, "Failed to load Word document context", e)
                doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."
                
            system_instruction = f"Você é um assistente especialista em legislação prestativo e polido. Use o seguinte texto do documento ativo no Word como contexto principal para responder às dúvidas do usuário:\n\n{doc_text}"
            
            # Check for custom model selection
            model = 'gemini-3.1-pro-preview' # Default
            try:
                user_profile = os.environ.get('USERPROFILE')
                if user_profile:
                    model_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'selected_model.txt'
                    if model_file.exists():
                        with open(model_file, 'r', encoding='utf-8') as f:
                            saved_model = f.read().strip()
                            if saved_model:
                                model = saved_model
            except Exception as e:
                log_exception(LOGGER, "Failed to load selected model for chat_ia", e)
                
            self.chat_session = self.client.chats.create(
                model=model,
                config=types.GenerateContentConfig(system_instruction=system_instruction)
            )
            LOGGER.info("Chat session started")
            
            self.root.after(0, self._on_ai_ready)
            
        except Exception as e:
            log_exception(LOGGER, "Failed to initialize chat AI", e)
            
            # Auto-reparo de chave caso seja um erro de autenticacao
            error_msg = str(e).lower()
            if "403" in error_msg or "401" in error_msg or "invalid api key" in error_msg or "api_key" in error_msg:
                try:
                    user_profile = os.environ.get('USERPROFILE')
                    key_file = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers' / 'gemini.key'
                    if key_file.exists():
                        key_file.unlink()
                        LOGGER.info("Invalid API key file deleted")
                except Exception as del_e:
                    pass
                self.root.after(0, lambda: self.status_lbl.config(text="Erro: Chave Inválida. Reinicie."))
                self.root.after(0, lambda: self.append_message("Sistema", "Sua chave da API é inválida ou expirou. O arquivo de chave foi deletado. Feche e abra o chat novamente para inserir uma nova chave."))
            else:
                self.root.after(0, lambda: self.status_lbl.config(text="Erro na inicialização."))
                self.root.after(0, lambda: self.append_message("Sistema", f"Erro crítico: {str(e)}"))

    def _on_ai_ready(self) -> None:
        self.status_lbl.config(text="Pronto para conversar")
        self.append_message("AI", "Olá! Li o seu documento atual. Como posso ajudar?")

    def send_message(self) -> None:
        if self.is_generating or not self.chat_session:
            LOGGER.info("Send skipped: generating=%s session_ready=%s", self.is_generating, self.chat_session is not None)
            return
            
        user_msg = self.input_text.get("1.0", tk.END).strip()
        if not user_msg:
            LOGGER.info("Send skipped: empty message")
            return
            
        self.input_text.delete("1.0", tk.END)
        self.append_message("User", user_msg)
        
        self.is_generating = True
        self.status_lbl.config(text="IA digitando...")
        self.send_btn.config(state=tk.DISABLED)
        
        threading.Thread(target=self._send_message_thread, args=(user_msg,), daemon=True).start()

    def _send_message_thread(self, user_msg: str) -> None:
        try:
            LOGGER.info("Sending message to Gemini chat")
            response = self.chat_session.send_message(user_msg)
            reply = response.text
        except Exception as e:
            log_exception(LOGGER, "Chat message request failed", e)
            reply = f"Erro de comunicação: {str(e)}"
            
        self.root.after(0, self._on_message_received, reply)

    def _on_message_received(self, reply: str) -> None:
        self.append_message("AI", reply)
        self.is_generating = False
        self.status_lbl.config(text="Pronto para conversar")
        self.send_btn.config(state=tk.NORMAL)
        self.input_text.focus_set()

def main() -> None:
    root = tk.Tk()
    app = ChatApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
