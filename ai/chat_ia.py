import os
import threading
import tkinter as tk
from tkinter import scrolledtext
from pathlib import Path

import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_gemini_key import get_api_key, delete_api_key

LOGGER = configure_component_logger("chat_ia")

_DEFAULT_MODEL = 'gemini-2.5-flash'
_MAX_CONTEXT_CHARS = 150_000

# ==========================================
# Classe Principal do Chat
# ==========================================
class ChatApp:
    def __init__(self, root: tk.Tk) -> None:
        LOGGER.info("Initializing ChatApp UI")
        self.root = root
        self.word_app = None
        self.root.title("Chat com a IA - Z7 StdProposers")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        chat_width = int((screen_width // 4) * 0.9)
        chat_height = int(screen_height * 0.75)
        chat_left = 0
        chat_top = (screen_height - chat_height) // 2 + int(screen_height * 0.05)  # Centraliza e desloca 5% para baixo
        
        self.chat_width_px = chat_width
        self.root.geometry(f"{chat_width}x{chat_height}+{chat_left}+{chat_top}")
        self.root.minsize(300, 500)
        self.root.attributes('-topmost', True)
        
        self.resize_word_window(screen_width, screen_height)
        
        self.mode = z7_theme.load_theme()
        self.client = None
        self.chat_session = None
        self.is_generating = False
        self._cancel_requested = False
        self.last_ai_reply = ""
        self._model = _DEFAULT_MODEL
        self.doc_text = ""
        self._doc_truncated = False
        self._context_pending = False  # True when initial context send failed; injected on first user message

        self.build_ui()
        self.apply_theme()
        
        # Privacy Warning before starting AI
        if not z7_theme.ask_privacy_warning(
            "Aviso de Privacidade - Z7 StdProposers",
            "O texto do seu documento atual será enviado para a API do Google Gemini para servir de contexto do chat.\n\n"
            "Certifique-se de que não há dados sigilosos e que o uso está de acordo com as diretrizes do seu órgão.\n\n"
            "Deseja iniciar o assistente?",
            key="chat_ia",
            parent=self.root
        ):
            LOGGER.info("User cancelled chat at privacy warning")
            self.root.destroy()
            return

        self._load_doc_text_main_thread()
        self.init_ai()

    def resize_word_window(self, screen_width: int, screen_height: int) -> None:
        try:
            import win32com.client
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                word = win32com.client.GetObject(Class="Word.Application")

            self.word_app = word
                
            word.WindowState = 1  # wdWindowStateNormal
            word.Left = 0
            word.Top = 0
            word.StatusBar = "Z7: Aguardando interação no Chat..."
            
            # A API COM do Word trabalha em 'Points', enquanto o Tkinter usa 'Pixels'.
            # Precisamos converter para que a janela não fique gigante ou pequena demais.
            target_width_px = screen_width - self.chat_width_px
            target_height_px = screen_height
            
            try:
                word.Left = word.PixelsToPoints(self.chat_width_px)
            except Exception:
                word.Left = self.chat_width_px * 0.75
            
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

    def _set_word_status(self, message: str) -> None:
        """Atualiza StatusBar do Word sem interromper o fluxo do chat."""
        try:
            if self.word_app is not None:
                self.word_app.StatusBar = message
        except Exception:
            pass
        
    def apply_theme(self) -> None:
        colors = z7_theme.get_theme_colors(self.mode)
        bg = colors["bg"]
        fg = colors["fg"]
        fg_muted = colors["fg_muted"]
        text_bg = colors["text_bg"]
        btn_sec_hover = colors["btn_sec_hover"]
        btn_primary_bg = colors["btn_primary_bg"]
        user_tag_color = colors["user_tag"]
        ai_tag_color = colors["ai_tag"]
            
        self.root.configure(bg=bg)
        self.top_frame.configure(bg=bg)
        self.title_lbl.configure(bg=bg, fg=fg)
        self.status_lbl.configure(bg=bg, fg=fg_muted)
        
        self.chat_area.configure(bg=text_bg, fg=fg, insertbackground=fg)
        self.input_frame.configure(bg=bg)
        self.input_text.configure(bg=text_bg, fg=fg, insertbackground=fg)
        self.send_btn.configure(bg=btn_primary_bg, fg="white", activebackground=btn_sec_hover, activeforeground="white")

        if hasattr(self, 'action_frame'):
            self.action_frame.configure(bg=bg)
        for btn in [getattr(self, 'new_conv_btn', None), getattr(self, 'copy_btn', None),
                    getattr(self, 'grammar_btn', None), getattr(self, 'consistency_btn', None)]:
            if btn:
                btn.configure(bg=bg, fg=fg_muted, activebackground=btn_sec_hover, activeforeground=fg)

        self.chat_area.tag_config("user_tag", font=("Segoe UI", 11, "bold"), foreground=user_tag_color)
        self.chat_area.tag_config("ai_tag", font=("Segoe UI", 11, "bold"), foreground=ai_tag_color)
        self.chat_area.tag_config("sys_tag", font=("Segoe UI", 10, "italic"), foreground=fg_muted)

    def build_ui(self) -> None:
        # Top Frame
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=(15, 10), padx=20)

        self.title_lbl = tk.Label(self.top_frame, text="Assistente de IA", font=("Segoe UI", 16, "bold"))
        self.title_lbl.pack(side=tk.TOP, anchor="w")

        self.status_lbl = tk.Label(self.top_frame, text="Carregando contexto...", font=("Segoe UI", 10, "italic"))
        self.status_lbl.pack(side=tk.TOP, anchor="w", pady=(2, 0))

        # Botões de ação rápida
        self.action_frame = tk.Frame(self.top_frame)
        self.action_frame.pack(side=tk.TOP, anchor="w", pady=(6, 0))

        self.new_conv_btn = tk.Button(
            self.action_frame, text="↺ Nova Conversa",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            command=self.new_conversation
        )
        self.new_conv_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.copy_btn = tk.Button(
            self.action_frame, text="⎘ Copiar Última Resposta",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            command=self.copy_last_reply
        )
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.grammar_btn = tk.Button(
            self.action_frame, text="📝 Corrigir Gramática",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            command=self.run_grammar_check
        )
        self.grammar_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.consistency_btn = tk.Button(
            self.action_frame, text="🔍 Verificar Consistência",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            command=self.run_consistency_check
        )
        self.consistency_btn.pack(side=tk.LEFT)

        # Input Area (packed before Chat Area to stay visible at the bottom)
        self.input_frame = tk.Frame(self.root)
        self.input_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(10, 20))

        self.input_text = tk.Text(self.input_frame, wrap=tk.WORD, height=3, font=("Segoe UI", 11), relief=tk.FLAT, padx=12, pady=12)
        self.input_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        self.input_text.bind("<Return>", self.on_enter)
        self.input_text.bind("<Shift-Return>", self.on_shift_enter)

        # Botão único: "Enviar" em repouso, "Cancelar" enquanto gera
        self.send_btn = tk.Button(self.input_frame, text="Enviar", width=12, font=("Segoe UI", 11, "bold"), relief=tk.FLAT, cursor="hand2", command=self.send_or_cancel)
        self.send_btn.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))

        # Chat Area
        self.chat_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, font=("Segoe UI", 11), relief=tk.FLAT, padx=15, pady=15, state=tk.DISABLED)
        self.chat_area.pack(expand=True, fill=tk.BOTH, padx=20, pady=5)

    def append_message(self, role: str, message: str) -> None:
        self.chat_area.config(state=tk.NORMAL)
        if role == "User":
            self.chat_area.insert(tk.END, "Você:\n", "user_tag")
            self.chat_area.insert(tk.END, f"{message}\n\n", "user_msg")
        elif role == "Sistema":
            self.chat_area.insert(tk.END, f"⚠ {message}\n\n", "sys_tag")
        else:
            self.chat_area.insert(tk.END, "Gemini:\n", "ai_tag")
            self.chat_area.insert(tk.END, f"{message}\n\n", "ai_msg")

        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def send_or_cancel(self) -> None:
        if self.is_generating:
            self._cancel_requested = True
            self.send_btn.config(text="Enviar")
            self.status_lbl.config(text="Cancelando...")
        else:
            self.send_message()

    def copy_last_reply(self) -> None:
        if not self.last_ai_reply:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_ai_reply)
        old_status = self.status_lbl.cget("text")
        self.status_lbl.config(text="Resposta copiada!")
        self.root.after(2000, lambda: self.status_lbl.config(text=old_status))

    def _reload_doc_text(self) -> bool:
        """Tenta atualizar self.doc_text com o conteúdo atual do Word na thread principal."""
        try:
            if not self.word_app:
                return False
            raw_text = self.word_app.ActiveDocument.Content.Text
            if len(raw_text) > _MAX_CONTEXT_CHARS:
                cut = raw_text.rfind(' ', 0, _MAX_CONTEXT_CHARS)
                if cut == -1:
                    cut = _MAX_CONTEXT_CHARS
                self.doc_text = raw_text[:cut]
                self._doc_truncated = True
            else:
                self.doc_text = raw_text
                self._doc_truncated = False
            return True
        except Exception as e:
            LOGGER.warning(f"Failed to reload document text: {e}")
            return False

    def run_grammar_check(self) -> None:
        if self.is_generating or not self.chat_session:
            return
        
        self._reload_doc_text()
        
        self.input_text.delete("1.0", tk.END)
        display_msg = "Por favor, corrija a gramática do documento."
        self.append_message("User", display_msg)

        self.is_generating = True
        self._cancel_requested = False
        self.status_lbl.config(text="IA corrigindo gramática...")
        self.send_btn.config(text="Cancelar")

        threading.Thread(target=self._run_task_thread, args=("grammar",), daemon=True).start()

    def run_consistency_check(self) -> None:
        if self.is_generating or not self.chat_session:
            return

        self._reload_doc_text()

        self.input_text.delete("1.0", tk.END)
        display_msg = "Por favor, verifique a consistência do documento."
        self.append_message("User", display_msg)

        self.is_generating = True
        self._cancel_requested = False
        self.status_lbl.config(text="IA verificando consistência...")
        self.send_btn.config(text="Cancelar")

        threading.Thread(target=self._run_task_thread, args=("consistency",), daemon=True).start()

    def _run_task_thread(self, task_type: str) -> None:
        try:
            from config_prompt import load_prompt, load_consistency_prompt
            
            if not self.doc_text or "nenhum documento" in self.doc_text.lower():
                reply = "Não há contexto de documento carregado para realizar esta tarefa."
                self.root.after(0, self._on_message_received, reply)
                return

            if task_type == "grammar":
                base_prompt = load_prompt()
            else:
                base_prompt = load_consistency_prompt()
                
            prompt = f"{base_prompt}\n\n---INICIO DO DOCUMENTO---\n{self.doc_text}\n---FIM DO DOCUMENTO---\n"
            
            LOGGER.info(f"Sending {task_type} task to Gemini chat")
            
            response = self.chat_session.send_message(prompt)
            reply = response.text
        except Exception as e:
            log_exception(LOGGER, f"{task_type} task failed", e)
            self._set_word_status(f"Z7: Erro na tarefa {task_type}.")
            reply = f"Erro ao processar a solicitação: {str(e)}"
            
        self.root.after(0, self._on_message_received, reply)

    def new_conversation(self) -> None:
        if self.is_generating or not self.client:
            return
        self.chat_area.config(state=tk.NORMAL)
        self.chat_area.delete("1.0", tk.END)
        self.chat_area.config(state=tk.DISABLED)
        self.last_ai_reply = ""
        self.chat_session = None
        self.status_lbl.config(text="Iniciando nova conversa...")
        threading.Thread(target=self._new_conversation_thread, daemon=True).start()

    def _new_conversation_thread(self) -> None:
        try:
            from google.genai import types
            system_instruction = "Você é um assistente especialista em legislação prestativo e polido. Auxilie o usuário alterando, revisando ou tirando dúvidas."
            self.chat_session = self.client.chats.create(
                model=self._model,
                config=types.GenerateContentConfig(system_instruction=system_instruction)
            )
            try:
                if self.doc_text and "nenhum documento" not in self.doc_text.lower():
                    ctx = (
                        f"Abaixo está o texto do documento no Word (contexto desta conversa):\n\n"
                        f"{self.doc_text}\n\n"
                        "Esta é uma nova conversa. Confirme brevemente que está pronto."
                    )
                    greeting = self.chat_session.send_message(ctx).text
                else:
                    greeting = "Nova conversa iniciada. Como posso ajudar?"
            except Exception as ctx_e:
                log_exception(LOGGER, "Failed to send context in new conversation", ctx_e)
                greeting = "Nova conversa iniciada. Como posso ajudar?"
            self.root.after(0, lambda g=greeting: self._on_new_conversation_ready(g))
        except Exception as e:
            log_exception(LOGGER, "Failed to start new conversation", e)
            self.root.after(0, lambda: self.status_lbl.config(text="Erro ao iniciar nova conversa."))

    def _on_new_conversation_ready(self, greeting: str) -> None:
        self.status_lbl.config(text="Pronto para conversar")
        self.append_message("AI", greeting)
        LOGGER.info("New conversation started")

    def on_enter(self, event: tk.Event) -> str:
        self.send_message()
        return "break"
        
    def on_shift_enter(self, event: tk.Event) -> str:
        self.input_text.insert(tk.INSERT, "\n")
        return "break"

    def _load_doc_text_main_thread(self) -> None:
        """Lê o texto do documento ativo na thread principal (COM funciona corretamente aqui)."""
        import win32com.client
        try:
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                word = win32com.client.GetObject(Class="Word.Application")

            self.word_app = word

            try:
                word.Application.Run("CreateDocumentBackup", word.ActiveDocument)
                LOGGER.info("Document backup created successfully.")
            except Exception as backup_e:
                LOGGER.warning("Could not run CreateDocumentBackup macro: %s", str(backup_e))

            raw_text = word.ActiveDocument.Content.Text
            if len(raw_text) > _MAX_CONTEXT_CHARS:
                cut = raw_text.rfind(' ', 0, _MAX_CONTEXT_CHARS)
                if cut == -1:
                    cut = _MAX_CONTEXT_CHARS
                self.doc_text = raw_text[:cut]
                self._doc_truncated = True
                LOGGER.warning("Document context truncated at %d chars", cut)
            else:
                self.doc_text = raw_text
            LOGGER.info("Loaded Word active document context for chat")
        except Exception as e:
            log_exception(LOGGER, "Failed to load Word document context", e)
            self._set_word_status("Z7: Erro ao carregar contexto do documento no Chat.")
            self.doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."

    def init_ai(self) -> None:
        # Inicialização em background para não travar a UI
        threading.Thread(target=self._init_ai_thread, daemon=True).start()
        
    def _init_ai_thread(self) -> None:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client
            import google.genai as genai
            from google.genai import types

            api_key = get_api_key(self.root)
            if not api_key:
                LOGGER.error("Chat initialization aborted: missing API key")
                self._set_word_status("Z7: Erro no Chat - chave da API indisponivel.")
                self.root.after(0, lambda: self.status_lbl.config(text="Erro: Chave API ausente."))
                return

            self.client = genai.Client(api_key=api_key, http_options={'timeout': 60_000})

            # --- Texto do documento (carregado na thread principal via _load_doc_text_main_thread) ---
            _doc_truncated = self._doc_truncated

            # --- Carrega modelo ---
            try:
                from config_prompt import load_ai_model
                self._model = load_ai_model()
                LOGGER.info("Model loaded via config_prompt: %s", self._model)
            except Exception as e:
                log_exception(LOGGER, "Failed to load selected model for chat_ia", e)
                self._model = _DEFAULT_MODEL

            system_instruction = "Você é um assistente especialista em legislação prestativo e polido. Auxilie o usuário alterando, revisando ou tirando dúvidas."
            self.chat_session = self.client.chats.create(
                model=self._model,
                config=types.GenerateContentConfig(system_instruction=system_instruction)
            )
            LOGGER.info("Chat session started with model: %s", self._model)

            # --- Envia contexto do documento ---
            # Isolado em try/except próprio: falha aqui não deve travar o chat (#3)
            _truncation_notice = _doc_truncated
            try:
                doc_context = self.doc_text
                if doc_context and "nenhum documento" not in doc_context.lower():
                    context_msg = (
                        f"Abaixo está o texto atual do meu documento no Word para ser usado como base e contexto dessa conversa:\n\n"
                        f"{doc_context}\n\n"
                        "Apresente-se brevemente, confirme que leu o documento e aguarde meu primeiro pedido."
                    )
                    context_response = self.chat_session.send_message(context_msg)
                    self.initial_greeting = context_response.text
                    LOGGER.info("Document context sent to AI during initialization")
                else:
                    self.initial_greeting = "Olá! Não consegui acessar o documento atual. Como posso ajudar?"
                    _truncation_notice = False
                    LOGGER.warning("No document context available for AI initialization")
            except Exception as ctx_e:
                log_exception(LOGGER, "Failed to send document context to AI", ctx_e)
                self._set_word_status("Z7: Erro ao enviar contexto do documento ao Chat.")
                self.initial_greeting = "Olá! Não foi possível enviar o contexto do documento agora, mas ele será incluído na sua primeira mensagem. Como posso ajudar?"
                self._context_pending = True
                _truncation_notice = False

            def _on_ready_with_notice(notice: bool) -> None:
                self._on_ai_ready()
                if notice:
                    self.append_message(
                        "Sistema",
                        f"O documento excede o limite de contexto ({_MAX_CONTEXT_CHARS:,} caracteres). "
                        "Apenas o trecho inicial foi enviado à IA."
                    )

            self.root.after(0, lambda n=_truncation_notice: _on_ready_with_notice(n))

        except Exception as e:
            log_exception(LOGGER, "Failed to initialize chat AI", e)
            self._set_word_status("Z7: Erro ao inicializar o Chat IA.")
            error_msg = str(e).lower()
            if "403" in error_msg or "401" in error_msg or "invalid api key" in error_msg or "api_key" in error_msg:
                self.root.after(0, lambda: self.status_lbl.config(text="Erro: Chave Inválida."))
                self.root.after(0, lambda: self.append_message("Sistema", "Sua chave da API parece inválida ou expirou. Abra as Configurações da IA para atualizar ou remover a chave."))
            else:
                self.root.after(0, lambda: self.status_lbl.config(text="Erro na inicialização."))
                self.root.after(0, lambda: self.append_message("Sistema", f"Erro crítico: {str(e)}"))
        finally:
            pythoncom.CoUninitialize()

    def _on_ai_ready(self) -> None:
        self.status_lbl.config(text="Pronto para conversar")
        self.append_message("AI", getattr(self, 'initial_greeting', "Olá! Como posso ajudar?"))

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
        self._cancel_requested = False
        self.status_lbl.config(text="IA digitando...")
        self.send_btn.config(text="Cancelar")

        threading.Thread(target=self._send_message_thread, args=(user_msg,), daemon=True).start()

    def _send_message_thread(self, user_msg: str) -> None:
        try:
            LOGGER.info("Sending message to Gemini chat")
            if self._context_pending and self.doc_text and "nenhum documento" not in self.doc_text.lower():
                self._context_pending = False
                LOGGER.info("Injecting deferred document context with first user message")
                combined = (
                    f"Contexto do documento (Word):\n\n{self.doc_text}\n\n---\n\n{user_msg}"
                )
                response = self.chat_session.send_message(combined)
            else:
                response = self.chat_session.send_message(user_msg)
            reply = response.text
        except Exception as e:
            log_exception(LOGGER, "Chat message request failed", e)
            self._set_word_status("Z7: Erro de comunicacao no Chat IA.")
            reply = f"Erro de comunicação: {str(e)}"
            
        self.root.after(0, self._on_message_received, reply)

    def _on_message_received(self, reply: str) -> None:
        self.is_generating = False
        self.send_btn.config(text="Enviar")
        self.input_text.focus_set()
        if self._cancel_requested:
            self._cancel_requested = False
            self.status_lbl.config(text="Pronto para conversar")
            LOGGER.info("Response discarded after user cancel")
            return
        self.last_ai_reply = reply
        self.append_message("AI", reply)
        self.status_lbl.config(text="Pronto para conversar")

def main() -> None:
    root = tk.Tk()
    app = ChatApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
