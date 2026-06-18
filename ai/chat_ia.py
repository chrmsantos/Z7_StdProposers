import threading
import tkinter as tk
from tkinter import scrolledtext

import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_gemini_key import get_api_key

LOGGER = configure_component_logger("chat_ia")

_DEFAULT_MODEL = 'gemini-3.5-flash'
_MAX_CONTEXT_CHARS = 150_000

_APP_VERSION = "7.9.8"
_APP_AUTHOR  = "CMS"
_ORG         = "Câmara Municipal de Santa Bárbara d'Oeste"
_LICENSE     = "GPL-3.0"
_MOTTO       = "Dharma, virtude e gratidão."

def get_today_date_text() -> str:
    """Retorna a data atual formatada como 'Hoje é DD de MM de AAAA.'"""
    import datetime
    now = datetime.datetime.now()
    months = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    month_name = months[now.month]
    return f"Hoje é {now.day} de {month_name} de {now.year}."

# ==========================================
# Classe Principal do Chat
# ==========================================
class ChatApp:
    def __init__(self, root: tk.Tk) -> None:
        LOGGER.info("Initializing ChatApp UI")
        self.root = root
        self.word_app = None
        self.root.title(f"Chat com a IA — Z7 StdProposers v{_APP_VERSION}")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        chat_width = int(screen_width * 2 / 3 * 1.15)
        chat_height = int(screen_height * 0.92 * 0.90)
        chat_left = 0
        chat_top = int(screen_height * 0.02)
        
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
        self._context_pending = False
        self.current_status_text = "Carregando contexto..."

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
            word.StatusBar = "Z7: Aguardando interacao no Chat..."
            
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
        bg        = colors["bg"]
        fg        = colors["fg"]
        fg_muted  = colors["fg_muted"]
        text_bg   = colors["text_bg"]
        border    = colors["border"]
        btn_sec_hover    = colors["btn_sec_hover"]
        btn_sec_bg       = colors["btn_sec_bg"]
        btn_primary_bg   = colors["btn_primary_bg"]
        user_tag_color   = colors["user_tag"]
        ai_tag_color     = colors["ai_tag"]

        self.root.configure(bg=bg)

        # Header
        self.top_frame.configure(bg=bg)
        self.title_lbl.configure(bg=bg, fg=fg)
        self.action_frame.configure(bg=bg)
        self.header_sep.configure(bg=border)

        # Action buttons (top)
        for btn in [self.new_conv_btn, self.copy_btn]:
            btn.configure(bg=btn_sec_bg, fg=fg, activebackground=btn_sec_hover, activeforeground=fg)

        # Analysis buttons (bottom bar)
        self.analysis_frame.configure(bg=bg)
        self.analysis_sep.configure(bg=border)
        for btn in [self.grammar_btn, self.consistency_btn]:
            btn.configure(bg=btn_sec_bg, fg=fg, activebackground=btn_sec_hover, activeforeground=fg)

        # Chat area
        self.chat_border.configure(bg=border)
        self.chat_area.configure(bg=text_bg, fg=fg, insertbackground=fg)

        # Input area
        self.input_outer.configure(bg=bg)
        self.input_sep.configure(bg=border)
        self.input_border.configure(bg=border)
        self.input_text.configure(bg=text_bg, fg=fg, insertbackground=fg)
        self.send_btn.configure(
            bg=btn_primary_bg, fg="white",
            activebackground=colors["btn_primary_hover"], activeforeground="white"
        )

        # Rodapé
        self.footer_lbl.configure(bg=bg, fg=fg_muted)

        # Chat message tags
        self.chat_area.tag_config("user_tag", font=("Segoe UI", 10, "bold"), foreground=user_tag_color)
        self.chat_area.tag_config("user_msg", font=("Segoe UI", 11), foreground=fg,
            background=colors["user_bubble_bg"], lmargin1=10, lmargin2=10, rmargin=10,
            spacing1=4, spacing3=4)
        self.chat_area.tag_config("ai_tag",   font=("Segoe UI", 10, "bold"), foreground=ai_tag_color)
        self.chat_area.tag_config("ai_msg",   font=("Segoe UI", 11), foreground=fg,
            background=colors["ai_bubble_bg"], lmargin1=10, lmargin2=10, rmargin=10,
            spacing1=4, spacing3=4)
        self.chat_area.tag_config("sys_tag",  font=("Segoe UI", 10, "italic"), foreground=fg_muted)

        # Refresh the status badge with current mode
        self.update_status(self.current_status_text)

    def update_status(self, text: str) -> None:
        """Atualiza o texto e a aparência do status de acordo com o estado atual."""
        self.current_status_text = text
        
        # Determina a cor baseada no conteúdo do status
        colors = z7_theme.get_theme_colors(self.mode)
        text_lower = text.lower()
        
        if "erro" in text_lower or "inválida" in text_lower or "ausente" in text_lower:
            # Estado de erro: vermelho destacado
            fg_color = "#ef4444"
            bg_color = "#fef2f2" if self.mode == "light" else "#450a0a"
            border_color = "#fca5a5" if self.mode == "light" else "#991b1b"
            indicator = "● "
        elif "pronto" in text_lower or "copiada" in text_lower or "recebi" in text_lower:
            # Estado pronto/sucesso: verde
            fg_color = "#10b981" if self.mode == "light" else "#34d399"
            bg_color = "#ecfdf5" if self.mode == "light" else "#064e3b"
            border_color = "#a7f3d0" if self.mode == "light" else "#065f46"
            indicator = "● "
        elif "digitando" in text_lower or "corrigindo" in text_lower or "verificando" in text_lower or "carregando" in text_lower or "iniciando" in text_lower or "cancelando" in text_lower:
            # Estado ocupado/trabalhando: azul/indigo
            fg_color = "#6366f1" if self.mode == "light" else "#a5b4fc"
            bg_color = "#f5f3ff" if self.mode == "light" else "#1e1b4b"
            border_color = "#c7d2fe" if self.mode == "light" else "#312e81"
            indicator = "◌ "
        else:
            # Estado padrão
            fg_color = colors["fg_muted"]
            bg_color = colors["bg"]
            border_color = colors["border"]
            indicator = ""
            
        def _apply():
            self.status_lbl.config(text=f"{indicator}{text}", fg=fg_color, bg=bg_color)
            if hasattr(self, "status_frame"):
                self.status_frame.config(bg=bg_color, highlightbackground=border_color, highlightcolor=border_color)
        
        try:
            self.root.after(0, _apply)
        except Exception:
            pass

    def build_ui(self) -> None:
        # ── Cabeçalho ────────────────────────────────────────────────────────
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=(18, 0), padx=24)

        self.title_lbl = tk.Label(
            self.top_frame, text="✨ Assistente de IA",
            font=("Segoe UI", 15, "bold")
        )
        self.title_lbl.pack(side=tk.TOP, anchor="w")

        self.status_frame = tk.Frame(
            self.top_frame,
            padx=8, pady=3,
            highlightthickness=1,
            bd=0
        )
        self.status_frame.pack(side=tk.TOP, anchor="w", pady=(6, 12))

        self.status_lbl = tk.Label(
            self.status_frame, text="Carregando contexto...",
            font=("Segoe UI", 9, "bold")
        )
        self.status_lbl.pack()

        self.action_frame = tk.Frame(self.top_frame)
        self.action_frame.pack(side=tk.TOP, anchor="w", pady=(0, 14))

        self.new_conv_btn = tk.Button(
            self.action_frame, text="↺  Nova Conversa",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            padx=10, pady=5, command=self.new_conversation
        )
        self.new_conv_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.copy_btn = tk.Button(
            self.action_frame, text="⎘  Copiar Resposta",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            padx=10, pady=5, command=self.copy_last_reply
        )
        self.copy_btn.pack(side=tk.LEFT)

        # Linha separadora sob o cabeçalho
        self.header_sep = tk.Frame(self.root, height=1)
        self.header_sep.pack(fill=tk.X)

        # ── Rod apé (empacotado BOTTOM primeiro: fica na base absoluta) ────────────────
        _footer_text = f"{_ORG}  ·  {_APP_AUTHOR}  ·  {_LICENSE}  ·  {_MOTTO}"
        self.footer_lbl = tk.Label(
            self.root, text=_footer_text,
            font=("Segoe UI", 8), anchor="center"
        )
        self.footer_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 4))

        # ── Área de entrada (empacotada BOTTOM após o rod apé) ─────────────────
        self.input_outer = tk.Frame(self.root)
        self.input_outer.pack(side=tk.BOTTOM, fill=tk.X, padx=24, pady=(8, 12))

        # Wrapper com borda visual para o campo de texto
        self.input_border = tk.Frame(self.input_outer, bd=1, relief=tk.SOLID)
        self.input_border.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        self.input_text = tk.Text(
            self.input_border, wrap=tk.WORD, height=3,
            font=("Segoe UI", 11), relief=tk.FLAT,
            padx=12, pady=10, bd=0
        )
        self.input_text.pack(expand=True, fill=tk.BOTH)
        self.input_text.bind("<Return>", self.on_enter)
        self.input_text.bind("<Shift-Return>", self.on_shift_enter)

        self.send_btn = tk.Button(
            self.input_outer, text="Enviar ➤", width=10,
            font=("Segoe UI", 11, "bold"), relief=tk.FLAT,
            cursor="hand2", padx=0, pady=0,
            command=self.send_or_cancel
        )
        self.send_btn.pack(side=tk.RIGHT, fill=tk.Y, padx=(12, 0))

        # Linha separadora acima da área de entrada
        self.input_sep = tk.Frame(self.root, height=1)
        self.input_sep.pack(side=tk.BOTTOM, fill=tk.X)

        # ── Botões de análise rápida (acima do input, abaixo do chat) ────────
        self.analysis_frame = tk.Frame(self.root)
        self.analysis_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=24, pady=(6, 6))

        self.grammar_btn = tk.Button(
            self.analysis_frame, text="📝  Corrigir Gramática",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            padx=10, pady=5, command=self.run_grammar_check
        )
        self.grammar_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.consistency_btn = tk.Button(
            self.analysis_frame, text="🔍  Verificar Consistência",
            font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
            padx=10, pady=5, command=self.run_consistency_check
        )
        self.consistency_btn.pack(side=tk.LEFT)

        # Linha separadora acima dos botões de análise
        self.analysis_sep = tk.Frame(self.root, height=1)
        self.analysis_sep.pack(side=tk.BOTTOM, fill=tk.X)

        # ── Área de chat ──────────────────────────────────────────────────────
        self.chat_border = tk.Frame(self.root, bd=1, relief=tk.SOLID)
        self.chat_border.pack(expand=True, fill=tk.BOTH, padx=24, pady=(8, 0))

        self.chat_area = scrolledtext.ScrolledText(
            self.chat_border, wrap=tk.WORD,
            font=("Segoe UI", 11), relief=tk.FLAT,
            padx=15, pady=12, state=tk.DISABLED, bd=0
        )
        self.chat_area.pack(expand=True, fill=tk.BOTH)

        # Ajusta a largura da janela para acomodar os botões lado a lado
        self.root.update_idletasks()
        btn_row_w = (
            self.grammar_btn.winfo_reqwidth()
            + 8
            + self.consistency_btn.winfo_reqwidth()
            + 48  # padx da janela
            + 20  # margem de conforto
        )
        if self.root.winfo_width() < btn_row_w:
            h = self.root.winfo_height()
            self.chat_width_px = btn_row_w
            self.root.geometry(f"{btn_row_w}x{h}+0+{self.root.winfo_y()}")

    def append_message(self, role: str, message: str) -> None:
        self.chat_area.config(state=tk.NORMAL)
        if role == "User":
            self.chat_area.insert(tk.END, "Você:\n", "user_tag")
            self.chat_area.insert(tk.END, f"{message}\n", "user_msg")
            self.chat_area.insert(tk.END, "\n")
        elif role == "Sistema":
            self.chat_area.insert(tk.END, f"⚠ {message}\n\n", "sys_tag")
        else:
            self.chat_area.insert(tk.END, "Gemini:\n", "ai_tag")
            self.chat_area.insert(tk.END, f"{message}\n", "ai_msg")
            self.chat_area.insert(tk.END, "\n")

        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def send_or_cancel(self) -> None:
        if self.is_generating:
            self._cancel_requested = True
            self.send_btn.config(text="Enviar ➤")
            self.update_status("Cancelando...")
        else:
            self.send_message()

    def copy_last_reply(self) -> None:
        if not self.last_ai_reply:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_ai_reply)
        old_status = self.current_status_text
        self.update_status("Resposta copiada!")
        self.root.after(2000, lambda: self.update_status(old_status))

    def _reload_doc_text(self) -> bool:
        """Tenta atualizar self.doc_text com o conteúdo atual do Word na thread principal."""
        try:
            if not self.word_app:
                import win32com.client
                try:
                    self.word_app = win32com.client.GetActiveObject("Word.Application")
                except Exception:
                    self.word_app = win32com.client.GetObject(Class="Word.Application")
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

    def load_context(self) -> None:
        """Recarrega o texto do documento e envia para a IA confirmar o novo contexto."""
        if self.is_generating or not self.chat_session:
            return

        success = self._reload_doc_text()

        if not success:
            self.append_message("Sistema", "Não foi possível conectar ao Word. Certifique-se de que há um documento aberto e tente novamente.")
            return

        if not self.doc_text or not self.doc_text.strip() or "nenhum documento" in self.doc_text.lower():
            self.append_message("Sistema", "Nenhum documento ativo encontrado no Word.")
            return

        self.append_message("Sistema", "📄 Enviando contexto do documento para a IA...")

        self.is_generating = True
        self._cancel_requested = False
        self.update_status("Carregando contexto na IA...")
        self.send_btn.config(text="Cancelar")

        threading.Thread(target=self._load_context_thread, daemon=True).start()

    def _load_context_thread(self) -> None:
        try:
            truncation_note = (
                f"\n\n⚠ Documento truncado em {_MAX_CONTEXT_CHARS:,} caracteres."
                if self._doc_truncated else ""
            )
            ctx_msg = (
                f"O conteúdo do documento foi atualizado. Abaixo está a versão mais recente:"
                f"{truncation_note}\n\n"
                f"---INICIO DO DOCUMENTO---\n{self.doc_text}\n---FIM DO DOCUMENTO---\n\n"
                "Por favor, confirme que recebeu e processou o contexto atualizado do documento."
            )
            LOGGER.info("Sending updated document context to Gemini chat")
            response = self.chat_session.send_message(ctx_msg)
            reply = response.text
        except Exception as e:
            log_exception(LOGGER, "Failed to send context to AI", e)
            self._set_word_status("Z7: Erro ao carregar contexto no Chat IA.")
            reply = f"Falha ao enviar o contexto para a IA: {str(e)}"

        self.root.after(0, self._on_message_received, reply)

    def run_grammar_check(self) -> None:
        if self.is_generating or not self.chat_session:
            return
        
        self._reload_doc_text()
        
        self.input_text.delete("1.0", tk.END)
        display_msg = "Por favor, corrija a gramática do documento."
        self.append_message("User", display_msg)

        self.is_generating = True
        self._cancel_requested = False
        self.update_status("IA corrigindo gramática...")
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
        self.update_status("IA verificando consistência...")
        self.send_btn.config(text="Cancelar")

        threading.Thread(target=self._run_task_thread, args=("consistency",), daemon=True).start()

    def _run_task_thread(self, task_type: str) -> None:
        try:
            from config_prompt import load_prompt, load_consistency_prompt
            
            if not self.doc_text or not self.doc_text.strip() or "nenhum documento" in self.doc_text.lower():
                reply = "Não há contexto de documento carregado para realizar esta tarefa."
                self.root.after(0, self._on_message_received, reply)
                return

            if task_type == "grammar":
                base_prompt = load_prompt()
            else:
                base_prompt = load_consistency_prompt()
                
            today_prefix = get_today_date_text()
            ignore_instruction = (
                "As strings \"$ANO$\" e \"$DATAATUALEXTENSO$\" devem ser ignoradas no processo de verificação de consistência "
                "de datas, não devendo ser comparadas com outras datas no restante do documento."
            )
            grammar_instruction = (
                "A verificação de consistência deve também verificar e apontar erros gramaticais graves."
            )
            normative_instruction = (
                "A verificação de consistência deverá verificar as referências normativas do documento sob os seguintes requisitos:\n"
                "- Se o documento/propositura for uma indicação, o texto deverá fazer referência expressa ao Art. 108 do Regimento Interno;\n"
                "- Se o documento/propositura for um Requerimento de Informações, o texto deverá fazer referência expressa ao Art. 10, Inciso X, da Lei Orgânica do município de Santa Bárbara d’Oeste, combinado com o Art. 63, Inciso IX, do mesmo diploma legal;\n"
                "- Se o documento/propositura for um Requerimento de Pesar, o texto deverá fazer referência expressa ao Art. 102, Inciso IV, do Regimento Interno;\n"
                "- Se o documento/propositura for uma Moção, o texto deverá fazer referência expressa ao Art. 92, do Capítulo IV, Título V, do Regimento Interno."
            )
            prompt = f"{today_prefix}\n{ignore_instruction}\n{grammar_instruction}\n{normative_instruction}\n\n{base_prompt}\n\n---INICIO DO DOCUMENTO---\n{self.doc_text}\n---FIM DO DOCUMENTO---\n"
            
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
        self.update_status("Iniciando nova conversa...")
        threading.Thread(target=self._new_conversation_thread, daemon=True).start()

    def _new_conversation_thread(self) -> None:
        try:
            from google.genai import types
            today_prefix = get_today_date_text()
            system_instruction = f"{today_prefix} Você é um assistente especialista em legislação prestativo e polido. Auxilie o usuário alterando, revisando ou tirando dúvidas."

            # Pré-popula histórico com contexto do documento sem chamada de API
            self._reload_doc_text()
            history = []
            if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
                ctx_user_msg = (
                    f"Abaixo está o texto do documento no Word (contexto desta conversa):\n\n"
                    f"{self.doc_text}"
                )
                history = [
                    types.Content(role="user",  parts=[types.Part(text=ctx_user_msg)]),
                    types.Content(role="model", parts=[types.Part(text="Entendido! Contexto atualizado.")]),
                ]
                greeting = "✅ Contexto atualizado! Como posso ajudar?"
                LOGGER.info("New conversation: document context pre-seeded in history")
            else:
                greeting = "💬 Nova conversa reiniciada! Como posso ajudar?"

            self.chat_session = self.client.chats.create(
                model=self._model,
                config=types.GenerateContentConfig(system_instruction=system_instruction),
                history=history,
            )
            self.root.after(0, lambda g=greeting: self._on_new_conversation_ready(g))
        except Exception as e:
            log_exception(LOGGER, "Failed to start new conversation", e)
            self.root.after(0, lambda: self.update_status("Erro ao iniciar nova conversa."))

    def _on_new_conversation_ready(self, greeting: str) -> None:
        self.update_status("Pronto para conversar")
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
            import google.genai as genai
            from google.genai import types

            api_key = get_api_key(self.root)
            if not api_key:
                LOGGER.error("Chat initialization aborted: missing API key")
                self._set_word_status("Z7: Erro no Chat - chave da API indisponivel.")
                self.root.after(0, lambda: self.update_status("Erro: Chave API ausente."))
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

            today_prefix = get_today_date_text()
            system_instruction = f"{today_prefix} Você é um assistente especialista em legislação prestativo e polido. Auxilie o usuário alterando, revisando ou tirando dúvidas."

            # --- Pré-popula o histórico com o contexto do documento (sem chamada de API) ---
            # Usar history=[] evita um round-trip de rede no início, eliminando a principal
            # causa de falha: timeout/erro de rede ao enviar um payload grande na inicialização.
            _truncation_notice = _doc_truncated
            doc_context = self.doc_text
            history = []
            if doc_context and doc_context.strip() and "nenhum documento" not in doc_context.lower():
                ctx_user_msg = (
                    f"Abaixo está o texto atual do meu documento no Word para ser usado como base e contexto dessa conversa:\n\n"
                    f"{doc_context}"
                )
                history = [
                    types.Content(role="user",  parts=[types.Part(text=ctx_user_msg)]),
                    types.Content(role="model", parts=[types.Part(text="Entendido! Recebi o contexto do documento e estou pronto para ajudar.")]),
                ]
                self.initial_greeting = "✅ Contexto do documento carregado! Como posso ajudar?"
                LOGGER.info("Document context pre-seeded in chat history (no API call required)")
            else:
                self.initial_greeting = "Olá! Não consegui acessar o documento atual. Como posso ajudar?"
                _truncation_notice = False
                LOGGER.warning("No document context available for AI initialization")

            self.chat_session = self.client.chats.create(
                model=self._model,
                config=types.GenerateContentConfig(system_instruction=system_instruction),
                history=history,
            )
            LOGGER.info("Chat session started with model: %s", self._model)

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
                self.root.after(0, lambda: self.update_status("Erro: Chave Inválida."))
                self.root.after(0, lambda: self.append_message("Sistema", "Sua chave da API parece inválida ou expirou. Abra as Configurações da IA para atualizar ou remover a chave."))
            else:
                self.root.after(0, lambda: self.update_status("Erro na inicialização."))
                self.root.after(0, lambda err=str(e): self.append_message("Sistema", f"Erro crítico: {err}"))
        finally:
            pythoncom.CoUninitialize()

    def _on_ai_ready(self) -> None:
        self.update_status("Pronto para conversar")
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
        self.update_status("IA digitando...")
        self.send_btn.config(text="Cancelar")

        threading.Thread(target=self._send_message_thread, args=(user_msg,), daemon=True).start()

    def _send_message_thread(self, user_msg: str) -> None:
        try:
            actual_msg = user_msg
            if getattr(self, '_context_pending', False):
                if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
                    actual_msg = (
                        f"[Contexto do documento]\n{self.doc_text}\n\n"
                        f"[Mensagem do usu\u00e1rio]\n{user_msg}"
                    )
                self._context_pending = False
            LOGGER.info("Sending message to Gemini chat")
            response = self.chat_session.send_message(actual_msg)
            reply = response.text
        except Exception as e:
            log_exception(LOGGER, "Chat message request failed", e)
            self._set_word_status("Z7: Erro de comunicacao no Chat IA.")
            reply = f"Erro de comunicação: {str(e)}"
            
        self.root.after(0, self._on_message_received, reply)

    def _on_message_received(self, reply: str) -> None:
        self.is_generating = False
        self.send_btn.config(text="Enviar ➤")
        self.input_text.focus_set()
        if self._cancel_requested:
            self._cancel_requested = False
            self.update_status("Pronto para conversar")
            LOGGER.info("Response discarded after user cancel")
            return
        self.last_ai_reply = reply
        self.append_message("AI", reply)
        self.update_status("Pronto para conversar")

def main() -> None:
    root = tk.Tk()
    ChatApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
