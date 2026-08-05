import threading
import tkinter as tk
from tkinter import scrolledtext

import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_api_key import get_api_key

LOGGER = configure_component_logger("chat_ia")

_DEFAULT_MODEL = 'deepseek/deepseek-chat'
_OPENROUTER_BASE_URL = 'https://openrouter.ai/api/v1'
_MAX_CONTEXT_CHARS = 150_000

_APP_VERSION = "8.0.6"
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
        
        chat_width = int(screen_width * 2 / 3 * 1.15 * 1.10 * 0.70)
        chat_height = int(screen_height * 0.92 * 0.90 * 1.05)
        chat_left = 0
        chat_top = int(screen_height * 0.02)
        
        self.chat_width_px = chat_width
        self.root.geometry(f"{chat_width}x{chat_height}+{chat_left}+{chat_top}")
        self.root.minsize(300, 500)
        self.root.attributes('-topmost', True)
        
        self.resize_word_window(screen_width, screen_height)
        
        self.mode = z7_theme.load_theme()
        self.client = None
        self.messages = []  # Histórico de mensagens no formato OpenAI
        self.system_instruction = ""
        self.is_generating = False
        self._cancel_requested = False
        self.last_ai_reply = ""
        self._model = _DEFAULT_MODEL
        self.doc_text = ""
        self._doc_truncated = False
        self._context_pending = False
        self._doc_load_error = ""
        self.current_status_text = "Carregando contexto..."
        self._update_status = "checking"

        self.build_ui()
        self.apply_theme()
        
        self._update_ai_status()
        self._check_for_updates_silent()

        self._load_doc_text_main_thread()
        self.init_ai()

    def resize_word_window(self, screen_width: int, screen_height: int) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize()
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
        btn_primary_bg   = colors["btn_primary_bg"]
        user_tag_color   = colors["user_tag"]
        ai_tag_color     = colors["ai_tag"]

        self.root.configure(bg=bg)

        # Header
        self.top_frame.configure(bg=bg)
        self.title_lbl.configure(bg=bg, fg=fg)
        if hasattr(self, 'settings_btn'):
            self.settings_btn.configure(
                bg=colors["btn_sec_bg"], fg=colors["btn_sec_fg"],
                activebackground=colors["btn_sec_hover"], activeforeground=fg
            )
        # Header sub-frames
        if hasattr(self, 'header_top'):
            self.header_top.configure(bg=bg)
        if hasattr(self, 'header_badges'):
            self.header_badges.configure(bg=bg)

        # Toolbar
        if hasattr(self, 'toolbar_frame'):
            self.toolbar_frame.configure(bg=bg)
            for btn_attr in ('new_chat_btn', 'reload_ctx_btn', 'copy_reply_btn'):
                btn = getattr(self, btn_attr, None)
                if btn:
                    btn.configure(
                        bg=colors["btn_sec_bg"], fg=colors["btn_sec_fg"],
                        activebackground=colors["btn_sec_hover"], activeforeground=fg
                    )

        # Separadores
        for sep_attr in ('header_sep', 'toolbar_sep'):
            sep = getattr(self, sep_attr, None)
            if sep:
                sep.configure(bg=border)

        # Chat area
        self.chat_border.configure(bg=border)
        select_bg = "#6366f1" if self.mode == "dark" else "#7c3aed"
        select_fg = "#ffffff" if self.mode == "dark" else "#ffffff"
        self.chat_area.configure(
            bg=text_bg, fg=fg, insertbackground=fg,
            selectbackground=select_bg, selectforeground=select_fg,
            inactiveselectbackground=select_bg
        )

        # Input area
        self.input_outer.configure(bg=bg)
        self.input_sep.configure(bg=border)
        self.input_border.configure(bg=border)
        self.input_text.configure(
            bg=text_bg, fg=fg, insertbackground=fg,
            selectbackground=select_bg, selectforeground=select_fg,
            inactiveselectbackground=select_bg
        )
        self.send_btn.configure(
            bg=btn_primary_bg, fg="white",
            activebackground=colors["btn_primary_hover"], activeforeground="white"
        )

        # Rodapé
        self.footer_lbl.configure(bg=bg, fg=fg_muted)

        # Input hint
        if hasattr(self, 'input_hint'):
            self.input_hint.configure(bg=bg, fg=fg_muted)

        # Chat message tags
        self.chat_area.tag_config("user_tag", font=("Segoe UI", 10, "bold"),
            foreground=user_tag_color, spacing1=6, spacing3=2)
        self.chat_area.tag_config("user_msg", font=("Segoe UI", 11), foreground=fg,
            background=colors["user_bubble_bg"], lmargin1=14, lmargin2=14, rmargin=14,
            spacing1=6, spacing3=8, relief="flat", borderwidth=0)
        self.chat_area.tag_config("ai_tag", font=("Segoe UI", 10, "bold"),
            foreground=ai_tag_color, spacing1=6, spacing3=2)
        self.chat_area.tag_config("ai_msg", font=("Segoe UI", 11), foreground=fg,
            background=colors["ai_bubble_bg"], lmargin1=14, lmargin2=14, rmargin=14,
            spacing1=6, spacing3=8, relief="flat", borderwidth=0)
        self.chat_area.tag_config("sys_tag", font=("Segoe UI", 10, "italic"),
            foreground=fg_muted, spacing1=4, spacing3=4)
        self.chat_area.tag_config("msg_sep", foreground=border,
            font=("Segoe UI", 6), spacing1=2, spacing3=6)

        # Ensure selection highlight is visible above all other tags
        self.chat_area.tag_config("sel", background=select_bg, foreground=select_fg)
        self.chat_area.tag_raise("sel")

        # Version badge
        if hasattr(self, 'version_badge'):
            self.version_badge.configure(bg=colors["btn_primary_bg"], fg="white")

        # Refresh all status badges with current mode
        self._set_update_status(self._update_status)
        self._update_ai_status()
        self.update_status(self.current_status_text)

    def update_status(self, text: str) -> None:
        """Atualiza o texto e a aparência do status de acordo com o estado atual."""
        self.current_status_text = text
        
        colors = z7_theme.get_theme_colors(self.mode)
        text_lower = text.lower()
        
        if "erro" in text_lower or "inválida" in text_lower or "ausente" in text_lower:
            fg_color = "#ef4444"
            bg_color = "#fef2f2" if self.mode == "light" else "#450a0a"
            border_color = "#fca5a5" if self.mode == "light" else "#991b1b"
            indicator = "● "
        elif "pronto" in text_lower or "copiada" in text_lower or "recebi" in text_lower:
            fg_color = "#10b981" if self.mode == "light" else "#34d399"
            bg_color = "#ecfdf5" if self.mode == "light" else "#064e3b"
            border_color = "#a7f3d0" if self.mode == "light" else "#065f46"
            indicator = "● "
        elif "digitando" in text_lower or "corrigindo" in text_lower or "verificando" in text_lower or "carregando" in text_lower or "iniciando" in text_lower or "cancelando" in text_lower:
            fg_color = "#6366f1" if self.mode == "light" else "#a5b4fc"
            bg_color = "#f5f3ff" if self.mode == "light" else "#1e1b4b"
            border_color = "#c7d2fe" if self.mode == "light" else "#312e81"
            indicator = "◌ "
        else:
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

    def _set_update_status(self, status: str) -> None:
        """Atualiza o badge de status de atualização."""
        self._update_status = status
        if not hasattr(self, 'update_status_lbl'):
            return

        colors = z7_theme.get_theme_colors(self.mode)

        if status == "checking":
            text = "⏳ Checando..."
            fg_color = colors["fg_muted"]
        elif status == "up_to_date":
            text = "✔ App atualizado"
            fg_color = "#10b981" if self.mode == "light" else "#34d399"
        elif status == "update_available":
            text = "🔄 Atualização disponível"
            fg_color = "#f59e0b" if self.mode == "light" else "#fbbf24"
        elif status == "error":
            text = "⚠ Erro ao checar"
            fg_color = "#ef4444"
        else:
            text = ""
            fg_color = colors["fg_muted"]

        def _apply():
            try:
                self.update_status_lbl.config(text=text, fg=fg_color, bg=colors["bg"])
            except Exception:
                pass

        try:
            self.root.after(0, _apply)
        except Exception:
            pass

    def _update_ai_status(self) -> None:
        """Atualiza o badge de status da IA (modelo + validação)."""
        if not hasattr(self, 'ai_status_lbl'):
            return

        model = self._model or _DEFAULT_MODEL
        has_client = self.client is not None

        colors = z7_theme.get_theme_colors(self.mode)

        if has_client:
            text = f"🤖 {model}  •  ✔ Validado"
            color = "#10b981" if self.mode == "light" else "#34d399"
        else:
            text = f"🤖 {model}  •  ⚠ Não validado"
            color = "#f59e0b" if self.mode == "light" else "#fbbf24"

        def _apply():
            try:
                self.ai_status_lbl.config(text=text, fg=color, bg=colors["bg"])
            except Exception:
                pass

        try:
            self.root.after(0, _apply)
        except Exception:
            pass

    def _check_for_updates_silent(self) -> None:
        """Verifica atualizações silenciosamente em background."""
        def _run():
            try:
                from config_prompt import get_latest_github_release, compare_versions
                data = get_latest_github_release()
                tag_name = data.get("tag_name", "").strip()
                if not tag_name:
                    tag_name = data.get("name", "").strip()
                if not tag_name:
                    self.root.after(0, lambda: self._set_update_status("error"))
                    return
                comparison = compare_versions(tag_name, _APP_VERSION)
                if comparison > 0:
                    self.root.after(0, lambda: self._set_update_status("update_available"))
                else:
                    self.root.after(0, lambda: self._set_update_status("up_to_date"))
            except Exception as e:
                LOGGER.warning("Silent update check failed: %s", e)
                self.root.after(0, lambda: self._set_update_status("error"))

        threading.Thread(target=_run, daemon=True).start()

    def build_ui(self) -> None:
        # ── Cabeçalho ────────────────────────────────────────────────────────
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=(16, 0), padx=20)

        # Linha 1: Título + botões de ação
        self.header_top = tk.Frame(self.top_frame)
        self.header_top.pack(fill=tk.X)

        self.title_lbl = tk.Label(
            self.header_top, text="✨ Assistente de IA",
            font=("Segoe UI", 16, "bold")
        )
        self.title_lbl.pack(side=tk.LEFT, anchor="w")

        self.settings_btn = tk.Button(
            self.header_top, text="⚙ Config",
            font=("Segoe UI", 9), relief=tk.FLAT,
            cursor="hand2", padx=8, pady=3,
            command=self._open_config_prompt
        )
        self.settings_btn.pack(side=tk.RIGHT, anchor="e")

        # Linha 2: Badges de status
        self.header_badges = tk.Frame(self.top_frame)
        self.header_badges.pack(fill=tk.X, pady=(6, 0))

        self.version_badge = tk.Label(
            self.header_badges, text=f"v{_APP_VERSION}",
            font=("Segoe UI", 9, "bold"), padx=6, pady=1
        )
        self.version_badge.pack(side=tk.LEFT, anchor="w")

        self.ai_status_lbl = tk.Label(
            self.header_badges, text="",
            font=("Segoe UI", 9)
        )
        self.ai_status_lbl.pack(side=tk.LEFT, anchor="w", padx=(8, 0))

        self.update_status_lbl = tk.Label(
            self.header_badges, text="⏳ Checando...",
            font=("Segoe UI", 9)
        )
        self.update_status_lbl.pack(side=tk.LEFT, anchor="w", padx=(8, 0))

        self.status_frame = tk.Frame(
            self.header_badges,
            padx=10, pady=4,
            highlightthickness=1,
            bd=0
        )
        self.status_frame.pack(side=tk.RIGHT, anchor="e")

        self.status_lbl = tk.Label(
            self.status_frame, text="Carregando contexto...",
            font=("Segoe UI", 9, "bold")
        )
        self.status_lbl.pack()

        # Separador após cabeçalho
        self.header_sep = tk.Frame(self.top_frame, height=1)
        self.header_sep.pack(fill=tk.X, pady=(12, 0))

        # ── Toolbar de ações rápidas ─────────────────────────────────────────
        self.toolbar_frame = tk.Frame(self.root)
        self.toolbar_frame.pack(side=tk.TOP, fill=tk.X, padx=20, pady=(8, 0))

        self.new_chat_btn = tk.Button(
            self.toolbar_frame, text="💬 Nova Conversa",
            font=("Segoe UI", 9), relief=tk.FLAT,
            cursor="hand2", padx=8, pady=3,
            command=self.new_conversation
        )
        self.new_chat_btn.pack(side=tk.LEFT)

        self.reload_ctx_btn = tk.Button(
            self.toolbar_frame, text="📄 Recarregar Contexto",
            font=("Segoe UI", 9), relief=tk.FLAT,
            cursor="hand2", padx=8, pady=3,
            command=self.load_context
        )
        self.reload_ctx_btn.pack(side=tk.LEFT, padx=(6, 0))

        self.copy_reply_btn = tk.Button(
            self.toolbar_frame, text="📋 Copiar Resposta",
            font=("Segoe UI", 9), relief=tk.FLAT,
            cursor="hand2", padx=8, pady=3,
            command=self.copy_last_reply
        )
        self.copy_reply_btn.pack(side=tk.LEFT, padx=(6, 0))

        # Separador após toolbar
        self.toolbar_sep = tk.Frame(self.toolbar_frame, height=1)
        self.toolbar_sep.pack(fill=tk.X, pady=(8, 0))

        # ── Rodapé ───────────────────────────────────────────────────────────
        _footer_text = f"{_ORG}  ·  {_APP_AUTHOR}  ·  {_LICENSE}  ·  {_MOTTO}"
        self.footer_lbl = tk.Label(
            self.root, text=_footer_text,
            font=("Segoe UI", 8), anchor="center"
        )
        self.footer_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 6))

        # ── Área de entrada ──────────────────────────────────────────────────
        self.input_outer = tk.Frame(self.root)
        self.input_outer.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(4, 10))

        self.input_hint = tk.Label(
            self.input_outer, text="Enter para enviar  ·  Shift+Enter para nova linha",
            font=("Segoe UI", 8), anchor="w"
        )
        self.input_hint.pack(side=tk.BOTTOM, anchor="w", pady=(2, 0))

        self.input_border = tk.Frame(self.input_outer, bd=1, relief=tk.SOLID)
        self.input_border.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        self.input_text = tk.Text(
            self.input_border, wrap=tk.WORD, height=2,
            font=("Segoe UI", 11), relief=tk.FLAT,
            padx=12, pady=10, bd=0
        )
        self.input_text.pack(expand=True, fill=tk.BOTH)
        self.input_text.bind("<Return>", self.on_enter)
        self.input_text.bind("<Shift-Return>", self.on_shift_enter)
        self.input_text.bind("<Key>", self._on_input_key, add="+")

        self.send_btn = tk.Button(
            self.input_outer, text="Enviar ➤",
            font=("Segoe UI", 11, "bold"), relief=tk.FLAT,
            cursor="hand2", padx=14, pady=0,
            command=self.send_or_cancel
        )
        self.send_btn.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))

        self.input_sep = tk.Frame(self.root, height=1)
        self.input_sep.pack(side=tk.BOTTOM, fill=tk.X)

        # ── Área de chat ──────────────────────────────────────────────────────
        self.chat_border = tk.Frame(self.root, bd=1, relief=tk.SOLID)
        self.chat_border.pack(expand=True, fill=tk.BOTH, padx=20, pady=(4, 0))

        self.chat_area = scrolledtext.ScrolledText(
            self.chat_border, wrap=tk.WORD,
            font=("Segoe UI", 11), relief=tk.FLAT,
            padx=14, pady=12, state=tk.DISABLED, bd=0
        )
        self.chat_area.pack(expand=True, fill=tk.BOTH)

    def _on_input_key(self, event=None):
        """Redimensiona a área de entrada automaticamente conforme o conteúdo."""
        self.root.after_idle(self._auto_resize_input)

    def _auto_resize_input(self):
        """Ajusta a altura da entrada com base no número de linhas (máx. 6)."""
        try:
            num_lines = int(self.input_text.index('end-1c').split('.')[0])
            new_height = max(2, min(6, num_lines))
            if self.input_text.cget('height') != new_height:
                self.input_text.configure(height=new_height)
        except Exception:
            pass

    def append_message(self, role: str, message: str) -> None:
        self.chat_area.config(state=tk.NORMAL)
        if role == "User":
            self.chat_area.insert(tk.END, "Você:\n", "user_tag")
            self.chat_area.insert(tk.END, f"{message}\n", "user_msg")
        elif role == "Sistema":
            self.chat_area.insert(tk.END, f"⚠ {message}\n", "sys_tag")
        else:
            self.chat_area.insert(tk.END, "IA:\n", "ai_tag")
            self.chat_area.insert(tk.END, f"{message}\n", "ai_msg")

        # Separador visual entre mensagens
        self.chat_area.insert(tk.END, "─" * 60 + "\n\n", "msg_sep")

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

    def _read_word_doc_text(self, word) -> str:
        """Lê o texto do documento ativo do Word com retry e múltiplos fallbacks.

        Lida com RPC_E_CALL_REJECTED (Word ocupado/modal) tentando novamente,
        e usa ActiveDocument -> Documents(1) -> ActiveWindow -> Selection.
        """
        import time

        last_err = None

        for attempt in range(1, 6):
            try:
                # Método 1: ActiveDocument
                try:
                    return word.ActiveDocument.Content.Text or ""
                except Exception as e_ad:
                    last_err = e_ad
                    # Método 2: Documents collection
                    try:
                        if word.Documents.Count > 0:
                            return word.Documents(1).Content.Text or ""
                    except Exception as e_docs:
                        last_err = e_docs
                        # Método 3: ActiveWindow.Document
                        try:
                            return word.ActiveWindow.Document.Content.Text or ""
                        except Exception as e_aw:
                            last_err = e_aw
                            # Método 4: Selection.Document
                            try:
                                return word.Selection.Document.Content.Text or ""
                            except Exception as e_sel:
                                last_err = e_sel
                                raise
            except Exception as e:
                last_err = e
                hresult = getattr(e, 'hresult', None)
                # Se o Word rejeitou a chamada (ocupado/modal), espera e tenta de novo
                if hresult is not None and (hresult & 0xFFFF) == 0x2711:  # RPC_E_CALL_REJECTED
                    LOGGER.info("Word busy (RPC_E_CALL_REJECTED), retry %d/5", attempt)
                    time.sleep(0.6 * attempt)
                    continue
                # Se é erro de "nenhum documento aberto", não adianta repetir
                err_str = str(e).lower()
                if "nenhum documento" in err_str or "no document" in err_str:
                    raise Exception(
                        "O Microsoft Word está aberto, mas nenhum documento foi encontrado. "
                        "Abra um documento no Word e tente novamente."
                    )
                raise
        raise last_err if last_err else Exception("Falha desconhecida ao ler documento do Word")

    def _find_word_with_documents(self):
        """Varre a Running Object Table e retorna a instância do Word que tem documentos abertos.

        Quando há múltiplos processos WINWORD.EXE, GetActiveObject pode retornar
        uma instância vazia (tela inicial), causando 'nenhum documento foi aberto'.
        """
        import pythoncom
        import win32com.client

        candidates = []
        try:
            ctx = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            enum = rot.EnumRunning()
            while True:
                monikers = enum.Next(1)
                if not monikers:
                    break
                mk = monikers[0]
                try:
                    name = mk.GetDisplayName(ctx, None)
                except Exception:
                    continue
                if "word.application" not in name.lower():
                    continue
                try:
                    unk = rot.GetObject(mk)
                    disp = unk.QueryInterface(pythoncom.IID_IDispatch)
                    word = win32com.client.Dispatch(disp)
                    docs_count = int(word.Documents.Count)
                    LOGGER.info("ROT Word instance '%s': Documents.Count=%d", name, docs_count)
                    candidates.append((docs_count, word))
                except Exception as e_cand:
                    LOGGER.warning("Failed to inspect ROT Word instance '%s': %s", name, e_cand)
        except Exception as e_rot:
            LOGGER.warning("ROT enumeration failed: %s", e_rot)

        # Prefere instância com documentos abertos
        candidates.sort(key=lambda c: c[0], reverse=True)
        for docs_count, word in candidates:
            if docs_count > 0:
                LOGGER.info("Selected Word instance with %d open document(s)", docs_count)
                return word
        # Se nenhuma tem documento, retorna a primeira (se houver) para erro coerente
        if candidates:
            return candidates[0][1]
        return None

    def _get_word_app(self):
        """Obtém uma referência COM para o Word, preferindo instância com documentos abertos."""
        import win32com.client

        # 1) Procura na ROT uma instância do Word com documentos abertos
        try:
            word = self._find_word_with_documents()
            if word is not None:
                try:
                    if int(word.Documents.Count) > 0:
                        return word
                except Exception:
                    pass
        except Exception as e_rot:
            LOGGER.warning("_find_word_with_documents failed: %s", e_rot)

        # 2) Fallbacks tradicionais
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            LOGGER.info("Got active Word object via GetActiveObject")
            return word
        except Exception as e1:
            LOGGER.warning("GetActiveObject failed: %s", e1)
        try:
            word = win32com.client.GetObject(Class="Word.Application")
            LOGGER.info("Got Word object via GetObject")
            return word
        except Exception as e2:
            LOGGER.warning("GetObject failed: %s", e2)
        word = win32com.client.Dispatch("Word.Application")
        LOGGER.info("Got Word object via Dispatch")
        return word

    def _reload_doc_text(self) -> bool:
        """Tenta atualizar self.doc_text com o conteúdo atual do Word na thread principal."""
        import pythoncom
        # Garante COM inicializado nesta thread (crucial em executável compilado)
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            if not self.word_app:
                self.word_app = self._get_word_app()
            try:
                raw_text = self._read_word_doc_text(self.word_app)
            except Exception:
                # Referência pode estar obsoleta (Word reiniciado); reconecta
                LOGGER.info("word_app stale, reconnecting to Word")
                self.word_app = self._get_word_app()
                raw_text = self._read_word_doc_text(self.word_app)
            if raw_text is None:
                raw_text = ""
                LOGGER.warning("Document text returned None from COM, treating as empty")
            self._doc_load_error = ""

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
        if self.is_generating or not self.client:
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
            LOGGER.info("Sending updated document context to AI chat")
            self.messages.append({"role": "user", "content": ctx_msg})
            reply = self._call_api()
            self.messages.append({"role": "assistant", "content": reply})
        except Exception as e:
            log_exception(LOGGER, "Failed to send context to AI", e)
            self._set_word_status("Z7: Erro ao carregar contexto no Chat IA.")
            reply = f"Falha ao enviar o contexto para a IA: {str(e)}"

        self.root.after(0, self._on_message_received, reply)

    def run_grammar_check(self) -> None:
        if self.is_generating or not self.client:
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
        if self.is_generating or not self.client:
            return

        self._reload_doc_text()

        self.input_text.delete("1.0", tk.END)
        # Exibe o texto exato do prompt de consistência enviado à IA
        from config_prompt import load_consistency_prompt
        today_prefix = get_today_date_text()
        display_msg = f"{today_prefix}\n{load_consistency_prompt()}"
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
            prompt = f"{today_prefix}\n{base_prompt}\n\n---INICIO DO DOCUMENTO---\n{self.doc_text}\n---FIM DO DOCUMENTO---\n"
            
            LOGGER.info(f"Sending {task_type} task to AI chat")
            
            self.messages.append({"role": "user", "content": prompt})
            reply = self._call_api()
            self.messages.append({"role": "assistant", "content": reply})
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
        self.messages = []
        self.update_status("Iniciando nova conversa...")
        threading.Thread(target=self._new_conversation_thread, daemon=True).start()

    def _new_conversation_thread(self) -> None:
        try:
            from config_prompt import load_chat_system_prompt
            today_prefix = get_today_date_text()
            chat_system_prompt = load_chat_system_prompt()
            self.system_instruction = f"{today_prefix} {chat_system_prompt}"
            self.initial_prompt_text = chat_system_prompt

            # Pré-popula histórico com contexto do documento
            self._reload_doc_text()
            self.messages = []
            if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
                ctx_user_msg = (
                    f"Abaixo está o texto do documento no Word (contexto desta conversa):\n\n"
                    f"{self.doc_text}"
                )
                self.messages.append({"role": "user", "content": ctx_user_msg})
                self.messages.append({"role": "assistant", "content": "Entendido! Contexto atualizado."})
                greeting = "✅ Contexto atualizado! Como posso ajudar?"
                LOGGER.info("New conversation: document context pre-seeded in history")
            else:
                greeting = "💬 Nova conversa reiniciada! Como posso ajudar?"

            self.root.after(0, lambda g=greeting: self._on_new_conversation_ready(g))
        except Exception as e:
            log_exception(LOGGER, "Failed to start new conversation", e)
            self.root.after(0, lambda: self.update_status("Erro ao iniciar nova conversa."))

    def _on_new_conversation_ready(self, greeting: str) -> None:
        self.update_status("Pronto para conversar")
        self._update_ai_status()
        # Exibe o texto exato do prompt enviado como instrução do sistema
        if self.system_instruction:
            self.append_message("User", self.system_instruction)
        self.append_message("AI", greeting)
        if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
            self.run_consistency_check()
        LOGGER.info("New conversation started")

    def _reinit_client(self, api_key: str, model: str) -> None:
        """Reinstancia o client OpenAI com nova chave e modelo."""
        import pythoncom
        pythoncom.CoInitialize()
        try:
            from openai import OpenAI
            self.client = OpenAI(
                base_url=_OPENROUTER_BASE_URL,
                api_key=api_key,
                timeout=60.0,
            )
            self._model = model
            LOGGER.info("OpenAI client reinitialized with model %s", model)
            self.root.after(0, lambda: self.update_status("Pronto para conversar"))
            self.root.after(50, self._update_ai_status)
        except Exception as e:
            log_exception(LOGGER, "Failed to reinitialize OpenAI client", e)
            self.root.after(0, lambda: self.update_status("Erro ao atualizar API."))
        finally:
            pythoncom.CoUninitialize()

    def _open_config_prompt(self) -> None:
        """Abre a janela de configuração do prompt em um processo separado."""
        import subprocess
        import sys
        from pathlib import Path

        if getattr(sys, 'frozen', False):
            # Compilado com PyInstaller: o executável config_prompt.exe
            # está em um diretório irmão: ai/config_prompt/config_prompt.exe
            current_exe = Path(sys.executable)
            config_exe = current_exe.parent.parent / "config_prompt" / "config_prompt.exe"
            if not config_exe.exists():
                # Tenta na raiz do diretório pai (caso a estrutura seja diferente)
                config_exe = current_exe.parent / "config_prompt.exe"
            if not config_exe.exists():
                from z7_theme import show_error
                show_error("Erro", f"Executável não encontrado:\n{config_exe}", parent=self.root)
                return
            try:
                subprocess.Popen([str(config_exe)])
                LOGGER.info("Config prompt executable opened: %s", config_exe)
            except Exception as e:
                log_exception(LOGGER, "Failed to open config_prompt", e)
                from z7_theme import show_error
                show_error("Erro", f"Não foi possível abrir as configurações:\n{e}", parent=self.root)
        else:
            # Executando via interpretador Python (desenvolvimento)
            config_script = Path(__file__).resolve().parent / "config_prompt.py"
            if not config_script.exists():
                from z7_theme import show_error
                show_error("Erro", f"Arquivo não encontrado:\n{config_script}", parent=self.root)
                return
            try:
                subprocess.Popen([sys.executable, str(config_script)])
                LOGGER.info("Config prompt window opened")
            except Exception as e:
                log_exception(LOGGER, "Failed to open config_prompt", e)
                from z7_theme import show_error
                show_error("Erro", f"Não foi possível abrir as configurações:\n{e}", parent=self.root)

    def on_enter(self, event: tk.Event) -> str:
        self.send_message()
        return "break"
        
    def on_shift_enter(self, event: tk.Event) -> str:
        self.input_text.insert(tk.INSERT, "\n")
        return "break"

    def _load_doc_text_main_thread(self) -> None:
        """Lê o texto do documento ativo na thread principal (COM funciona corretamente aqui)."""
        import pythoncom
        
        # Inicializa COM na thread principal (necessário para executável compilado)
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass  # Já pode estar inicializado
        
        try:
            # Tenta obter instância ativa do Word
            try:
                word = self._get_word_app()
            except Exception as e_conn:
                raise Exception(
                    f"Não foi possível conectar ao Microsoft Word. "
                    f"Verifique se o Word está aberto com um documento.\n"
                    f"Detalhes: {e_conn}"
                )

            self.word_app = word

            # Lê o texto do documento com retry + fallbacks (ActiveDocument, Documents, etc.)
            # IMPORTANTE: lê o texto ANTES de rodar macros, pois o backup pode trocar o documento ativo
            raw_text = self._read_word_doc_text(word)

            try:
                word.Application.Run("CreateDocumentBackup", word.ActiveDocument)
                LOGGER.info("Document backup created successfully.")
            except Exception as backup_e:
                LOGGER.warning("Could not run CreateDocumentBackup macro: %s", str(backup_e))
            LOGGER.info("Got document text (%d chars)", len(raw_text))

            if raw_text is None:

                raw_text = ""
                LOGGER.warning("Document text returned None from COM, treating as empty")
            
            if len(raw_text) > _MAX_CONTEXT_CHARS:
                cut = raw_text.rfind(' ', 0, _MAX_CONTEXT_CHARS)
                if cut == -1:
                    cut = _MAX_CONTEXT_CHARS
                self.doc_text = raw_text[:cut]
                self._doc_truncated = True
                LOGGER.warning("Document context truncated at %d chars", cut)
            else:
                self.doc_text = raw_text
            self._doc_load_error = ""
            LOGGER.info("Loaded Word active document context for chat (%d chars)", len(self.doc_text))
        except Exception as e:
            error_detail = str(e)
            log_exception(LOGGER, "Failed to load Word document context", e)
            self._set_word_status("Z7: Erro ao carregar contexto do documento no Chat.")
            self.doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."
            self._doc_load_error = error_detail
            LOGGER.error("Document load error details: %s", error_detail)

    def _call_api(self) -> str:
        """Envia o histórico completo de mensagens para a API e retorna a resposta."""
        api_messages = [{"role": "system", "content": self.system_instruction}]
        api_messages.extend(self.messages)

        response = self.client.chat.completions.create(
            model=self._model,
            messages=api_messages,
            timeout=120,
        )
        return response.choices[0].message.content

    def init_ai(self) -> None:
        # Inicialização em background para não travar a UI
        threading.Thread(target=self._init_ai_thread, daemon=True).start()
        
    def _init_ai_thread(self) -> None:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            from openai import OpenAI

            api_key = get_api_key(self.root)
            if not api_key:
                LOGGER.error("Chat initialization aborted: missing API key")
                self._set_word_status("Z7: Erro no Chat - chave da API indisponivel.")
                self.root.after(0, lambda: self.update_status("Erro: Chave API ausente."))
                return

            self.client = OpenAI(
                base_url=_OPENROUTER_BASE_URL,
                api_key=api_key,
                timeout=60.0,
            )

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

            from config_prompt import load_chat_system_prompt
            today_prefix = get_today_date_text()
            chat_system_prompt = load_chat_system_prompt()
            self.system_instruction = f"{today_prefix} {chat_system_prompt}"
            self.initial_prompt_text = chat_system_prompt

            # --- Pré-popula o histórico com o contexto do documento ---
            _truncation_notice = _doc_truncated
            doc_context = self.doc_text
            self.messages = []
            if doc_context and doc_context.strip() and "nenhum documento" not in doc_context.lower():
                ctx_user_msg = (
                    f"Abaixo está o texto atual do meu documento no Word para ser usado como base e contexto dessa conversa:\n\n"
                    f"{doc_context}"
                )
                self.messages.append({"role": "user", "content": ctx_user_msg})
                self.messages.append({"role": "assistant", "content": "Entendido! Recebi o contexto do documento e estou pronto para ajudar."})
                self.initial_greeting = "✅ Contexto do documento carregado! Como posso ajudar?"
                LOGGER.info("Document context pre-seeded in chat history (no API call required)")
            else:
                error_detail = getattr(self, '_doc_load_error', '')
                if error_detail:
                    self.initial_greeting = (
                        "⚠ Não consegui acessar o documento atual.\n\n"
                        f"Erro: {error_detail}\n\n"
                        "💡 Dica: Certifique-se de que o Word está aberto com um documento ativo. "
                        "Você pode digitar 'recarregar contexto' para tentar novamente."
                    )
                elif doc_context is not None and not doc_context.strip():
                    self.initial_greeting = (
                        "📄 O documento no Word está em branco.\n\n"
                        "Escreva ou cole o conteúdo da propositura no Word e, em seguida, "
                        "use o comando \"recarregar contexto\" aqui no chat para que eu possa analisá-lo."
                    )
                    LOGGER.info("Document is blank/empty")
                else:
                    self.initial_greeting = (
                        "⚠ Não foi possível obter o conteúdo do documento.\n\n"
                        "💡 Dica: Certifique-se de que o Word está aberto com um documento ativo. "
                        "Você pode digitar 'recarregar contexto' para tentar novamente."
                    )
                _truncation_notice = False
                LOGGER.warning("No document context available for AI initialization. Error: %s", error_detail)

            LOGGER.info("Chat session ready with model: %s", self._model)

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
            if "401" in error_msg or "403" in error_msg or "invalid api key" in error_msg or "api_key" in error_msg:
                self.root.after(0, lambda: self.update_status("Erro: Chave Inválida."))
                self.root.after(0, lambda: self.append_message("Sistema", "Sua chave da API parece inválida ou expirou. Abra as Configurações da IA para atualizar ou remover a chave."))
            else:
                self.root.after(0, lambda: self.update_status("Erro na inicialização."))
                self.root.after(0, lambda err=str(e): self.append_message("Sistema", f"Erro crítico: {err}"))
        finally:
            pythoncom.CoUninitialize()

    def _on_ai_ready(self) -> None:
        self.update_status("Pronto para conversar")
        self._update_ai_status()
        # Exibe o texto exato do prompt enviado como instrução do sistema
        if self.system_instruction:
            self.append_message("User", self.system_instruction)
        self.append_message("AI", getattr(self, 'initial_greeting', "Olá! Como posso ajudar?"))
        if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
            self.run_consistency_check()

    def send_message(self) -> None:
        if self.is_generating or not self.client:
            LOGGER.info("Send skipped: generating=%s client_ready=%s", self.is_generating, self.client is not None)
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
                        f"[Mensagem do usuário]\n{user_msg}"
                    )
                self._context_pending = False
            LOGGER.info("Sending message to AI chat")
            self.messages.append({"role": "user", "content": actual_msg})
            reply = self._call_api()
            self.messages.append({"role": "assistant", "content": reply})
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