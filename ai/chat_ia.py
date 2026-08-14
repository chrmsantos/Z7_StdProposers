import threading
import tkinter as tk
from tkinter import scrolledtext

import z7_theme
from z7_logging import configure_component_logger, log_exception
from z7_api_key import get_api_key

LOGGER = configure_component_logger("chat_ia")


def _configure_ssl_certifi() -> None:
    """Ensure the SSL module can locate the CA bundle shipped with *certifi*.

    PyInstaller bundles *certifi``s ``cacert.pem`` inside ``_internal/``, but
    the system OpenSSL used by :func:`ssl.create_default_context` may not find
    it automatically (especially on Windows with Python ≥ 3.14).  Setting the
    ``SSL_CERT_FILE`` environment variable **before** any HTTPS connection is
    created resolves ``FileNotFoundError: [Errno 2] No such file or directory``
    raised by ``httpx`` → ``ssl.create_default_context``.
    """
    import os
    try:
        import certifi
        os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    except ImportError:
        LOGGER.warning("certifi not available; SSL may fail in frozen environment")


_DEFAULT_MODEL = 'meta-llama/llama-3.3-70b-instruct:free'
_FALLBACK_MODEL = 'google/gemma-2-9b-it:free'
_OPENROUTER_BASE_URL = 'https://openrouter.ai/api/v1'
_MAX_CONTEXT_CHARS = 150_000

_APP_VERSION = "8.7.1"
_APP_AUTHOR  = "CMS"
_ORG         = "Câmara Municipal de Santa Bárbara d'Oeste"
_LICENSE     = "GPL-3.0"
_MOTTO       = "Dharma, virtude e gratidão."

# Fixed first user message displayed when chat starts
_FIRST_USER_MSG = (
    "Olá, Léia! Por gentileza, leia a propositura, identifique erros e proponha sugestões de correção. "
    "Atenda às tarefas descritas e detalhadas nas configurações de prompt."
)

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
        self.root.title(f"Chat com a LÉIA — Z7 StdProposers v{_APP_VERSION}")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        chat_width = int(screen_width * 2 / 3 * 1.15 * 1.10 * 0.70 * 1.10)
        chat_height = int(screen_height * 0.92 * 0.90 * 1.05)
        chat_left = 0
        chat_top = int(screen_height * 0.02)
        
        self.chat_width_px = chat_width
        self.root.geometry(f"{chat_width}x{chat_height}+{chat_left}+{chat_top}")
        self.root.minsize(300, 500)
        
        self.resize_word_window(screen_width, screen_height)
        
        self.mode = z7_theme.load_theme()
        self.client = None
        self.messages = []  # Histórico de mensagens no formato OpenAI
        self.system_instruction = ""
        self.is_generating = False
        self._cancel_requested = False
        self.last_ai_reply = ""
        self._model = _DEFAULT_MODEL
        self._fallback_model = _FALLBACK_MODEL
        self.doc_text = ""
        self._doc_truncated = False
        self._context_pending = False
        self._doc_load_error = ""
        self.current_status_text = "Carregando contexto..."
        self._last_chat_status = ""  # Deduplicação de status no chat
        self._streaming_active = False
        self._streamed_text = ""
        self._streaming_start = None
        self._streamed_reply = False

        self.build_ui()
        self.apply_theme()
        
        self._load_doc_text_main_thread()
        self.init_ai()

        # Bring window to front on startup without staying always-on-top
        self.root.lift()
        self.root.focus_force()

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

        # Separadores
        for sep_attr in ('header_sep',):
            sep = getattr(self, sep_attr, None)
            if sep:
                sep.configure(bg=border)

        # Chat area
        self.chat_border.configure(bg=border)
        select_bg = "#8b5cf6" if self.mode == "dark" else "#7c3aed"
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
        self.chat_area.tag_config("mediator_msg", font=("Segoe UI", 9, "italic"),
            foreground=fg_muted, lmargin1=14, lmargin2=14, rmargin=14,
            spacing1=1, spacing3=1)
        self.chat_area.tag_config("msg_sep", foreground=border,
            font=("Segoe UI", 6), spacing1=2, spacing3=6)

        # Ensure selection highlight is visible above all other tags
        self.chat_area.tag_config("sel", background=select_bg, foreground=select_fg)
        self.chat_area.tag_raise("sel")

        # Version badge
        if hasattr(self, 'version_badge'):
            self.version_badge.configure(bg=colors["btn_primary_bg"], fg="white")

        self.update_status(self.current_status_text)

    def update_status(self, text: str) -> None:
        """Atualiza o status da IA publicando uma mensagem de mediador na conversa.

        Publica uma mensagem de mediador no chat para que o usuário
        acompanhe o estado da IA em tempo real (com deduplicação para
        evitar repetição consecutiva).
        """
        self.current_status_text = text

        # "Pronto para conversar" é um estado interno — não exibir no chat
        if text.lower().strip() == "pronto para conversar":
            self._last_chat_status = text.lower().strip()
            return

        # Publica status como mensagem de mediador no chat (evita duplicatas consecutivas)
        text_lower_for_dedup = text.lower().strip()
        if text_lower_for_dedup != self._last_chat_status:
            self._last_chat_status = text_lower_for_dedup
            # Usa after(0) para garantir execução na thread principal
            try:
                self.root.after(0, self.append_status_message, text)
            except Exception:
                pass

    def build_ui(self) -> None:
        # ── Cabeçalho ────────────────────────────────────────────────────────
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=(16, 0), padx=20)

        # Linha 1: Título + botões de ação
        self.header_top = tk.Frame(self.top_frame)
        self.header_top.pack(fill=tk.X)

        self.title_lbl = tk.Label(
            self.header_top, text="✨ LÉIA — Assistente Legislativa de IA",
            font=("Segoe UI", 16, "bold")
        )
        self.title_lbl.pack(side=tk.LEFT, anchor="w")

        self.version_badge = tk.Label(
            self.header_top, text=f"v{_APP_VERSION}",
            font=("Segoe UI", 10, "bold"), padx=6, pady=1
        )
        self.version_badge.pack(side=tk.LEFT, anchor="w", padx=(8, 0))

        self.settings_btn = tk.Button(
            self.header_top, text="⚙ Configurações",
            font=("Segoe UI", 10, "bold"), relief=tk.FLAT,
            cursor="hand2", padx=12, pady=4,
            command=self._open_config_prompt
        )
        self.settings_btn.pack(side=tk.RIGHT, anchor="e", padx=(8, 0))

        # Separador após cabeçalho
        self.header_sep = tk.Frame(self.top_frame, height=1)
        self.header_sep.pack(fill=tk.X, pady=(12, 0))

        # ── Rodapé ───────────────────────────────────────────────────────────
        _footer_text = f"{_ORG}  ·  {_APP_AUTHOR}  ·  {_LICENSE}  ·  {_MOTTO}"
        self.footer_lbl = tk.Label(
            self.root, text=_footer_text,
            font=("Segoe UI", 8), anchor="e"
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
            self.chat_area.insert(tk.END, "LÉIA:\n", "ai_tag")
            self.chat_area.insert(tk.END, f"{message}\n", "ai_msg")

        # Separador visual entre mensagens
        self.chat_area.insert(tk.END, "─" * 60 + "\n\n", "msg_sep")

        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def append_status_message(self, text: str) -> None:
        """Exibe uma mensagem de status da IA como mediador na conversa."""
        # Remove prefixo "📋 Mediador:" caso presente no texto
        for prefix in ("📋 Mediador:\n", "📋 Mediador:", "Mediador:"):
            if text.startswith(prefix):
                text = text[len(prefix):].lstrip()
                break
        self.chat_area.config(state=tk.NORMAL)
        self.chat_area.insert(tk.END, f"{text}\n", "mediator_msg")
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

    # ------------------------------------------------------------------
    # Leitura robusta do documento do Word
    # ------------------------------------------------------------------
    @staticmethod
    def _word_doc_counts(word) -> tuple:
        """Retorna (docs_normais, docs_protected_view) de forma defensiva.

        Documentos abertos a partir de pastas temporárias (%LocalAppData%\\Temp),
        anexos de e-mail ou downloads são abertos pelo Word em Protected View e
        NÃO constam em Documents.Count — apenas em ProtectedViewWindows.Count.

        IMPORTANTE (diagnóstico do bug 'nenhum documento'): quando a leitura de
        uma contagem falha (ex.: RPC_E_CALL_REJECTED, Word ocupado/modal), a
        contagem é reportada como 0 e o erro é registrado em WARNING — nunca
        silenciosamente ignorado. Isso evita confundir 'instância vazia' com
        'não foi possível ler a contagem'.
        """
        normal = 0
        protected = 0
        try:
            normal = int(word.Documents.Count)
        except Exception as e:
            LOGGER.warning("Could not read Documents.Count (hresult=%s): %s",
                           getattr(e, "hresult", None), e)
        try:
            protected = int(word.ProtectedViewWindows.Count)
        except Exception as e:
            LOGGER.warning("Could not read ProtectedViewWindows.Count (hresult=%s): %s",
                           getattr(e, "hresult", None), e)
        return normal, protected

    def _word_doc_snapshot(self, word) -> dict:
        """Coleta um diagnóstico completo do estado de documentos da instância Word.

        Diferente de :meth:`_word_doc_counts` (que retorna apenas números), aqui
        cada contagem carrega também o status de leitura (``'ok'`` ou ``'error'``),
        permitindo distinguir três situações críticas para o bug recorrente:

        * ``(value=0, status='ok')``        → instância de fato vazia;
        * ``(value=1, status='ok')``        → instância com documento acessível;
        * ``(value=None, status='error')``  → leitura falhou (COM/RPC) — a
          instância PODE ter documentos não visíveis na contagem.

        Retorna:
            dict com as chaves ``documents``, ``protected_view``, ``windows`` e
            ``active_document``; cada uma com ``status`` e, quando aplicável,
            ``value``/``present``/``error``/``hresult``.
        """
        snapshot = {
            "documents": {"status": "not_read", "value": None},
            "protected_view": {"status": "not_read", "value": None},
            "windows": {"status": "not_read", "value": None},
            "active_document": {"status": "not_read", "present": None},
        }
        for key, getter in (
            ("documents", lambda: int(word.Documents.Count)),
            ("protected_view", lambda: int(word.ProtectedViewWindows.Count)),
            ("windows", lambda: int(word.Windows.Count)),
        ):
            try:
                snapshot[key]["value"] = getter()
                snapshot[key]["status"] = "ok"
            except Exception as e:
                snapshot[key]["status"] = "error"
                snapshot[key]["error"] = str(e)
                snapshot[key]["hresult"] = getattr(e, "hresult", None)
        try:
            snapshot["active_document"]["present"] = word.ActiveDocument is not None
            snapshot["active_document"]["status"] = "ok"
        except Exception as e:
            snapshot["active_document"]["status"] = "error"
            snapshot["active_document"]["error"] = str(e)
            snapshot["active_document"]["hresult"] = getattr(e, "hresult", None)
        return snapshot

    @staticmethod
    def _fmt_count(snap_entry: dict) -> str:
        """Formata uma entrada de contagem do snapshot para o log."""
        if snap_entry["status"] == "ok":
            return str(snap_entry["value"])
        return f"ERR(hresult={snap_entry.get('hresult')})"

    def _log_word_snapshot(self, word, label: str) -> dict:
        """Registra o diagnóstico da instância Word em nível INFO e o retorna.

        Centraliza a emissão do snapshot no log, garantindo consistência de
        formato e nível entre todos os pontos que inspecionam o Word.
        """
        snap = self._word_doc_snapshot(word)
        LOGGER.info(
            "Word snapshot [%s]: Documents=%s, ProtectedView=%s, Windows=%s, ActiveDocument=%s",
            label,
            self._fmt_count(snap["documents"]),
            self._fmt_count(snap["protected_view"]),
            self._fmt_count(snap["windows"]),
            (str(snap["active_document"]["present"])
             if snap["active_document"]["status"] == "ok" else "ERR"),
        )
        return snap

    @staticmethod
    def _is_no_doc_error(err: Exception) -> bool:
        """Detecta o erro COM 'nenhum documento foi aberto' (PT/EN) ou coleção vazia."""
        msg = str(err).lower()
        return ("nenhum documento" in msg or "no document" in msg
                or "count == 0" in msg)

    @staticmethod
    def _is_rpc_busy_error(err: Exception) -> bool:
        """Detecta RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER (Word ocupado/modal)."""
        hresult = getattr(err, 'hresult', None)
        if hresult is None:
            return False
        try:
            hresult = int(hresult)
        except (TypeError, ValueError):
            return False
        # FACILITY_RPC (0x8001xxxx): CALL_REJECTED, SERVERCALL_RETRYLATER, SERVERCALL_REJECTED
        if (hresult & 0xFFFF0000) == 0x80010000:
            return True
        return (hresult & 0xFFFF) == 0x2711  # compatibilidade com verificação legada

    def _iter_doc_read_methods(self, word) -> list:
        """Retorna lista ordenada de (nome, callable) para extrair o texto do documento.

        Além dos caminhos tradicionais (ActiveDocument, Documents, ActiveWindow,
        Selection), cobre Protected View — cenário típico de documentos abertos
        a partir de %LocalAppData%\\Temp, anexos de e-mail ou downloads, nos quais
        ActiveDocument/Documents lançam 'nenhum documento foi aberto'
        (hresult interno -2146824040) mesmo havendo documento visível na tela.
        """
        def _active_document():
            return word.ActiveDocument.Content.Text

        def _documents_collection():
            if int(word.Documents.Count) > 0:
                return word.Documents(1).Content.Text
            raise RuntimeError("Documents.Count == 0")

        def _active_window():
            return word.ActiveWindow.Document.Content.Text

        def _selection():
            return word.Selection.Document.Content.Text

        def _protected_view():
            if int(word.ProtectedViewWindows.Count) > 0:
                return word.ProtectedViewWindows(1).Document.Content.Text
            raise RuntimeError("ProtectedViewWindows.Count == 0")

        def _windows_collection():
            if int(word.Windows.Count) > 0:
                return word.Windows(1).Document.Content.Text
            raise RuntimeError("Windows.Count == 0")

        return [
            ("ActiveDocument", _active_document),
            ("Documents(1)", _documents_collection),
            ("ActiveWindow.Document", _active_window),
            ("Selection.Document", _selection),
            ("ProtectedViewWindows(1).Document", _protected_view),
            ("Windows(1).Document", _windows_collection),
        ]

    def _read_word_doc_text(self, word) -> str:
        """Lê o texto do documento do Word com múltiplos fallbacks e retries.

        Estratégia em camadas:
        - Em cada tentativa, percorre TODOS os métodos de leitura, incluindo
          Protected View (essencial para documentos abertos do Temp/Internet,
          que não aparecem em ActiveDocument/Documents).
        - Entre tentativas, aguarda progressivamente — cobre RPC_E_CALL_REJECTED
          (Word ocupado/modal) e documento ainda em carregamento na abertura do chat.
        - Erros 'nenhum documento' são determinísticos: se TODOS os métodos
          falham assim, o retry é interrompido cedo (sem esperar 5 rodadas),
          reduzindo o tempo de abertura do chat quando o Word está vazio.
        - Só desiste após esgotar as tentativas, produzindo mensagem amigável
          quando a causa raiz é 'nenhum documento aberto'.
        """
        import time

        max_attempts = 5
        last_err = None
        saw_no_doc = False
        method_failures = {}  # name -> (categoria, exceção)
        start = time.perf_counter()

        for attempt in range(1, max_attempts + 1):
            round_busy = False
            for name, read_fn in self._iter_doc_read_methods(word):
                try:
                    text = read_fn()
                    if text is None:
                        LOGGER.warning("Read method %s returned None; treating as empty", name)
                        text = ""
                    LOGGER.info(
                        "Document text read OK via %s (%d chars, attempt %d/%d, %.2fs)",
                        name, len(text), attempt, max_attempts, time.perf_counter() - start,
                    )
                    return text
                except Exception as e:
                    last_err = e
                    if self._is_no_doc_error(e):
                        saw_no_doc = True
                        cat = "no_doc"
                    elif self._is_rpc_busy_error(e):
                        round_busy = True
                        cat = "rpc_busy"
                    else:
                        cat = "other"
                    method_failures[name] = (cat, e)
                    LOGGER.info(
                        "Read method %s failed [%s] (attempt %d/%d): %s",
                        name, cat, attempt, max_attempts, e,
                    )

            # 'nenhum documento' em TODOS os métodos = estado determinístico:
            # parar cedo evita ~6s de backoff inútil por ciclo de carga.
            if saw_no_doc and not round_busy and attempt >= 2:
                LOGGER.info(
                    "All read methods consistently report 'no document'; stopping retries early (attempt %d/%d)",
                    attempt, max_attempts,
                )
                break

            if attempt < max_attempts:
                delay = 0.6 * attempt
                if last_err is not None and self._is_rpc_busy_error(last_err):
                    LOGGER.info("Word busy (RPC), retry %d/%d in %.1fs",
                                attempt, max_attempts, delay)
                else:
                    LOGGER.info("No readable document yet (attempt %d/%d); retrying in %.1fs",
                                attempt, max_attempts, delay)
                time.sleep(delay)

        total = time.perf_counter() - start
        cats = sorted({c for c, _ in method_failures.values()}) or ["none"]
        failed_names = ", ".join(f"{n}[{c}]" for n, (c, _) in method_failures.items())
        LOGGER.warning(
            "Document read FAILED after %.2fs (%d attempts). Methods: %s | categories: %s",
            total, attempt, failed_names or "-", ", ".join(cats),
        )

        if saw_no_doc:
            raise Exception(
                "O Microsoft Word está aberto, mas nenhum documento foi encontrado. "
                "Abra um documento no Word e tente novamente."
            )
        raise last_err if last_err else Exception("Falha desconhecida ao ler documento do Word")

    def _find_word_with_documents(self):
        """Varre a Running Object Table e retorna a instância do Word que tem documentos.

        Quando há múltiplos processos WINWORD.EXE, GetActiveObject pode retornar
        uma instância vazia (tela inicial), causando 'nenhum documento foi aberto'.
        Documentos em Protected View (Temp/Internet/anexos) também são contados,
        pois NÃO aparecem em Documents.Count.
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
                    normal, protected = self._word_doc_counts(word)
                    score = normal + protected
                    LOGGER.info("ROT Word instance '%s': Documents=%d, ProtectedView=%d",
                                name, normal, protected)
                    candidates.append((score, word))
                except Exception as e_cand:
                    LOGGER.warning("Failed to inspect ROT Word instance '%s': %s", name, e_cand)
        except Exception as e_rot:
            LOGGER.warning("ROT enumeration failed: %s", e_rot)

        # Prefere instância com documentos acessíveis (normais ou Protected View)
        candidates.sort(key=lambda c: c[0], reverse=True)
        for score, word in candidates:
            if score > 0:
                LOGGER.info("Selected Word instance with %d accessible document(s)", score)
                return word
        # Se nenhuma tem documento, retorna a primeira (se houver) para erro coerente
        if candidates:
            LOGGER.warning("No Word instance reports accessible documents; using first ROT instance")
            return candidates[0][1]
        return None

    def _get_word_app(self):
        """Obtém uma referência COM para o Word, preferindo instância com documentos acessíveis.

        Cada fallback é verificado: se a instância retornada não tem documentos
        acessíveis (normais ou Protected View), tenta o próximo método.
        Isso evita retornar uma instância vazia quando o documento está em
        outro processo WINWORD.EXE ou em Protected View.
        """
        import win32com.client

        # 1) Procura na ROT uma instância do Word com documentos (normais ou Protected View)
        try:
            word = self._find_word_with_documents()
            if word is not None:
                snap = self._log_word_snapshot(word, "ROT")
                if (snap["documents"]["status"] == "ok" and snap["documents"]["value"] > 0) or \
                   (snap["protected_view"]["status"] == "ok" and snap["protected_view"]["value"] > 0):
                    return word
                LOGGER.warning("ROT Word instance has no accessible documents yet; trying classic fallbacks")
        except Exception as e_rot:
            LOGGER.warning("_find_word_with_documents failed: %s", e_rot)

        # 2) Fallbacks tradicionais — cada um verificado antes de retornar
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            snap = self._log_word_snapshot(word, "GetActiveObject")
            if (snap["documents"]["status"] == "ok" and snap["documents"]["value"] > 0) or \
               (snap["protected_view"]["status"] == "ok" and snap["protected_view"]["value"] > 0):
                return word
            LOGGER.warning("GetActiveObject returned Word instance with no accessible documents; trying next fallback")
        except Exception as e1:
            LOGGER.warning("GetActiveObject failed: %s", e1)
        try:
            word = win32com.client.GetObject(Class="Word.Application")
            snap = self._log_word_snapshot(word, "GetObject")
            if (snap["documents"]["status"] == "ok" and snap["documents"]["value"] > 0) or \
               (snap["protected_view"]["status"] == "ok" and snap["protected_view"]["value"] > 0):
                return word
            LOGGER.warning("GetObject returned Word instance with no accessible documents; trying next fallback")
        except Exception as e2:
            LOGGER.warning("GetObject failed: %s", e2)
        # Dispatch SEMPRE cria uma nova instância vazia — não resolve 'nenhum documento'.
        # Se chegamos aqui, o Word pode estar aberto mas nenhum método encontrou
        # uma instância com documentos. Usa GetActiveObject como último recurso
        # (a instância pode ter documentos que _word_doc_counts não conseguiu contar).
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            self._log_word_snapshot(word, "last-resort GetActiveObject")
            LOGGER.info("Last-resort GetActiveObject for Word (doc counts previously reported 0)")
            return word
        except Exception:
            raise Exception(
                "Não foi possível encontrar uma instância do Word com documentos. "
                "Certifique-se de que o Word está aberto com um documento."
            )

    def _reload_doc_text(self) -> bool:
        """Tenta atualizar self.doc_text com o conteúdo atual do Word na thread principal.

        Faz até 2 ciclos completos de conexão+leitura; em cada ciclo, uma falha
        de leitura provoca reconexão (referência COM pode estar obsoleta).
        """
        import pythoncom
        # Garante COM inicializado nesta thread (crucial em executável compilado)
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        last_err = None
        for attempt in range(1, 3):
            try:
                if not self.word_app:
                    self.word_app = self._get_word_app()
                try:
                    raw_text = self._read_word_doc_text(self.word_app)
                except Exception:
                    # Referência pode estar obsoleta (Word reiniciado); reconecta
                    LOGGER.info("word_app stale or unreadable, reconnecting to Word")
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
                last_err = e
                LOGGER.warning("Failed to reload document text (attempt %d/2): %s", attempt, e)
                self.word_app = None  # força reconexão completa na próxima tentativa
        LOGGER.warning("Failed to reload document text: %s", last_err)
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

    def _start_initial_analysis(self) -> None:
        """Envia o prompt configurado + documento para a IA realizar a análise inicial."""
        self.is_generating = True
        self._cancel_requested = False
        self.update_status("IA analisando documento...")
        self.send_btn.config(text="Cancelar")
        threading.Thread(target=self._initial_analysis_thread, daemon=True).start()

    def _initial_analysis_thread(self) -> None:
        try:
            if not self.doc_text or not self.doc_text.strip() or "nenhum documento" in self.doc_text.lower():
                reply = "Não há contexto de documento carregado para realizar a análise."
                self.root.after(0, self._on_message_received, reply)
                return

            today_prefix = get_today_date_text()
            prompt = (
                f"{today_prefix}\n"
                f"{_FIRST_USER_MSG}\n\n"
                f"---INICIO DO DOCUMENTO---\n{self.doc_text}\n---FIM DO DOCUMENTO---\n"
            )

            LOGGER.info("Sending initial analysis request to AI")
            self.messages.append({"role": "user", "content": prompt})
            reply = self._call_api()
            self.messages.append({"role": "assistant", "content": reply})
        except Exception as e:
            log_exception(LOGGER, "Initial analysis failed", e)
            self._set_word_status("Z7: Erro na análise inicial.")
            reply = f"Erro ao processar a análise: {str(e)}"

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

            # Pré-popula histórico com contexto do documento
            self._reload_doc_text()
            self.messages = []
            has_doc = (self.doc_text and self.doc_text.strip()
                       and "nenhum documento" not in self.doc_text.lower())
            if has_doc:
                ctx_user_msg = (
                    f"Abaixo está o texto do documento no Word (contexto desta conversa):\n\n"
                    f"{self.doc_text}"
                )
                self.messages.append({"role": "user", "content": ctx_user_msg})
                self.messages.append({"role": "assistant", "content": "Entendido! Contexto atualizado."})
                LOGGER.info("New conversation: document context pre-seeded in history")

            self.root.after(0, self._on_new_conversation_ready)
        except Exception as e:
            log_exception(LOGGER, "Failed to start new conversation", e)
            self.root.after(0, lambda: self.update_status("Erro ao iniciar nova conversa."))

    def _on_new_conversation_ready(self) -> None:
        self.update_status("Pronto para conversar")
        # Exibe a mensagem fixa pré-definida
        self.append_message("User", _FIRST_USER_MSG)
        # Inicia a análise inicial se houver documento
        if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
            self._start_initial_analysis()
        else:
            self.append_message("AI", "⚠ Nenhum documento ativo encontrado. Abra um documento no Word e reinicie a conversa.")
        LOGGER.info("New conversation started")

    def _reinit_client(self, api_key: str, model: str) -> None:
        """Reinstancia o client OpenAI com nova chave e modelo."""
        import pythoncom
        pythoncom.CoInitialize()
        try:
            _configure_ssl_certifi()
            from openai import OpenAI
            self.client = OpenAI(
                base_url=_OPENROUTER_BASE_URL,
                api_key=api_key,
                timeout=60.0,
            )
            self._model = model
            # Recarrega modelo fallback salvo no diálogo
            try:
                from config_prompt import load_fallback_model
                self._fallback_model = load_fallback_model()
                LOGGER.info("Fallback model reloaded: %s", self._fallback_model)
            except Exception as fb_exc:
                LOGGER.warning("Could not reload fallback model: %s", fb_exc)
            LOGGER.info("OpenAI client reinitialized with model %s", model)
            self.root.after(0, lambda: self.update_status("Pronto para conversar"))
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
        """Lê o texto do documento ativo na thread principal (COM funciona corretamente aqui).

        Programação defensiva em camadas para garantir que o inteiro teor do
        documento ativo — salvo ou não, normal ou em Protected View — seja
        sempre obtido como contexto inicial:
        1. Conexão: _get_word_app escolhe a instância do Word com documentos
           acessíveis (incluindo Protected View) via Running Object Table.
        2. Leitura: _read_word_doc_text tenta 6 métodos em até 5 rodadas de retry.
        3. Reconexão: se conectar+ler falhar, descarta a referência e repete
           todo o ciclo (até 3 vezes) — cobre o caso de o documento ainda
           estar sendo carregado quando o chat abre.

        Toda a jornada é instrumentada com tempo (perf_counter) e um resumo
        estruturado ao final, permitindo diagnosticar a duração de cada fase e
        a causa exata de uma falha ('nenhum documento' vs. erro COM/RPC).
        """
        import pythoncom
        import time

        # Inicializa COM na thread principal (necessário para executável compilado)
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass  # Já pode estar inicializado

        raw_text = None
        last_error = None
        max_cycles = 3
        start_total = time.perf_counter()
        cycle_times = []
        LOGGER.info("[doc-load] START (max_cycles=%d)", max_cycles)

        for cycle in range(1, max_cycles + 1):
            cycle_start = time.perf_counter()
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
                self._log_word_snapshot(word, f"cycle {cycle} connected")

                # Lê o texto com retry + fallbacks (ActiveDocument, Documents, Protected View, etc.)
                raw_text = self._read_word_doc_text(word)

                cycle_times.append(time.perf_counter() - cycle_start)
                LOGGER.info("Got document text (%d chars, cycle %d/%d)",
                            len(raw_text), cycle, max_cycles)
                break  # sucesso — sai do loop de ciclos
            except Exception as e:
                last_error = e
                cycle_elapsed = time.perf_counter() - cycle_start
                cycle_times.append(cycle_elapsed)
                cat = ("no_doc" if self._is_no_doc_error(e)
                       else "rpc_busy" if self._is_rpc_busy_error(e)
                       else "connection" if "conectar" in str(e) else "other")
                LOGGER.warning("Document load cycle %d/%d failed [%s] after %.2fs: %s",
                               cycle, max_cycles, cat, cycle_elapsed, e)
                self.word_app = None  # força reconexão completa no próximo ciclo
                raw_text = None
                if cycle < max_cycles:
                    time.sleep(1.0 * cycle)

        total_elapsed = time.perf_counter() - start_total
        try:
            if raw_text is None:
                raise last_error if last_error else Exception("Falha desconhecida ao ler documento do Word")

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
            LOGGER.info(
                "[doc-load] SUCCESS: %d chars loaded in %.2fs (cycles=%s, truncated=%s)",
                len(self.doc_text), total_elapsed,
                ", ".join(f"{t:.2f}s" for t in cycle_times), self._doc_truncated,
            )
        except Exception as e:
            error_detail = str(e)
            log_exception(LOGGER, "Failed to load Word document context", e)
            self._set_word_status("Z7: Erro ao carregar contexto do documento no Chat.")
            self.doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."
            self._doc_load_error = error_detail
            LOGGER.error(
                "[doc-load] FAILURE after %.2fs (cycles=%s): %s",
                total_elapsed,
                ", ".join(f"{t:.2f}s" for t in cycle_times), error_detail,
            )

    def _start_streaming(self) -> None:
        """Begin a streaming AI response: insert the 'LÉIA:' header and prepare for chunks."""
        self._streaming_active = True
        self._streamed_text = ""

        self.chat_area.config(state=tk.NORMAL)
        # Check if we need a separator before the AI header
        content = self.chat_area.get("1.0", "end-1c")
        if content and not content.endswith("\n"):
            self.chat_area.insert(tk.END, "\n")
        self.chat_area.insert(tk.END, "LÉIA:\n", "ai_tag")
        self._streaming_start = self.chat_area.index(tk.END + "-1c")
        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def _append_streaming_chunk(self, text: str) -> None:
        """Append a text chunk to the current streaming AI response."""
        if not self._streaming_active:
            return
        self._streamed_text += text
        self.chat_area.config(state=tk.NORMAL)
        self.chat_area.insert(tk.END, text, "ai_msg")
        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

    def _finalize_streaming(self) -> None:
        """Finalize the streaming AI response: add the separator and clean up state."""
        self.chat_area.config(state=tk.NORMAL)
        if not self.chat_area.get("1.0", "end-1c").endswith("\n"):
            self.chat_area.insert(tk.END, "\n")
        self.chat_area.insert(tk.END, "─" * 60 + "\n\n", "msg_sep")
        self.chat_area.see(tk.END)
        self.chat_area.config(state=tk.DISABLED)

        self._streaming_active = False
        self._streaming_start = None

    def _call_api(self) -> str:
        """Envia o histórico completo de mensagens para a API e retorna a resposta.

        Utiliza streaming para exibir a resposta em tempo real na interface.
        Se o modelo primário falhar, tenta automaticamente o modelo fallback.
        """
        api_messages = [{"role": "system", "content": self.system_instruction}]
        api_messages.extend(self.messages)

        self.root.after(0, self._start_streaming)

        try:
            stream = self.client.chat.completions.create(
                model=self._model,
                messages=api_messages,
                stream=True,
                timeout=120,
            )
            accumulated = ""
            for chunk in stream:
                if self._cancel_requested:
                    LOGGER.info("Streaming cancelled by user")
                    break
                if (chunk.choices
                        and chunk.choices[0].delta
                        and chunk.choices[0].delta.content):
                    piece = chunk.choices[0].delta.content
                    accumulated += piece
                    self.root.after(0, self._append_streaming_chunk, piece)

            self.root.after(0, self._finalize_streaming)
            self._streamed_reply = True
            return accumulated

        except Exception as primary_err:
            LOGGER.warning("Primary model streaming failed: %s", primary_err)

            # Se foi cancelamento, não tenta fallback
            if self._cancel_requested:
                self.root.after(0, self._finalize_streaming)
                self._streamed_reply = True
                return ""

            # Limpa texto parcial que pode ter sido inserido antes do erro
            def _cleanup_partial() -> None:
                try:
                    if (self._streaming_active
                            and self._streaming_start is not None):
                        self.chat_area.config(state=tk.NORMAL)
                        self.chat_area.delete(self._streaming_start, tk.END)
                        self.chat_area.config(state=tk.DISABLED)
                    self._streaming_active = False
                    self._streaming_start = None
                except Exception:
                    self._streaming_active = False
                    self._streaming_start = None

            self.root.after(0, _cleanup_partial)

            # ── Fallback: tenta modelo alternativo (sem streaming) ──
            try:
                LOGGER.info("Attempting fallback model: %s", self._fallback_model)
                response = self.client.chat.completions.create(
                    model=self._fallback_model,
                    messages=api_messages,
                    timeout=120,
                )
                fallback_text = response.choices[0].message.content or ""
                # Exibe a resposta do fallback via streaming simulado
                self.root.after(0, self._start_streaming)
                self.root.after(0, self._append_streaming_chunk, fallback_text)
                self.root.after(0, self._finalize_streaming)
                self._streamed_reply = True
                return fallback_text
            except Exception as fallback_err:
                LOGGER.error("Fallback model also failed: %s", fallback_err)
                raise primary_err

    def init_ai(self) -> None:
        # Inicialização em background para não travar a UI
        threading.Thread(target=self._init_ai_thread, daemon=True).start()
        
    def _init_ai_thread(self) -> None:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            _configure_ssl_certifi()
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

            # --- Carrega modelo fallback ---
            try:
                from config_prompt import load_fallback_model
                self._fallback_model = load_fallback_model()
                LOGGER.info("Fallback model loaded via config_prompt: %s", self._fallback_model)
            except Exception as e:
                log_exception(LOGGER, "Failed to load fallback model for chat_ia", e)
                self._fallback_model = _FALLBACK_MODEL

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
                LOGGER.info(
                    "Document context pre-seeded in chat history: %d chars (truncated=%s)",
                    len(doc_context), _doc_truncated,
                )
            else:
                error_detail = getattr(self, '_doc_load_error', '')
                # Classifica a ausência de contexto para diagnóstico preciso
                if error_detail:
                    reason = f"load_error: {error_detail}"
                    self.initial_greeting = (
                        "⚠ Não consegui acessar o documento atual.\n\n"
                        f"Erro: {error_detail}\n\n"
                        "💡 Dica: Certifique-se de que o Word está aberto com um documento ativo. "
                        "Você pode digitar 'recarregar contexto' para tentar novamente."
                    )
                elif doc_context is not None and not doc_context.strip():
                    reason = "blank_document"
                    self.initial_greeting = (
                        "📄 O documento no Word está em branco.\n\n"
                        "Escreva ou cole o conteúdo da propositura no Word e, em seguida, "
                        "use o comando \"recarregar contexto\" aqui no chat para que eu possa analisá-lo."
                    )
                    LOGGER.info("Document is blank/empty")
                else:
                    reason = "no_document_context"
                    self.initial_greeting = (
                        "⚠ Não foi possível obter o conteúdo do documento.\n\n"
                        "💡 Dica: Certifique-se de que o Word está aberto com um documento ativo. "
                        "Você pode digitar 'recarregar contexto' para tentar novamente."
                    )
                _truncation_notice = False
                LOGGER.warning(
                    "No document context available for AI initialization (reason=%s, doc_text_len=%d)",
                    reason, len(doc_context or ""),
                )

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
        # Exibe a mensagem fixa pré-definida
        self.append_message("User", _FIRST_USER_MSG)
        # Inicia a análise inicial se houver documento
        if self.doc_text and self.doc_text.strip() and "nenhum documento" not in self.doc_text.lower():
            self._start_initial_analysis()
        else:
            error_detail = getattr(self, '_doc_load_error', '')
            if error_detail:
                self.append_message("AI", f"⚠ Não consegui acessar o documento.\n\nErro: {error_detail}")
            else:
                self.append_message("AI", "⚠ Nenhum documento ativo encontrado. Abra um documento no Word.")

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
        if not getattr(self, '_streamed_reply', False):
            self.append_message("AI", reply)
        self._streamed_reply = False
        self.update_status("Pronto para conversar")

def _activate_existing_window(title_prefix: str) -> bool:
    """Check if a window with *title_prefix* already exists.

    If found, bring it to front (restoring from minimized state if needed)
    and return ``True``.  Otherwise return ``False``.
    """
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32

    WNDENUMPROC = ctypes.WINFUNCTYPE(
        wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
    )

    found_hwnd: list[wintypes.HWND | None] = [None]

    def _enum_cb(hwnd: wintypes.HWND, _lparam: wintypes.LPARAM) -> bool:
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                if buf.value.startswith(title_prefix):
                    found_hwnd[0] = hwnd
                    return False  # stop enumeration
        return True

    user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)

    if found_hwnd[0] is not None:
        hwnd = found_hwnd[0]
        # SW_RESTORE (9) if minimized, otherwise SW_SHOW (5)
        SW_RESTORE = 9
        SW_SHOW = 5
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        else:
            user32.ShowWindow(hwnd, SW_SHOW)
        user32.SetForegroundWindow(hwnd)
        return True
    return False


def main() -> None:
    # Prevent duplicate instances: if a chat window is already open, just
    # bring it to front and exit.
    if _activate_existing_window("Chat com a LÉIA"):
        return

    root = tk.Tk()
    ChatApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
