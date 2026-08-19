import json
import tkinter as tk
from tkinter import ttk
from pathlib import Path
import z7_theme
from z7_logging import configure_component_logger, get_data_dir, log_exception
import threading

LOGGER = configure_component_logger("config_prompt")

_APP_VERSION = "8.7.2"
_APP_AUTHOR  = "CMS"
_ORG         = "Câmara Municipal de Santa Bárbara d'Oeste"
_LICENSE     = "GPL-3.0"
_MOTTO       = "Dharma, virtude e gratidão."

_DEFAULT_PREFIX = """As strings \"$NUMERO$/$ANO$\", \"$ANO$\" e \"$DATAATUALEXTENSO$\" são placeholders de template utilizados pelo sistema de padronização automática e estão CORRETAS no documento. Elas NÃO devem ser consideradas erros, inconsistências ou problemas de qualquer tipo. Ignore-as completamente na verificação de consistência de datas e não as compare com outras datas do documento. Também não as aponte como erros ortográficos, de formatação ou de conteúdo. Em particular, NÃO analise, NÃO critique e NÃO sugira alterações envolvendo a string \"$NUMERO$/$ANO$\" nem qualquer combinação numérica/ano que esteja nesse formato de placeholder.
A verificação de consistência deve também verificar e apontar erros gramaticais graves.
A verificação de consistência deverá verificar as referências normativas do documento sob os seguintes requisitos:
- Se o documento/propositura for uma indicação, o texto deverá fazer referência expressa ao Art. 108 do Regimento Interno;
- Se o documento/propositura for um Requerimento de Informações, o texto deverá fazer referência expressa ao Art. 10, Inciso X, da Lei Orgânica do município de Santa Bárbara d'Oeste, combinado com o Art. 63, Inciso IX, do mesmo diploma legal;
- Se o documento/propositura for um Requerimento de Pesar, o texto deverá fazer referência expressa ao Art. 102, Inciso IV, do Regimento Interno;
- Se o documento/propositura for uma Moção, o texto deverá fazer referência expressa ao Art. 92, do Capítulo IV, Título V, do Regimento Interno.
Se houver perguntas no documento, a verificação de consistência deve verificar se elas são consistentes e coerentes ao contexto do documento."""

DEFAULT_PROMPT = f"""{_DEFAULT_PREFIX}

Você é um especialista em análise jurídica e legislativa no idioma Português do Brasil.
Analise criteriosamente a propositura legislativa abaixo em busca de inconsistências.

Considere como inconsistências:
1. Divergências entre o conteúdo da ementa e o restante do texto;
2. Contradições entre nomes de pessoas, lugares, logradouros ou endereços;
3. Contradições entre tipos de moção ou natureza do ato legislativo;
4. Contradições lógicas ou semânticas internas ao texto;
5. Qualquer incoerência que comprometa a validade ou o sentido jurídico do documento;
6. Erros gramaticais graves (por exemplo, falhas graves de concordância nominal ou verbal, erros ortográficos crassos, desvios sérios de regência, ou problemas de pontuação que prejudiquem a clareza e compreensão da matéria);
7. Falha ou ausência de referências normativas obrigatórias de acordo com o tipo de propositura:
   - Se a propositura for uma Indicação, o texto deverá fazer referência expressa ao "Art. 108 do Regimento Interno";
   - Se for um Requerimento de Informações, o texto deverá fazer referência expressa ao "Art. 10, Inciso X, da Lei Orgânica do município de Santa Bárbara d'Oeste, combinado com o Art. 63, Inciso IX, do mesmo diploma legal";
   - Se for um Requerimento de Pesar, o texto deverá fazer referência expressa ao "Art. 102, Inciso IV, do Regimento Interno";
   - Se for uma Moção, o texto deverá fazer referência expressa ao "Art. 92, do Capítulo IV, Título V, do Regimento Interno".
8. Perguntas no documento que não sejam consistentes ou coerentes ao contexto do documento.


Não reporte como problemas:
- Pequenas divergências de grafia ou acentuação;
- Diferenças na ordem de palavras que não alterem o sentido;
- Pequenos erros formais ou desvios gramaticais leves que não comprometam a estrutura ou a lógica do texto (erros gramaticais graves, contudo, devem ser apontados);
- As strings "$NUMERO$/$ANO$", "$ANO$" e "$DATAATUALEXTENSO$" são placeholders de template do sistema de padronização e estão corretas. NÃO as reporte como erros, inconsistências ou problemas de qualquer natureza.

Se encontrar inconsistências, classifique-as por grau de gravidade (Crítica, Alta, Média, Baixa) e apresente-as ordenadas da mais grave para a menos grave. Cada item deve indicar claramente o nível de severidade e o trecho conflitante, de forma sucinta e direta, sem alongar-se em explicações teóricas ou doutrinárias.
Priorize sempre a apresentação das sugestões de correção em detrimento de explicações detalhadas dos erros encontrados. Para cada problema, apresente imediatamente a correção sugerida — preferencialmente na forma de texto revisado pronto para substituição.
Se nenhum problema for encontrado, informe que o documento está consistente.

Responda em Português do Brasil."""

DEFAULT_CONSISTENCY_PROMPT = f"""{_DEFAULT_PREFIX}

Regras de Classificação e Limites:
titulo: A primeira linha do documento, geralmente em caixa alta, contendo a natureza da propositura e as marcações de número/ano.
ementa: O parágrafo logo abaixo do título, que resume o objeto da matéria e geralmente começa com um verbo de ação (Indica, Requer, Manifesta).
vocativo: O cumprimento formal direcionado à autoridade ou aos pares. Pode ter uma ou múltiplas linhas.
proposicao: O corpo principal e variável do texto. Inicia logo após o vocativo e termina imediatamente antes do título da justificativa. Pode conter parágrafos de contextualização (CONSIDERANDO) e os pedidos ou apelos em si.
titulo_da_justificativa: Exatamente a marcação textual que introduz a argumentação.
justificativa: O texto argumentativo completo. Inicia logo após o título da justificativa e vai até antes da data.
data: A linha que marca o local, o nome do plenário e a data de emissão.
assinatura: O bloco final do documento contendo a indicação de autoria e cargo.

Exemplo de Entrada e Saída Esperada:
Entrada:
INDICAÇÃO Nº $NUMERO$/$ANO$
Indica ao Poder Executivo Municipal a ampliação da rede de creches nos bairros com maior demanda por vagas.
Excelentíssimo Senhor Prefeito Municipal,
Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excelência para indicar que seja realizado um estudo técnico para ampliação da rede de creches públicas, com prioridade aos bairros com maior número de crianças em lista de espera, como o Jardim São Fernando e o Parque Zabani, neste Município.
Justificativa:
A falta de vagas em creches tem afetado diretamente as famílias, em especial mães que dependem do serviço para poder trabalhar. A ampliação do número de unidades ou convênios com instituições qualificadas atenderá à demanda crescente e garantirá o direito à educação infantil.
Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$.
AUTORIA
– Vereador –

Saída JSON:
{{"titulo": "INDICAÇÃO Nº $NUMERO$/$ANO$","ementa": "Indica ao Poder Executivo Municipal a ampliação da rede de creches nos bairros com maior demanda por vagas.","vocativo": "Excelentíssimo Senhor Prefeito Municipal,","proposicao": "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excelência para indicar que seja realizado um estudo técnico para ampliação da rede de creches públicas, com prioridade aos bairros com maior número de crianças em lista de espera, como o Jardim São Fernando e o Parque Zabani, neste Município.","titulo_da_justificativa": "Justificativa:","justificativa": "A falta de vagas em creches tem afetado diretamente as famílias, em especial mães que dependem do serviço para poder trabalhar. A ampliação do número de unidades ou convênios com instituições qualificadas atenderá à demanda crescente e garantirá o direito à educação infantil.","data": "Plenário \\"Dr. Tancredo Neves\\", $DATAATUALEXTENSO$.","assinatura": "AUTORIA\\n– Vereador –"}}"""


DEFAULT_CHAT_SYSTEM_PROMPT = "Você é a LÉIA — Assistente Legislativa de IA. Sempre se apresente como LÉIA. Leia o documento ativo. Identifique erros. Auxilie o usuário alterando, revisando ou tirando dúvidas. Ao apontar problemas, apresente-os de forma sucinta, ordenados por severidade (Crítica → Alta → Média → Baixa), e priorize sempre as sugestões de correção em vez de explicações detalhadas dos erros. As strings `$NUMERO$/$ANO$`, `$ANO$` e `$DATAATUALEXTENSO$` são placeholders de template do sistema de padronização automática — estão corretas e NÃO devem ser apontadas como erros."


def get_chat_system_prompt_file_path() -> Path:
    return get_data_dir() / "chat_system_prompt.txt"


def load_chat_system_prompt() -> str:
    prompt_file = get_chat_system_prompt_file_path()
    if prompt_file.exists():
        try:
            text = prompt_file.read_text(encoding='utf-8').strip()
            LOGGER.info("Loaded custom chat system prompt file")
            return text
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom chat system prompt", e)
    return DEFAULT_CHAT_SYSTEM_PROMPT


def save_chat_system_prompt(prompt_text: str) -> None:
    prompt_file = get_chat_system_prompt_file_path()
    try:
        prompt_file.write_text(prompt_text.strip(), encoding='utf-8')
        LOGGER.info("Chat system prompt saved successfully")
    except Exception as e:
        log_exception(LOGGER, "Failed to save chat system prompt", e)


def get_prompt_file_path() -> Path:
    return get_data_dir() / "gemini_prompt.txt"


def get_consistency_prompt_file_path() -> Path:
    return get_data_dir() / "consistency_prompt.txt"


def load_api_key() -> str:
    from z7_api_key import read_stored_api_key
    return read_stored_api_key()

def save_api_key(api_key: str) -> None:
    from z7_api_key import write_api_key
    write_api_key(api_key)

def load_prompt() -> str:
    prompt_file = get_prompt_file_path()
    if prompt_file.exists():
        try:
            text = prompt_file.read_text(encoding='utf-8')
            LOGGER.info("Loaded custom prompt file")
            return text
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom prompt", e)
    return DEFAULT_PROMPT


def load_consistency_prompt() -> str:
    consistency_file = get_consistency_prompt_file_path()
    if consistency_file.exists():
        try:
            text = consistency_file.read_text(encoding='utf-8')
            LOGGER.info("Loaded custom consistency prompt file")
            return text
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom consistency prompt", e)
    return DEFAULT_CONSISTENCY_PROMPT

def get_model_file_path() -> Path:
    return get_data_dir() / "selected_model.txt"

def load_ai_model() -> str:
    model_file = get_model_file_path()
    if model_file.exists():
        try:
            return model_file.read_text(encoding='utf-8').strip()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom model", e)
    return "deepseek/deepseek-v4-pro"

def save_ai_model(model_name: str) -> None:
    model_file = get_model_file_path()
    try:
        model_file.write_text(model_name, encoding='utf-8')
        LOGGER.info("Model saved successfully: %s", model_name)
    except Exception as e:
        log_exception(LOGGER, "Failed to save model", e)

def get_fallback_model_file_path() -> Path:
    return get_data_dir() / "selected_fallback_model.txt"

def load_fallback_model() -> str:
    fallback_file = get_fallback_model_file_path()
    if fallback_file.exists():
        try:
            return fallback_file.read_text(encoding='utf-8').strip()
        except Exception as e:
            log_exception(LOGGER, "Failed to load fallback model", e)
    return "deepseek/deepseek-v4-flash"

def save_fallback_model(model_name: str) -> None:
    fallback_file = get_fallback_model_file_path()
    try:
        fallback_file.write_text(model_name, encoding='utf-8')
        LOGGER.info("Fallback model saved successfully: %s", model_name)
    except Exception as e:
        log_exception(LOGGER, "Failed to save fallback model", e)

def save_prompt(consistency_text: str, root: tk.Tk) -> None:
    if not consistency_text.strip():
        LOGGER.warning("Consistency prompt save blocked because text is empty")
        z7_theme.show_warning("Aviso", "O prompt do Verificador de Consistência não pode estar vazio.", parent=root)
        return

    consistency_file = get_consistency_prompt_file_path()
    try:
        consistency_file.write_text(consistency_text.strip(), encoding='utf-8')
        LOGGER.info("Consistency prompt saved successfully")
    except Exception as e:
        log_exception(LOGGER, "Failed to save consistency prompt", e)
        z7_theme.show_error("Erro", f"Erro ao salvar prompt de consistência:\n{e}", parent=root)
        return

    root.destroy()


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
        btn_primary_bg = colors["btn_primary_bg"]
        btn_primary_fg = colors["btn_primary_fg"]
        btn_primary_hover = colors["btn_primary_hover"]
        select_bg = "#8b5cf6" if self.mode == "dark" else "#7c3aed"
        select_fg = "#ffffff"
            
        self.root.configure(bg=bg)
        
        # ── Header area ───────────────────────────────────────────────────
        if 'header_frame' in self.widgets:
            self.widgets['header_frame'].configure(bg=bg)
        if 'title_row' in self.widgets:
            self.widgets['title_row'].configure(bg=bg)
        if 'title_lbl' in self.widgets:
            self.widgets['title_lbl'].configure(bg=bg, fg=fg)
        if 'version_badge' in self.widgets:
            self.widgets['version_badge'].configure(bg=colors["btn_primary_bg"], fg="white")
        if 'info_lbl' in self.widgets:
            self.widgets['info_lbl'].configure(bg=bg, fg=fg_muted)
        if 'separator' in self.widgets:
            self.widgets['separator'].configure(bg=border)

        # ── API button ───────────────────────────────────────────────────
        if 'api_btn' in self.widgets:
            self.widgets['api_btn'].configure(bg=btn_primary_bg, fg=btn_primary_fg,
                                              activebackground=btn_primary_hover, activeforeground=btn_primary_fg)
        if 'api_btn_frame' in self.widgets:
            self.widgets['api_btn_frame'].configure(bg=bg)

        # ── Text area ─────────────────────────────────────────────────────
        if 'border_frame' in self.widgets:
            self.widgets['border_frame'].configure(bg=border)
        if 'prompt_label' in self.widgets:
            self.widgets['prompt_label'].configure(bg=bg, fg=fg_muted)
        if 'text_inner' in self.widgets:
            self.widgets['text_inner'].configure(bg=border)
        if 'text_area' in self.widgets:
            self.widgets['text_area'].configure(
                bg=text_bg, fg=fg, insertbackground=fg,
                selectbackground=select_bg, selectforeground=select_fg,
                inactiveselectbackground=select_bg)
        if 'chat_text_area' in self.widgets:
            self.widgets['chat_text_area'].configure(
                bg=text_bg, fg=fg, insertbackground=fg,
                selectbackground=select_bg, selectforeground=select_fg,
                inactiveselectbackground=select_bg)
        if 'notebook' in self.widgets:
            try:
                style = ttk.Style()
                style.configure('TNotebook', background=bg)
                style.configure('TNotebook.Tab', background=btn_sec_bg, foreground=btn_sec_fg, padding=[10, 4])
                style.map('TNotebook.Tab',
                    background=[('selected', bg)],
                    foreground=[('selected', fg)])
            except Exception:
                pass

        # ── Buttons area ──────────────────────────────────────────────────
        if 'btn_frame' in self.widgets:
            self.widgets['btn_frame'].configure(bg=bg)

        for btn in self.widgets.get('sec_btns', []):
            btn.configure(bg=btn_sec_bg, fg=btn_sec_fg, activebackground=btn_sec_hover, activeforeground=fg)
        
        if 'toggle_btn' in self.widgets:
            icon = "🌙 Modo Escuro" if self.mode == 'light' else "☀️ Modo Claro"
            self.widgets['toggle_btn'].configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)

        # ── Footer ────────────────────────────────────────────────────────
        if 'footer_lbl' in self.widgets:
            self.widgets['footer_lbl'].configure(bg=bg, fg=fg_muted)


def open_ai_api_dialog(
    parent: tk.Tk,
    theme_mode: str,
    on_saved: "Callable[[str, str], None] | None" = None,
) -> None:
    """Open the AI API settings dialog.

    Shows fields for the OpenRouter API key and AI model name. A "Salvar"
    button persists both values. A "Testar Modelo" button performs a live
    connection test and displays the result in an output area.

    Args:
        parent: The root or toplevel window to centre the dialog over.
        theme_mode: Current theme mode ('light' or 'dark').
        on_saved: Optional callback invoked with ``(api_key, model)``
                  after a successful save.
    """
    import re as _re
    import threading
    import webbrowser

    colors = z7_theme.get_theme_colors(theme_mode)
    bg           = colors["bg"]
    fg           = colors["fg"]
    fg_muted     = colors["fg_muted"]
    text_bg      = colors["text_bg"]
    border       = colors["border"]
    btn_sec_bg   = colors["btn_sec_bg"]
    btn_sec_fg   = colors["btn_sec_fg"]
    btn_sec_hover = colors["btn_sec_hover"]
    btn_primary  = colors.get("btn_primary_bg", "#2563eb")
    btn_primary_hover = colors.get("btn_primary_hover", "#6d28d9")
    select_bg = "#8b5cf6" if theme_mode == "dark" else "#7c3aed"
    select_fg = "#ffffff"

    _RE_API_KEY = _re.compile(
        r"^(sk-or-v1-[0-9a-zA-F]{64}|sk-[0-9A-Za-z\-_]{16,}|[0-9A-Za-z\-_]{16,})$"
    )

    apikey_visible: list[bool] = [False]

    dialog = tk.Toplevel(parent)
    dialog.title("API de IA (OpenRouter)")
    dialog.geometry("520x620")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()
    dialog.configure(bg=bg)

    # Bring dialog to front without staying always-on-top
    dialog.after(10, dialog.lift)

    # Centres dialog over parent
    dialog.update_idletasks()
    px, py = parent.winfo_x(), parent.winfo_y()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    dialog.geometry(f"520x620+{px + (pw - 520) // 2}+{py + (ph - 620) // 2}")

    # ── Status label ──────────────────────────────────────────────────────────
    _current_key = load_api_key()
    _status_lbl = tk.Label(
        dialog,
        text="✔  Chave configurada" if _current_key else "⚠  Chave não configurada",
        font=("Segoe UI", 11, "bold"),
        fg="#10b981" if _current_key else "#f59e0b",
        bg=bg, anchor="w",
    )
    _status_lbl.pack(fill=tk.X, padx=22, pady=(16, 2))

    # ── Section: API Key ──────────────────────────────────────────────────────
    tk.Label(
        dialog, text="CHAVE OPENROUTER API",
        font=("Segoe UI", 10, "bold"), fg=fg_muted, bg=bg, anchor="w",
    ).pack(fill=tk.X, padx=22, pady=(10, 2))
    tk.Frame(dialog, height=1, bg=border).pack(fill=tk.X, padx=22, pady=(0, 8))

    api_frame = tk.Frame(dialog, bg=bg)
    api_frame.pack(fill=tk.X, padx=22)

    api_var = tk.StringVar(value=_current_key)
    api_entry = tk.Entry(
        api_frame, textvariable=api_var, font=("Segoe UI", 11),
        relief=tk.FLAT, bd=2, show="•", bg=text_bg, fg=fg,
        insertbackground=fg,
        selectbackground=select_bg, selectforeground=select_fg,
    )
    api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    api_entry.focus_set()

    def _toggle_api_visibility() -> None:
        apikey_visible[0] = not apikey_visible[0]
        api_entry.config(show="" if apikey_visible[0] else "•")
        vis_btn.config(text="👁 Ocultar" if apikey_visible[0] else "👁 Mostrar")

    vis_btn = tk.Button(
        api_frame, text="👁 Mostrar", font=("Segoe UI", 9),
        relief=tk.FLAT, cursor="hand2",
        bg=btn_sec_bg, fg=btn_sec_fg,
        activebackground=btn_sec_hover, activeforeground=fg,
        command=_toggle_api_visibility,
    )
    vis_btn.pack(side=tk.LEFT, padx=(6, 0))

    # ── Section: AI Model ─────────────────────────────────────────────────────
    tk.Label(
        dialog, text="MODELO IA",
        font=("Segoe UI", 10, "bold"), fg=fg_muted, bg=bg, anchor="w",
    ).pack(fill=tk.X, padx=22, pady=(14, 2))
    tk.Frame(dialog, height=1, bg=border).pack(fill=tk.X, padx=22, pady=(0, 8))

    model_var = tk.StringVar(value=load_ai_model())
    model_entry = tk.Entry(
        dialog, textvariable=model_var, font=("Segoe UI", 11),
        relief=tk.FLAT, bd=2, bg=text_bg, fg=fg, insertbackground=fg,
        selectbackground=select_bg, selectforeground=select_fg,
    )
    model_entry.pack(fill=tk.X, padx=22)

    # ── Section: Fallback Model ────────────────────────────────────────────────
    tk.Label(
        dialog, text="MODELO FALLBACK (alternativo)",
        font=("Segoe UI", 10, "bold"), fg=fg_muted, bg=bg, anchor="w",
    ).pack(fill=tk.X, padx=22, pady=(14, 2))
    tk.Frame(dialog, height=1, bg=border).pack(fill=tk.X, padx=22, pady=(0, 8))

    fallback_var = tk.StringVar(value=load_fallback_model())
    fallback_entry = tk.Entry(
        dialog, textvariable=fallback_var, font=("Segoe UI", 11),
        relief=tk.FLAT, bd=2, bg=text_bg, fg=fg, insertbackground=fg,
        selectbackground=select_bg, selectforeground=select_fg,
    )
    fallback_entry.pack(fill=tk.X, padx=22)

    # ── Output area ───────────────────────────────────────────────────────────
    tk.Label(
        dialog, text="SAÍDA",
        font=("Segoe UI", 9, "bold"), fg=fg_muted, bg=bg, anchor="w",
    ).pack(fill=tk.X, padx=22, pady=(14, 2))

    output_frame = tk.Frame(dialog, bg=border)
    output_frame.pack(fill=tk.BOTH, expand=True, padx=22)

    output_box = tk.Text(
        output_frame, wrap=tk.WORD, font=("Consolas", 10),
        relief=tk.FLAT, bd=0, padx=10, pady=8,
        bg=text_bg, fg=fg, insertbackground=fg,
        selectbackground=select_bg, selectforeground=select_fg,
        inactiveselectbackground=select_bg,
        state=tk.DISABLED, height=6,
    )
    output_box.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

    output_box.tag_config("success", foreground="#10b981")
    output_box.tag_config("error", foreground="#ef4444")
    output_box.tag_config("warn", foreground="#f59e0b")
    output_box.tag_config("dim", foreground=fg_muted)

    def _append(text: str, tag: str = "") -> None:
        output_box.config(state=tk.NORMAL)
        if tag:
            output_box.insert(tk.END, text + "\n", tag)
        else:
            output_box.insert(tk.END, text + "\n")
        output_box.see(tk.END)
        output_box.config(state=tk.DISABLED)

    def _clear_output() -> None:
        output_box.config(state=tk.NORMAL)
        output_box.delete("1.0", tk.END)
        output_box.config(state=tk.DISABLED)

    # ── Buttons ───────────────────────────────────────────────────────────────
    save_btn = tk.Button(
        dialog, text="💾  Salvar", font=("Segoe UI", 12, "bold"),
        relief=tk.FLAT, cursor="hand2", pady=8,
        bg=btn_primary, fg="white",
        activebackground=btn_primary_hover, activeforeground="white",
    )
    save_btn.pack(fill=tk.X, padx=22, pady=(14, 4))

    test_btn = tk.Button(
        dialog, text="🧪  Testar Modelo", font=("Segoe UI", 10),
        relief=tk.FLAT, cursor="hand2", pady=6,
        bg=btn_sec_bg, fg=btn_sec_fg,
        activebackground=btn_sec_hover, activeforeground=fg,
    )
    test_btn.pack(fill=tk.X, padx=22, pady=(0, 4))

    web_btn = tk.Button(
        dialog, text="🌐  OpenRouter Keys", font=("Segoe UI", 10),
        relief=tk.FLAT, cursor="hand2", pady=4,
        bg=bg, fg=fg_muted,
        activebackground=bg, activeforeground=fg,
        command=lambda: webbrowser.open("https://openrouter.ai/keys"),
    )
    web_btn.pack(fill=tk.X, padx=22, pady=(0, 16))

    # ── Validation ────────────────────────────────────────────────────────────
    def _validate_inputs() -> "tuple[str, str, str] | None":
        _clear_output()
        key = api_var.get().strip()
        model = model_var.get().strip()
        fallback = fallback_var.get().strip()
        effective_key = key or load_api_key()

        if not effective_key:
            _append("⚠  Informe uma chave de API.", "warn")
            return None
        if not model:
            _append("⚠  Informe um nome de modelo.", "warn")
            return None
        if not fallback:
            _append("⚠  Informe um nome de modelo fallback.", "warn")
            return None
        if key and not _RE_API_KEY.match(key):
            _append(
                "✘  Formato de chave inválido.\n"
                '   Chaves OpenRouter costumam ter o formato "sk-or-v1-…".',
                "error",
            )
            return None
        return effective_key, model, fallback

    # ── Save ──────────────────────────────────────────────────────────────────
    def _on_save() -> None:
        validated = _validate_inputs()
        if validated is None:
            return
        effective_key, model, fallback = validated

        save_btn.config(state=tk.DISABLED)
        test_btn.config(state=tk.DISABLED)
        _append("Salvando chave e modelos…", "dim")

        def _do_save() -> None:
            try:
                if effective_key != load_api_key():
                    save_api_key(effective_key)
                    dialog.after(0, lambda: _append("✔  Chave salva.", "success"))
                else:
                    dialog.after(0, lambda: _append("ℹ  Usando chave já armazenada.", "dim"))
                save_ai_model(model)
                dialog.after(0, lambda: _append("✔  Modelo salvo.", "success"))
                save_fallback_model(fallback)
                dialog.after(0, lambda: _append("✔  Modelo fallback salvo.", "success"))
                dialog.after(0, lambda: _close_and_save(effective_key, model))
            except Exception as exc:
                err_msg = str(exc)
                dialog.after(0, lambda: _append(f"✘  Falha ao salvar: {err_msg}", "error"))
            finally:
                try:
                    dialog.after(0, lambda: (
                        save_btn.config(state=tk.NORMAL),
                        test_btn.config(state=tk.NORMAL),
                    ))
                except Exception:
                    pass

        threading.Thread(target=_do_save, daemon=True).start()

    # ── Test Model ────────────────────────────────────────────────────────────
    def _on_test() -> None:
        validated = _validate_inputs()
        if validated is None:
            return
        effective_key, model, fallback = validated

        test_btn.config(state=tk.DISABLED)
        save_btn.config(state=tk.DISABLED)

        def _do_test() -> None:
            try:
                import os as _os
                try:
                    import certifi as _certifi
                    _os.environ.setdefault("SSL_CERT_FILE", _certifi.where())
                except ImportError:
                    LOGGER.warning("certifi not available; SSL may fail in frozen environment")

                from openai import OpenAI

                dialog.after(0, lambda: _append("Testando conexão com a API OpenRouter…", "dim"))

                client = OpenAI(
                    base_url="https://openrouter.ai/api/v1",
                    api_key=effective_key,
                )

                # Testa modelo primário
                dialog.after(0, lambda: _append(f"  Modelo primário: {model}", "dim"))
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": "Responda apenas com a palavra: OK"}],
                )
                resp_text = ""
                if getattr(response, "choices", None) and len(response.choices) > 0:
                    resp_text = (response.choices[0].message.content or "").strip()

                if not resp_text:
                    raise ValueError("A IA não retornou conteúdo na resposta.")

                dialog.after(0, lambda: _append("✔  Modelo primário respondeu — configuração válida.", "success"))
                dialog.after(0, lambda: _append(f"   Resposta: {resp_text}"))

                # Testa modelo fallback
                dialog.after(0, lambda: _append(f"  Modelo fallback: {fallback}", "dim"))
                try:
                    fb_response = client.chat.completions.create(
                        model=fallback,
                        messages=[{"role": "user", "content": "Responda apenas com a palavra: OK"}],
                    )
                    fb_text = ""
                    if getattr(fb_response, "choices", None) and len(fb_response.choices) > 0:
                        fb_text = (fb_response.choices[0].message.content or "").strip()
                    if fb_text:
                        dialog.after(0, lambda: _append("✔  Modelo fallback respondeu — configuração válida.", "success"))
                        dialog.after(0, lambda: _append(f"   Resposta: {fb_text}"))
                    else:
                        dialog.after(0, lambda: _append("⚠  Modelo fallback não retornou conteúdo.", "warn"))
                except Exception as fb_exc:
                    dialog.after(0, lambda: _append(f"⚠  Modelo fallback falhou: {fb_exc}", "warn"))

                dialog.after(0, lambda: _status_lbl.configure(
                    text="✔  Chave configurada", fg="#10b981",
                ))

            except Exception as exc:
                err_msg = str(exc)
                dialog.after(0, lambda: _append(f"✘  Falha na validação: {err_msg}", "error"))
            finally:
                try:
                    dialog.after(0, lambda: (
                        test_btn.config(state=tk.NORMAL),
                        save_btn.config(state=tk.NORMAL),
                    ))
                except Exception:
                    pass

        threading.Thread(target=_do_test, daemon=True).start()

    # ── Close handler ─────────────────────────────────────────────────────────
    def _close_and_save(effective_key: str, model: str) -> None:
        if on_saved is not None:
            on_saved(effective_key, model)
        dialog.destroy()

    save_btn.config(command=_on_save)
    test_btn.config(command=_on_test)

    def _on_close() -> None:
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", _on_close)


# Keep backward-compatible alias for config_prompt internal usage
def open_api_key_dialog(parent: tk.Tk, theme_mode: str) -> None:
    """Legacy wrapper – opens the full AI API dialog (backward compatibility)."""
    open_ai_api_dialog(parent, theme_mode)


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
    # Prevent duplicate instances: if a config window is already open, just
    # bring it to front and exit.
    if _activate_existing_window("Configurar Prompt"):
        return

    LOGGER.info("Starting prompt configuration UI")

    try:
        import win32com.client
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.GetObject(Class="Word.Application")
        word.StatusBar = "Z7: Abrindo Configuracoes..."
    except Exception as e:
        LOGGER.warning("Could not connect to Word to update status bar: %s", str(e))

    root = tk.Tk()
    root.title(f"Configurar Prompt — Z7 StdProposers v{_APP_VERSION}")

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    win_width = int(screen_width * 2 / 3 * 1.15 * 1.10 * 0.70 * 1.10)
    win_height = int(screen_height * 0.92 * 0.90 * 1.05)
    win_left = 0
    win_top = int(screen_height * 0.02)

    root.geometry(f"{win_width}x{win_height}+{win_left}+{win_top}")
    root.minsize(300, 500)

    theme = AppTheme(root)

    # ── Cabeçalho ─────────────────────────────────────────────────────────
    _c = z7_theme.get_theme_colors(theme.mode)
    header_frame = tk.Frame(root, bg=_c["bg"])
    header_frame.pack(fill=tk.X, padx=25, pady=(20, 0))
    theme.widgets['header_frame'] = header_frame

    # Linha 1: Título + versão + tema
    title_row = tk.Frame(header_frame, bg=_c["bg"])
    title_row.pack(fill=tk.X)
    theme.widgets['title_row'] = title_row

    lbl = tk.Label(title_row, text="⚙ Configurações do Prompt",
                   font=("Segoe UI", 16, "bold"), bg=_c["bg"], fg=_c["fg"])
    lbl.pack(side=tk.LEFT, anchor="w")
    theme.widgets['title_lbl'] = lbl

    version_badge = tk.Label(title_row, text=f"v{_APP_VERSION}",
                             font=("Segoe UI", 10, "bold"), padx=6, pady=1,
                             bg=_c["btn_primary_bg"], fg="white")
    version_badge.pack(side=tk.LEFT, anchor="w", padx=(8, 0))
    theme.widgets['version_badge'] = version_badge

    toggle_btn = tk.Button(title_row, font=("Segoe UI", 9), relief=tk.FLAT,
                           cursor="hand2", command=theme.toggle, bd=0,
                           bg=_c["bg"], fg=_c["fg"],
                           activebackground=_c["bg"], activeforeground=_c["fg"])
    toggle_btn.pack(side=tk.RIGHT, anchor="e")
    theme.widgets['toggle_btn'] = toggle_btn

    info_lbl = tk.Label(header_frame,
                        text="Defina o prompt inicial enviado à IA ao analisar o documento.",
                        font=("Segoe UI", 10), bg=_c["bg"], fg=_c["fg_muted"])
    info_lbl.pack(anchor="w", pady=(4, 0))
    theme.widgets['info_lbl'] = info_lbl

    # Separador
    separator = tk.Frame(header_frame, height=1, bg=_c["border"])
    separator.pack(fill=tk.X, pady=(12, 0))
    theme.widgets['separator'] = separator

    # ── Área de texto única ───────────────────────────────────────────────
    frame = tk.Frame(root, bg=_c["border"])
    theme.widgets['border_frame'] = frame

    # Label da área de texto
    prompt_label = tk.Label(frame, text="PROMPT INICIAL COMPLEMENTAR PARA A LÉIA",
                            font=("Segoe UI", 10, "bold"), anchor="w",
                            bg=_c["bg"], fg=_c["fg_muted"])
    prompt_label.pack(fill=tk.X, padx=2, pady=(4, 2))
    theme.widgets['prompt_label'] = prompt_label

    text_inner = tk.Frame(frame, bg=_c["border"])
    text_inner.pack(expand=True, fill=tk.BOTH)
    theme.widgets['text_inner'] = text_inner

    _select_bg = "#8b5cf6" if theme.mode == "dark" else "#7c3aed"
    _select_fg = "#ffffff"
    text_area = tk.Text(text_inner, wrap=tk.WORD, font=("Consolas", 11),
                        relief=tk.FLAT, padx=12, pady=12,
                        bg=_c["text_bg"], fg=_c["fg"], insertbackground=_c["fg"],
                        selectbackground=_select_bg, selectforeground=_select_fg,
                        inactiveselectbackground=_select_bg)
    text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=1, pady=1)
    scrollbar = tk.Scrollbar(text_inner, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)
    text_area.insert(tk.END, load_chat_system_prompt())
    theme.widgets['text_area'] = text_area

    # ── Botões de ação ────────────────────────────────────────────────────
    btn_frame = tk.Frame(root)
    theme.widgets['btn_frame'] = btn_frame

    btn_font = ("Segoe UI", 10, "bold")

    def do_save() -> None:
        prompt_text = text_area.get("1.0", tk.END).strip()
        if not prompt_text:
            z7_theme.show_warning("Aviso",
                                 "O prompt não pode estar vazio.", parent=root)
            return
        save_chat_system_prompt(prompt_text)
        root.destroy()

    save_btn = tk.Button(btn_frame, text="💾  Salvar", width=18,
                         bg=_c["btn_primary_bg"], fg=_c["btn_primary_fg"],
                         font=btn_font, relief=tk.FLAT,
                         activebackground=_c["btn_primary_hover"],
                         activeforeground=_c["btn_primary_fg"],
                         cursor="hand2", command=do_save)
    save_btn.pack(side=tk.RIGHT, padx=25)

    cancel_btn = tk.Button(btn_frame, text="Cancelar", width=12,
                           font=btn_font, relief=tk.FLAT, cursor="hand2",
                           bg=_c["btn_sec_bg"], fg=_c["btn_sec_fg"],
                           activebackground=_c["btn_sec_hover"],
                           activeforeground=_c["fg"],
                           command=root.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=5)

    def do_restore() -> None:
        text_area.delete("1.0", tk.END)
        text_area.insert(tk.END, DEFAULT_CHAT_SYSTEM_PROMPT)

    restore_btn = tk.Button(btn_frame, text="Restaurar Padrão", width=16,
                            font=btn_font, relief=tk.FLAT, cursor="hand2",
                            bg=_c["btn_sec_bg"], fg=_c["btn_sec_fg"],
                            activebackground=_c["btn_sec_hover"],
                            activeforeground=_c["fg"],
                            command=do_restore)
    restore_btn.pack(side=tk.LEFT, padx=25)

    api_btn = tk.Button(btn_frame, text="🔑 API de IA", font=("Segoe UI", 10, "bold"),
                        relief=tk.FLAT, cursor="hand2", padx=12, pady=4,
                        bg=_c["btn_primary_bg"], fg=_c["btn_primary_fg"],
                        activebackground=_c["btn_primary_hover"],
                        activeforeground=_c["btn_primary_fg"],
                        command=lambda: open_ai_api_dialog(root, theme.mode))
    api_btn.pack(side=tk.LEFT, padx=(0, 5))
    theme.widgets['api_btn'] = api_btn

    theme.widgets['sec_btns'] = [cancel_btn, restore_btn]

    # ── Rodapé ────────────────────────────────────────────────────────────
    _footer_text = f"{_ORG}  ·  {_APP_AUTHOR}  ·  {_LICENSE}  ·  {_MOTTO}"
    footer_lbl = tk.Label(root, text=_footer_text, font=("Segoe UI", 8),
                          anchor="e", bg=_c["bg"], fg=_c["fg_muted"])
    theme.widgets['footer_lbl'] = footer_lbl

    # Aplica o tema inicial
    theme.apply()

    # Pack order: footer and buttons fixed at bottom, text area fills rest
    footer_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 5))
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 5))
    frame.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=25, pady=(4, 10))

    # Bring window to front on startup without staying always-on-top
    root.lift()
    root.focus_force()

    root.mainloop()

if __name__ == "__main__":
    main()
