import os
import json
import tkinter as tk
from functools import lru_cache
from pathlib import Path

def get_theme_file_path() -> Path | None:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'theme_config.json'

@lru_cache(maxsize=None)
def load_theme() -> str:
    theme_file = get_theme_file_path()
    if theme_file and theme_file.exists():
        try:
            with open(theme_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get('theme', 'light')
        except Exception:
            pass
    return 'light'

def save_theme(theme_mode: str) -> None:
    theme_file = get_theme_file_path()
    if theme_file:
        try:
            with open(theme_file, 'w', encoding='utf-8') as f:
                json.dump({'theme': theme_mode}, f)
        except Exception:
            pass
    load_theme.cache_clear()
    get_theme_colors.cache_clear()

@lru_cache(maxsize=None)
def get_theme_colors(mode=None):
    if mode is None:
        mode = load_theme()
    if mode == 'dark':
        return {
            "bg": "#1e1e1e",
            "fg": "#e4e4e4",
            "fg_muted": "#a0a0a0",
            "text_bg": "#252526",
            "border": "#333333",
            "btn_sec_bg": "#333333",
            "btn_sec_fg": "#cccccc",
            "btn_sec_hover": "#444444",
            "btn_primary_bg": "#005a9e",
            "btn_primary_fg": "#ffffff",
            "btn_primary_hover": "#004578",
            "user_tag": "#60a5fa",
            "ai_tag": "#34d399"
        }
    else:
        return {
            "bg": "#f3f4f6",
            "fg": "#111827",
            "fg_muted": "#4b5563",
            "text_bg": "#ffffff",
            "border": "#d1d5db",
            "btn_sec_bg": "#e5e7eb",
            "btn_sec_fg": "#374151",
            "btn_sec_hover": "#d1d5db",
            "btn_primary_bg": "#2563eb",
            "btn_primary_fg": "#ffffff",
            "btn_primary_hover": "#1d4ed8",
            "user_tag": "#2563eb",
            "ai_tag": "#10b981"
        }

def _center_window(win, parent=None):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    if parent and parent.winfo_viewable():
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (width // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (height // 2)
    else:
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f'+{x}+{y}')

def _create_dialog(title, message, parent=None, is_prompt=False, is_cancelable=False, show_char=""):
    colors = get_theme_colors()
    
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.attributes('-topmost', True)
    dlg.configure(bg=colors["bg"], padx=20, pady=20)
    dlg.resizable(False, False)
    
    if parent and parent.winfo_viewable():
        dlg.transient(parent)
    dlg.grab_set()
    dlg.after(100, lambda: (dlg.lift(), dlg.focus_force()))
    
    tk.Label(dlg, text=message, font=("Segoe UI", 10), bg=colors["bg"], fg=colors["fg"], justify=tk.LEFT, wraplength=400).pack(pady=(0, 15))
    
    result = {"value": None}
    
    entry = None
    if is_prompt:
        entry = tk.Entry(dlg, font=("Segoe UI", 10), bg=colors["text_bg"], fg=colors["fg"], insertbackground=colors["fg"], relief=tk.FLAT, show=show_char)
        entry.pack(fill=tk.X, pady=(0, 15))
        # Add a border frame effect
        tk.Frame(dlg, bg=colors["border"], height=1).pack(fill=tk.X, pady=(0, 15))
        entry.focus_set()
        
    btn_frame = tk.Frame(dlg, bg=colors["bg"])
    btn_frame.pack(fill=tk.X)
    
    def on_ok(event=None):
        if is_prompt:
            result["value"] = entry.get()
        else:
            result["value"] = True
        dlg.destroy()
        
    def on_cancel(event=None):
        result["value"] = None if is_prompt else False
        dlg.destroy()
        
    dlg.bind("<Return>", on_ok)
    dlg.bind("<Escape>", on_cancel)
    dlg.protocol("WM_DELETE_WINDOW", on_cancel)
    
    ok_text = "OK" if not is_cancelable else "Confirmar"
    tk.Button(btn_frame, text=ok_text, font=("Segoe UI", 10, "bold"), bg=colors["btn_primary_bg"], fg=colors["btn_primary_fg"], activebackground=colors["btn_primary_hover"], activeforeground=colors["btn_primary_fg"], relief=tk.FLAT, cursor="hand2", command=on_ok, width=10).pack(side=tk.RIGHT, padx=(10, 0))
    
    if is_cancelable:
        tk.Button(btn_frame, text="Cancelar", font=("Segoe UI", 10), bg=colors["btn_sec_bg"], fg=colors["btn_sec_fg"], activebackground=colors["btn_sec_hover"], activeforeground=colors["btn_sec_fg"], relief=tk.FLAT, cursor="hand2", command=on_cancel, width=10).pack(side=tk.RIGHT)
        
    _center_window(dlg, parent)
    dlg.wait_window()
    return result["value"]

def ask_string(title, message, parent=None, show=""):
    return _create_dialog(title, message, parent, is_prompt=True, is_cancelable=True, show_char=show)

def ask_ok_cancel(title, message, parent=None):
    return _create_dialog(title, message, parent, is_prompt=False, is_cancelable=True)

_privacy_prefs_cache: dict | None = None

def _get_privacy_prefs_path() -> Path | None:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        return None
    key_dir = Path(user_profile) / 'AppData' / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
    key_dir.mkdir(parents=True, exist_ok=True)
    return key_dir / 'privacy_prefs.json'

def _load_privacy_prefs() -> dict:
    global _privacy_prefs_cache
    if _privacy_prefs_cache is not None:
        return _privacy_prefs_cache
    path = _get_privacy_prefs_path()
    if path and path.exists():
        try:
            with open(path, 'r', encoding='utf-8') as f:
                _privacy_prefs_cache = json.load(f)
                return _privacy_prefs_cache
        except Exception:
            pass
    _privacy_prefs_cache = {}
    return _privacy_prefs_cache

def _save_privacy_pref(key: str) -> None:
    global _privacy_prefs_cache
    path = _get_privacy_prefs_path()
    if not path:
        return
    prefs = _load_privacy_prefs()
    prefs[key] = True
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(prefs, f)
        _privacy_prefs_cache = prefs
    except Exception:
        pass

def load_privacy_prefs() -> dict:
    return _load_privacy_prefs()

def save_privacy_prefs(prefs: dict) -> None:
    global _privacy_prefs_cache
    path = _get_privacy_prefs_path()
    if not path:
        return
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(prefs, f)
        _privacy_prefs_cache = prefs
    except Exception:
        pass

def ask_privacy_warning(title: str, message: str, key: str, parent=None) -> bool:
    """Exibe aviso de privacidade com opção 'Não mostrar novamente'.
    Retorna True se o usuário confirmou (ou já optou por não ver mais), False se cancelou.
    'key' identifica o componente (ex.: 'chat_ia', 'correct_grammar').
    """
    prefs = _load_privacy_prefs()
    if prefs.get(key):
        return True

    colors = get_theme_colors()

    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.attributes('-topmost', True)
    dlg.configure(bg=colors["bg"], padx=20, pady=20)
    dlg.resizable(False, False)
    if parent and parent.winfo_viewable():
        dlg.transient(parent)
    dlg.grab_set()
    dlg.after(100, lambda: (dlg.lift(), dlg.focus_force()))

    tk.Label(dlg, text=message, font=("Segoe UI", 10), bg=colors["bg"], fg=colors["fg"],
             justify=tk.LEFT, wraplength=400).pack(pady=(0, 15))

    dont_show_var = tk.BooleanVar(value=False)
    tk.Checkbutton(
        dlg, text="Não mostrar este aviso novamente",
        variable=dont_show_var,
        font=("Segoe UI", 9),
        bg=colors["bg"], fg=colors["fg_muted"],
        activebackground=colors["bg"], activeforeground=colors["fg"],
        selectcolor=colors["text_bg"],
        relief=tk.FLAT, cursor="hand2"
    ).pack(anchor="w", pady=(0, 15))

    result = {"value": False}

    def on_ok(event=None):
        result["value"] = True
        if dont_show_var.get():
            _save_privacy_pref(key)
        dlg.destroy()

    def on_cancel(event=None):
        result["value"] = False
        dlg.destroy()

    dlg.bind("<Return>", on_ok)
    dlg.bind("<Escape>", on_cancel)
    dlg.protocol("WM_DELETE_WINDOW", on_cancel)

    btn_frame = tk.Frame(dlg, bg=colors["bg"])
    btn_frame.pack(fill=tk.X)

    tk.Button(btn_frame, text="Confirmar", font=("Segoe UI", 10, "bold"),
              bg=colors["btn_primary_bg"], fg=colors["btn_primary_fg"],
              activebackground=colors["btn_primary_hover"], activeforeground=colors["btn_primary_fg"],
              relief=tk.FLAT, cursor="hand2", command=on_ok, width=10).pack(side=tk.RIGHT, padx=(10, 0))
    tk.Button(btn_frame, text="Cancelar", font=("Segoe UI", 10),
              bg=colors["btn_sec_bg"], fg=colors["btn_sec_fg"],
              activebackground=colors["btn_sec_hover"], activeforeground=colors["btn_sec_fg"],
              relief=tk.FLAT, cursor="hand2", command=on_cancel, width=10).pack(side=tk.RIGHT)

    _center_window(dlg, parent)
    dlg.wait_window()
    return result["value"]

def show_info(title, message, parent=None):
    _create_dialog(title, message, parent, is_prompt=False, is_cancelable=False)

def show_warning(title, message, parent=None):
    _create_dialog(title, message, parent, is_prompt=False, is_cancelable=False)

def show_error(title, message, parent=None):
    _create_dialog(title, message, parent, is_prompt=False, is_cancelable=False)
