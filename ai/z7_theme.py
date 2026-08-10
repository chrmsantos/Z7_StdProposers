import json
import tkinter as tk
from functools import lru_cache
from pathlib import Path

__all__ = [
    "load_theme",
    "save_theme",
    "get_theme_colors",
    "ask_string",
    "ask_ok_cancel",
    "show_info",
    "show_warning",
    "show_error",
]


def _get_data_dir() -> Path:
    """Diretório de dados do usuário (reutiliza z7_logging para evitar duplicação)."""
    from z7_logging import get_data_dir
    return get_data_dir()


def get_theme_file_path() -> Path:
    return _get_data_dir() / "theme_config.json"


@lru_cache(maxsize=None)
def load_theme() -> str:
    theme_file = get_theme_file_path()
    if theme_file.exists():
        try:
            with open(theme_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get('theme', 'light')
        except Exception:
            pass
    return 'light'

def save_theme(theme_mode: str) -> None:
    theme_file = get_theme_file_path()
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
            "bg": "#0f111a",
            "fg": "#cdd6f4",
            "fg_muted": "#6c7086",
            "text_bg": "#1a1d2e",
            "border": "#2e3150",
            "btn_sec_bg": "#22253a",
            "btn_sec_fg": "#cdd6f4",
            "btn_sec_hover": "#2e3150",
            "btn_primary_bg": "#8b5cf6",
            "btn_primary_fg": "#ffffff",
            "btn_primary_hover": "#7c3aed",
            "user_tag": "#8b5cf6",
            "ai_tag": "#57c77c",
            "user_bubble_bg": "#22253a",
            "ai_bubble_bg": "#064e3b",
        }
    else:
        return {
            "bg": "#f0f2f8",
            "fg": "#1e2030",
            "fg_muted": "#6b7280",
            "text_bg": "#ffffff",
            "border": "#c8cedf",
            "btn_sec_bg": "#e8ecf6",
            "btn_sec_fg": "#6d28d9",
            "btn_sec_hover": "#ddd6fe",
            "btn_primary_bg": "#7c3aed",
            "btn_primary_fg": "#ffffff",
            "btn_primary_hover": "#6d28d9",
            "user_tag": "#7c3aed",
            "ai_tag": "#16a34a",
            "user_bubble_bg": "#e8ecf6",
            "ai_bubble_bg": "#ecfdf5",
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


def show_info(title, message, parent=None):
    _create_dialog(title, f"\u2139\ufe0f  {message}", parent, is_prompt=False, is_cancelable=False)

def show_warning(title, message, parent=None):
    _create_dialog(title, f"\u26a0\ufe0f  {message}", parent, is_prompt=False, is_cancelable=False)

def show_error(title, message, parent=None):
    _create_dialog(title, f"\u274c  {message}", parent, is_prompt=False, is_cancelable=False)
