import os
import json
import tkinter as tk
from pathlib import Path

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
    
    if parent:
        dlg.transient(parent)
    dlg.grab_set()
    
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
    _create_dialog(title, message, parent, is_prompt=False, is_cancelable=False)

def show_warning(title, message, parent=None):
    _create_dialog(title, message, parent, is_prompt=False, is_cancelable=False)

def show_error(title, message, parent=None):
    _create_dialog(title, message, parent, is_prompt=False, is_cancelable=False)
