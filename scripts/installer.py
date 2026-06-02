import os
import sys
import shutil
import json
import threading
import urllib.request
import urllib.error
import subprocess
import tkinter as tk
from pathlib import Path

# Adiciona o diretório 'ai' ao sys.path para que os imports locais funcionem no desenvolvimento e compilação
current_dir = Path(__file__).resolve().parent
project_root = current_dir.parent
ai_dir = project_root / 'ai'
if ai_dir.exists():
    sys.path.insert(0, str(ai_dir))

try:
    import z7_theme
    from z7_logging import configure_component_logger, log_exception
except ImportError:
    # Fallback se executado fora do ambiente estruturado antes da compilação
    class MockLogger:
        def info(self, msg, *args): print(f"INFO: {msg % args if args else msg}")
        def warning(self, msg, *args): print(f"WARNING: {msg % args if args else msg}")
        def error(self, msg, *args): print(f"ERROR: {msg % args if args else msg}")
    LOGGER = MockLogger()
    log_exception = lambda l, m, e: print(f"EXCEPTION: {m} - {e}")
else:
    LOGGER = configure_component_logger("installer")

_APP_VERSION = "7.8.7"
GITHUB_REPO_URL = "https://github.com/chrmsantos/Z7_StdProposers"

class InstallerTheme:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        try:
            self.mode = z7_theme.load_theme()
        except Exception:
            self.mode = 'light'
        self.widgets = {}

    def toggle(self) -> None:
        self.mode = 'dark' if self.mode == 'light' else 'light'
        try:
            z7_theme.save_theme(self.mode)
        except Exception:
            pass
        self.apply()

    def apply(self) -> None:
        try:
            colors = z7_theme.get_theme_colors(self.mode)
        except Exception:
            # Fallback de cores se z7_theme falhar
            if self.mode == 'dark':
                colors = {
                    "bg": "#0f172a", "fg": "#e2e8f0", "fg_muted": "#94a3b8",
                    "text_bg": "#1e293b", "border": "#334155",
                    "btn_sec_bg": "#1e293b", "btn_sec_fg": "#cbd5e1", "btn_sec_hover": "#334155",
                    "btn_primary_bg": "#6366f1", "btn_primary_fg": "#ffffff", "btn_primary_hover": "#4f46e5"
                }
            else:
                colors = {
                    "bg": "#f5f3ff", "fg": "#1e1b4b", "fg_muted": "#6b7280",
                    "text_bg": "#ffffff", "border": "#c4b5fd",
                    "btn_sec_bg": "#ede9fe", "btn_sec_fg": "#4c1d95", "btn_sec_hover": "#ddd6fe",
                    "btn_primary_bg": "#6366f1", "btn_primary_fg": "#ffffff", "btn_primary_hover": "#4f46e5"
                }
            
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
            
        self.root.configure(bg=bg)
        
        if 'title_lbl' in self.widgets:
            self.widgets['title_lbl'].configure(bg=bg, fg=fg)
        if 'info_lbl' in self.widgets:
            self.widgets['info_lbl'].configure(bg=bg, fg=fg_muted)
        if 'card_frame' in self.widgets:
            self.widgets['card_frame'].configure(bg=text_bg, highlightbackground=border)
        if 'desc_lbl' in self.widgets:
            self.widgets['desc_lbl'].configure(bg=text_bg, fg=fg)
        if 'status_lbl' in self.widgets:
            self.widgets['status_lbl'].configure(bg=text_bg, fg=fg_muted)
        if 'progress_canvas' in self.widgets:
            self.widgets['progress_canvas'].configure(bg=bg)
        if 'btn_frame' in self.widgets:
            self.widgets['btn_frame'].configure(bg=bg)
            
        if 'install_btn' in self.widgets:
            self.widgets['install_btn'].configure(bg=btn_primary_bg, fg=btn_primary_fg,
                                                 activebackground=btn_primary_hover, activeforeground=btn_primary_fg)
        if 'cancel_btn' in self.widgets:
            self.widgets['cancel_btn'].configure(bg=btn_sec_bg, fg=btn_sec_fg,
                                                activebackground=btn_sec_hover, activeforeground=fg)
        if 'toggle_btn' in self.widgets:
            icon = "🌙 Modo Escuro" if self.mode == 'light' else "☀️ Modo Claro"
            self.widgets['toggle_btn'].configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)


def get_latest_github_release() -> dict:
    url = "https://api.github.com/repos/chrmsantos/Z7_StdProposers/releases/latest"
    req = urllib.request.Request(
        url,
        headers={'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}
    )
    with urllib.request.urlopen(req, timeout=10) as response:
        return json.loads(response.read().decode('utf-8'))


def main() -> None:
    LOGGER.info("Starting standalone installer wizard")
    
    root = tk.Tk()
    root.title(f"Instalador do Z7 StdProposers")
    root.geometry("600x420")
    root.resizable(False, False)
    root.attributes('-topmost', True)
    
    theme = InstallerTheme(root)
    
    # Alternador de tema superior
    toggle_btn = tk.Button(root, font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2", command=theme.toggle, bd=0)
    toggle_btn.place(relx=0.04, rely=0.04)
    theme.widgets['toggle_btn'] = toggle_btn
    
    # Cabeçalho principal
    title_lbl = tk.Label(root, text="Instalador Z7 StdProposers", font=("Segoe UI", 18, "bold"))
    title_lbl.pack(pady=(45, 5))
    theme.widgets['title_lbl'] = title_lbl
    
    info_lbl = tk.Label(root, text="Assistente de Instalação do Sistema de Padronização Legislativa", font=("Segoe UI", 10))
    info_lbl.pack(pady=(0, 20))
    theme.widgets['info_lbl'] = info_lbl
    
    # Card Central
    colors = z7_theme.get_theme_colors(theme.mode) if 'z7_theme' in sys.modules else None
    card_frame = tk.Frame(root, relief=tk.SINGLE, bd=1, highlightthickness=1)
    card_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=(0, 25))
    theme.widgets['card_frame'] = card_frame
    
    desc_lbl = tk.Label(card_frame, text=(
        "Este utilitário fará o download da versão estável mais recente do sistema diretamente\n"
        "do GitHub e configurará automaticamente todos os executáveis, ferramentas auxiliares,\n"
        "e modelos de automação do Microsoft Word no seu computador de forma livre de privilégios admin.\n\n"
        "Clique em 'Instalar' para começar. O Microsoft Word será fechado durante a instalação."
    ), font=("Segoe UI", 10), justify=tk.CENTER, wraplength=480)
    desc_lbl.pack(pady=(25, 10), padx=20)
    theme.widgets['desc_lbl'] = desc_lbl
    
    status_lbl = tk.Label(card_frame, text="", font=("Segoe UI", 10, "italic"))
    theme.widgets['status_lbl'] = status_lbl
    
    canvas_width = 460
    progress_canvas = tk.Canvas(card_frame, width=canvas_width, height=10, highlightthickness=0)
    theme.widgets['progress_canvas'] = progress_canvas
    progress_bar = None
    
    # Rodapé / Botões
    btn_frame = tk.Frame(root)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 25))
    theme.widgets['btn_frame'] = btn_frame
    
    def on_cancel():
        root.destroy()
        sys.exit(0)
        
    cancel_btn = tk.Button(btn_frame, text="Cancelar", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, width=12, cursor="hand2", command=on_cancel)
    cancel_btn.pack(side=tk.RIGHT, padx=(0, 40))
    theme.widgets['cancel_btn'] = cancel_btn
    
    def start_install():
        desc_lbl.pack_forget()
        status_lbl.pack(pady=(35, 5))
        progress_canvas.pack(fill=tk.X, padx=25, pady=(5, 10))
        
        nonlocal progress_bar
        colors_curr = z7_theme.get_theme_colors(theme.mode) if 'z7_theme' in sys.modules else None
        p_bg = colors_curr["btn_primary_bg"] if colors_curr else "#6366f1"
        t_bg = colors_curr["text_bg"] if colors_curr else "#ffffff"
        progress_canvas.configure(bg=t_bg)
        progress_bar = progress_canvas.create_rectangle(0, 0, 0, 10, fill=p_bg, width=0)
        
        install_btn.config(state=tk.DISABLED)
        cancel_btn.config(state=tk.DISABLED)
        toggle_btn.config(state=tk.DISABLED)
        
        threading.Thread(target=run_installation, daemon=True).start()

    install_btn = tk.Button(btn_frame, text="Instalar", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, width=15, cursor="hand2", command=start_install)
    install_btn.pack(side=tk.RIGHT, padx=15)
    theme.widgets['install_btn'] = install_btn
    
    theme.apply()
    
    def update_progress(ratio: float, message: str):
        def _gui_update():
            if root.winfo_exists():
                status_lbl.config(text=message)
                progress_canvas.coords(progress_bar, 0, 0, int(canvas_width * ratio), 10)
                root.update()
        root.after(0, _gui_update)

    def run_installation():
        # Pasta de instalação definitiva em LOCALAPPDATA
        install_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "Z7" / "Apps" / "Z7_StdProposers"
        temp_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "Temp" / "Z7_Installation"
        
        # Pasta de backups
        backup_dir = temp_dir / "backup"
        
        try:
            update_progress(0.05, "Verificando versão estável mais recente no GitHub...")
            data = get_latest_github_release()
            latest_version = data.get("tag_name", "").strip().lower().lstrip('v')
            if not latest_version:
                latest_version = data.get("name", "").strip().lower().lstrip('v')
            
            if not latest_version:
                raise Exception("Não foi possível resolver a tag da versão mais recente no GitHub.")
            
            update_progress(0.1, f"Versão estável encontrada: v{latest_version}. Preparando pastas...")
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            # 1. Download de Ativos de Release
            assets = data.get("assets", [])
            total_items = len(assets) + 3
            current_item = 0
            
            headers = {'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}
            
            for asset in assets:
                name = asset.get("name")
                url = asset.get("browser_download_url")
                if not name or not url:
                    continue
                current_item += 1
                update_progress(0.1 + (current_item / total_items) * 0.4, f"Baixando {name}...")
                
                dest_file = temp_dir / name
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=25) as resp, open(dest_file, "wb") as f:
                    f.write(resp.read())
            
            # 2. Download Fallback de arquivos ausentes do repositório Raw
            raw_base = "https://raw.githubusercontent.com/chrmsantos/Z7_StdProposers/main"
            fallback_files = [
                ("dist/Normal.dotm", "Normal.dotm"),
                ("dist/Word.officeUI", "Word.officeUI"),
                ("scripts/import_word.exe", "import_word.exe")
            ]
            
            for repo_path, local_name in fallback_files:
                dest_file = temp_dir / local_name
                current_item += 1
                
                if not dest_file.exists():
                    update_progress(0.1 + (current_item / total_items) * 0.4, f"Verificando {local_name}...")
                    url = f"{raw_base}/{repo_path}"
                    try:
                        req = urllib.request.Request(url, headers=headers)
                        with urllib.request.urlopen(req, timeout=15) as resp, open(dest_file, "wb") as f:
                            f.write(resp.read())
                    except Exception as download_err:
                        LOGGER.warning(f"Failed to download raw {local_name}: {download_err}")
                        # Fallback: se o arquivo já existir localmente no diretório de compilação, copia dele
                        local_source = None
                        if local_name == "import_word.exe":
                            local_source = project_root / "scripts" / "import_word.exe"
                        elif local_name == "Normal.dotm":
                            local_source = project_root / "dist" / "Normal.dotm"
                        elif local_name == "Word.officeUI":
                            local_source = project_root / "dist" / "Word.officeUI"
                        
                        if local_source and local_source.exists():
                            shutil.copy2(local_source, dest_file)
                            LOGGER.info(f"Copied local {local_name} as fallback during installation")
            
            update_progress(0.55, "Detectando e fechando o Microsoft Word...")
            
            # Detecta documentos abertos para reabertura automática
            docs_to_reopen = []
            try:
                import win32com.client
                word_app = win32com.client.GetActiveObject("Word.Application")
                for d in word_app.Documents:
                    try:
                        if d.FullName and Path(d.FullName).exists():
                            docs_to_reopen.append(d.FullName)
                    except Exception:
                        pass
                
                # Fecha de forma amigável
                word_app.Quit()
                import time
                time.sleep(3)
            except Exception:
                pass
            
            # Encerramento garantido via taskkill se winword ainda ativo
            subprocess.run("taskkill /f /im winword.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            update_progress(0.65, "Criando backup da instalação atual (se houver)...")
            
            # Se já existir uma instalação prévia, cria backup de segurança
            if install_dir.exists():
                backup_dir.mkdir(parents=True, exist_ok=True)
                if (install_dir / "ai" / "chat_ia").exists():
                    shutil.copytree(install_dir / "ai" / "chat_ia", backup_dir / "chat_ia", dirs_exist_ok=True)
                if (install_dir / "ai" / "config_prompt").exists():
                    shutil.copytree(install_dir / "ai" / "config_prompt", backup_dir / "config_prompt", dirs_exist_ok=True)
                if (install_dir / "scripts" / "import_word.exe").exists():
                    shutil.copy2(install_dir / "scripts" / "import_word.exe", backup_dir / "import_word.exe")
                if (install_dir / "dist" / "Normal.dotm").exists():
                    shutil.copy2(install_dir / "dist" / "Normal.dotm", backup_dir / "Normal.dotm")
                if (install_dir / "dist" / "Word.officeUI").exists():
                    shutil.copy2(install_dir / "dist" / "Word.officeUI", backup_dir / "Word.officeUI")
            
            update_progress(0.75, "Extraindo e aplicando novos binários...")
            
            # Cria a estrutura final de pastas
            install_dir.mkdir(parents=True, exist_ok=True)
            (install_dir / "ai").mkdir(exist_ok=True)
            (install_dir / "scripts").mkdir(exist_ok=True)
            (install_dir / "dist").mkdir(exist_ok=True)
            
            # Extrai os pacotes zip baixados
            import zipfile
            zips = list(temp_dir.glob("*.zip"))
            for zip_file in zips:
                extract_temp = temp_dir / f"extract_{zip_file.stem}"
                extract_temp.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    zip_ref.extractall(extract_temp)
                
                # Trata pacote unificado vs pacotes individuais
                if (extract_temp / "ai").exists():
                    shutil.copytree(extract_temp / "ai", install_dir / "ai", dirs_exist_ok=True)
                    if (extract_temp / "scripts").exists():
                        shutil.copytree(extract_temp / "scripts", install_dir / "scripts", dirs_exist_ok=True)
                    if (extract_temp / "dist").exists():
                        shutil.copytree(extract_temp / "dist", install_dir / "dist", dirs_exist_ok=True)
                elif "chat_ia" in zip_file.name.lower():
                    shutil.copytree(extract_temp, install_dir / "ai" / "chat_ia", dirs_exist_ok=True)
                elif "config_prompt" in zip_file.name.lower():
                    shutil.copytree(extract_temp, install_dir / "ai" / "config_prompt", dirs_exist_ok=True)
            
            # Copia recursos avulsos
            if (temp_dir / "import_word.exe").exists():
                shutil.copy2(temp_dir / "import_word.exe", install_dir / "scripts" / "import_word.exe")
            if (temp_dir / "Normal.dotm").exists():
                shutil.copy2(temp_dir / "Normal.dotm", install_dir / "dist" / "Normal.dotm")
            if (temp_dir / "Word.officeUI").exists():
                shutil.copy2(temp_dir / "Word.officeUI", install_dir / "dist" / "Word.officeUI")
                
            update_progress(0.9, "Executando import_word.exe para configurar os templates no Word...")
            
            import_exe = install_dir / "scripts" / "import_word.exe"
            if not import_exe.exists():
                raise Exception("Arquivo 'import_word.exe' não foi encontrado na instalação.")
            
            # Roda import_word em segundo plano silenciada e aguarda finalizar
            res = subprocess.run([str(import_exe)], cwd=str(install_dir / "scripts"), capture_output=True, text=True)
            if res.returncode != 0:
                raise Exception(f"Falha ao executar import_word.exe: {res.stderr or res.stdout}")
            
            update_progress(1.0, "Instalação concluída com sucesso!")
            
            # Limpa vestígios temporários
            shutil.rmtree(temp_dir, ignore_errors=True)
            
            def success_action():
                root.destroy()
                try:
                    import z7_theme
                    z7_theme.show_info("Sucesso", f"O Z7 StdProposers v{latest_version} foi instalado com sucesso!", parent=None)
                except Exception:
                    tk.messagebox.showinfo("Sucesso", f"O Z7 StdProposers v{latest_version} foi instalado com sucesso!")
                
                # Reabre os documentos originais no Word (ou abre ele limpo)
                if docs_to_reopen:
                    for doc_path in docs_to_reopen:
                        subprocess.Popen(["cmd.exe", "/c", "start", "winword.exe", doc_path])
                else:
                    subprocess.Popen(["cmd.exe", "/c", "start", "winword.exe"])
                sys.exit(0)
                
            root.after(500, success_action)
            
        except Exception as err:
            LOGGER.error(f"Erro crítico durante instalação: {err}")
            
            # Rollback total de segurança
            update_progress(1.0, "Ocorreu um erro. Revertendo arquivos originais...")
            if backup_dir.exists():
                if (backup_dir / "chat_ia").exists():
                    shutil.rmtree(install_dir / "ai" / "chat_ia", ignore_errors=True)
                    shutil.copytree(backup_dir / "chat_ia", install_dir / "ai" / "chat_ia", dirs_exist_ok=True)
                if (backup_dir / "config_prompt").exists():
                    shutil.rmtree(install_dir / "ai" / "config_prompt", ignore_errors=True)
                    shutil.copytree(backup_dir / "config_prompt", install_dir / "ai" / "config_prompt", dirs_exist_ok=True)
                if (backup_dir / "import_word.exe").exists():
                    shutil.copy2(backup_dir / "import_word.exe", install_dir / "scripts" / "import_word.exe")
                if (backup_dir / "Normal.dotm").exists():
                    shutil.copy2(backup_dir / "Normal.dotm", install_dir / "dist" / "Normal.dotm")
                if (backup_dir / "Word.officeUI").exists():
                    shutil.copy2(backup_dir / "Word.officeUI", install_dir / "dist" / "Word.officeUI")
            
            # Limpa temporários
            shutil.rmtree(temp_dir, ignore_errors=True)
            
            def error_action():
                root.destroy()
                try:
                    import z7_theme
                    z7_theme.show_error("Falha na Instalação", f"A instalação falhou. As configurações anteriores foram restauradas.\n\nErro: {err}", parent=None)
                except Exception:
                    tk.messagebox.showerror("Falha na Instalação", f"A instalação falhou. As configurações anteriores foram restauradas.\n\nErro: {err}")
                
                # Reabre os documentos no Word se foram fechados
                if docs_to_reopen:
                    for doc_path in docs_to_reopen:
                        subprocess.Popen(["cmd.exe", "/c", "start", "winword.exe", doc_path])
                else:
                    subprocess.Popen(["cmd.exe", "/c", "start", "winword.exe"])
                sys.exit(1)
                
            root.after(500, error_action)

    root.mainloop()

if __name__ == "__main__":
    main()
