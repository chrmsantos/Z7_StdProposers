import os
import sys
import shutil
import json
import time
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
    def log_exception(logger, msg, exc, *, reraise=False):
        print(f"EXCEPTION: {msg} - {exc}")
        if reraise:
            raise exc
else:
    LOGGER = configure_component_logger("installer")

_APP_VERSION = "8.6.0"
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
                    "bg": "#0f111a", "fg": "#cdd6f4", "fg_muted": "#6c7086",
                    "text_bg": "#1a1d2e", "border": "#2e3150",
                    "btn_sec_bg": "#22253a", "btn_sec_fg": "#cdd6f4", "btn_sec_hover": "#2e3150",
                    "btn_primary_bg": "#8b5cf6", "btn_primary_fg": "#ffffff", "btn_primary_hover": "#7c3aed"
                }
            else:
                colors = {
                    "bg": "#f0f2f8", "fg": "#1e2030", "fg_muted": "#6b7280",
                    "text_bg": "#ffffff", "border": "#c8cedf",
                    "btn_sec_bg": "#e8ecf6", "btn_sec_fg": "#6d28d9", "btn_sec_hover": "#ddd6fe",
                    "btn_primary_bg": "#7c3aed", "btn_primary_fg": "#ffffff", "btn_primary_hover": "#6d28d9"
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
    root.title("Instalador do Z7 StdProposers")
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
    card_frame = tk.Frame(root, relief=tk.SOLID, bd=1, highlightthickness=1)
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
        p_bg = colors_curr["btn_primary_bg"] if colors_curr else "#7c3aed"
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
        
        LOGGER.info("Iniciando processo de instalacao/atualizacao. Pasta destino: %s, Pasta temp: %s", install_dir, temp_dir)
        
        try:
            LOGGER.info("Fase 1: Verificando versao estavel no GitHub...")
            update_progress(0.05, "Verificando versão estável mais recente no GitHub...")
            
            data = get_latest_github_release()
            latest_version = data.get("tag_name", "").strip().lower().lstrip('v')
            if not latest_version:
                latest_version = data.get("name", "").strip().lower().lstrip('v')
            
            if not latest_version:
                LOGGER.error("Falha ao resolver tag de versao remota. Dados da API: %s", data)
                raise Exception("Não foi possível resolver a tag da versão mais recente no GitHub.")
            
            LOGGER.info("Versao remota identificada com sucesso: v%s", latest_version)
            update_progress(0.1, f"Versão estável encontrada: v{latest_version}. Preparando pastas...")
            
            LOGGER.info("Preparando diretorio temporario: %s", temp_dir)
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            # 1. Download de Ativos de Release
            assets = data.get("assets", [])
            LOGGER.info("Encontrados %d ativos na release do GitHub", len(assets))
            
            total_items = len(assets) + 3
            current_item = 0
            
            headers = {'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}
            
            for asset in assets:
                name = asset.get("name")
                url = asset.get("browser_download_url")
                if not name or not url:
                    LOGGER.warning("Ativo ignorado devido a nome ou URL ausente: %s", asset)
                    continue
                current_item += 1
                LOGGER.info("Baixando ativo %d/%d: %s de %s", current_item, len(assets), name, url)
                update_progress(0.1 + (current_item / total_items) * 0.4, f"Baixando {name}...")
                
                dest_file = temp_dir / name
                req = urllib.request.Request(url, headers=headers)
                t0 = time.monotonic()
                with urllib.request.urlopen(req, timeout=25) as resp:
                    content = resp.read()
                with open(dest_file, "wb") as f:
                    f.write(content)
                elapsed = time.monotonic() - t0
                LOGGER.info("Download concluido para %s (%d bytes) em %.2fs", name, len(content), elapsed)
            
            # 2. Download Fallback de arquivos ausentes do repositório Raw
            raw_base = "https://raw.githubusercontent.com/chrmsantos/Z7_StdProposers/main"
            fallback_files = [
                ("dist/Normal.dotm", "Normal.dotm"),
                ("dist/Word.officeUI", "Word.officeUI"),
                ("scripts/import_word.exe", "import_word.exe"),
                ("scripts/installer.exe", "installer.exe")
            ]
            
            LOGGER.info("Verificando downloads de fallback necessarios")
            for repo_path, local_name in fallback_files:
                dest_file = temp_dir / local_name
                current_item += 1
                
                if not dest_file.exists():
                    update_progress(0.1 + (current_item / total_items) * 0.4, f"Verificando {local_name}...")
                    url = f"{raw_base}/{repo_path}"
                    LOGGER.info("Ativo de fallback %s nao encontrado localmente. Tentando baixar de: %s", local_name, url)
                    try:
                        req = urllib.request.Request(url, headers=headers)
                        t0 = time.monotonic()
                        with urllib.request.urlopen(req, timeout=15) as resp:
                            content = resp.read()
                        with open(dest_file, "wb") as f:
                            f.write(content)
                        elapsed = time.monotonic() - t0
                        LOGGER.info("Download de fallback concluido para %s (%d bytes) em %.2fs", local_name, len(content), elapsed)
                    except Exception as download_err:
                        LOGGER.warning("Falha ao baixar %s via URL Raw: %s. Tentando usar copia local do projeto ou instalacao existente...", local_name, download_err)
                        # Fallback: se o arquivo já existir localmente no diretório de compilação ou na instalacao existente, copia dele
                        local_source = None
                        if local_name == "import_word.exe":
                            local_source = project_root / "scripts" / "import_word.exe"
                        elif local_name == "installer.exe":
                            local_source = project_root / "scripts" / "installer.exe"
                        elif local_name == "Normal.dotm":
                            local_source = project_root / "dist" / "Normal.dotm"
                            # Se nao existir no projeto, tenta a instalacao existente
                            if not local_source or not local_source.exists():
                                existing_normal = install_dir / "dist" / "Normal.dotm"
                                if existing_normal.exists():
                                    local_source = existing_normal
                        elif local_name == "Word.officeUI":
                            local_source = project_root / "dist" / "Word.officeUI"
                            # Se nao existir no projeto, tenta a instalacao existente
                            if not local_source or not local_source.exists():
                                existing_officeui = install_dir / "dist" / "Word.officeUI"
                                if existing_officeui.exists():
                                    local_source = existing_officeui
                        
                        if local_source and local_source.exists():
                            shutil.copy2(local_source, dest_file)
                            LOGGER.info("Copia local do fallback %s realizada com sucesso a partir de %s", local_name, local_source)
                        else:
                            LOGGER.error("Falha critica: arquivo de fallback %s nao pode ser obtido", local_name)
                else:
                    LOGGER.info("Arquivo de fallback %s ja existe na pasta temporaria, download dispensado", local_name)
            
            LOGGER.info("Fase 2: Detectando e fechando o Microsoft Word...")
            update_progress(0.55, "Detectando e fechando o Microsoft Word...")
            
            # Detecta documentos abertos para reabertura automática
            docs_to_reopen = []
            try:
                import win32com.client
                LOGGER.info("Buscando instancia ativa do Word via COM...")
                word_app = win32com.client.GetActiveObject("Word.Application")
                LOGGER.info("Instancia do Word encontrada. Mapeando documentos abertos...")
                for d in word_app.Documents:
                    try:
                        if d.FullName and Path(d.FullName).exists():
                            docs_to_reopen.append(d.FullName)
                            LOGGER.info("Documento aberto mapeado para reabertura: %s", d.FullName)
                    except Exception as doc_err:
                        LOGGER.warning("Falha ao obter informacoes do documento Word: %s", doc_err)
                
                LOGGER.info("Solicitando encerramento amigavel do Word...")
                word_app.Quit()
                time.sleep(3)
            except Exception as com_err:
                LOGGER.info("Nenhuma instancia ativa do Word COM detectada ou falha ao fechar amigavelmente: %s", com_err)
            
            # Encerramento garantido via taskkill se winword ainda ativo
            LOGGER.info("Executando taskkill winword.exe para garantir finalizacao...")
            res_kill = subprocess.run("taskkill /f /im winword.exe", shell=True, capture_output=True, text=True)
            LOGGER.info("Resultado do taskkill: returncode=%d, stdout=%s, stderr=%s", res_kill.returncode, res_kill.stdout.strip(), res_kill.stderr.strip())
            
            LOGGER.info("Fase 3: Criando backup da instalacao atual (se houver)...")
            update_progress(0.65, "Criando backup da instalação atual (se houver)...")
            
            # Se já existir uma instalação prévia, cria backup de segurança
            if install_dir.exists():
                LOGGER.info("Pasta de instalacao previa detectada em %s. Criando backup em %s", install_dir, backup_dir)
                backup_dir.mkdir(parents=True, exist_ok=True)
                
                backed_up_items = []
                if (install_dir / "ai" / "chat_ia").exists():
                    shutil.copytree(install_dir / "ai" / "chat_ia", backup_dir / "chat_ia", dirs_exist_ok=True)
                    backed_up_items.append("ai/chat_ia")
                if (install_dir / "ai" / "config_prompt").exists():
                    shutil.copytree(install_dir / "ai" / "config_prompt", backup_dir / "config_prompt", dirs_exist_ok=True)
                    backed_up_items.append("ai/config_prompt")
                if (install_dir / "scripts" / "import_word.exe").exists():
                    shutil.copy2(install_dir / "scripts" / "import_word.exe", backup_dir / "import_word.exe")
                    backed_up_items.append("scripts/import_word.exe")
                if (install_dir / "dist" / "Normal.dotm").exists():
                    shutil.copy2(install_dir / "dist" / "Normal.dotm", backup_dir / "Normal.dotm")
                    backed_up_items.append("dist/Normal.dotm")
                if (install_dir / "dist" / "Word.officeUI").exists():
                    shutil.copy2(install_dir / "dist" / "Word.officeUI", backup_dir / "Word.officeUI")
                    backed_up_items.append("dist/Word.officeUI")
                LOGGER.info("Backup dos seguintes itens concluido com sucesso: %s", backed_up_items)
            else:
                LOGGER.info("Nenhuma instalacao previa encontrada em %s. Backup dispensado", install_dir)
            
            LOGGER.info("Fase 4: Extraindo e aplicando novos binarios...")
            update_progress(0.75, "Extraindo e aplicando novos binários...")
            
            # Cria a estrutura final de pastas
            LOGGER.info("Garantindo estrutura de pastas finais em %s", install_dir)
            install_dir.mkdir(parents=True, exist_ok=True)
            (install_dir / "ai").mkdir(exist_ok=True)
            (install_dir / "scripts").mkdir(exist_ok=True)
            (install_dir / "dist").mkdir(exist_ok=True)
            
            # Extrai os pacotes zip baixados
            import zipfile
            zips = list(temp_dir.glob("*.zip"))
            LOGGER.info("Encontrados %d arquivos ZIP para extrair", len(zips))
            for zip_file in zips:
                extract_temp = temp_dir / f"extract_{zip_file.stem}"
                LOGGER.info("Extraindo %s para %s", zip_file.name, extract_temp)
                extract_temp.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    zip_ref.extractall(extract_temp)
                
                # Trata pacote unificado vs pacotes individuais
                if (extract_temp / "ai").exists():
                    LOGGER.info("Pacote unificado detectado em %s", zip_file.name)
                    shutil.copytree(extract_temp / "ai", install_dir / "ai", dirs_exist_ok=True)
                    if (extract_temp / "scripts").exists():
                        shutil.copytree(extract_temp / "scripts", install_dir / "scripts", dirs_exist_ok=True)
                    if (extract_temp / "dist").exists():
                        shutil.copytree(extract_temp / "dist", install_dir / "dist", dirs_exist_ok=True)
                elif "chat_ia" in zip_file.name.lower():
                    LOGGER.info("Pacote individual chat_ia detectado")
                    shutil.copytree(extract_temp, install_dir / "ai" / "chat_ia", dirs_exist_ok=True)
                elif "config_prompt" in zip_file.name.lower():
                    LOGGER.info("Pacote individual config_prompt detectado")
                    shutil.copytree(extract_temp, install_dir / "ai" / "config_prompt", dirs_exist_ok=True)
            
            # Copia recursos avulsos
            copied_assets = []
            if (temp_dir / "import_word.exe").exists():
                shutil.copy2(temp_dir / "import_word.exe", install_dir / "scripts" / "import_word.exe")
                copied_assets.append("import_word.exe")
            if (temp_dir / "Normal.dotm").exists():
                shutil.copy2(temp_dir / "Normal.dotm", install_dir / "dist" / "Normal.dotm")
                copied_assets.append("Normal.dotm")
            if (temp_dir / "Word.officeUI").exists():
                shutil.copy2(temp_dir / "Word.officeUI", install_dir / "dist" / "Word.officeUI")
                copied_assets.append("Word.officeUI")
            LOGGER.info("Arquivos avulsos copiados para a instalacao final: %s", copied_assets)
                
            LOGGER.info("Fase 5: Executando import_word.exe para configurar os templates no Word...")
            update_progress(0.9, "Executando import_word.exe para configurar os templates no Word...")
            
            import_exe = install_dir / "scripts" / "import_word.exe"
            if not import_exe.exists():
                LOGGER.error("Arquivo import_word.exe nao encontrado no destino final: %s", import_exe)
                raise Exception("Arquivo 'import_word.exe' não foi encontrado na instalação.")
            
            # Roda import_word em segundo plano silenciada e aguarda finalizar
            LOGGER.info("Disparando subprocesso: %s", import_exe)
            res = subprocess.run([str(import_exe)], cwd=str(install_dir / "scripts"), capture_output=True, text=True)
            LOGGER.info("Resultado de import_word.exe: returncode=%d, stdout=%s, stderr=%s", res.returncode, res.stdout.strip(), res.stderr.strip())
            if res.returncode != 0:
                raise Exception(f"Falha ao executar import_word.exe: {res.stderr or res.stdout}")
            
            LOGGER.info("Instalacao e configuracao concluidas com sucesso!")
            update_progress(1.0, "Instalação concluída com sucesso!")
            
            # Limpa vestígios temporários
            LOGGER.info("Limpando diretorio temporario de instalacao: %s", temp_dir)
            shutil.rmtree(temp_dir, ignore_errors=True)
            
            def success_action():
                root.destroy()
                try:
                    import z7_theme
                    z7_theme.show_info("Sucesso", f"O Z7 StdProposers v{latest_version} foi instalado com sucesso!", parent=None)
                except Exception:
                    tk.messagebox.showinfo("Sucesso", f"O Z7 StdProposers v{latest_version} foi instalado com sucesso!")
                
                # Reabre os documentos originais no Word (ou abre ele limpo)
                LOGGER.info("Abrindo Microsoft Word. Documentos a reabrir: %s", docs_to_reopen)
                if docs_to_reopen:
                    for doc_path in docs_to_reopen:
                        try:
                            os.startfile(doc_path)
                        except Exception as start_err:
                            LOGGER.warning("Falha ao abrir documento via startfile: %s. Tentando via Popen...", start_err)
                            subprocess.Popen(["cmd.exe", "/c", "start", "", doc_path])
                else:
                    try:
                        os.startfile("winword.exe")
                    except Exception as start_err:
                        LOGGER.warning("Falha ao abrir Word via startfile: %s. Tentando via Popen...", start_err)
                        subprocess.Popen(["cmd.exe", "/c", "start", "winword.exe"])
                sys.exit(0)
                
            root.after(500, success_action)
            
        except Exception as err:
            err_msg = str(err)
            LOGGER.error("Erro critico durante o processo de instalacao: %s", err, exc_info=True)
            
            # Rollback total de segurança
            LOGGER.info("Iniciando ROLLBACK de seguranca para restaurar arquivos originais...")
            update_progress(1.0, "Ocorreu um erro. Revertendo arquivos originais...")
            
            if backup_dir.exists():
                try:
                    if (backup_dir / "chat_ia").exists():
                        shutil.rmtree(install_dir / "ai" / "chat_ia", ignore_errors=True)
                        shutil.copytree(backup_dir / "chat_ia", install_dir / "ai" / "chat_ia", dirs_exist_ok=True)
                        LOGGER.info("Rollback: ai/chat_ia restaurado")
                    if (backup_dir / "config_prompt").exists():
                        shutil.rmtree(install_dir / "ai" / "config_prompt", ignore_errors=True)
                        shutil.copytree(backup_dir / "config_prompt", install_dir / "ai" / "config_prompt", dirs_exist_ok=True)
                        LOGGER.info("Rollback: ai/config_prompt restaurado")
                    if (backup_dir / "import_word.exe").exists():
                        shutil.copy2(backup_dir / "import_word.exe", install_dir / "scripts" / "import_word.exe")
                        LOGGER.info("Rollback: scripts/import_word.exe restaurado")
                    if (backup_dir / "Normal.dotm").exists():
                        shutil.copy2(backup_dir / "Normal.dotm", install_dir / "dist" / "Normal.dotm")
                        LOGGER.info("Rollback: dist/Normal.dotm restaurado")
                    if (backup_dir / "Word.officeUI").exists():
                        shutil.copy2(backup_dir / "Word.officeUI", install_dir / "dist" / "Word.officeUI")
                        LOGGER.info("Rollback: dist/Word.officeUI restaurado")
                    LOGGER.info("Rollback de seguranca executado com sucesso.")
                except Exception as rollback_err:
                    LOGGER.error("FALHA CRITICA NO ROLLBACK: %s", rollback_err, exc_info=True)
            else:
                LOGGER.info("Nenhum backup disponivel para restaurar, rollback nao realizado")
            
            # Limpa temporários
            LOGGER.info("Limpando pasta temporaria apos falha: %s", temp_dir)
            shutil.rmtree(temp_dir, ignore_errors=True)
            
            def error_action():
                root.destroy()
                try:
                    import z7_theme
                    z7_theme.show_error("Falha na Instalação", f"A instalação falhou. As configurações anteriores foram restauradas.\n\nErro: {err_msg}", parent=None)
                except Exception:
                    tk.messagebox.showerror("Falha na Instalação", f"A instalação falhou. As configurações anteriores foram restauradas.\n\nErro: {err_msg}")
                
                # Reabre os documentos no Word se foram fechados
                LOGGER.info("Reabrindo Microsoft Word apos falha. Documentos a reabrir: %s", docs_to_reopen)
                if docs_to_reopen:
                    for doc_path in docs_to_reopen:
                        try:
                            os.startfile(doc_path)
                        except Exception as start_err:
                            LOGGER.warning("Falha ao abrir documento via startfile: %s. Tentando via Popen...", start_err)
                            subprocess.Popen(["cmd.exe", "/c", "start", "", doc_path])
                else:
                    try:
                        os.startfile("winword.exe")
                    except Exception as start_err:
                        LOGGER.warning("Falha ao abrir Word via startfile: %s. Tentando via Popen...", start_err)
                        subprocess.Popen(["cmd.exe", "/c", "start", "winword.exe"])
                sys.exit(1)
                
            root.after(500, error_action)

    root.mainloop()

if __name__ == "__main__":
    main()
