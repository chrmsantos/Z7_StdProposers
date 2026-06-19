import json
import tkinter as tk
from pathlib import Path
import z7_theme
from z7_logging import configure_component_logger, get_data_dir, log_exception
import os
import sys
import shutil
import threading
import urllib.request
import urllib.error
import subprocess

LOGGER = configure_component_logger("config_prompt")

GITHUB_REPO_URL = "https://github.com/chrmsantos/Z7_StdProposers"

_APP_VERSION = "7.9.8"
_APP_AUTHOR  = "CMS"
_ORG         = "Câmara Municipal de Santa Bárbara d'Oeste"
_LICENSE     = "GPL-3.0"
_MOTTO       = "Dharma, virtude e gratidão."

DEFAULT_PROMPT = """Você é um especialista em análise jurídica e legislativa no idioma Português do Brasil.
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
   - Se for um Requerimento de Informações, o texto deverá fazer referência expressa ao "Art. 10, Inciso X, da Lei Orgânica do município de Santa Bárbara d’Oeste, combinado com o Art. 63, Inciso IX, do mesmo diploma legal";
   - Se for um Requerimento de Pesar, o texto deverá fazer referência expressa ao "Art. 102, Inciso IV, do Regimento Interno";
   - Se for uma Moção, o texto deverá fazer referência expressa ao "Art. 92, do Capítulo IV, Título V, do Regimento Interno".

Não reporte como problemas:
- Pequenas divergências de grafia ou acentuação;
- Diferenças na ordem de palavras que não alterem o sentido;
- Pequenos erros formais ou desvios gramaticais leves que não comprometam a estrutura ou a lógica do texto (erros gramaticais graves, contudo, devem ser apontados);
- As strings "$ANO$" e "$DATAATUALEXTENSO$" (que devem ser ignoradas no processo de verificação de consistência de datas, não devendo ser comparadas com outras datas no restante do documento).

Se encontrar inconsistências, liste-as de forma clara, sucinta e objetiva, indicando os trechos conflitantes e explicando o problema.
Se NÃO encontrar inconsistências, responda APENAS com: "Sem inconsistências detectadas no documento."

Responda em Português do Brasil."""

DEFAULT_CONSISTENCY_PROMPT = """Regras de Classificação e Limites:
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
Plenário “Dr. Tancredo Neves”, $DATAATUALEXTENSO$.
AUTORIA
– Vereador –

Saída JSON:
{"titulo": "INDICAÇÃO Nº $NUMERO$/$ANO$","ementa": "Indica ao Poder Executivo Municipal a ampliação da rede de creches nos bairros com maior demanda por vagas.","vocativo": "Excelentíssimo Senhor Prefeito Municipal,","proposicao": "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, dirijo-me a Vossa Excelência para indicar que seja realizado um estudo técnico para ampliação da rede de creches públicas, com prioridade aos bairros com maior número de crianças em lista de espera, como o Jardim São Fernando e o Parque Zabani, neste Município.","titulo_da_justificativa": "Justificativa:","justificativa": "A falta de vagas em creches tem afetado diretamente as famílias, em especial mães que dependem do serviço para poder trabalhar. A ampliação do número de unidades ou convênios com instituições qualificadas atenderá à demanda crescente e garantirá o direito à educação infantil.","data": "Plenário “Dr. Tancredo Neves”, $DATAATUALEXTENSO$.","assinatura": "AUTORIA\\n– Vereador –"}"""


def get_prompt_file_path() -> Path:
    return get_data_dir() / "gemini_prompt.txt"


def get_consistency_prompt_file_path() -> Path:
    return get_data_dir() / "consistency_prompt.txt"


def load_api_key() -> str:
    from z7_gemini_key import read_stored_api_key
    return read_stored_api_key()

def save_api_key(api_key: str) -> None:
    from z7_gemini_key import write_api_key
    write_api_key(api_key)

def load_prompt() -> str:
    prompt_file = get_prompt_file_path()
    if prompt_file.exists():
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                LOGGER.info("Loaded custom prompt file")
                return f.read()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom prompt", e)
    return DEFAULT_PROMPT


def load_consistency_prompt() -> str:
    consistency_file = get_consistency_prompt_file_path()
    if consistency_file.exists():
        try:
            with open(consistency_file, 'r', encoding='utf-8') as f:
                LOGGER.info("Loaded custom consistency prompt file")
                return f.read()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom consistency prompt", e)
    return DEFAULT_CONSISTENCY_PROMPT

def get_model_file_path() -> Path:
    return get_data_dir() / "selected_model.txt"

def load_ai_model() -> str:
    model_file = get_model_file_path()
    if model_file.exists():
        try:
            with open(model_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
        except Exception as e:
            log_exception(LOGGER, "Failed to load custom model", e)
    return "gemini-3.5-flash"

def save_ai_model(model_name: str) -> None:
    model_file = get_model_file_path()
    try:
        with open(model_file, 'w', encoding='utf-8') as f:
            f.write(model_name)
        LOGGER.info("Model saved successfully: %s", model_name)
    except Exception as e:
        log_exception(LOGGER, "Failed to save model", e)

def save_prompt(grammar_text: str, consistency_text: str, root: tk.Tk, model_var: tk.StringVar,
               privacy_chat_var: tk.BooleanVar | None = None) -> None:
    if not grammar_text.strip():
        LOGGER.warning("Grammar prompt save blocked because text is empty")
        z7_theme.show_warning("Aviso", "O prompt do Corretor Gramatical não pode estar vazio.", parent=root)
        return
    if not consistency_text.strip():
        LOGGER.warning("Consistency prompt save blocked because text is empty")
        z7_theme.show_warning("Aviso", "O prompt do Verificador de Consistência não pode estar vazio.", parent=root)
        return

    prompt_file = get_prompt_file_path()
    try:
        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(grammar_text.strip())
        LOGGER.info("Grammar prompt saved successfully")
    except Exception as e:
        log_exception(LOGGER, "Failed to save grammar prompt", e)
        z7_theme.show_error("Erro", f"Erro ao salvar prompt gramatical:\n{e}", parent=root)
        return

    consistency_file = get_consistency_prompt_file_path()
    try:
        with open(consistency_file, 'w', encoding='utf-8') as f:
            f.write(consistency_text.strip())
        LOGGER.info("Consistency prompt saved successfully")
    except Exception as e:
        log_exception(LOGGER, "Failed to save consistency prompt", e)
        z7_theme.show_error("Erro", f"Erro ao salvar prompt de consistência:\n{e}", parent=root)
        return

    save_ai_model(model_var.get())

    # Save privacy prefs (True = aviso desativado; ausente/False = aviso ativo)
    if privacy_chat_var is not None:
        prefs = z7_theme.load_privacy_prefs()
        if privacy_chat_var.get():
            prefs.pop('chat_ia', None)       # reativar aviso
        else:
            prefs['chat_ia'] = True          # suprimir aviso
        z7_theme.save_privacy_prefs(prefs)
        LOGGER.info("Privacy prefs saved")

    z7_theme.show_info("Sucesso", "Configurações salvas com sucesso!", parent=root)
    root.destroy()

def compare_versions(v1: str, v2: str) -> int:
    try:
        p1 = [int(x) for x in v1.strip().lower().lstrip('v').split('.')]
        p2 = [int(x) for x in v2.strip().lower().lstrip('v').split('.')]
        for i in range(max(len(p1), len(p2))):
            n1 = p1[i] if i < len(p1) else 0
            n2 = p2[i] if i < len(p2) else 0
            if n1 > n2:
                return 1
            elif n1 < n2:
                return -1
        return 0
    except Exception:
        return 0

def get_latest_github_release() -> dict:
    url = "https://api.github.com/repos/chrmsantos/Z7_StdProposers/releases/latest"
    req = urllib.request.Request(
        url,
        headers={'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}
    )
    with urllib.request.urlopen(req, timeout=10) as response:
        return json.loads(response.read().decode('utf-8'))

def check_for_updates_ui(parent_root: tk.Tk) -> None:
    colors = z7_theme.get_theme_colors()
    
    checking_dialog = tk.Toplevel(parent_root)
    checking_dialog.title("Verificando Atualizações")
    checking_dialog.configure(bg=colors["bg"], padx=20, pady=20)
    checking_dialog.resizable(False, False)
    checking_dialog.transient(parent_root)
    checking_dialog.grab_set()
    
    z7_theme._center_window(checking_dialog, parent_root)
    
    lbl = tk.Label(checking_dialog, text="🔄 Verificando atualizações no GitHub...", font=("Segoe UI", 10), bg=colors["bg"], fg=colors["fg"])
    lbl.pack(pady=10)
    
    checking_dialog.update()

    def run_check():
        try:
            data = get_latest_github_release()
            tag_name = data.get("tag_name", "").strip()
            if not tag_name:
                tag_name = data.get("name", "").strip()
            
            checking_dialog.destroy()
            
            if not tag_name:
                parent_root.after(0, lambda: z7_theme.show_error("Erro de Atualização", "Não foi possível identificar a versão no GitHub.", parent=parent_root))
                return
            
            local_ver = _APP_VERSION
            comparison = compare_versions(tag_name, local_ver)
            
            if comparison <= 0:
                parent_root.after(0, lambda: z7_theme.show_info("Sistema Atualizado", f"O Z7 StdProposers já está na versão mais recente (v{local_ver}).", parent=parent_root))
            else:
                parent_root.after(0, lambda: prompt_update_confirmation(parent_root, tag_name, data))
                
        except Exception as e:
            checking_dialog.destroy()
            parent_root.after(0, lambda err=str(e): z7_theme.show_error("Erro de Conectividade", f"Falha ao buscar atualizações:\n{err}", parent=parent_root))
            
    threading.Thread(target=run_check, daemon=True).start()

def prompt_update_confirmation(parent_root: tk.Tk, latest_version: str, release_data: dict) -> None:
    msg = (
        f"Uma nova versão ({latest_version}) está disponível no GitHub!\n\n"
        "IMPORTANTE: Salve qualquer trabalho pendente antes de prosseguir, pois o Microsoft Word será fechado e reiniciado durante o processo.\n\n"
        "Deseja iniciar a atualização agora?"
    )
    if z7_theme.ask_ok_cancel("Atualização Disponível", msg, parent=parent_root):
        start_download_and_update(parent_root, latest_version, release_data)

def start_download_and_update(parent_root: tk.Tk, latest_version: str, release_data: dict) -> None:
    colors = z7_theme.get_theme_colors()
    
    dl_dialog = tk.Toplevel(parent_root)
    dl_dialog.title("Instalando Atualização")
    dl_dialog.configure(bg=colors["bg"], padx=20, pady=20)
    dl_dialog.resizable(False, False)
    dl_dialog.transient(parent_root)
    dl_dialog.grab_set()
    
    z7_theme._center_window(dl_dialog, parent_root)
    
    lbl_title = tk.Label(dl_dialog, text=f"Baixando Z7 StdProposers v{latest_version}", font=("Segoe UI", 11, "bold"), bg=colors["bg"], fg=colors["fg"])
    lbl_title.pack(anchor="w", pady=(0, 10))
    
    status_lbl = tk.Label(dl_dialog, text="Iniciando downloads...", font=("Segoe UI", 10), bg=colors["bg"], fg=colors["fg_muted"])
    status_lbl.pack(anchor="w", pady=(0, 15))
    
    canvas_width = 360
    progress_canvas = tk.Canvas(dl_dialog, width=canvas_width, height=12, bg=colors["text_bg"], highlightthickness=0)
    progress_canvas.pack(fill=tk.X, pady=(0, 10))
    progress_bar = progress_canvas.create_rectangle(0, 0, 0, 12, fill=colors["btn_primary_bg"], width=0)
    
    dl_dialog.update()

    def update_progress(ratio: float, message: str):
        def _gui_update():
            if dl_dialog.winfo_exists():
                status_lbl.config(text=message)
                progress_canvas.coords(progress_bar, 0, 0, int(canvas_width * ratio), 12)
                dl_dialog.update()
        parent_root.after(0, _gui_update)

    def run_downloads():
        try:
            temp_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "Temp" / "Z7_Update"
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            assets = release_data.get("assets", [])
            total_items = len(assets) + 3
            current_item = 0
            
            headers = {'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}
            
            for asset in assets:
                name = asset.get("name")
                url = asset.get("browser_download_url")
                if not name or not url:
                    continue
                current_item += 1
                update_progress(current_item / total_items, f"Baixando {name}...")
                
                dest_file = temp_dir / name
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=20) as resp, open(dest_file, "wb") as f:
                    f.write(resp.read())
            
            raw_base = "https://raw.githubusercontent.com/chrmsantos/Z7_StdProposers/main"
            fallback_files = [
                ("dist/Normal.dotm", "Normal.dotm"),
                ("dist/Word.officeUI", "Word.officeUI"),
                ("scripts/import_word.exe", "import_word.exe")
            ]
            
            if getattr(sys, 'frozen', False):
                current_exe_dir = Path(sys.executable).parent
                if current_exe_dir.name.lower() == 'config_prompt':
                    install_dir = current_exe_dir.parent.parent
                else:
                    install_dir = current_exe_dir.parent
            else:
                install_dir = Path(__file__).resolve().parent.parent

            for repo_path, local_name in fallback_files:
                dest_file = temp_dir / local_name
                current_item += 1
                
                if not dest_file.exists():
                    update_progress(current_item / total_items, f"Verificando {local_name}...")
                    url = f"{raw_base}/{repo_path}"
                    
                    local_source = None
                    if local_name == "import_word.exe":
                        local_source = install_dir / "scripts" / "import_word.exe"
                    elif local_name == "Normal.dotm":
                        local_source = install_dir / "dist" / "Normal.dotm"
                    elif local_name == "Word.officeUI":
                        local_source = install_dir / "dist" / "Word.officeUI"

                    try:
                        req = urllib.request.Request(url, headers=headers)
                        with urllib.request.urlopen(req, timeout=15) as resp, open(dest_file, "wb") as f:
                            f.write(resp.read())
                    except Exception as download_err:
                        LOGGER.warning(f"Failed to download raw {local_name}: {download_err}")
                        if local_source and local_source.exists():
                            shutil.copy2(local_source, dest_file)
                            LOGGER.info(f"Copia local do fallback {local_name} realizada com sucesso a partir de {local_source}")
                        else:
                            LOGGER.error(f"Falha critica: arquivo de fallback {local_name} nao pode ser obtido")
                        
            # Detecta documentos atualmente abertos para reabri-los após a atualização
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
            except Exception:
                pass

            if docs_to_reopen:
                ps_docs_array = ", ".join([f'"{doc.replace("\\", "\\\\")}"' for doc in docs_to_reopen])
                ps_docs_def = f"$DocsToReopen = @({ps_docs_array})"
            else:
                ps_docs_def = "$DocsToReopen = @()"

            worker_path = temp_dir / "update_worker.ps1"
            parent_pid = os.getpid()
            
            ps_script = f"""$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ParentPid   = {parent_pid}
$InstallDir  = "{install_dir}"
$SourceDir   = "{temp_dir}"
$BackupDir   = Join-Path $SourceDir "backup"
{ps_docs_def}

try {{
    Write-Output "Aguardando encerramento do Z7..."
    Wait-Process -Id $ParentPid -Timeout 10 -ErrorAction SilentlyContinue

    Write-Output "Fechando o Microsoft Word..."
    try {{
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        if ($word) {{
            $word.Quit()
            Start-Sleep -Seconds 3
        }}
    }} catch {{}}
    
    $wordProcesses = Get-Process -Name "winword" -ErrorAction SilentlyContinue
    if ($wordProcesses) {{
        Stop-Process -Name "winword" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }}

    Write-Output "Criando copia de seguranca..."
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
    if (Test-Path "$InstallDir\\ai\\chat_ia") {{ Copy-Item -Path "$InstallDir\\ai\\chat_ia" -Destination "$BackupDir\\chat_ia" -Recurse -Force }}
    if (Test-Path "$InstallDir\\ai\\config_prompt") {{ Copy-Item -Path "$InstallDir\\ai\\config_prompt" -Destination "$BackupDir\\config_prompt" -Recurse -Force }}
    if (Test-Path "$InstallDir\\scripts\\import_word.exe") {{ Copy-Item -Path "$InstallDir\\scripts\\import_word.exe" -Destination "$BackupDir\\import_word.exe" -Force }}
    if (Test-Path "$InstallDir\\dist\\Normal.dotm") {{ Copy-Item -Path "$InstallDir\\dist\\Normal.dotm" -Destination "$BackupDir\\Normal.dotm" -Force }}
    if (Test-Path "$InstallDir\\dist\\Word.officeUI") {{ Copy-Item -Path "$InstallDir\\dist\\Word.officeUI" -Destination "$BackupDir\\Word.officeUI" -Force }}

    Write-Output "Instalando novos arquivos..."
    $zips = Get-ChildItem -Path $SourceDir -Filter "*.zip"
    foreach ($zip in $zips) {{
        $extractTemp = Join-Path $SourceDir "extract_$($zip.BaseName)"
        New-Item -ItemType Directory -Force -Path $extractTemp | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zip.FullName, $extractTemp)
        
        if (Test-Path "$extractTemp\\ai") {{
            if (Test-Path "$InstallDir\\ai") {{ Remove-Item -Path "$InstallDir\\ai" -Recurse -Force -ErrorAction SilentlyContinue }}
            Copy-Item -Path "$extractTemp\\ai" -Destination "$InstallDir" -Recurse -Force
            if (Test-Path "$extractTemp\\scripts") {{ Copy-Item -Path "$extractTemp\\scripts\\*" -Destination "$InstallDir\\scripts" -Recurse -Force }}
            if (Test-Path "$extractTemp\\dist") {{ Copy-Item -Path "$extractTemp\\dist\\*" -Destination "$InstallDir\\dist" -Recurse -Force }}
        }} elseif ($zip.Name -like "*chat_ia*") {{
            $destChat = "$InstallDir\\ai\\chat_ia"
            if (Test-Path $destChat) {{ Remove-Item -Path $destChat -Recurse -Force -ErrorAction SilentlyContinue }}
            Copy-Item -Path $extractTemp -Destination $destChat -Recurse -Force
        }} elseif ($zip.Name -like "*config_prompt*") {{
            $destConfig = "$InstallDir\\ai\\config_prompt"
            if (Test-Path $destConfig) {{ Remove-Item -Path $destConfig -Recurse -Force -ErrorAction SilentlyContinue }}
            Copy-Item -Path $extractTemp -Destination $destConfig -Recurse -Force
        }}
    }}

    if (Test-Path "$SourceDir\\import_word.exe") {{ Copy-Item -Path "$SourceDir\\import_word.exe" -Destination "$InstallDir\\scripts\\import_word.exe" -Force }}
    if (Test-Path "$SourceDir\\Normal.dotm") {{ Copy-Item -Path "$SourceDir\\Normal.dotm" -Destination "$InstallDir\\dist\\Normal.dotm" -Force }}
    if (Test-Path "$SourceDir\\Word.officeUI") {{ Copy-Item -Path "$SourceDir\\Word.officeUI" -Destination "$InstallDir\\dist\\Word.officeUI" -Force }}

    Write-Output "Aplicando os modelos do Word..."
    $importExe = Join-Path $InstallDir "scripts\\import_word.exe"
    if (Test-Path $importExe) {{
        $importProcess = Start-Process -FilePath $importExe -WorkingDirectory "$InstallDir\\scripts" -NoNewWindow -Wait -PassThru
        if ($importProcess.ExitCode -ne 0) {{
            throw "Falha ao executar import_word.exe (codigo de saida: $($importProcess.ExitCode))."
        }}
    }} else {{
        throw "import_word.exe nao encontrado para importar os templates."
    }}

    [System.Windows.Forms.MessageBox]::Show("Atualizacao concluida com sucesso para a versao v{latest_version}!", "Z7 StdProposers", 0, 64)
    if ($DocsToReopen.Count -gt 0) {{
        foreach ($docPath in $DocsToReopen) {{
            if (Test-Path $docPath) {{
                Start-Process -FilePath "winword.exe" -ArgumentList "`"$docPath`""
            }}
        }}
    }} else {{
        Start-Process -FilePath "winword.exe"
    }}
    
    Remove-Item -Path $SourceDir -Recurse -Force -ErrorAction SilentlyContinue
    
}} catch {{
    Write-Output "Erro detectado. Executando rollback de seguranca..."
    if (Test-Path "$BackupDir\\chat_ia") {{
        Remove-Item -Path "$InstallDir\\ai\\chat_ia" -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -Path "$BackupDir\\chat_ia" -Destination "$InstallDir\\ai\\chat_ia" -Recurse -Force
    }}
    if (Test-Path "$BackupDir\\config_prompt") {{
        Remove-Item -Path "$InstallDir\\ai\\config_prompt" -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -Path "$BackupDir\\config_prompt" -Destination "$InstallDir\\ai\\config_prompt" -Recurse -Force
    }}
    if (Test-Path "$BackupDir\\import_word.exe") {{ Copy-Item -Path "$BackupDir\\import_word.exe" -Destination "$InstallDir\\scripts\\import_word.exe" -Force }}
    if (Test-Path "$BackupDir\\Normal.dotm") {{ Copy-Item -Path "$BackupDir\\Normal.dotm" -Destination "$InstallDir\\dist\\Normal.dotm" -Force }}
    if (Test-Path "$BackupDir\\Word.officeUI") {{ Copy-Item -Path "$BackupDir\\Word.officeUI" -Destination "$InstallDir\\dist\\Word.officeUI" -Force }}

    [System.Windows.Forms.MessageBox]::Show("Falha na atualizacao.`n`nOs arquivos originais foram totalmente restaurados.`n`nErro: $_", "Erro de Atualizacao", 0, 16)
    if ($DocsToReopen.Count -gt 0) {{
        foreach ($docPath in $DocsToReopen) {{
            if (Test-Path $docPath) {{
                Start-Process -FilePath "winword.exe" -ArgumentList "`"$docPath`""
            }}
        }}
    }} else {{
        Start-Process -FilePath "winword.exe"
    }}
}}
"""
            with open(worker_path, "w", encoding="utf-8") as f:
                f.write(ps_script)
            
            update_progress(1.0, "Iniciando instalador em segundo plano...")
            
            subprocess.Popen([
                "powershell.exe",
                "-ExecutionPolicy", "Bypass",
                "-WindowStyle", "Hidden",
                "-File", str(worker_path)
            ])
            
            parent_root.after(1000, lambda: (dl_dialog.destroy(), parent_root.destroy(), sys.exit(0)))
            
        except Exception as e:
            err_msg = str(e)
            def show_err():
                dl_dialog.destroy()
                z7_theme.show_error("Erro de Instalação", f"Ocorreu um erro ao preparar os arquivos:\n{err_msg}", parent=parent_root)
            parent_root.after(0, show_err)

    threading.Thread(target=run_downloads, daemon=True).start()

def restore_default(text_widget: tk.Text) -> None:
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, DEFAULT_PROMPT)


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
            
        self.root.configure(bg=bg)
        
        if 'title_lbl' in self.widgets:
            self.widgets['title_lbl'].configure(bg=bg, fg=fg)
        if 'info_lbl' in self.widgets:
            self.widgets['info_lbl'].configure(bg=bg, fg=fg_muted)
        if 'btn_frame' in self.widgets:
            self.widgets['btn_frame'].configure(bg=bg)
        if 'border_frame' in self.widgets:
            self.widgets['border_frame'].configure(bg=border)
        if 'text_area' in self.widgets:
            self.widgets['text_area'].configure(bg=text_bg, fg=fg, insertbackground=fg)
        if 'model_frame' in self.widgets:
            self.widgets['model_frame'].configure(bg=bg)
        if 'model_lbl' in self.widgets:
            self.widgets['model_lbl'].configure(bg=bg, fg=fg)
        if 'model_dropdown' in self.widgets:
            self.widgets['model_dropdown'].configure(bg=text_bg, fg=fg, insertbackground=fg)

        if 'api_frame' in self.widgets:
            self.widgets['api_frame'].configure(bg=bg)
        if 'api_btn' in self.widgets:
            self.widgets['api_btn'].configure(bg=btn_sec_bg, fg=btn_sec_fg,
                                              activebackground=btn_sec_hover, activeforeground=fg)

        if 'privacy_frame' in self.widgets:
            self.widgets['privacy_frame'].configure(bg=bg)
        if 'privacy_lbl' in self.widgets:
            self.widgets['privacy_lbl'].configure(bg=bg, fg=fg)
        for cb in self.widgets.get('privacy_checks', []):
            cb.configure(bg=bg, fg=fg_muted, activebackground=bg, activeforeground=fg,
                         selectcolor=text_bg)

        if 'tab_frame' in self.widgets:
            self.widgets['tab_frame'].configure(bg=bg)
        active_idx = self.widgets.get('active_tab_idx', 0)
        for i, btn in enumerate(self.widgets.get('tab_btns', [])):
            if i == active_idx:
                btn.configure(
                    bg=colors["btn_primary_bg"], fg=colors["btn_primary_fg"],
                    activebackground=colors["btn_primary_hover"],
                    activeforeground=colors["btn_primary_fg"],
                )
            else:
                btn.configure(
                    bg=btn_sec_bg, fg=btn_sec_fg,
                    activebackground=btn_sec_hover, activeforeground=fg,
                )

        for btn in self.widgets.get('sec_btns', []):
            btn.configure(bg=btn_sec_bg, fg=btn_sec_fg, activebackground=btn_sec_hover, activeforeground=fg)
        
        if 'toggle_btn' in self.widgets:
            icon = "🌙 Modo Escuro" if self.mode == 'light' else "☀️ Modo Claro"
            self.widgets['toggle_btn'].configure(text=icon, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)

        if 'footer_lbl' in self.widgets:
            self.widgets['footer_lbl'].configure(bg=bg, fg=fg_muted)


def open_api_key_dialog(parent: tk.Tk, theme_mode: str) -> None:
    colors = z7_theme.get_theme_colors(theme_mode)
    bg          = colors["bg"]
    fg          = colors["fg"]
    text_bg     = colors["text_bg"]
    btn_sec_bg  = colors["btn_sec_bg"]
    btn_sec_fg  = colors["btn_sec_fg"]
    btn_sec_hover = colors["btn_sec_hover"]
    btn_primary = colors.get("btn_primary_bg", "#2563eb")

    dialog = tk.Toplevel(parent)
    dialog.title("Chave de API Gemini")
    dialog.geometry("460x155")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes('-topmost', True)
    dialog.configure(bg=bg)

    tk.Label(dialog, text="Chave de API do Gemini:", font=("Segoe UI", 10, "bold"),
             bg=bg, fg=fg).pack(anchor="w", padx=25, pady=(20, 5))

    entry_frame = tk.Frame(dialog, bg=bg)
    entry_frame.pack(fill=tk.X, padx=25)

    entry_var = tk.StringVar(value=load_api_key())
    entry = tk.Entry(entry_frame, textvariable=entry_var, font=("Segoe UI", 11),
                     relief=tk.FLAT, bd=2, show="*", bg=text_bg, fg=fg,
                     insertbackground=fg)
    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def toggle_visibility() -> None:
        if entry.cget("show") == "*":
            entry.config(show="")
            show_btn.config(text="Ocultar")
        else:
            entry.config(show="*")
            show_btn.config(text="Mostrar")

    show_btn = tk.Button(entry_frame, text="Mostrar", font=("Segoe UI", 9),
                         relief=tk.FLAT, cursor="hand2",
                         bg=btn_sec_bg, fg=btn_sec_fg,
                         activebackground=btn_sec_hover, activeforeground=fg,
                         command=toggle_visibility)
    show_btn.pack(side=tk.LEFT, padx=(8, 0))

    def do_save() -> None:
        save_api_key(entry_var.get().strip())
        LOGGER.info("API key saved from dialog")
        dialog.destroy()

    btn_row = tk.Frame(dialog, bg=bg)
    btn_row.pack(fill=tk.X, padx=25, pady=(15, 0))

    tk.Button(btn_row, text="Cancelar", font=("Segoe UI", 10, "bold"),
              relief=tk.FLAT, cursor="hand2",
              bg=btn_sec_bg, fg=btn_sec_fg,
              activebackground=btn_sec_hover, activeforeground=fg,
              command=dialog.destroy).pack(side=tk.RIGHT, padx=(10, 0))

    tk.Button(btn_row, text="Salvar", font=("Segoe UI", 10, "bold"),
              relief=tk.FLAT, cursor="hand2",
              bg=btn_primary, fg="white",
              activebackground="#1d4ed8", activeforeground="white",
              command=do_save).pack(side=tk.RIGHT)


def main() -> None:
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
    root.title(f"Configurar Prompts — Z7 StdProposers v{_APP_VERSION}")
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    chat_width = int(screen_width * 2 / 3 * 1.15 * 1.10)
    chat_height = int(screen_height * 0.92 * 0.90 * 1.05)
    chat_left = 0
    chat_top = int(screen_height * 0.02)
    
    root.geometry(f"{chat_width}x{chat_height}+{chat_left}+{chat_top}")
    root.minsize(300, 500)
    
    theme = AppTheme(root)
    
    # Faz a janela aparecer na frente
    root.attributes('-topmost', True)
    
    # Botão de tema no canto superior
    toggle_btn = tk.Button(root, font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2", command=theme.toggle, bd=0)
    toggle_btn.place(relx=0.03, rely=0.03)
    theme.widgets['toggle_btn'] = toggle_btn
    
    lbl = tk.Label(root, text="Instruções para a Inteligência Artificial", font=("Segoe UI", 16, "bold"))
    lbl.pack(pady=(35, 5))  # Aumentado o pady superior para não sobrepor o botão de tema
    theme.widgets['title_lbl'] = lbl
    
    info_lbl = tk.Label(root, text="Personalize o comportamento do modelo ajustando o prompt abaixo.", font=("Segoe UI", 10))
    info_lbl.pack(pady=(0, 20))
    theme.widgets['info_lbl'] = info_lbl

    btn_frame = tk.Frame(root)
    theme.widgets['btn_frame'] = btn_frame

    # Seleção de Modelo
    model_frame = tk.Frame(root)
    model_frame.pack(fill=tk.X, padx=25, pady=(0, 10))
    theme.widgets['model_frame'] = model_frame

    model_lbl = tk.Label(model_frame, text="Modelo de IA:", font=("Segoe UI", 10, "bold"))
    model_lbl.pack(side=tk.LEFT)
    theme.widgets['model_lbl'] = model_lbl

    model_var = tk.StringVar(root)
    current_model = load_ai_model()
    model_var.set(current_model)

    model_entry = tk.Entry(model_frame, textvariable=model_var, font=("Segoe UI", 10), relief=tk.FLAT, bd=2)
    model_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
    theme.widgets['model_dropdown'] = model_entry

    # Chave de API
    api_frame = tk.Frame(root)
    api_frame.pack(fill=tk.X, padx=25, pady=(0, 10))
    theme.widgets['api_frame'] = api_frame

    api_btn = tk.Button(api_frame, text="🔑 Chave de API...", font=("Segoe UI", 10),
                        relief=tk.FLAT, cursor="hand2",
                        command=lambda: open_api_key_dialog(root, theme.mode))
    api_btn.pack(side=tk.LEFT)
    theme.widgets['api_btn'] = api_btn

    # Avisos de Privacidade
    privacy_prefs = z7_theme.load_privacy_prefs()

    privacy_frame = tk.Frame(root)
    privacy_frame.pack(fill=tk.X, padx=25, pady=(0, 10))
    theme.widgets['privacy_frame'] = privacy_frame

    privacy_lbl = tk.Label(privacy_frame, text="Avisos de Privacidade:", font=("Segoe UI", 10, "bold"))
    privacy_lbl.pack(side=tk.LEFT)
    theme.widgets['privacy_lbl'] = privacy_lbl

    # Checked = exibir aviso (pref ausente); Unchecked = não exibir (pref True)
    privacy_chat_var = tk.BooleanVar(value=not privacy_prefs.get('chat_ia', False))

    cb_chat = tk.Checkbutton(privacy_frame, text="Chat IA",
                             variable=privacy_chat_var, font=("Segoe UI", 10),
                             relief=tk.FLAT, cursor="hand2")
    cb_chat.pack(side=tk.LEFT, padx=(15, 0))

    theme.widgets['privacy_checks'] = [cb_chat]

    # Seletor de abas de prompt
    prompt_buffers = {
        'grammar': load_prompt(),
        'consistency': load_consistency_prompt(),
    }
    current_tab = tk.StringVar(value='grammar')

    tab_frame = tk.Frame(root)
    tab_frame.pack(fill=tk.X, padx=25, pady=(0, 4))
    theme.widgets['tab_frame'] = tab_frame
    theme.widgets['active_tab_idx'] = 0

    def switch_tab(tab_name: str, tab_idx: int) -> None:
        if current_tab.get() == tab_name:
            return
        prompt_buffers[current_tab.get()] = text_area.get("1.0", tk.END)
        current_tab.set(tab_name)
        theme.widgets['active_tab_idx'] = tab_idx
        text_area.delete("1.0", tk.END)
        text_area.insert(tk.END, prompt_buffers[tab_name].strip())
        theme.apply()

    grammar_tab_btn = tk.Button(
        tab_frame, text="Corretor Gramatical",
        font=("Segoe UI", 10, "bold"), relief=tk.FLAT, cursor="hand2",
        command=lambda: switch_tab('grammar', 0), padx=12,
    )
    grammar_tab_btn.pack(side=tk.LEFT)

    consistency_tab_btn = tk.Button(
        tab_frame, text="Verificador de Consistência",
        font=("Segoe UI", 10, "bold"), relief=tk.FLAT, cursor="hand2",
        command=lambda: switch_tab('consistency', 1), padx=12,
    )
    consistency_tab_btn.pack(side=tk.LEFT, padx=(4, 0))

    theme.widgets['tab_btns'] = [grammar_tab_btn, consistency_tab_btn]

    frame = tk.Frame(root) # Borda sutil
    theme.widgets['border_frame'] = frame
    
    text_area = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 11), relief=tk.FLAT, padx=12, pady=12)
    text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=1, pady=1)
    theme.widgets['text_area'] = text_area

    scrollbar = tk.Scrollbar(frame, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

    # Carrega o prompt da aba ativa (gramática por padrão)
    text_area.insert(tk.END, prompt_buffers['grammar'])

    # Estilos de botão
    btn_font = ("Segoe UI", 10, "bold")

    def do_save() -> None:
        prompt_buffers[current_tab.get()] = text_area.get("1.0", tk.END)
        save_prompt(
            grammar_text=prompt_buffers['grammar'],
            consistency_text=prompt_buffers['consistency'],
            root=root,
            model_var=model_var,
            privacy_chat_var=privacy_chat_var,
        )

    def do_restore_default() -> None:
        if current_tab.get() == 'grammar':
            restore_default(text_area)
        else:
            restore_default_consistency(text_area)

    save_btn = tk.Button(btn_frame, text="Salvar Configuração", width=20, bg="#2563eb", fg="white", font=btn_font, relief=tk.FLAT, activebackground="#1d4ed8", activeforeground="white", cursor="hand2", command=do_save)
    save_btn.pack(side=tk.RIGHT, padx=25)
    
    cancel_btn = tk.Button(btn_frame, text="Cancelar", width=15, font=btn_font, relief=tk.FLAT, cursor="hand2", command=root.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=5)

    restore_btn = tk.Button(btn_frame, text="Restaurar Padrão", width=18, font=btn_font, relief=tk.FLAT, cursor="hand2", command=do_restore_default)
    restore_btn.pack(side=tk.LEFT, padx=25)

    import webbrowser
    github_btn = tk.Button(btn_frame, text="⧉ GitHub", font=("Segoe UI", 9), relief=tk.FLAT, cursor="hand2",
                           command=lambda: webbrowser.open(GITHUB_REPO_URL))
    github_btn.pack(side=tk.LEFT, padx=(0, 5))
    theme.widgets.setdefault('sec_btns', []).append(github_btn)

    update_btn = tk.Button(btn_frame, text="🔄 Atualizar...", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, cursor="hand2",
                           command=lambda: check_for_updates_ui(root))
    update_btn.pack(side=tk.LEFT, padx=(5, 5))
    theme.widgets.setdefault('sec_btns', []).append(update_btn)

    theme.widgets['sec_btns'] = [cancel_btn, restore_btn, github_btn, update_btn]
    
    # Rodapé
    _footer_text = f"{_ORG}  ·  {_APP_AUTHOR}  ·  {_LICENSE}  ·  {_MOTTO}"
    footer_lbl = tk.Label(root, text=_footer_text, font=("Segoe UI", 8), anchor="center")
    theme.widgets['footer_lbl'] = footer_lbl

    # Aplica o tema inicial
    theme.apply()

    # Pack in order so btn_frame is fixed at the bottom
    footer_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 5))
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 5))
    frame.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=25, pady=(0, 10))

    root.mainloop()

if __name__ == "__main__":
    main()
