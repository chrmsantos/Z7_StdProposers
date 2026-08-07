import json
import tkinter as tk
from tkinter import ttk
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

_APP_VERSION = "8.4.1"
_APP_AUTHOR  = "CMS"
_ORG         = "Câmara Municipal de Santa Bárbara d'Oeste"
_LICENSE     = "GPL-3.0"
_MOTTO       = "Dharma, virtude e gratidão."

_DEFAULT_PREFIX = """As strings \"$ANO$\" e \"$DATAATUALEXTENSO$\" devem ser ignoradas no processo de verificação de consistência de datas, não devendo ser comparadas com outras datas no restante do documento.
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
- As strings "$ANO$" e "$DATAATUALEXTENSO$" (que devem ser ignoradas no processo de verificação de consistência de datas, não devendo ser comparadas com outras datas no restante do documento).

Se encontrar inconsistências, liste-as de forma clara, sucinta e objetiva, indicando os trechos conflitantes e explicando o problema. 
Elabore sugestões de correção para cada inconsistência identificada, mantendo a integridade e o sentido jurídico do documento.

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


DEFAULT_CHAT_SYSTEM_PROMPT = "Você é a LÉIA — Assistente Legislativa de IA. Sempre se apresente como LÉIA. Leia o documento ativo. Identifique erros. Auxilie o usuário alterando, revisando ou tirando dúvidas."


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
    return "deepseek/deepseek-chat"

def save_ai_model(model_name: str) -> None:
    model_file = get_model_file_path()
    try:
        model_file.write_text(model_name, encoding='utf-8')
        LOGGER.info("Model saved successfully: %s", model_name)
    except Exception as e:
        log_exception(LOGGER, "Failed to save model", e)

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

def read_update_status() -> str:
    """Lê o status de atualização escrito pelo chat_ia em arquivo compartilhado."""
    import json
    status_file = get_data_dir() / "update_status.json"
    if status_file.exists():
        try:
            data = json.loads(status_file.read_text(encoding="utf-8"))
            return data.get("status", "checking")
        except Exception:
            pass
    return "checking"


def get_latest_github_release() -> dict:
    url = "https://api.github.com/repos/chrmsantos/Z7_StdProposers/releases/latest"
    req = urllib.request.Request(
        url,
        headers={'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}
    )
    with urllib.request.urlopen(req, timeout=10) as response:
        return json.loads(response.read().decode('utf-8'))

def _download_with_retry(url: str, dest_file: Path, headers: dict, timeout: int = 30, retries: int = 3) -> None:
    """Baixa um arquivo com tentativas automáticas e verificação de tamanho mínimo."""
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = resp.read()
            if len(data) < 100:
                raise Exception(f"Arquivo muito pequeno ({len(data)} bytes), provavelmente erro no download.")
            dest_file.write_bytes(data)
            LOGGER.info("Downloaded %s (%d bytes, attempt %d)", dest_file.name, len(data), attempt)
            return
        except Exception as e:
            LOGGER.warning("Download attempt %d/%d failed for %s: %s", attempt, retries, url, e)
            if attempt == retries:
                raise
            import time
            time.sleep(2 * attempt)


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
            
            # Destruir diálogo e mostrar resultado na thread principal (thread-safe)
            if not tag_name:
                parent_root.after(0, lambda: _safe_destroy(checking_dialog))
                parent_root.after(0, lambda: z7_theme.show_error("Erro de Atualização", "Não foi possível identificar a versão no GitHub.", parent=parent_root))
                return
            
            local_ver = _APP_VERSION
            comparison = compare_versions(tag_name, local_ver)
            
            if comparison <= 0:
                parent_root.after(0, lambda: _safe_destroy(checking_dialog))
                parent_root.after(10, lambda: z7_theme.show_info("Sistema Atualizado", f"O Z7 StdProposers já está na versão mais recente (v{local_ver}).", parent=parent_root))
            else:
                parent_root.after(0, lambda: _safe_destroy(checking_dialog))
                parent_root.after(10, lambda: prompt_update_confirmation(parent_root, tag_name, data))
                
        except Exception as e:
            LOGGER.error("Update check failed: %s", e)
            parent_root.after(0, lambda: _safe_destroy(checking_dialog))
            parent_root.after(10, lambda err=str(e): z7_theme.show_error("Erro de Conectividade", f"Falha ao buscar atualizações:\n{err}", parent=parent_root))
            
    threading.Thread(target=run_check, daemon=True).start()


def _safe_destroy(window: tk.Toplevel) -> None:
    """Destrói uma janela Toplevel de forma segura, ignorando se já foi destruída."""
    try:
        if window.winfo_exists():
            window.destroy()
    except Exception:
        pass


def prompt_update_confirmation(parent_root: tk.Tk, latest_version: str, release_data: dict) -> None:
    msg = (
        f"Uma nova versão ({latest_version}) está disponível no GitHub!\n\n"
        "IMPORTANTE: Salve qualquer trabalho pendente antes de prosseguir, pois o Microsoft Word será fechado e reiniciado durante o processo.\n\n"
        "Deseja iniciar a atualização agora?"
    )
    if z7_theme.ask_ok_cancel("Atualização Disponível", msg, parent=parent_root):
        start_download_and_update(parent_root, latest_version, release_data)


def _resolve_install_dir() -> Path:
    """Determina o diretório de instalação do Z7 StdProposers de forma robusta."""
    if getattr(sys, 'frozen', False):
        current_exe_dir = Path(sys.executable).parent
        if current_exe_dir.name.lower() == 'config_prompt':
            return current_exe_dir.parent.parent
        return current_exe_dir.parent
    return Path(__file__).resolve().parent.parent


def _escape_ps_string(value: str) -> str:
    """Escapa uma string para uso seguro dentro de um script PowerShell entre aspas duplas."""
    return value.replace('`', '``').replace('"', '`"').replace('$', '`$')


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
            try:
                if dl_dialog.winfo_exists():
                    status_lbl.config(text=message)
                    progress_canvas.coords(progress_bar, 0, 0, int(canvas_width * ratio), 12)
                    dl_dialog.update()
            except Exception:
                pass
        parent_root.after(0, _gui_update)

    def run_downloads():
        try:
            temp_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "Temp" / "Z7_Update"
            # Limpa diretório temporário anterior para evitar conflitos
            if temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            install_dir = _resolve_install_dir()
            headers = {'User-Agent': f'Z7_StdProposers/{_APP_VERSION}'}

            # ── Fase 1: Baixar assets do release ───────────────────────────
            assets = release_data.get("assets", [])
            # Assets do release + 4 fallbacks (Normal.dotm, Word.officeUI, import_word.exe, installer.exe)
            total_items = len(assets) + 4
            current_item = 0

            for asset in assets:
                name = asset.get("name")
                url = asset.get("browser_download_url")
                if not name or not url:
                    continue
                current_item += 1
                update_progress(current_item / total_items, f"Baixando {name}...")
                
                dest_file = temp_dir / name
                _download_with_retry(url, dest_file, headers, timeout=30, retries=3)

            # ── Fase 2: Garantir arquivos complementares (fallback) ────────
            raw_base = "https://raw.githubusercontent.com/chrmsantos/Z7_StdProposers/main"
            fallback_files = [
                ("dist/Normal.dotm", "Normal.dotm"),
                ("dist/Word.officeUI", "Word.officeUI"),
                ("scripts/import_word.exe", "import_word.exe"),
                ("scripts/installer.exe", "installer.exe"),
            ]
            
            for repo_path, local_name in fallback_files:
                dest_file = temp_dir / local_name
                current_item += 1
                
                # Pula se já foi baixado como asset do release e tem tamanho razoável
                if dest_file.exists() and dest_file.stat().st_size > 100:
                    LOGGER.info("Fallback file %s already present from release assets (%d bytes)", local_name, dest_file.stat().st_size)
                    continue

                update_progress(current_item / total_items, f"Baixando {local_name}...")
                url = f"{raw_base}/{repo_path}"
                
                # Determina fonte local como fallback
                local_source = None
                if local_name == "import_word.exe":
                    local_source = install_dir / "scripts" / "import_word.exe"
                elif local_name == "installer.exe":
                    local_source = install_dir / "scripts" / "installer.exe"
                elif local_name == "Normal.dotm":
                    local_source = install_dir / "dist" / "Normal.dotm"
                elif local_name == "Word.officeUI":
                    local_source = install_dir / "dist" / "Word.officeUI"

                try:
                    _download_with_retry(url, dest_file, headers, timeout=30, retries=2)
                except Exception as download_err:
                    LOGGER.warning("Failed to download %s from GitHub: %s", local_name, download_err)
                    if local_source and local_source.exists():
                        shutil.copy2(local_source, dest_file)
                        LOGGER.info("Used local copy for %s from %s", local_name, local_source)
                    else:
                        LOGGER.error("CRITICAL: %s could not be obtained from any source", local_name)

            # ── Fase 3: Verificar que todos os arquivos essenciais existem ─
            # Quando o release nao possui assets (zips compilados), a atualizacao
            # prossegue apenas com os arquivos complementares (installer, templates).
            has_release_assets = len(assets) > 0
            
            # Arquivos obrigatorios independentes de haver assets no release
            core_required_files = [
                temp_dir / "installer.exe",
                temp_dir / "Normal.dotm",
                temp_dir / "Word.officeUI",
                temp_dir / "import_word.exe",
            ]
            
            # Zips compilados so sao obrigatorios se o release possuia assets
            zip_required_files = []
            if has_release_assets:
                zip_required_files = [
                    temp_dir / f"config_prompt-v{latest_version}.zip",
                    temp_dir / f"chat_ia-v{latest_version}.zip",
                ]
            
            # Aceita nomes de zip sem prefixo "v" tambem
            alt_zip_names = [
                temp_dir / f"config_prompt-v{latest_version.lstrip('v')}.zip",
                temp_dir / f"chat_ia-v{latest_version.lstrip('v')}.zip",
            ]

            missing = []
            skipped_zips = []
            
            # Verifica arquivos nucleares
            for f in core_required_files:
                if not f.exists() or f.stat().st_size < 100:
                    missing.append(f.name)
            
            # Verifica zips compilados (apenas se o release tinha assets)
            for f in zip_required_files:
                if not f.exists() or f.stat().st_size < 100:
                    alt_found = False
                    for alt in alt_zip_names:
                        if alt.name.startswith(f.stem.split("-v")[0]) and alt.exists() and alt.stat().st_size > 100:
                            alt_found = True
                            break
                    if not alt_found:
                        skipped_zips.append(f.name)

            if missing:
                raise Exception(
                    "Arquivos obrigatórios não puderam ser baixados:\n"
                    + "\n".join(f"  • {m}" for m in missing)
                )
            
            if skipped_zips:
                LOGGER.warning("Pacotes compilados (IA) nao disponiveis no release: %s. "
                               "Ferramentas de IA nao serao atualizadas nesta versao.", skipped_zips)

            # ── Fase 4: Copiar VERSION para o temp_dir ─────────────────────
            version_file = install_dir / "VERSION"
            if version_file.exists():
                shutil.copy2(version_file, temp_dir / "VERSION")

            # ── Fase 5: Detectar documentos abertos no Word ────────────────
            docs_to_reopen: list[str] = []
            try:
                import pythoncom
                pythoncom.CoInitialize()
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

            # ── Fase 6: Gerar script PowerShell worker ─────────────────────
            ps_docs_array = ", ".join(
                [f'"{_escape_ps_string(doc)}"' for doc in docs_to_reopen]
            )
            ps_docs_def = f"$DocsToReopen = @({ps_docs_array})" if docs_to_reopen else "$DocsToReopen = @()"

            worker_path = temp_dir / "update_worker.ps1"
            parent_pid = os.getpid()
            log_path = temp_dir / "update_worker.log"
            escaped_install_dir = _escape_ps_string(str(install_dir))
            escaped_source_dir = _escape_ps_string(str(temp_dir))
            escaped_log_path = _escape_ps_string(str(log_path))
            
            ps_script = f"""$ErrorActionPreference = "Stop"
Start-Transcript -Path "{escaped_log_path}" -Force
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ParentPid   = {parent_pid}
$InstallDir  = "{escaped_install_dir}"
$SourceDir   = "{escaped_source_dir}"
$BackupDir   = Join-Path $SourceDir "backup"
{ps_docs_def}

try {{
    Write-Output "Aguardando encerramento do Z7 (PID $ParentPid)..."
    $waited = 0
    while ((Get-Process -Id $ParentPid -ErrorAction SilentlyContinue) -and ($waited -lt 30)) {{
        Start-Sleep -Seconds 1
        $waited++
    }}
    if (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue) {{
        Write-Output "AVISO: processo pai ainda ativo apos 30s. Prosseguindo mesmo assim..."
    }}

    Write-Output "Fechando o Microsoft Word..."
    $retries = 0
    while ($retries -lt 3) {{
        try {{
            $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
            if ($word) {{
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
                Start-Sleep -Seconds 3
            }}
        }} catch {{}}
        
        $wordProcesses = Get-Process -Name "winword" -ErrorAction SilentlyContinue
        if (-not $wordProcesses) {{ break }}
        
        $retries++
        if ($retries -lt 3) {{
            Write-Output "Word ainda aberto. Tentativa $retries/3 de fechar..."
            Stop-Process -Name "winword" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }} else {{
            Write-Output "AVISO: Nao foi possivel fechar todas as instancias do Word."
            Stop-Process -Name "winword" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        }}
    }}

    Write-Output "Criando copia de seguranca..."
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
    $backupTargets = @{{
        "ai\\chat_ia"           = "$BackupDir\\chat_ia"
        "ai\\config_prompt"     = "$BackupDir\\config_prompt"
        "scripts\\import_word.exe" = "$BackupDir\\import_word.exe"
        "dist\\Normal.dotm"     = "$BackupDir\\Normal.dotm"
        "dist\\Word.officeUI"   = "$BackupDir\\Word.officeUI"
        "VERSION"               = "$BackupDir\\VERSION"
    }}
    foreach ($kv in $backupTargets.GetEnumerator()) {{
        $src = Join-Path $InstallDir $kv.Key
        if (Test-Path $src) {{
            $dstDir = Split-Path $kv.Value -Parent
            New-Item -ItemType Directory -Force -Path $dstDir | Out-Null
            Copy-Item -Path $src -Destination $kv.Value -Recurse -Force
        }}
    }}
    Write-Output "Backup concluido."

    Write-Output "Instalando novos arquivos..."
    $zips = Get-ChildItem -Path $SourceDir -Filter "*.zip"
    foreach ($zip in $zips) {{
        Write-Output "Extraindo $($zip.Name)..."
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
        Remove-Item -Path $extractTemp -Recurse -Force -ErrorAction SilentlyContinue
    }}

    if (Test-Path "$SourceDir\\import_word.exe") {{ Copy-Item -Path "$SourceDir\\import_word.exe" -Destination "$InstallDir\\scripts\\import_word.exe" -Force }}
    if (Test-Path "$SourceDir\\Normal.dotm") {{ Copy-Item -Path "$SourceDir\\Normal.dotm" -Destination "$InstallDir\\dist\\Normal.dotm" -Force }}
    if (Test-Path "$SourceDir\\Word.officeUI") {{ Copy-Item -Path "$SourceDir\\Word.officeUI" -Destination "$InstallDir\\dist\\Word.officeUI" -Force }}
    if (Test-Path "$SourceDir\\installer.exe") {{ Copy-Item -Path "$SourceDir\\installer.exe" -Destination "$InstallDir\\scripts\\installer.exe" -Force }}
    if (Test-Path "$SourceDir\\VERSION") {{
        $newVersion = (Get-Content "$SourceDir\\VERSION" -Raw).Trim()
        Write-Output "Atualizando VERSION para $newVersion"
        Set-Content -Path "$InstallDir\\VERSION" -Value $newVersion -NoNewline
    }}

    Write-Output "Aplicando os modelos do Word..."
    $importExe = Join-Path $InstallDir "scripts\\import_word.exe"
    if (Test-Path $importExe) {{
        $importProcess = Start-Process -FilePath $importExe -WorkingDirectory "$InstallDir\\scripts" -NoNewWindow -Wait -PassThru
        if ($importProcess.ExitCode -ne 0) {{
            throw "Falha ao executar import_word.exe (codigo de saida: $($importProcess.ExitCode))."
        }}
        Write-Output "Templates do Word aplicados com sucesso."
    }} else {{
        Write-Output "AVISO: import_word.exe nao encontrado. Templates nao atualizados."
    }}

    Write-Output "Atualizacao para v{latest_version} concluida!"
    Stop-Transcript

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
    $errMsg = $_.Exception.Message
    Write-Output "ERRO FATAL: $errMsg"
    Write-Output "Executando rollback de seguranca..."
    Stop-Transcript

    if (Test-Path $BackupDir) {{
        $restoreTargets = @{{
            "chat_ia"               = "$InstallDir\\ai\\chat_ia"
            "config_prompt"         = "$InstallDir\\ai\\config_prompt"
            "import_word.exe"       = "$InstallDir\\scripts\\import_word.exe"
            "Normal.dotm"           = "$InstallDir\\dist\\Normal.dotm"
            "Word.officeUI"         = "$InstallDir\\dist\\Word.officeUI"
            "VERSION"               = "$InstallDir\\VERSION"
        }}
        foreach ($kv in $restoreTargets.GetEnumerator()) {{
            $bakFile = Join-Path $BackupDir $kv.Key
            if (Test-Path $bakFile) {{
                $dstDir = Split-Path $kv.Value -Parent
                if ($dstDir -and -not (Test-Path $dstDir)) {{ New-Item -ItemType Directory -Force -Path $dstDir | Out-Null }}
                Remove-Item -Path $kv.Value -Recurse -Force -ErrorAction SilentlyContinue
                Copy-Item -Path $bakFile -Destination $kv.Value -Recurse -Force
            }}
        }}
        Write-Output "Rollback concluido."
    }}

    [System.Windows.Forms.MessageBox]::Show("Falha na atualizacao.`n`nOs arquivos originais foram totalmente restaurados.`n`nErro: $errMsg`n`nLog detalhado: {escaped_log_path}", "Erro de Atualizacao", 0, 16)
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
            worker_path.write_text(ps_script, encoding='utf-8')
            
            LOGGER.info("Update worker script written to %s", worker_path)
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
            LOGGER.error("Update preparation failed: %s", err_msg)
            def show_err():
                _safe_destroy(dl_dialog)
                z7_theme.show_error("Erro de Instalação", f"Ocorreu um erro ao preparar os arquivos:\n{err_msg}", parent=parent_root)
            parent_root.after(0, show_err)

    threading.Thread(target=run_downloads, daemon=True).start()

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
        select_bg = "#6366f1" if self.mode == "dark" else "#7c3aed"
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

        # ── Status bar ────────────────────────────────────────────────────
        if 'status_frame' in self.widgets:
            self.widgets['status_frame'].configure(bg=bg)
        if 'api_btn' in self.widgets:
            self.widgets['api_btn'].configure(bg=btn_sec_bg, fg=btn_sec_fg,
                                              activebackground=btn_sec_hover, activeforeground=fg)
        if 'update_status_lbl' in self.widgets:
            self.widgets['update_status_lbl'].configure(bg=bg)
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
    btn_primary_hover = colors.get("btn_primary_hover", "#4f46e5")
    select_bg = "#6366f1" if theme_mode == "dark" else "#7c3aed"
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
    def _validate_inputs() -> "tuple[str, str] | None":
        _clear_output()
        key = api_var.get().strip()
        model = model_var.get().strip()
        effective_key = key or load_api_key()

        if not effective_key:
            _append("⚠  Informe uma chave de API.", "warn")
            return None
        if not model:
            _append("⚠  Informe um nome de modelo.", "warn")
            return None
        if key and not _RE_API_KEY.match(key):
            _append(
                "✘  Formato de chave inválido.\n"
                '   Chaves OpenRouter costumam ter o formato "sk-or-v1-…".',
                "error",
            )
            return None
        return effective_key, model

    # ── Save ──────────────────────────────────────────────────────────────────
    def _on_save() -> None:
        validated = _validate_inputs()
        if validated is None:
            return
        effective_key, model = validated

        save_btn.config(state=tk.DISABLED)
        test_btn.config(state=tk.DISABLED)
        _append("Salvando chave e modelo…", "dim")

        def _do_save() -> None:
            try:
                if effective_key != load_api_key():
                    save_api_key(effective_key)
                    dialog.after(0, lambda: _append("✔  Chave salva.", "success"))
                else:
                    dialog.after(0, lambda: _append("ℹ  Usando chave já armazenada.", "dim"))
                save_ai_model(model)
                dialog.after(0, lambda: _append("✔  Modelo salvo.", "success"))
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
        effective_key, model = validated

        test_btn.config(state=tk.DISABLED)
        save_btn.config(state=tk.DISABLED)

        def _do_test() -> None:
            try:
                from openai import OpenAI

                dialog.after(0, lambda: _append("Testando conexão com a API OpenRouter…", "dim"))

                client = OpenAI(
                    base_url="https://openrouter.ai/api/v1",
                    api_key=effective_key,
                )
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": "Responda apenas com a palavra: OK"}],
                )
                resp_text = ""
                if getattr(response, "choices", None) and len(response.choices) > 0:
                    resp_text = (response.choices[0].message.content or "").strip()

                if not resp_text:
                    raise ValueError("A IA não retornou conteúdo na resposta.")

                dialog.after(0, lambda: _append("✔  IA respondeu — configuração válida.", "success"))
                dialog.after(0, lambda: _append(f"   Resposta: {resp_text}"))
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

    # ── Status bar (API + Update) ─────────────────────────────────────────
    status_frame = tk.Frame(root, bg=_c["bg"])
    status_frame.pack(fill=tk.X, padx=25, pady=(10, 0))
    theme.widgets['status_frame'] = status_frame

    # API button
    api_btn = tk.Button(status_frame, text="🔑 API de IA", font=("Segoe UI", 9, "bold"),
                        relief=tk.FLAT, cursor="hand2", padx=8, pady=3,
                        bg=_c["btn_sec_bg"], fg=_c["btn_sec_fg"],
                        activebackground=_c["btn_sec_hover"], activeforeground=_c["fg"],
                        command=lambda: open_ai_api_dialog(root, theme.mode))
    api_btn.pack(side=tk.LEFT)
    theme.widgets['api_btn'] = api_btn

    # Update status badge (independent check)
    _colors = z7_theme.get_theme_colors(theme.mode)
    update_status_lbl = tk.Label(status_frame, text="⏳ Checando atualizações...",
                                 font=("Segoe UI", 9), fg=_colors["fg_muted"],
                                 bg=_colors["bg"])
    update_status_lbl.pack(side=tk.LEFT, padx=(12, 0))
    theme.widgets['update_status_lbl'] = update_status_lbl

    # Run independent update check in background
    def _check_updates_bg() -> None:
        try:
            data = get_latest_github_release()
            tag_name = data.get("tag_name", "").strip()
            if not tag_name:
                tag_name = data.get("name", "").strip()
            if not tag_name:
                root.after(0, lambda: update_status_lbl.config(
                    text="⚠ Erro ao checar atualização", fg="#ef4444"))
                return
            comparison = compare_versions(tag_name, _APP_VERSION)
            if comparison > 0:
                root.after(0, lambda: update_status_lbl.config(
                    text="🔄 Atualização disponível", fg="#f59e0b"))
            else:
                root.after(0, lambda: update_status_lbl.config(
                    text="✔ App atualizado",
                    fg="#10b981"))
        except Exception as ex:
            LOGGER.warning("Background update check failed: %s", ex)
            root.after(0, lambda: update_status_lbl.config(
                text="⚠ Erro ao checar atualização", fg="#ef4444"))

    threading.Thread(target=_check_updates_bg, daemon=True).start()

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

    _select_bg = "#6366f1" if theme.mode == "dark" else "#7c3aed"
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

    import webbrowser
    github_btn = tk.Button(btn_frame, text="⧉ GitHub", font=("Segoe UI", 9),
                           relief=tk.FLAT, cursor="hand2",
                           command=lambda: webbrowser.open(GITHUB_REPO_URL))
    github_btn.pack(side=tk.LEFT, padx=(0, 5))
    theme.widgets.setdefault('sec_btns', []).append(github_btn)

    update_btn = tk.Button(btn_frame, text="🔄 Atualizar...",
                           font=("Segoe UI", 9, "bold"), relief=tk.FLAT,
                           cursor="hand2",
                           command=lambda: check_for_updates_ui(root))
    update_btn.pack(side=tk.LEFT, padx=(5, 5))
    theme.widgets.setdefault('sec_btns', []).append(update_btn)

    theme.widgets['sec_btns'] = [cancel_btn, restore_btn, github_btn, update_btn]

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
    status_frame.pack(fill=tk.X, padx=25, pady=(10, 0))
    frame.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=25, pady=(4, 10))

    # Bring window to front on startup without staying always-on-top
    root.lift()
    root.focus_force()

    root.mainloop()

if __name__ == "__main__":
    main()
