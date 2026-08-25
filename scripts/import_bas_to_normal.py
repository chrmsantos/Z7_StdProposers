"""
import_bas_to_normal.py
Imports .bas VBA module files from source/main/ into the user's Normal.dotm.

Workflow:
  1. Locates Normal.dotm in %APPDATA%\\Microsoft\\Templates
  2. Creates a timestamped backup (Normal_backup.dotm) in the same folder
  3. Opens Normal.dotm via Word COM (win32com)
  4. Removes any existing Z7 modules and re-imports all .bas files
  5. Saves and closes

Requires: pywin32 (pip install pywin32), Microsoft Word installed.
No admin privileges required -- operates entirely in the user's profile.

Usage:
  python scripts/import_bas_to_normal.py [--dry-run] [--no-backup] [--verbose]
"""
import argparse
import datetime
import logging
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
LOGGER_NAME = "import_bas"
BACKUP_NAME = "Normal_backup.dotm"
NORMAL_DOTM_RELATIVE = os.path.join("Microsoft", "Templates", "Normal.dotm")
Z7_MODULE_PREFIX = "Mod_"

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

def setup_logging(verbose: bool = False) -> logging.Logger:
    """Configure a logger with console + optional file handler."""
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(console)

    try:
        log_dir = _get_logs_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        fh = logging.FileHandler(
            log_dir / f"import_bas_{ts}.log", encoding="utf-8"
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        logger.addHandler(fh)
    except OSError as e:
        logger.warning("Nao foi possivel criar arquivo de log: %s", e)

    return logger


def _get_logs_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return Path(local) / "Z7" / "Tmp" / "StdProposers" / "logs"
    profile = os.environ.get("USERPROFILE", "")
    if profile:
        return Path(profile) / "AppData" / "Local" / "Z7" / "Tmp" / "StdProposers" / "logs"
    return Path.cwd() / "logs"

# PLACEHOLDER_PART2

# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------

def resolve_project_root() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        return exe_dir.parent
    return Path(__file__).resolve().parent.parent


def resolve_normal_dotm() -> Path:
    """Locate Normal.dotm in %APPDATA%\\Microsoft\\Templates."""
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / NORMAL_DOTM_RELATIVE
    profile = os.environ.get("USERPROFILE", "")
    if profile:
        return Path(profile) / "AppData" / "Roaming" / NORMAL_DOTM_RELATIVE
    raise FileNotFoundError(
        "Nao foi possivel determinar APPDATA. "
        "Defina a variavel de ambiente APPDATA ou USERPROFILE."
    )


def discover_bas_files(source_dir: Path) -> list:
    if not source_dir.is_dir():
        raise FileNotFoundError(f"Diretorio de origem nao encontrado: {source_dir}")
    files = sorted(source_dir.glob("*.bas"))
    return [str(f) for f in files]

# PLACEHOLDER_PART3

# ---------------------------------------------------------------------------
# Backup
# ---------------------------------------------------------------------------

def create_backup(normal_path: Path, logger: logging.Logger) -> Path:
    """Create a timestamped backup of Normal.dotm alongside the original.

    The backup is named Normal_backup.dotm. If it already exists, a
    timestamped variant is created instead (Normal_backup_YYYYMMDD_HHMMSS.dotm).
    """
    if not normal_path.exists():
        raise FileNotFoundError(f"Normal.dotm nao encontrado: {normal_path}")

    backup_dir = normal_path.parent
    backup_path = backup_dir / BACKUP_NAME

    if backup_path.exists():
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        rotated = backup_dir / f"Normal_backup_{ts}.dotm"
        logger.info("Rotacionando backup existente para: %s", rotated.name)
        shutil.copy2(str(backup_path), str(rotated))

    shutil.copy2(str(normal_path), str(backup_path))
    logger.info("Backup criado: %s", backup_path)

    orig_size = normal_path.stat().st_size
    backup_size = backup_path.stat().st_size
    if backup_size != orig_size:
        raise IOError(
            f"Falha na verificacao de integridade do backup: "
            f"original={orig_size} bytes, backup={backup_size} bytes"
        )
    logger.debug("Verificacao de integridade: OK (%d bytes)", backup_size)
    return backup_path

# PLACEHOLDER_PART4

# ---------------------------------------------------------------------------
# Word COM operations
# ---------------------------------------------------------------------------

def import_modules_to_normal(
    normal_dotm: Path,
    bas_files: list,
    logger: logging.Logger,
    dry_run: bool = False,
) -> dict:
    """Import all .bas files into Normal.dotm's VBProject.

    Always creates a new isolated Word instance (DispatchEx) and works on a
    temporary copy of Normal.dotm to avoid conflicts with any running Word
    process that may have the file locked or be in a modal state.
    """
    try:
        import win32com.client
    except ImportError:
        raise ImportError("pywin32 nao instalado. Execute: pip install pywin32")

    stats = {"imported": 0, "skipped": 0, "errors": 0, "error_details": []}

    if dry_run:
        logger.info("[DRY-RUN] Simulacao: %d modulos seriam importados", len(bas_files))
        for f in bas_files:
            logger.info("  [DRY-RUN] %s", Path(f).name)
        return stats

    word = None
    doc = None
    tmp_copy = None

    try:
        # Work on a temporary copy so we never conflict with a running Word.
        fd, tmp_str = tempfile.mkstemp(suffix=".dotm")
        os.close(fd)
        tmp_copy = Path(tmp_str)
        shutil.copy2(str(normal_dotm), str(tmp_copy))
        logger.info("Copia temporaria criada: %s", tmp_copy)

        logger.info("Iniciando Microsoft Word (instancia isolada)...")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        logger.info("Abrindo copia temporaria de Normal.dotm...")
        doc = word.Documents.Open(str(tmp_copy))
        vb_project = doc.VBProject
        logger.info("VBProject: %s | Modulos existentes: %d",
                     vb_project.Name, vb_project.VBComponents.Count)

        # Step 1: Remove ALL existing Z7 modules (including stale ones).
        removed = _remove_z7_modules(vb_project, logger)
        if removed:
            logger.info("Modulos Z7 removidos: %d", removed)

        # Step 2: Import all .bas files.
        for bas_path in bas_files:
            bas_name = Path(bas_path).stem
            try:
                _import_single_module(vb_project, bas_path, bas_name, logger)
                stats["imported"] += 1
            except Exception as e:
                stats["errors"] += 1
                stats["error_details"].append((bas_name, str(e)))
                logger.error("  [ERRO] %s - %s", bas_name, e)

        logger.info("Salvando Normal.dotm...")
        # Use SaveAs2 with correct format for .dotm (macro-enabled template)
        # wdFormatXMLTemplateMacroEnabled = 15
        doc.SaveAs2(str(tmp_copy), FileFormat=15)
        logger.info("Normal.dotm salvo com sucesso!")

    except Exception as e:
        stats["errors"] += 1
        stats["error_details"].append(("_global", str(e)))
        logger.error("Erro durante processamento: %s", e)
        raise
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(0.5)

    # Copy the updated temp file back over Normal.dotm.
    if tmp_copy is not None and tmp_copy.exists():
        _copy_back_to_normal(tmp_copy, normal_dotm, logger)

    return stats


def _copy_back_to_normal(tmp_copy: Path, normal_dotm: Path,
                         logger: logging.Logger):
    """Replace Normal.dotm with the updated temporary copy.

    Retries a few times with a short delay because Word may not have fully
    released the file lock yet after quitting.
    """
    max_retries = 5
    for attempt in range(1, max_retries + 1):
        try:
            shutil.copy2(str(tmp_copy), str(normal_dotm))
            logger.info("Normal.dotm atualizado a partir da copia temporaria.")
            # Clean up temp file on success.
            try:
                tmp_copy.unlink(missing_ok=True)
            except OSError:
                pass
            return
        except PermissionError:
            if attempt < max_retries:
                logger.debug("Arquivo bloqueado, tentativa %d/%d...",
                             attempt, max_retries)
                time.sleep(1)
            else:
                logger.warning(
                    "Nao foi possivel substituir Normal.dotm (arquivo bloqueado). "
                    "Feche o Word e execute novamente, ou copie manualmente:\n"
                    "  de: %s\n  para: %s", tmp_copy, normal_dotm
                )

# PLACEHOLDER_PART5

def _remove_z7_modules(vb_project, logger: logging.Logger) -> int:
    """Remove all Z7 modules (names starting with Z7_MODULE_PREFIX) from the VBProject.

    Returns the number of modules removed.  Modules are collected first and
    removed in a second pass to avoid mutating the VBComponents collection
    while iterating over it.
    """
    to_remove = []
    for comp in vb_project.VBComponents:
        if comp.Name.startswith(Z7_MODULE_PREFIX):
            to_remove.append(comp)

    for comp in to_remove:
        try:
            logger.info("  [REM] %s", comp.Name)
            vb_project.VBComponents.Remove(comp)
        except Exception as e:
            logger.warning("  [REM] %s - falha ao remover: %s", comp.Name, e)

    return len(to_remove)


def _import_single_module(vb_project, bas_path: str, bas_name: str, logger: logging.Logger):
    """Import a single .bas module into the VBProject.

    Assumes that any pre-existing module with the same name has already been
    removed by _remove_z7_modules().
    """
    raw = Path(bas_path).read_bytes()

    # Decodifica o conteudo do .bas tentando UTF-8 primeiro e CP1252 como
    # fallback. Isso garante que modulos salvos acidentalmente em UTF-8 (sem
    # BOM) sejam lidos corretamente, preservando o legado CP1252 dos modulos
    # gerados por fix_bas_encoding.py.
    try:
        content = raw.decode("utf-8")
    except UnicodeDecodeError:
        content = raw.decode("cp1252")

    # Arquivos .bas podem conter U+FFFD (caracter de substituicao) fruto de
    # conversao de encoding anterior mal-sucedida. Esse codepoint nao existe em
    # CP1252, entao trocamos por '?' antes de gravar o temporario.
    content = content.replace("\ufffd", "?")

    # Write temp file in CP1252 so VBComponents.Import() (which reads using the
    # system ANSI codepage) sees the exact bytes expected.
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=".bas")
        os.close(fd)
        Path(tmp_path).write_text(content, encoding="cp1252")
        vb_project.VBComponents.Import(tmp_path)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    logger.info("  [OK] %s - importado", bas_name)

# PLACEHOLDER_PART6

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Importa modulos VBA (.bas) no Normal.dotm do usuario"
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Simula a importacao sem alterar Normal.dotm")
    parser.add_argument("--no-backup", action="store_true",
                        help="Nao cria backup do Normal.dotm antes de importar")
    parser.add_argument("--verbose", action="store_true",
                        help="Exibe mensagens de debug detalhadas")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logger = setup_logging(verbose=args.verbose)
    project_root = resolve_project_root()
    exit_code = 0

    logger.info("=" * 60)
    logger.info("Z7 StdProposers - Importador de Modulos VBA")
    logger.info("=" * 60)

    try:
        source_dir = project_root / "source" / "main"
        normal_dotm = resolve_normal_dotm()

        logger.info("Projeto: %s", project_root)
        logger.info("Origem: %s", source_dir)
        logger.info("Normal.dotm: %s", normal_dotm)

        bas_files = discover_bas_files(source_dir)
        if not bas_files:
            logger.error("Nenhum arquivo .bas encontrado em: %s", source_dir)
            return 1
        logger.info("Encontrados %d arquivo(s) .bas", len(bas_files))

        if not normal_dotm.exists():
            logger.error("Normal.dotm nao encontrado: %s", normal_dotm)
            logger.info("Dica: Abra o Word uma vez para criar o Normal.dotm padrao")
            return 1

        if not args.no_backup and not args.dry_run:
            try:
                backup_path = create_backup(normal_dotm, logger)
                logger.info("Backup concluido: %s", backup_path)
            except Exception as e:
                logger.error("Falha ao criar backup: %s", e)
                logger.error("Abortando para preservar o Normal.dotm original")
                return 1
        elif args.no_backup:
            logger.warning("Backup desabilitado por --no-backup")
        else:
            logger.info("[DRY-RUN] Backup seria criado em: %s / %s",
                        normal_dotm.parent, BACKUP_NAME)

        stats = import_modules_to_normal(
            normal_dotm, bas_files, logger, dry_run=args.dry_run
        )

        logger.info("-" * 60)
        logger.info("Resumo: %d importados, %d ignorados, %d erros",
                     stats["imported"], stats["skipped"], stats["errors"])

        if stats["errors"] > 0:
            logger.warning("Detalhes dos erros:")
            for name, detail in stats["error_details"]:
                logger.warning("  - %s: %s", name, detail)
            exit_code = 1
        else:
            logger.info("Importacao concluida com sucesso!")

    except FileNotFoundError as e:
        logger.error("Arquivo nao encontrado: %s", e)
        exit_code = 1
    except PermissionError as e:
        logger.error("Permissao negada: %s", e)
        logger.info("Dica: Feche o Word e tente novamente")
        exit_code = 1
    except ImportError as e:
        logger.error("Dependencia ausente: %s", e)
        exit_code = 1
    except Exception as e:
        logger.error("Erro inesperado: %s: %s", type(e).__name__, e)
        exit_code = 1

    logger.info("=" * 60)
    return exit_code


if __name__ == "__main__":
    _exit_code = main()
    print()
    print("Pressione qualquer tecla para sair...")
    try:
        import msvcrt
        msvcrt.getch()
    except ImportError:
        input()
    sys.exit(_exit_code)