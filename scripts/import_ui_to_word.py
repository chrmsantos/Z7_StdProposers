"""
import_ui_to_word.py
Imports Quick Access Toolbar (QAT) customization from an .exportedUI file
into the user's Word.officeUI configuration.

Workflow:
  1. Locates the correct Word.officeUI path based on Office version
     - Office 2016+/365: %APPDATA%\\Microsoft\\Office\\Word.officeUI (Roaming)
     - Older versions:      %LOCALAPPDATA%\\Microsoft\\Office\\Word.officeUI (Local)
  2. Detects running Word instances and closes them gracefully (or forcefully)
     — Word locks the .officeUI file while running, so it must be closed
  3. Creates a timestamped backup of any existing Word.officeUI
  4. Copies the .exportedUI file to Word.officeUI
  5. (Optional) Adjusts HKCU registry entries to prevent Word's internal
     security from blocking the customization:
     - Adds %APPDATA%\\Microsoft\\Templates as a Trusted Location
     - Lowers VBA macro security to allow signed/trusted macros
  6. Verifies the file was written correctly

Requires: pywin32 (pip install pywin32) for COM-based Word operations.
No admin privileges required — operates entirely in the user's profile.

Usage:
  python scripts/import_ui_to_word.py [--source PATH] [--dry-run] [--force]
                                     [--no-backup] [--no-trust] [--verbose]
"""

# ---------------------------------------------------------------------------
# Standard library imports
# ---------------------------------------------------------------------------
import argparse
import datetime
import logging
import os
import shutil
import json as _json
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
LOGGER_NAME = "import_ui"
BACKUP_SUFFIX = ".backup"
DEFAULT_SOURCE_NAME = "word_ui.exportedUI"
# Office version → registry key prefix for trusted locations / VBA security
OFFICE_REGISTRY_VERSIONS = ["16.0", "15.0", "14.0", "12.0", "11.0"]
WORD_OFFICEUI_NAME = "Word.officeUI"

# HKCU registry paths (no admin required)
REG_VBA_WARNINGS_TEMPLATE = (
    r"Software\Microsoft\Office\{version}\Word\Security"
)
REG_TRUSTED_LOCATIONS_TEMPLATE = (
    r"Software\Microsoft\Office\{version}\Word\Security\Trusted Locations"
)

# Macro security values
VBA_WARNINGS_ENABLE_ALL = 1
VBA_WARNINGS_DISABLE_SIGNED = 2
VBA_WARNINGS_DISABLE_EXCEPT_TRUSTED = 3
VBA_WARNINGS_DISABLE_ALL = 4
# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

def _get_local_app_data() -> str:
    """Resolve LOCALAPPDATA with fallback."""
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return local
    profile = os.environ.get("USERPROFILE", "")
    if profile:
        return str(Path(profile) / "AppData" / "Local")
    return str(Path.cwd())


def _get_logs_dir() -> Path:
    """Return and create the logs directory."""
    logs_dir = Path(_get_local_app_data()) / "Z7" / "Apps" / "StdProposers" / "LocalConfigs" / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir


def setup_logging(verbose: bool = False) -> logging.Logger:
    """
    Configure a logger with rotating file + console output.

    Args:
        verbose: If True, set console level to DEBUG; otherwise INFO.

    Returns:
        Configured logger instance.
    """
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    if logger.handlers:
        return logger

    fmt = logging.Formatter(
        "%(asctime)s.%(msecs)03d | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(fmt)
    logger.addHandler(console)

    try:
        log_path = _get_logs_dir() / "import_ui_to_word.log"
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=2 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)
    except OSError as exc:
        logger.warning("Could not create log file: %s", exc)

    return logger
# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------

def resolve_project_root() -> Path:
    """
    Determine the project root directory.

    When running as a PyInstaller frozen executable, the .exe is inside
    scripts/dist/; go two levels up. In script mode, go one level up.
    """
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        return exe_dir.parent.parent
    return Path(__file__).resolve().parent.parent


def resolve_default_source() -> Path:
    """Resolve the default .exportedUI source file in dist/."""
    return resolve_project_root() / "dist" / DEFAULT_SOURCE_NAME


def _detect_office_version_from_registry() -> Optional[str]:
    """
    Detect installed Office version from the registry (HKLM or HKCU).

    Returns the highest version string found (e.g. '16.0') or None.
    """
    import winreg

    candidate_roots = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Office"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Office"),
    ]

    highest = None
    for root_key, sub in candidate_roots:
        try:
            with winreg.OpenKey(root_key, sub) as key:
                i = 0
                while True:
                    try:
                        name = winreg.EnumKey(key, i)
                    except OSError:
                        break
                    parts = name.split(".")
                    if len(parts) == 2 and all(p.isdigit() for p in parts):
                        try:
                            with winreg.OpenKey(key, name + r"\Word") as _w:
                                pass
                        except OSError:
                            i += 1
                            continue
                        major = int(parts[0])
                        if highest is None or major > int(highest.split(".")[0]):
                            highest = name
                    i += 1
        except OSError:
            continue

    return highest


def resolve_word_officeui_path() -> Path:
    """
    Determine the correct Word.officeUI path.

    Office 2016+ (16.0+) uses Roaming (%APPDATA%);
    older versions use Local (%LOCALAPPDATA%).
    """
    version = _detect_office_version_from_registry()
    if version:
        try:
            major = int(version.split(".")[0])
        except (ValueError, IndexError):
            major = 0
    else:
        major = 16

    if major >= 16:
        base = os.environ.get("APPDATA")
        if not base:
            profile = os.environ.get("USERPROFILE", "")
            base = str(Path(profile) / "AppData" / "Roaming") if profile else ""
    else:
        base = _get_local_app_data()

    if not base:
        raise FileNotFoundError(
            "Could not resolve %%APPDATA%% or %%LOCALAPPDATA%% for Word.officeUI"
        )

    return Path(base) / "Microsoft" / "Office" / WORD_OFFICEUI_NAME
# ---------------------------------------------------------------------------
# Word process management
# ---------------------------------------------------------------------------

def find_word_processes() -> list:
    """
    Find running WINWORD.EXE processes via PowerShell.

    Returns:
        List of dicts with Id, ProcessName, MainWindowTitle.
    """
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-Process -Name WINWORD -ErrorAction SilentlyContinue | "
                "Select-Object Id, ProcessName, MainWindowTitle | "
                "ConvertTo-Json",
            ],
            capture_output=True,
            text=True,
            timeout=15,
        )
        if result.returncode != 0 or not result.stdout.strip():
            return []
        data = _json.loads(result.stdout.strip())
        if isinstance(data, dict):
            data = [data]
        return data
    except (subprocess.TimeoutExpired, _json.JSONDecodeError, OSError):
        return []


def close_word_gracefully(logger: logging.Logger, timeout: int = 15) -> bool:
    """
    Attempt to close Word gracefully via COM.

    Returns True if all instances were closed, False otherwise.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            word.Quit(SaveChanges=0)
            logger.info("Sent Quit to running Word instance.")
        except Exception:
            logger.debug("No active Word instance found via COM.")
            return True

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if not find_word_processes():
                logger.info("Word closed gracefully.")
                return True
            time.sleep(0.5)
        return False
    except ImportError as exc:
        logger.warning("pywin32 not available for graceful close: %s", exc)
        return False
    except Exception as exc:
        logger.warning("Graceful Word close failed: %s", exc)
        return False


def kill_word_processes(logger: logging.Logger) -> bool:
    """
    Forcefully terminate all WINWORD.EXE processes.

    Returns True if no Word processes remain, False otherwise.
    """
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
            capture_output=True,
            timeout=15,
        )
    except (subprocess.TimeoutExpired, OSError) as exc:
        logger.error("Failed to kill Word processes: %s", exc)
        return False

    remaining = find_word_processes()
    if remaining:
        logger.error(
            "Word processes still running after kill: %s",
            [(p.get("Id"), p.get("MainWindowTitle", "")) for p in remaining],
        )
        return False
    logger.info("All Word processes terminated.")
    return True


def ensure_word_closed(logger: logging.Logger, force: bool = False) -> bool:
    """
    Ensure Word is closed before modifying Word.officeUI.

    Tries graceful close first, then forceful kill if --force is set.
    Returns True if Word is confirmed closed, False otherwise.
    """
    processes = find_word_processes()
    if not processes:
        logger.debug("No Word processes detected - safe to proceed.")
        return True

    logger.warning(
        "Word is running (%d instance(s)). UI file is locked while Word is open.",
        len(processes),
    )
    for p in processes:
        logger.info(
            "  PID %s - %s",
            p.get("Id", "?"),
            p.get("MainWindowTitle", ""),
        )

    if close_word_gracefully(logger):
        return True

    if force:
        logger.warning("Graceful close failed. Forcefully terminating Word...")
        return kill_word_processes(logger)

    logger.error(
        "Could not close Word gracefully. Use --force to terminate forcefully, "
        "or close Word manually and try again."
    )
    return False
# ---------------------------------------------------------------------------
# Registry operations (HKCU - no admin required)
# ---------------------------------------------------------------------------

def _get_office_version_for_registry(logger: logging.Logger) -> str:
    """Determine the Office version to use for registry operations."""
    version = _detect_office_version_from_registry()
    if version:
        return version
    import winreg
    for v in OFFICE_REGISTRY_VERSIONS:
        try:
            key_path = REG_VBA_WARNINGS_TEMPLATE.format(version=v)
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ
            ):
                logger.debug("Found Office registry version: %s", v)
                return v
        except OSError:
            continue
    logger.debug("No Office registry version found; defaulting to 16.0")
    return "16.0"


def set_vba_security(logger: logging.Logger, dry_run: bool = False) -> bool:
    """
    Lower VBA macro security to allow trusted macros.

    Sets HKCU VBAWarnings to 3 (Disable all macros except digitally signed).

    Returns True on success.
    """
    try:
        import winreg
        version = _get_office_version_for_registry(logger)
        key_path = REG_VBA_WARNINGS_TEMPLATE.format(version=version)
        if dry_run:
            logger.info(
                "[DRY-RUN] Would set VBAWarnings=3 at HKCU\\%s", key_path
            )
            return True
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 3)
        logger.info("VBA macro security set to 'Disable except digitally signed'.")
        return True
    except ImportError:
        logger.warning("winreg not available - skipping VBA security config.")
        return False
    except OSError as exc:
        logger.warning("Could not set VBA security: %s", exc)
        return False


def add_trusted_location(
    logger: logging.Logger, path: str, dry_run: bool = False
) -> bool:
    """
    Add a directory to Word's Trusted Locations in HKCU.

    Prevents Word from blocking macros loaded from templates in that directory.

    Args:
        path: Absolute path to the trusted directory.
        dry_run: If True, only logs what would be done.

    Returns True on success.
    """
    try:
        import winreg
        version = _get_office_version_for_registry(logger)
        base_key = REG_TRUSTED_LOCATIONS_TEMPLATE.format(version=version)
        if dry_run:
            logger.info(
                "[DRY-RUN] Would add trusted location '%s' at HKCU\\%s",
                path, base_key,
            )
            return True

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base_key) as parent:
            existing_indices = []
            i = 0
            while True:
                try:
                    sub_name = winreg.EnumKey(parent, i)
                except OSError:
                    break
                if sub_name.startswith("Location"):
                    try:
                        existing_indices.append(int(sub_name[len("Location"):]))
                    except ValueError:
                        pass
                i += 1

            norm_path = path.rstrip("\\").lower()
            for idx in existing_indices:
                try:
                    with winreg.OpenKey(parent, f"Location{idx}") as loc:
                        existing_path, _ = winreg.QueryValueEx(loc, "Path")
                        if str(existing_path).rstrip("\\").lower() == norm_path:
                            logger.debug(
                                "Path '%s' already trusted at Location%d.", path, idx
                            )
                            return True
                except OSError:
                    continue

            next_idx = max(existing_indices) + 1 if existing_indices else 0

        location_key = f"{base_key}\\Location{next_idx}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, location_key) as key:
            winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, path)
            winreg.SetValueEx(key, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(
                key, "Description", 0, winreg.REG_SZ,
                "Z7 StdProposers - QAT macro template"
            )
        logger.info("Added trusted location: %s (Location%d)", path, next_idx)
        return True
    except ImportError:
        logger.warning("winreg not available - skipping trusted locations.")
        return False
    except OSError as exc:
        logger.warning("Could not add trusted location: %s", exc)
        return False
# ---------------------------------------------------------------------------
# File operations
# ---------------------------------------------------------------------------

def backup_existing_officeui(
    target: Path, logger: logging.Logger, dry_run: bool = False
) -> Optional[Path]:
    """
    Create a timestamped backup of the existing Word.officeUI file.

    Returns the backup path, or None if no backup was needed.
    """
    if not target.exists():
        logger.info("No existing Word.officeUI to backup.")
        return None

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = target.with_name(f"{WORD_OFFICEUI_NAME}{BACKUP_SUFFIX}.{timestamp}")

    if dry_run:
        logger.info("[DRY-RUN] Would backup: %s -> %s", target, backup_path)
        return backup_path

    try:
        shutil.copy2(target, backup_path)
        logger.info("Backup created: %s", backup_path)
        return backup_path
    except OSError as exc:
        logger.error("Failed to create backup: %s", exc)
        raise


def validate_exportedui_file(source: Path, logger: logging.Logger) -> bool:
    """
    Validate that the .exportedUI file contains expected Word UI content.

    exportedUI files may have unbound namespace prefixes and multiple
    top-level elements. We use heuristics rather than strict XML parsing.

    Returns True if the file appears valid.
    """
    if not source.exists():
        logger.error("Source file not found: %s", source)
        return False

    if not source.is_file():
        logger.error("Source path is not a file: %s", source)
        return False

    try:
        content = source.read_text(encoding="utf-8")
        if not content.strip():
            logger.error("Source file is empty: %s", source)
            return False

        # exportedUI files are not strictly valid XML (unbound prefixes,
        # multiple top-level elements). Use heuristic checks instead.
        has_custom_ui = "customUI" in content
        has_qat = "<mso:qat>" in content or "mso:qat" in content
        has_word = 'app="Word"' in content

        if has_custom_ui and has_qat and has_word:
            logger.info(
                "Source file validated: %s (%d bytes)",
                source.name,
                len(content.encode("utf-8")),
            )
            return True

        # Still try XML as a fallback for cleaner exports
        try:
            wrapped = "<root>" + content + "</root>"
            ET.fromstring(wrapped)
            logger.info(
                "Source file validated (XML): %s (%d bytes)",
                source.name,
                len(content.encode("utf-8")),
            )
            return True
        except ET.ParseError:
            pass

        logger.error(
            "Source file does not appear to be a valid Word UI export: %s",
            source,
        )
        return False
    except OSError as exc:
        logger.error("Could not read source file: %s", exc)
        return False


def copy_ui_file(
    source: Path,
    target: Path,
    logger: logging.Logger,
    dry_run: bool = False,
) -> bool:
    """
    Copy the .exportedUI file to the Word.officeUI location.

    Args:
        source: Path to the .exportedUI file.
        target: Path to Word.officeUI.
        logger: Logger instance.
        dry_run: If True, simulate without writing.

    Returns True on success.
    """
    if dry_run:
        logger.info("[DRY-RUN] Would copy: %s -> %s", source, target)
        return True

    try:
        content = source.read_bytes()
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(content)

        if not target.exists():
            logger.error("Copy verification failed: target file does not exist.")
            return False

        target_content = target.read_bytes()
        if len(target_content) != len(content):
            logger.error(
                "Copy verification failed: size mismatch (%d vs %d bytes).",
                len(target_content), len(content),
            )
            return False

        logger.info(
            "Successfully copied UI file: %s -> %s (%d bytes)",
            source.name, target, len(content),
        )
        return True
    except PermissionError as exc:
        logger.error(
            "Permission denied writing to %s. Is Word still running?", target
        )
        logger.error("Details: %s", exc)
        return False
    except OSError as exc:
        logger.error("File operation failed: %s", exc)
        return False
# ---------------------------------------------------------------------------
# CLI parsing
# ---------------------------------------------------------------------------

def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Import Word Quick Access Toolbar customization from .exportedUI",
    )
    parser.add_argument(
        "--source",
        type=str,
        default=None,
        help="Path to .exportedUI file (default: dist/word_ui.exportedUI)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Simulate the import without modifying any files or registry",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Forcefully terminate Word if graceful close fails",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Skip backup of existing Word.officeUI",
    )
    parser.add_argument(
        "--no-trust",
        action="store_true",
        help="Skip registry changes for trusted locations and VBA security",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Show detailed debug messages",
    )
    return parser.parse_args(argv)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _print_banner(logger: logging.Logger) -> None:
    """Print the application banner."""
    logger.info("=" * 60)
    logger.info("Z7 StdProposers - Importador de UI do Word (QAT)")
    logger.info("=" * 60)


def _print_summary(logger: logging.Logger, stats: dict) -> None:
    """Print a summary of the operation."""
    logger.info("-" * 60)
    success = bool(stats.get("success"))
    status = "SUCESSO" if success else "FALHA"
    logger.info("Status: %s", status)
    for key, value in stats.items():
        if key == "success":
            continue
        logger.info("  %s: %s", key, value)
    logger.info("=" * 60)


def main(argv: Optional[List[str]] = None) -> int:
    """
    Entry point for the import_ui_to_word script.

    Args:
        argv: Command-line arguments (uses sys.argv if None).

    Returns:
        Exit code (0 = success, 1 = error).
    """
    args = parse_args(argv)
    logger = setup_logging(verbose=args.verbose)
    exit_code = 0
    stats: dict = {"success": False}

    _print_banner(logger)

    try:
        # Resolve source file
        if args.source:
            source = Path(args.source).resolve()
        else:
            source = resolve_default_source()
        logger.info("Source: %s", source)

        # Validate source file
        if not validate_exportedui_file(source, logger):
            return 1

        # Resolve target Word.officeUI path
        target = resolve_word_officeui_path()
        logger.info("Target: %s", target)
        stats["target"] = str(target)

        # Ensure Word is closed
        if not args.dry_run:
            if not ensure_word_closed(logger, force=args.force):
                logger.error(
                    "Cannot proceed while Word is still running. "
                    "Close Word manually or use --force."
                )
                return 1

        # Backup existing Word.officeUI
        if not args.no_backup:
            try:
                backup_path = backup_existing_officeui(
                    target, logger, dry_run=args.dry_run
                )
                stats["backup"] = str(backup_path) if backup_path else "(none)"
            except Exception as exc:
                logger.error("Backup failed: %s", exc)
                logger.error("Aborting to protect existing configuration.")
                return 1
        else:
            logger.warning("Backup skipped (--no-backup).")
            stats["backup"] = "(skipped)"

        # Copy the UI file
        if not copy_ui_file(source, target, logger, dry_run=args.dry_run):
            stats["success"] = False
            _print_summary(logger, stats)
            return 1
        stats["copied"] = True

        # Configure registry for security bypass
        if not args.no_trust:
            templates_dir = os.environ.get("APPDATA")
            if templates_dir:
                templates_dir = str(
                    Path(templates_dir) / "Microsoft" / "Templates"
                )
                stats["trusted_location"] = add_trusted_location(
                    logger, templates_dir, dry_run=args.dry_run
                )
            stats["vba_security"] = set_vba_security(
                logger, dry_run=args.dry_run
            )
        else:
            logger.warning("Registry configuration skipped (--no-trust).")
            stats["trusted_location"] = "(skipped)"
            stats["vba_security"] = "(skipped)"

        stats["success"] = True
        _print_summary(logger, stats)
        logger.info("Next steps: Launch Word to load the new QAT configuration.")

    except FileNotFoundError as exc:
        logger.error("File not found: %s", exc)
        exit_code = 1
    except PermissionError as exc:
        logger.error("Permission denied: %s", exc)
        logger.info(
            "Tip: Close Word and try again, or right-click -> Run as Administrator."
        )
        exit_code = 1
    except ImportError as exc:
        logger.error("Missing dependency: %s", exc)
        logger.info("Tip: pip install pywin32")
        exit_code = 1
    except Exception as exc:
        logger.exception("Unexpected error: %s: %s", type(exc).__name__, exc)
        exit_code = 1

    return exit_code


if __name__ == "__main__":
    _exit_code = main()
    print()
    print("Press any key to exit...")
    try:
        import msvcrt
        msvcrt.getch()
    except ImportError:
        input()
    sys.exit(_exit_code)