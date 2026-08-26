import contextlib
import datetime
import logging
import os
import sys
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path

__all__ = [
    "is_frozen",
    "get_runtime_dir",
    "get_data_dir",
    "get_logs_dir",
    "build_log_path",
    "get_component_log_path",
    "get_session_id",
    "configure_component_logger",
    "log_context",
    "log_exception",
]


def is_frozen() -> bool:
    """Retorna True quando executando como binário PyInstaller compilado."""
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")


def _get_local_app_data() -> str:
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return local
    user_profile = os.environ.get("USERPROFILE", "")
    return str(Path(user_profile) / "AppData" / "Local") if user_profile else str(Path.cwd())


def get_runtime_dir() -> Path:
    """Retorna e cria o diretório raiz de instalação do Z7_StdProposers."""
    runtime_dir = Path(_get_local_app_data()) / "Z7" / "Apps" / "Z7_StdProposers"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir


def get_data_dir() -> Path:
    """Retorna e cria o diretório de dados do usuário (configurações, chave, etc.)."""
    user_profile = os.environ.get("USERPROFILE", "")
    if user_profile:
        data_dir = Path(user_profile) / "AppData" / "Local" / "Z7" / "Apps" / "Z7_StdProposers" / "setup"
    else:
        data_dir = get_runtime_dir() / "setup"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def get_logs_dir() -> Path:
    """Retorna e cria o diretório de logs."""
    logs_dir = get_data_dir() / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir


def build_log_path(component: str) -> Path:
    """Monta o caminho do arquivo de log para um componente, sanitizando o nome."""
    safe_component = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in component)
    return get_logs_dir() / f"{safe_component}.log"


def get_component_log_path(component: str) -> Path:
    """Alias público para build_log_path — retorna o caminho do log do componente."""
    return build_log_path(component)


def get_session_id() -> str:
    """Retorna um ID de sessão único baseado no timestamp atual (formato YYYYMMDD_HHMMSS).

    Útil para correlacionar entradas de log de múltiplos componentes na mesma sessão.
    """
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


@contextlib.contextmanager
def log_context(logger: logging.Logger, label: str, level: int = logging.INFO):
    """Context manager que registra início, fim e duração de um bloco nomeado.

    Em caso de exceção, registra em nível ERROR e relança a exceção.

    Exemplo::

        with log_context(LOGGER, "Chamada à API Gemini"):
            response = client.models.generate_content(...)
    """
    logger.log(level, "[START] %s", label)
    t0 = time.monotonic()
    try:
        yield
    except Exception as exc:
        elapsed = time.monotonic() - t0
        logger.error("[FAIL]  %s | %.3fs | %s: %s", label, elapsed, type(exc).__name__, exc)
        raise
    else:
        elapsed = time.monotonic() - t0
        logger.log(level, "[END]   %s | %.3fs", label, elapsed)


def configure_component_logger(
    component: str,
    level: int = logging.INFO,
    *,
    debug: bool = False,
) -> logging.Logger:
    """Configura e retorna o logger para o componente especificado.

    Idempotente: se o logger já possuir handlers configurados, é retornado sem
    modificação.

    Args:
        component: Nome do componente (ex: ``"correct_grammar"``).
        level:     Nível mínimo de log (padrão: ``logging.INFO``).
        debug:     Se True, adiciona também um ``StreamHandler`` para ``stderr``
                   com nível ``DEBUG`` — útil durante desenvolvimento.
    """
    logger = logging.getLogger(f"z7.{component}")
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG if debug else level)
    logger.propagate = False

    formatter = logging.Formatter(
        fmt="%(asctime)s.%(msecs)03d | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    log_path = build_log_path(component)
    file_handler = RotatingFileHandler(
        log_path, maxBytes=2 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(level)
    logger.addHandler(file_handler)

    if debug:
        stream_handler = logging.StreamHandler(sys.stderr)
        stream_handler.setFormatter(formatter)
        stream_handler.setLevel(logging.DEBUG)
        logger.addHandler(stream_handler)

    return logger


def log_exception(
    logger: logging.Logger,
    context: str,
    exc: Exception,
    *,
    reraise: bool = False,
) -> None:
    """Registra a exceção com traceback completo.

    Args:
        logger:  Logger do componente.
        context: Mensagem descritiva de contexto.
        exc:     Exceção capturada.
        reraise: Se True, relança a exceção após registrá-la.
    """
    logger.exception("%s: %s", context, exc)
    if reraise:
        raise exc
