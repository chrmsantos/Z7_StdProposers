import re
import tkinter as tk
from pathlib import Path
from typing import Optional

from z7_logging import configure_component_logger, get_data_dir, log_exception

LOGGER = configure_component_logger("z7_gemini_key")

__all__ = [
    "get_api_key",
    "delete_api_key",
    "read_stored_api_key",
    "write_api_key",
]

_KEY_FILE_NAME = "gemini.key"
_API_KEY_PATTERN = re.compile(r"^AIza[A-Za-z0-9_\-]{35}$")


def _get_key_file() -> Path:
    return get_data_dir() / _KEY_FILE_NAME


def _validate_api_key(key: str) -> bool:
    """Validação mínima da chave: formato AIza + 35 caracteres alfanuméricos."""
    return bool(key) and bool(_API_KEY_PATTERN.match(key))


def _decrypt_key_file(key_file: Path) -> Optional[str]:
    """Lê e descriptografa o arquivo de chave. Retorna a chave ou None em caso de falha."""
    try:
        import win32crypt
        encrypted = key_file.read_bytes()
        _, decrypted = win32crypt.CryptUnprotectData(encrypted, None, None, None, 0)
        return decrypted.decode("utf-8")
    except Exception as e:
        log_exception(LOGGER, "Failed to decrypt API key", e)
        return None


def _encrypt_and_persist(api_key: str, key_file: Path) -> bool:
    """Criptografa e persiste a chave. Retorna True em caso de sucesso."""
    try:
        import win32crypt
        encrypted = win32crypt.CryptProtectData(
            api_key.encode("utf-8"), "Z7_Gemini_Key", None, None, None, 0
        )
        key_file.parent.mkdir(parents=True, exist_ok=True)
        key_file.write_bytes(encrypted)
        LOGGER.info("API key encrypted and persisted")
        return True
    except Exception as e:
        log_exception(LOGGER, "Failed to persist API key", e)
        return False


def get_api_key(parent: Optional[tk.Misc] = None) -> Optional[str]:
    """Carrega a chave da API Gemini do armazenamento criptografado.

    Se ainda não existir, utiliza a chave padrão, criptografa e persiste.
    Retorna a chave da API.
    """
    LOGGER.info("Loading Gemini API key")
    key_file = _get_key_file()

    if key_file.exists():
        key = _decrypt_key_file(key_file)
        if key and _validate_api_key(key):
            LOGGER.info("API key loaded from encrypted file")
            return key
        LOGGER.warning("Stored key invalid or unreadable; using default API key")

    default_key = "AIzaSyDM66y2zHExKWLwwGwKbE82EzrteMmMMkk"
    _encrypt_and_persist(default_key, key_file)
    return default_key


def delete_api_key() -> None:
    """Remove o arquivo de chave armazenado (usado em auto-reparo após erro de auth)."""
    key_file = _get_key_file()
    if key_file.exists():
        try:
            key_file.unlink()
            LOGGER.info("API key file deleted")
        except Exception as e:
            log_exception(LOGGER, "Failed to delete API key file", e)


def read_stored_api_key() -> str:
    """Retorna a chave armazenada descriptografada, ou a chave padrão se não existir.

    Não exibe nenhum diálogo; use get_api_key() quando quiser solicitar ao usuário.
    """
    key_file = _get_key_file()
    if not key_file.exists():
        default_key = "AIzaSyDM66y2zHExKWLwwGwKbE82EzrteMmMMkk"
        _encrypt_and_persist(default_key, key_file)
        return default_key
    key = _decrypt_key_file(key_file)
    if key:
        LOGGER.info("Stored API key read successfully")
        return key
    default_key = "AIzaSyDM66y2zHExKWLwwGwKbE82EzrteMmMMkk"
    _encrypt_and_persist(default_key, key_file)
    return default_key


def write_api_key(api_key: str) -> None:
    """Criptografa e persiste uma chave da API Gemini sem interação com o usuário."""
    api_key = api_key.strip()
    if not api_key:
        return
    _encrypt_and_persist(api_key, _get_key_file())
