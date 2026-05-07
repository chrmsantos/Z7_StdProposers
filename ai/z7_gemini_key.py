import tkinter as tk
from pathlib import Path
from typing import Optional

import z7_theme
from z7_logging import configure_component_logger, get_data_dir, log_exception

LOGGER = configure_component_logger("z7_gemini_key")

__all__ = [
    "get_api_key",
    "delete_api_key",
    "read_stored_api_key",
    "write_api_key",
]

_KEY_FILE_NAME = "gemini.key"


def _get_key_file() -> Path:
    return get_data_dir() / _KEY_FILE_NAME


def _validate_api_key(key: str) -> bool:
    """Validação mínima da chave: não-vazia e com comprimento plausível."""
    return bool(key) and len(key) >= 20


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

    Se ainda não existir, solicita ao usuário e persiste.
    Retorna None se a chave não estiver disponível.
    """
    LOGGER.info("Loading Gemini API key")
    key_file = _get_key_file()

    if key_file.exists():
        key = _decrypt_key_file(key_file)
        if key and _validate_api_key(key):
            LOGGER.info("API key loaded from encrypted file")
            return key
        LOGGER.warning("Stored key invalid or unreadable; prompting user")

    api_key = z7_theme.ask_string(
        "Z7 StdProposers",
        "Insira sua chave da API do Google Gemini:\n(Ela será criptografada e salva localmente)",
        parent=parent,
        show="*",
    )

    if not api_key or not api_key.strip():
        LOGGER.warning("User did not provide API key")
        return None

    api_key = api_key.strip()
    _encrypt_and_persist(api_key, key_file)
    return api_key


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
    """Retorna a chave armazenada descriptografada, ou string vazia se não existir.

    Não exibe nenhum diálogo; use get_api_key() quando quiser solicitar ao usuário.
    """
    key_file = _get_key_file()
    if not key_file.exists():
        return ""
    key = _decrypt_key_file(key_file)
    if key is not None:
        LOGGER.info("Stored API key read successfully")
    return key or ""


def write_api_key(api_key: str) -> None:
    """Criptografa e persiste uma chave da API Gemini sem interação com o usuário."""
    api_key = api_key.strip()
    if not api_key:
        return
    _encrypt_and_persist(api_key, _get_key_file())
