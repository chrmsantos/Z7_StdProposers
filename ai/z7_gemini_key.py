import os
import tkinter as tk
from pathlib import Path

import z7_theme
from z7_logging import configure_component_logger, log_exception

LOGGER = configure_component_logger("z7_gemini_key")

_KEY_REL_PATH = Path('AppData') / 'Local' / 'Z7' / 'Tmp' / 'StdProposers'
_KEY_FILE_NAME = 'gemini.key'


def _get_key_file() -> Path | None:
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        LOGGER.error("USERPROFILE env var not found")
        return None
    return Path(user_profile) / _KEY_REL_PATH / _KEY_FILE_NAME


def get_api_key(parent: tk.Misc | None = None) -> str | None:
    """Carrega a chave da API Gemini do armazenamento criptografado.
    Se ainda não existir, solicita ao usuário e persiste.
    Retorna None se a chave não estiver disponível.
    """
    LOGGER.info("Loading Gemini API key")
    key_file = _get_key_file()
    if key_file is None:
        return None

    if key_file.exists():
        try:
            import win32crypt
            with open(key_file, 'rb') as f:
                encrypted_key = f.read()
            _, decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)
            LOGGER.info("API key loaded from encrypted file")
            return decrypted_key.decode('utf-8')
        except Exception as e:
            log_exception(LOGGER, "Failed to decrypt API key", e)

    api_key = z7_theme.ask_string(
        "Z7 StdProposers",
        "Insira sua chave da API do Google Gemini:\n(Ela será criptografada e salva localmente)",
        parent=parent,
        show="*"
    )

    if not api_key or not api_key.strip():
        LOGGER.warning("User did not provide API key")
        return None

    api_key = api_key.strip()

    try:
        import win32crypt
        encrypted_key = win32crypt.CryptProtectData(
            api_key.encode('utf-8'), 'Z7_Gemini_Key', None, None, None, 0
        )
        key_file.parent.mkdir(parents=True, exist_ok=True)
        with open(key_file, 'wb') as f:
            f.write(encrypted_key)
        LOGGER.info("API key encrypted and persisted")
    except Exception as e:
        log_exception(LOGGER, "Failed to persist API key", e)

    return api_key


def delete_api_key() -> None:
    """Remove o arquivo de chave armazenado (usado em auto-reparo após erro de auth)."""
    key_file = _get_key_file()
    if key_file and key_file.exists():
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
    if key_file is None or not key_file.exists():
        return ""
    try:
        import win32crypt
        with open(key_file, 'rb') as f:
            encrypted_key = f.read()
        _, decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)
        LOGGER.info("Stored API key read successfully")
        return decrypted_key.decode('utf-8')
    except Exception as e:
        log_exception(LOGGER, "Failed to decrypt stored API key", e)
        return ""


def write_api_key(api_key: str) -> None:
    """Criptografa e persiste uma chave da API Gemini sem interação com o usuário."""
    api_key = api_key.strip()
    if not api_key:
        return
    key_file = _get_key_file()
    if key_file is None:
        return
    try:
        import win32crypt
        encrypted_key = win32crypt.CryptProtectData(
            api_key.encode('utf-8'), 'Z7_Gemini_Key', None, None, None, 0
        )
        key_file.parent.mkdir(parents=True, exist_ok=True)
        with open(key_file, 'wb') as f:
            f.write(encrypted_key)
        LOGGER.info("API key encrypted and persisted via write_api_key")
    except Exception as e:
        log_exception(LOGGER, "Failed to persist API key via write_api_key", e)
