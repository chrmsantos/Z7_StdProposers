import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
PY_ROOT = ROOT / "ai"
if str(PY_ROOT) not in sys.path:
    sys.path.insert(0, str(PY_ROOT))


# ---------------------------------------------------------------------------
# Stub win32crypt para ambientes sem pywin32 instalado
# ---------------------------------------------------------------------------
class _FakeCrypt:
    """Implementação mínima de CryptProtectData / CryptUnprotectData para testes."""

    @staticmethod
    def CryptProtectData(data: bytes, *args, **kwargs) -> bytes:
        return b"ENC:" + data

    @staticmethod
    def CryptUnprotectData(data: bytes, *args, **kwargs):
        if not data.startswith(b"ENC:"):
            raise ValueError("Not encrypted by stub")
        return (None, data[4:])


_win32crypt_patcher = mock.patch.dict(
    sys.modules,
    {"win32crypt": _FakeCrypt},
)


class TestGetKeyFile(unittest.TestCase):
    def test_key_file_uses_data_dir(self):
        import z7_gemini_key
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            # Patch the get_data_dir reference in z7_gemini_key's namespace directly
            # (avoids reload, which would create a log FileHandler locking the temp dir on Windows)
            with mock.patch.object(z7_gemini_key, "get_data_dir", return_value=tmp_path):
                key_file = z7_gemini_key._get_key_file()
                self.assertEqual(key_file.parent, tmp_path)
                self.assertEqual(key_file.name, "gemini.key")


class TestValidateApiKey(unittest.TestCase):
    def setUp(self):
        import importlib
        import z7_gemini_key
        importlib.reload(z7_gemini_key)
        self.mod = z7_gemini_key

    def test_empty_key_is_invalid(self):
        self.assertFalse(self.mod._validate_api_key(""))

    def test_short_key_is_invalid(self):
        self.assertFalse(self.mod._validate_api_key("abc"))

    def test_valid_length_key(self):
        self.assertTrue(self.mod._validate_api_key("AIza" + "A" * 35))


class TestEncryptDecryptRoundtrip(unittest.TestCase):
    def setUp(self):
        _win32crypt_patcher.start()
        import importlib
        import z7_gemini_key
        importlib.reload(z7_gemini_key)
        self.mod = z7_gemini_key

    def tearDown(self):
        _win32crypt_patcher.stop()

    def test_write_and_read_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "gemini.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                self.mod.write_api_key("my-test-api-key-1234567890")
                self.assertTrue(key_file.exists())
                result = self.mod.read_stored_api_key()
                self.assertEqual(result, "my-test-api-key-1234567890")

    def test_read_stored_returns_empty_when_no_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "missing.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                result = self.mod.read_stored_api_key()
                self.assertEqual(result, "")

    def test_delete_removes_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "gemini.key"
            key_file.write_bytes(b"ENC:dummy")
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                self.mod.delete_api_key()
                self.assertFalse(key_file.exists())

    def test_delete_is_noop_when_no_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "no_file.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                self.mod.delete_api_key()  # deve ser no-op sem exceção

    def test_get_api_key_loads_from_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "gemini.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                valid_key = "AIza" + "B" * 35
                self.mod.write_api_key(valid_key)
                loaded = self.mod.get_api_key(parent=None)
                self.assertEqual(loaded, valid_key)

    def test_get_api_key_prompts_when_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "gemini.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                with mock.patch("z7_theme.ask_string", return_value="new-key-prompt-1234567890"):
                    result = self.mod.get_api_key(parent=None)
                    self.assertEqual(result, "new-key-prompt-1234567890")
                    self.assertTrue(key_file.exists())

    def test_get_api_key_returns_none_when_user_cancels(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "gemini.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                with mock.patch("z7_theme.ask_string", return_value=None):
                    result = self.mod.get_api_key(parent=None)
                    self.assertIsNone(result)

    def test_write_empty_key_is_noop(self):
        with tempfile.TemporaryDirectory() as tmp:
            key_file = Path(tmp) / "gemini.key"
            with mock.patch.object(self.mod, "_get_key_file", return_value=key_file):
                self.mod.write_api_key("  ")
                self.assertFalse(key_file.exists())


if __name__ == "__main__":
    unittest.main()
