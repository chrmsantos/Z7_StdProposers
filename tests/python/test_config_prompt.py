import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
PY_ROOT = ROOT / "ai"
if str(PY_ROOT) not in sys.path:
    sys.path.insert(0, str(PY_ROOT))


def _reload_config():
    import importlib
    import config_prompt
    importlib.reload(config_prompt)
    return config_prompt


class TestGetPromptFilePath(unittest.TestCase):
    def test_path_uses_data_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_config()
                path = mod.get_prompt_file_path()
                self.assertEqual(path.parent, Path(tmp))
                self.assertEqual(path.name, "gemini_prompt.txt")

    def test_model_file_path_uses_data_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_config()
                path = mod.get_model_file_path()
                self.assertEqual(path.parent, Path(tmp))
                self.assertEqual(path.name, "selected_model.txt")


class TestLoadPrompt(unittest.TestCase):
    def test_returns_default_when_no_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_config()
                prompt = mod.load_prompt()
                self.assertEqual(prompt, mod.DEFAULT_PROMPT)

    def test_returns_custom_prompt_when_file_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "gemini_prompt.txt").write_text("custom prompt", encoding="utf-8")
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                mod = _reload_config()
                self.assertEqual(mod.load_prompt(), "custom prompt")

    def test_falls_back_to_default_on_read_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            prompt_file = tmp_path / "gemini_prompt.txt"
            prompt_file.write_text("irrelevant", encoding="utf-8")
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                mod = _reload_config()
                with mock.patch("builtins.open", side_effect=OSError("disk error")):
                    result = mod.load_prompt()
                    self.assertEqual(result, mod.DEFAULT_PROMPT)


class TestLoadAiModel(unittest.TestCase):
    def test_returns_default_when_no_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_config()
                self.assertEqual(mod.load_ai_model(), "gemini-2.5-flash")

    def test_returns_saved_model(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "selected_model.txt").write_text("gemini-2.5-pro", encoding="utf-8")
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                mod = _reload_config()
                self.assertEqual(mod.load_ai_model(), "gemini-2.5-pro")


class TestSaveAiModel(unittest.TestCase):
    def test_persists_model_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                mod = _reload_config()
                mod.save_ai_model("gemini-2.5-pro")
                content = (tmp_path / "selected_model.txt").read_text(encoding="utf-8")
                self.assertEqual(content, "gemini-2.5-pro")


class TestLoadApiKey(unittest.TestCase):
    def test_delegates_to_z7_gemini_key(self):
        mod = _reload_config()
        with mock.patch("z7_gemini_key.read_stored_api_key", return_value="my-key") as m:
            result = mod.load_api_key()
            m.assert_called_once()
            self.assertEqual(result, "my-key")


class TestRestoreDefault(unittest.TestCase):
    def test_restores_default_text(self):
        import tkinter as tk
        mod = _reload_config()
        # Cria widget Tkinter sem exibir janela (modo headless)
        try:
            root = tk.Tk()
            root.withdraw()
            text_widget = tk.Text(root)
            text_widget.insert("1.0", "old text")
            mod.restore_default(text_widget)
            content = text_widget.get("1.0", tk.END).strip()
            self.assertEqual(content, mod.DEFAULT_PROMPT)
            root.destroy()
        except tk.TclError:
            self.skipTest("Tkinter display not available")


if __name__ == "__main__":
    unittest.main()
