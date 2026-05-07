import json
import os
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
PY_ROOT = ROOT / "ai"
if str(PY_ROOT) not in sys.path:
    sys.path.insert(0, str(PY_ROOT))


def _reload_theme():
    import importlib
    import z7_theme
    # Limpa caches de lru_cache e cache global de prefs
    z7_theme.load_theme.cache_clear()
    z7_theme.get_theme_colors.cache_clear()
    z7_theme._privacy_prefs_cache = None
    importlib.reload(z7_theme)
    return z7_theme


class TestLoadSaveTheme(unittest.TestCase):
    def test_load_theme_defaults_to_light(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                self.assertEqual(mod.load_theme(), "light")

    def test_save_and_load_dark_theme(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                mod.save_theme("dark")
                # Após save_theme os caches são limpos; próximo load deve retornar dark
                self.assertEqual(mod.load_theme(), "dark")

    def test_corrupted_theme_file_falls_back_to_light(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            theme_file = tmp_path / "theme_config.json"
            theme_file.write_text("{ INVALID JSON }", encoding="utf-8")
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                mod = _reload_theme()
                self.assertEqual(mod.load_theme(), "light")


class TestGetThemeColors(unittest.TestCase):
    def test_light_colors_have_required_keys(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                colors = mod.get_theme_colors("light")
                for key in ("bg", "fg", "fg_muted", "text_bg", "btn_primary_bg"):
                    self.assertIn(key, colors)

    def test_dark_colors_have_required_keys(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                colors = mod.get_theme_colors("dark")
                for key in ("bg", "fg", "fg_muted", "text_bg", "btn_primary_bg"):
                    self.assertIn(key, colors)

    def test_light_and_dark_are_different(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                self.assertNotEqual(
                    mod.get_theme_colors("light")["bg"],
                    mod.get_theme_colors("dark")["bg"],
                )


class TestPrivacyPrefs(unittest.TestCase):
    def test_load_returns_empty_when_no_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                prefs = mod.load_privacy_prefs()
                self.assertIsInstance(prefs, dict)
                self.assertEqual(len(prefs), 0)

    def test_save_and_load_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                mod.save_privacy_prefs({"chat_ia": True})
                # Invalida cache manualmente para forçar releitura
                mod._privacy_prefs_cache = None
                prefs = mod.load_privacy_prefs()
                self.assertTrue(prefs.get("chat_ia"))

    def test_corrupted_prefs_file_returns_empty(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "privacy_prefs.json").write_text("{ BAD", encoding="utf-8")
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                mod = _reload_theme()
                prefs = mod.load_privacy_prefs()
                self.assertEqual(prefs, {})

    def test_privacy_prefs_thread_safety(self):
        """Múltiplas threads não devem corromper o cache de prefs."""
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                errors = []

                def worker():
                    try:
                        for _ in range(50):
                            mod._privacy_prefs_cache = None
                            mod.load_privacy_prefs()
                    except Exception as e:
                        errors.append(e)

                threads = [threading.Thread(target=worker) for _ in range(8)]
                for t in threads:
                    t.start()
                for t in threads:
                    t.join()
                self.assertEqual(errors, [], msg=f"Thread errors: {errors}")


class TestAskPrivacyWarning(unittest.TestCase):
    def test_returns_true_when_already_accepted(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                mod = _reload_theme()
                mod.save_privacy_prefs({"my_component": True})
                mod._privacy_prefs_cache = None
                result = mod.ask_privacy_warning("T", "M", key="my_component", parent=None)
                self.assertTrue(result)


if __name__ == "__main__":
    unittest.main()
