import sys
import tempfile
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
    # Limpa caches de lru_cache
    z7_theme.load_theme.cache_clear()
    z7_theme.get_theme_colors.cache_clear()
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


if __name__ == "__main__":
    unittest.main()
