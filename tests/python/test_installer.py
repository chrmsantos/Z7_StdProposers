"""Testes unitários para scripts/installer.py.

Cobre:
- Leitura de versão (_read_version)
- Helpers de retry (_retry_copytree, _retry_copy2, _retry_rmtree)
- Chamada à API do GitHub (get_latest_github_release)
- Tema do instalador (InstallerTheme)
- Constantes e versão (_APP_VERSION)
"""
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

# Injeta os diretórios necessários no path
ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = ROOT / "scripts"
AI_DIR = ROOT / "ai"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(AI_DIR) not in sys.path:
    sys.path.insert(0, str(AI_DIR))

import tkinter as tk  # noqa: E402
import installer  # noqa: E402


# ── _read_version ────────────────────────────────────────────────────────────

class TestReadVersion(unittest.TestCase):
    def test_read_version_from_project_root(self):
        version = installer._read_version()
        self.assertRegex(version, r"^\d+\.\d+\.\d+$")

    def test_read_version_returns_string(self):
        self.assertIsInstance(installer._read_version(), str)


# ── _retry_copytree ─────────────────────────────────────────────────────────

class TestRetryCopytree(unittest.TestCase):
    def test_success_on_first_attempt(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            dst = Path(tmp) / "dst"
            src.mkdir()
            (src / "file.txt").write_text("hello")
            installer._retry_copytree(src, dst)
            self.assertTrue((dst / "file.txt").exists())

    def test_success_with_dirs_exist_ok(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            dst = Path(tmp) / "dst"
            src.mkdir(); dst.mkdir()
            (src / "a.txt").write_text("a")
            installer._retry_copytree(src, dst, dirs_exist_ok=True)
            self.assertTrue((dst / "a.txt").exists())

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.copytree")
    def test_retries_on_winerror32(self, mock_copytree, mock_sleep):
        lock_err = shutil.Error([
            ("src.dll", "dst.dll", "[WinError 32] O arquivo já está sendo usado por outro processo")
        ])
        mock_copytree.side_effect = [lock_err, None]
        installer._retry_copytree(Path("src"), Path("dst"))
        self.assertEqual(mock_copytree.call_count, 2)
        mock_sleep.assert_called_once()

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.copytree")
    def test_raises_after_max_retries(self, mock_copytree, mock_sleep):
        lock_err = shutil.Error([
            ("src.dll", "dst.dll", "[WinError 32] being used by another process")
        ])
        mock_copytree.side_effect = lock_err
        with self.assertRaises(shutil.Error):
            installer._retry_copytree(Path("src"), Path("dst"))
        self.assertEqual(mock_copytree.call_count, installer._RETRY_ATTEMPTS)

    @mock.patch("installer.shutil.copytree")
    def test_raises_non_lock_errors_immediately(self, mock_copytree):
        other_err = shutil.Error([("src", "dst", "disk full")])
        mock_copytree.side_effect = other_err
        with self.assertRaises(shutil.Error):
            installer._retry_copytree(Path("src"), Path("dst"))
        self.assertEqual(mock_copytree.call_count, 1)


# ── _retry_copy2 ────────────────────────────────────────────────────────────

class TestRetryCopy2(unittest.TestCase):
    def test_success_on_first_attempt(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src.txt"
            dst = Path(tmp) / "dst.txt"
            src.write_text("content")
            installer._retry_copy2(src, dst)
            self.assertEqual(dst.read_text(), "content")

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.copy2")
    def test_retries_on_winerror32(self, mock_copy2, mock_sleep):
        err = PermissionError("[WinError 32] O arquivo já está sendo usado por outro processo")
        mock_copy2.side_effect = [err, None]
        installer._retry_copy2(Path("src"), Path("dst"))
        self.assertEqual(mock_copy2.call_count, 2)

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.copy2")
    def test_raises_after_max_retries(self, mock_copy2, mock_sleep):
        err = PermissionError("[WinError 32] being used by another process")
        mock_copy2.side_effect = err
        with self.assertRaises(PermissionError):
            installer._retry_copy2(Path("src"), Path("dst"))
        self.assertEqual(mock_copy2.call_count, installer._RETRY_ATTEMPTS)

    @mock.patch("installer.shutil.copy2")
    def test_raises_non_lock_permission_error_immediately(self, mock_copy2):
        err = PermissionError("access denied")
        mock_copy2.side_effect = err
        with self.assertRaises(PermissionError):
            installer._retry_copy2(Path("src"), Path("dst"))
        self.assertEqual(mock_copy2.call_count, 1)


# ── _retry_rmtree ───────────────────────────────────────────────────────────

class TestRetryRmtree(unittest.TestCase):
    def test_success_returns_true(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "to_delete"
            target.mkdir()
            (target / "file.txt").write_text("bye")
            result = installer._retry_rmtree(target)
            self.assertTrue(result)
            self.assertFalse(target.exists())

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.rmtree")
    def test_retries_on_permission_error(self, mock_rmtree, mock_sleep):
        mock_rmtree.side_effect = [PermissionError("WinError 32"), None]
        result = installer._retry_rmtree(Path("some_dir"))
        self.assertTrue(result)
        self.assertEqual(mock_rmtree.call_count, 2)

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.rmtree")
    def test_returns_false_after_max_retries(self, mock_rmtree, mock_sleep):
        mock_rmtree.side_effect = PermissionError("WinError 32")
        result = installer._retry_rmtree(Path("some_dir"))
        self.assertFalse(result)
        self.assertEqual(mock_rmtree.call_count, installer._RETRY_ATTEMPTS)


# ── get_latest_github_release ────────────────────────────────────────────────

class TestGetLatestGithubRelease(unittest.TestCase):
    @mock.patch("urllib.request.urlopen")
    @mock.patch("urllib.request.Request")
    def test_calls_api_correctly(self, mock_request_cls, mock_urlopen):
        mock_response = mock.Mock()
        mock_response.read.return_value = b'{"tag_name": "v1.2.3", "assets": []}'
        mock_urlopen.return_value.__enter__.return_value = mock_response
        res = installer.get_latest_github_release()
        self.assertEqual(res["tag_name"], "v1.2.3")
        mock_request_cls.assert_called_once_with(
            "https://api.github.com/repos/chrmsantos/Z7_StdProposers/releases/latest",
            headers={"User-Agent": f"Z7_StdProposers/{installer._APP_VERSION}"}
        )

    @mock.patch("urllib.request.urlopen")
    def test_raises_on_network_error(self, mock_urlopen):
        import urllib.error
        mock_urlopen.side_effect = urllib.error.URLError("timeout")
        with self.assertRaises(urllib.error.URLError):
            installer.get_latest_github_release()

    @mock.patch("urllib.request.urlopen")
    def test_returns_parsed_json(self, mock_urlopen):
        data = {"tag_name": "v9.0.0", "assets": [{"name": "test.zip"}]}
        mock_response = mock.Mock()
        mock_response.read.return_value = json.dumps(data).encode("utf-8")
        mock_urlopen.return_value.__enter__.return_value = mock_response
        result = installer.get_latest_github_release()
        self.assertEqual(result["tag_name"], "v9.0.0")
        self.assertEqual(len(result["assets"]), 1)


# ── InstallerTheme ──────────────────────────────────────────────────────────

class TestInstallerTheme(unittest.TestCase):
    def setUp(self):
        self.root = mock.MagicMock()
        self.theme = installer.InstallerTheme(self.root)

    def test_init_sets_mode_and_widgets(self):
        self.assertIn(self.theme.mode, ["light", "dark"])
        self.assertEqual(self.theme.widgets, {})

    @mock.patch("z7_theme.save_theme")
    def test_toggle_switches_mode_and_saves(self, mock_save):
        initial_mode = self.theme.mode
        self.theme.toggle()
        self.assertNotEqual(initial_mode, self.theme.mode)
        mock_save.assert_called_once_with(self.theme.mode)

    def test_apply_configures_root_and_widgets(self):
        mock_widget1 = mock.Mock(spec=tk.Label)
        mock_widget2 = mock.Mock(spec=tk.Button)
        self.theme.widgets["title_lbl"] = mock_widget1
        self.theme.widgets["install_btn"] = mock_widget2
        self.theme.apply()
        self.assertTrue(mock_widget1.configure.called)
        self.assertTrue(mock_widget2.configure.called)

    def test_toggle_twice_returns_to_original_mode(self):
        original = self.theme.mode
        self.theme.toggle()
        self.theme.toggle()
        self.assertEqual(self.theme.mode, original)


# ── _APP_VERSION ─────────────────────────────────────────────────────────────

class TestAppVersion(unittest.TestCase):
    def test_app_version_matches_version_file(self):
        version_file = ROOT / "VERSION"
        if version_file.exists():
            expected = version_file.read_text(encoding="utf-8").strip()
            self.assertEqual(installer._APP_VERSION, expected)

    def test_app_version_format(self):
        self.assertRegex(installer._APP_VERSION, r"^\d+\.\d+\.\d+$")


# ── Constants ────────────────────────────────────────────────────────────────

class TestConstants(unittest.TestCase):
    def test_retry_attempts_is_positive(self):
        self.assertGreater(installer._RETRY_ATTEMPTS, 0)

    def test_retry_base_delay_is_positive(self):
        self.assertGreater(installer._RETRY_BASE_DELAY, 0)

    def test_github_repo_url_is_set(self):
        self.assertIn("github.com", installer.GITHUB_REPO_URL)
        self.assertIn("Z7_StdProposers", installer.GITHUB_REPO_URL)


# ── InstallerMainFlow ────────────────────────────────────────────────────────

class TestInstallerMainFlow(unittest.TestCase):
    @mock.patch("tkinter.Tk")
    @mock.patch("threading.Thread")
    @mock.patch("installer.get_latest_github_release")
    def test_main_initializes_gui(self, mock_get_release, mock_thread, mock_tk):
        mock_root = mock.MagicMock()
        mock_tk.return_value = mock_root
        installer.main()


if __name__ == "__main__":
    unittest.main()
