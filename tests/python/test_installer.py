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
import subprocess
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
            self.assertEqual((dst / "file.txt").read_text(), "hello")

    def test_success_with_dirs_exist_ok(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            dst = Path(tmp) / "dst"
            src.mkdir(); dst.mkdir()
            (src / "a.txt").write_text("a")
            installer._retry_copytree(src, dst, dirs_exist_ok=True)
            self.assertTrue((dst / "a.txt").exists())

    def test_copies_subdirectories(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            dst = Path(tmp) / "dst"
            (src / "sub").mkdir(parents=True)
            (src / "sub" / "deep.txt").write_text("deep")
            installer._retry_copytree(src, dst)
            self.assertTrue((dst / "sub" / "deep.txt").exists())

    @mock.patch("installer.time.sleep")
    @mock.patch("installer._copy_single_file")
    @mock.patch("installer.kill_process_by_name")
    def test_raises_after_kill_still_blocked(self, mock_kill, mock_copy, mock_sleep):
        """Deve matar processos e ainda lançar se arquivos persistirem bloqueados."""
        mock_copy.return_value = False
        mock_kill.return_value = 1

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            dst = Path(tmp) / "dst"
            src.mkdir()
            (src / "blocked.dll").write_text("x")
            with self.assertRaises(shutil.Error):
                installer._retry_copytree(src, dst)

    @mock.patch("installer.time.sleep")
    @mock.patch("installer._copy_single_file")
    @mock.patch("installer.kill_process_by_name")
    def test_succeeds_after_kill(self, mock_kill, mock_copy, mock_sleep):
        """Deve matar processos e então copiar com sucesso (sem levantar exceção)."""
        mock_copy.side_effect = [False, True]  # falha 1a, sucesso 2a
        mock_kill.return_value = 1

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            dst = Path(tmp) / "dst"
            src.mkdir()
            (src / "blocked.dll").write_text("x")
            installer._retry_copytree(src, dst)
            self.assertTrue(mock_kill.called)
            self.assertEqual(mock_copy.call_count, 2)


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


# ── kill_process_by_name ────────────────────────────────────────────────────

class TestKillProcessByName(unittest.TestCase):
    @mock.patch("installer.subprocess.run")
    def test_returns_count_on_success(self, mock_run):
        mock_run.return_value = mock.Mock(returncode=0, stdout="Sucesso: processo encerrado.", stderr="")
        result = installer.kill_process_by_name("chat_ia.exe")
        self.assertGreaterEqual(result, 1)
        mock_run.assert_called_once_with(
            ["taskkill", "/F", "/IM", "chat_ia.exe"],
            capture_output=True, text=True, timeout=15,
        )

    @mock.patch("installer.subprocess.run")
    def test_returns_zero_when_not_found(self, mock_run):
        mock_run.return_value = mock.Mock(returncode=128, stdout="", stderr="não foi encontrado")
        result = installer.kill_process_by_name("ghost.exe")
        self.assertEqual(result, 0)

    @mock.patch("installer.subprocess.run")
    def test_returns_zero_on_timeout(self, mock_run):
        mock_run.side_effect = subprocess.TimeoutExpired(cmd="taskkill", timeout=15)
        result = installer.kill_process_by_name("stuck.exe")
        self.assertEqual(result, 0)


# ── _copy_single_file ───────────────────────────────────────────────────────

class TestCopySingleFile(unittest.TestCase):
    def test_success_on_first_attempt(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src.txt"
            dst = Path(tmp) / "sub" / "dst.txt"
            src.write_text("content")
            result = installer._copy_single_file(src, dst)
            self.assertTrue(result)
            self.assertEqual(dst.read_text(), "content")

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.copy2")
    def test_retries_on_winerror32(self, mock_copy2, mock_sleep):
        err = PermissionError("[WinError 32] file in use")
        mock_copy2.side_effect = [err, None]
        result = installer._copy_single_file(Path("src"), Path("dst"))
        self.assertTrue(result)
        self.assertEqual(mock_copy2.call_count, 2)

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.copy2")
    def test_returns_false_when_all_retries_exhausted(self, mock_copy2, mock_sleep):
        mock_copy2.side_effect = PermissionError("[WinError 32] file in use")
        result = installer._copy_single_file(Path("src"), Path("dst"))
        self.assertFalse(result)
        self.assertEqual(mock_copy2.call_count, installer._RETRY_ATTEMPTS)

    @mock.patch("installer.shutil.copy2")
    def test_raises_non_lock_errors(self, mock_copy2):
        mock_copy2.side_effect = PermissionError("access denied")
        with self.assertRaises(PermissionError):
            installer._copy_single_file(Path("src"), Path("dst"))


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
    @mock.patch("installer.kill_process_by_name")
    def test_returns_false_after_kill_and_still_blocked(self, mock_kill, mock_rmtree, mock_sleep):
        """Após esgotar tentativas, mata processos e ainda falha => False."""
        mock_rmtree.side_effect = PermissionError("WinError 32")  # sempre falha
        mock_kill.return_value = 1

        result = installer._retry_rmtree(Path("some_dir"))
        self.assertFalse(result)
        # rmtree chamado: _RETRY_ATTEMPTS (retry loop) + 1 (tentativa pós-kill)
        self.assertEqual(mock_rmtree.call_count, installer._RETRY_ATTEMPTS + 1)

    @mock.patch("installer.time.sleep")
    @mock.patch("installer.shutil.rmtree")
    @mock.patch("installer.kill_process_by_name")
    def test_succeeds_after_kill(self, mock_kill, mock_rmtree, mock_sleep):
        """Após esgotar tentativas, mata processos e then rmtree succeeds."""
        # Falha nas primeiras _RETRY_ATTEMPTS, sucesso na pós-kill
        mock_rmtree.side_effect = [PermissionError("WinError 32")] * installer._RETRY_ATTEMPTS + [None]
        mock_kill.return_value = 1

        result = installer._retry_rmtree(Path("some_dir"))
        self.assertTrue(result)
        self.assertEqual(mock_rmtree.call_count, installer._RETRY_ATTEMPTS + 1)


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
