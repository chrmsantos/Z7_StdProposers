import sys
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


class TestGetLatestGithubRelease(unittest.TestCase):
    @mock.patch("urllib.request.urlopen")
    @mock.patch("urllib.request.Request")
    def test_get_latest_github_release_calls_api_correctly(self, mock_request_cls, mock_urlopen):
        mock_response = mock.Mock()
        mock_response.read.return_value = b'{"tag_name": "v1.2.3", "assets": []}'
        mock_urlopen.return_value.__enter__.return_value = mock_response

        res = installer.get_latest_github_release()

        self.assertEqual(res["tag_name"], "v1.2.3")
        mock_request_cls.assert_called_once_with(
            "https://api.github.com/repos/chrmsantos/Z7_StdProposers/releases/latest",
            headers={"User-Agent": f"Z7_StdProposers/{installer._APP_VERSION}"}
        )


class TestInstallerTheme(unittest.TestCase):
    def setUp(self):
        self.root = mock.MagicMock()
        self.theme = installer.InstallerTheme(self.root)

    def tearDown(self):
        pass

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
        
        # O apply deve configurar os estilos nos widgets cadastrados
        self.assertTrue(mock_widget1.configure.called)
        self.assertTrue(mock_widget2.configure.called)


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
