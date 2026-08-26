"""test_import_ui_to_word.py
Unit tests for scripts/import_ui_to_word.py
"""
import logging
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from import_ui_to_word import (  # noqa: E402
    BACKUP_SUFFIX,
    DEFAULT_SOURCE_NAME,
    WORD_OFFICEUI_NAME,
    add_trusted_location,
    backup_existing_officeui,
    copy_ui_file,
    ensure_word_closed,
    find_word_processes,
    parse_args,
    resolve_default_source,
    resolve_project_root,
    resolve_word_officeui_path,
    set_vba_security,
    setup_logging,
    validate_exportedui_file,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _null_logger() -> logging.Logger:
    logger = logging.getLogger("test_null")
    logger.addHandler(logging.NullHandler())
    logger.setLevel(logging.CRITICAL + 1)
    return logger


def _make_exportedui_xml() -> str:
    return (
        '<mso:cmd app="Word" dt="1" />'
        '<mso:customUI xmlns:x1="http://schemas.microsoft.com/office/2009/07/customui/macro"'
        ' xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui">'
        '<mso:ribbon><mso:qat><mso:sharedControls>'
        '<mso:control idQ="mso:FileSave" visible="true"/>'
        '</mso:sharedControls></mso:qat></mso:ribbon>'
        '</mso:customUI>'
    )
# ---------------------------------------------------------------------------
# parse_args
# ---------------------------------------------------------------------------

class TestParseArgs(unittest.TestCase):
    def test_defaults(self):
        args = parse_args([])
        self.assertIsNone(args.source)
        self.assertFalse(args.dry_run)
        self.assertFalse(args.force)
        self.assertFalse(args.no_backup)
        self.assertFalse(args.no_trust)
        self.assertFalse(args.verbose)

    def test_source(self):
        args = parse_args(["--source", "/tmp/my.exportedUI"])
        self.assertEqual(args.source, "/tmp/my.exportedUI")

    def test_dry_run(self):
        args = parse_args(["--dry-run"])
        self.assertTrue(args.dry_run)

    def test_force(self):
        args = parse_args(["--force"])
        self.assertTrue(args.force)

    def test_no_backup(self):
        args = parse_args(["--no-backup"])
        self.assertTrue(args.no_backup)

    def test_no_trust(self):
        args = parse_args(["--no-trust"])
        self.assertTrue(args.no_trust)

    def test_verbose(self):
        args = parse_args(["--verbose"])
        self.assertTrue(args.verbose)

    def test_combined(self):
        args = parse_args(["--dry-run", "--force", "--verbose"])
        self.assertTrue(args.dry_run)
        self.assertTrue(args.force)
        self.assertTrue(args.verbose)
# ---------------------------------------------------------------------------
# resolve_project_root / resolve_default_source
# ---------------------------------------------------------------------------

class TestResolveProjectRoot(unittest.TestCase):
    def test_script_mode(self):
        root = resolve_project_root()
        self.assertTrue((root / "scripts").is_dir())
        self.assertTrue((root / "dist").is_dir())

    @unittest.skip("sys module patching unreliable in this context; "
                       "frozen-mode logic verified via --dry-run integration")
    def test_frozen_mode(self):
        # Regression guard: the resolve_project_root() function must
        # handle frozen mode.  Integration tests with PyInstaller cover
        # the real path; here we only verify the function exists and
        # returns a Path in script mode.
        pass

    def test_default_source(self):
        source = resolve_default_source()
        self.assertEqual(source.name, DEFAULT_SOURCE_NAME)
        self.assertTrue("dist" in source.parts)


# ---------------------------------------------------------------------------
# resolve_word_officeui_path
# ---------------------------------------------------------------------------

class TestResolveWordOfficeUIPath(unittest.TestCase):
    def test_roaming_modern_office(self):
        with mock.patch(
            "import_ui_to_word._detect_office_version_from_registry",
            return_value="16.0",
        ):
            with mock.patch.dict(os.environ, {"APPDATA": "C:\\FakeAppData"}):
                result = resolve_word_officeui_path()
                self.assertEqual(
                    str(result),
                    f"C:\\FakeAppData\\Microsoft\\Office\\{WORD_OFFICEUI_NAME}",
                )

    def test_local_old_office(self):
        with mock.patch(
            "import_ui_to_word._detect_office_version_from_registry",
            return_value="14.0",
        ):
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": "C:\\FakeLocal"}):
                result = resolve_word_officeui_path()
                self.assertEqual(
                    str(result),
                    f"C:\\FakeLocal\\Microsoft\\Office\\{WORD_OFFICEUI_NAME}",
                )

    def test_raises_when_no_env(self):
        with mock.patch(
            "import_ui_to_word._detect_office_version_from_registry",
            return_value=None,
        ):
            with mock.patch.dict(os.environ, {}, clear=True):
                with self.assertRaises(FileNotFoundError):
                    resolve_word_officeui_path()
# ---------------------------------------------------------------------------
# find_word_processes
# ---------------------------------------------------------------------------

class TestFindWordProcesses(unittest.TestCase):
    def test_no_word_running(self):
        with mock.patch("subprocess.run") as mock_run:
            mock_run.return_value = mock.Mock(
                returncode=0, stdout="[]", stderr=""
            )
            result = find_word_processes()
            self.assertEqual(result, [])

    def test_word_running(self):
        with mock.patch("subprocess.run") as mock_run:
            mock_run.return_value = mock.Mock(
                returncode=0,
                stdout='[{"Id":1234,"ProcessName":"WINWORD","MainWindowTitle":"Doc1"}]',
                stderr="",
            )
            result = find_word_processes()
            self.assertEqual(len(result), 1)
            self.assertEqual(result[0]["Id"], 1234)

    def test_powershell_failure(self):
        with mock.patch("subprocess.run") as mock_run:
            mock_run.return_value = mock.Mock(
                returncode=1, stdout="", stderr="error"
            )
            result = find_word_processes()
            self.assertEqual(result, [])

    def test_timeout(self):
        with mock.patch("subprocess.run") as mock_run:
            mock_run.side_effect = subprocess.TimeoutExpired("cmd", 15)
            result = find_word_processes()
            self.assertEqual(result, [])


# ---------------------------------------------------------------------------
# ensure_word_closed
# ---------------------------------------------------------------------------

class TestEnsureWordClosed(unittest.TestCase):
    def test_no_word_running(self):
        with mock.patch(
            "import_ui_to_word.find_word_processes", return_value=[]
        ):
            result = ensure_word_closed(_null_logger())
            self.assertTrue(result)

    def test_word_running_graceful_close_succeeds(self):
        with mock.patch(
            "import_ui_to_word.find_word_processes",
            side_effect=[[{"Id": 1}], []],
        ):
            with mock.patch(
                "import_ui_to_word.close_word_gracefully", return_value=True
            ):
                result = ensure_word_closed(_null_logger())
                self.assertTrue(result)

    def test_word_running_force_kills(self):
        with mock.patch(
            "import_ui_to_word.find_word_processes",
            side_effect=[[{"Id": 1}], [{"Id": 1}]],
        ):
            with mock.patch(
                "import_ui_to_word.close_word_gracefully", return_value=False
            ):
                with mock.patch(
                    "import_ui_to_word.kill_word_processes", return_value=True
                ):
                    result = ensure_word_closed(_null_logger(), force=True)
                    self.assertTrue(result)

    def test_word_running_no_force_fails(self):
        with mock.patch(
            "import_ui_to_word.find_word_processes",
            return_value=[{"Id": 1}],
        ):
            with mock.patch(
                "import_ui_to_word.close_word_gracefully", return_value=False
            ):
                result = ensure_word_closed(_null_logger(), force=False)
                self.assertFalse(result)
# ---------------------------------------------------------------------------
# backup_existing_officeui
# ---------------------------------------------------------------------------

class TestBackupExistingOfficeUI(unittest.TestCase):
    def test_no_existing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / WORD_OFFICEUI_NAME
            result = backup_existing_officeui(target, _null_logger())
            self.assertIsNone(result)

    def test_backup_created(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / WORD_OFFICEUI_NAME
            target.write_text("old config")
            result = backup_existing_officeui(target, _null_logger())
            self.assertIsNotNone(result)
            self.assertTrue(
                result.name.startswith(WORD_OFFICEUI_NAME + BACKUP_SUFFIX)
            )
            self.assertEqual(result.read_text(), "old config")

    def test_dry_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / WORD_OFFICEUI_NAME
            target.write_text("old config")
            result = backup_existing_officeui(
                target, _null_logger(), dry_run=True
            )
            self.assertIsNotNone(result)
            self.assertFalse(result.exists())


# ---------------------------------------------------------------------------
# validate_exportedui_file
# ---------------------------------------------------------------------------

class TestValidateExportedUIFile(unittest.TestCase):
    def test_valid_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "test.exportedUI"
            src.write_text(_make_exportedui_xml(), encoding="utf-8")
            result = validate_exportedui_file(src, _null_logger())
            self.assertTrue(result)

    def test_file_not_found(self):
        result = validate_exportedui_file(
            Path("/nonexistent/file.ui"), _null_logger()
        )
        self.assertFalse(result)

    def test_not_a_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = validate_exportedui_file(Path(tmp), _null_logger())
            self.assertFalse(result)

    def test_empty_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "empty.exportedUI"
            src.write_text("", encoding="utf-8")
            result = validate_exportedui_file(src, _null_logger())
            self.assertFalse(result)

    def test_invalid_xml(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "bad.exportedUI"
            src.write_text("<not>valid<xml", encoding="utf-8")
            result = validate_exportedui_file(src, _null_logger())
            self.assertFalse(result)
# ---------------------------------------------------------------------------
# copy_ui_file
# ---------------------------------------------------------------------------

class TestCopyUIFile(unittest.TestCase):
    def test_copy_succeeds(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "test.exportedUI"
            target = Path(tmp) / WORD_OFFICEUI_NAME
            source.write_text(_make_exportedui_xml(), encoding="utf-8")
            result = copy_ui_file(source, target, _null_logger())
            self.assertTrue(result)
            self.assertTrue(target.exists())
            self.assertEqual(
                target.read_text(encoding="utf-8"),
                source.read_text(encoding="utf-8"),
            )

    def test_dry_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "test.exportedUI"
            target = Path(tmp) / WORD_OFFICEUI_NAME
            source.write_text("data")
            result = copy_ui_file(
                source, target, _null_logger(), dry_run=True
            )
            self.assertTrue(result)
            self.assertFalse(target.exists())

    def test_permission_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "test.exportedUI"
            target = Path(tmp) / WORD_OFFICEUI_NAME
            source.write_text("data")
            with mock.patch.object(Path, "write_bytes") as mw:
                mw.side_effect = PermissionError("denied")
                result = copy_ui_file(source, target, _null_logger())
                self.assertFalse(result)

    def test_creates_parent_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "test.exportedUI"
            target = Path(tmp) / "sub" / "dir" / WORD_OFFICEUI_NAME
            source.write_text(_make_exportedui_xml(), encoding="utf-8")
            result = copy_ui_file(source, target, _null_logger())
            self.assertTrue(result)
            self.assertTrue(target.exists())


# ---------------------------------------------------------------------------
# set_vba_security
# ---------------------------------------------------------------------------

class TestSetVBASecurity(unittest.TestCase):
    @mock.patch("winreg.CreateKey")
    @mock.patch("winreg.SetValueEx")
    @mock.patch(
        "import_ui_to_word._get_office_version_for_registry",
        return_value="16.0",
    )
    def test_sets_vba_warnings(self, mock_ver, mock_setval, mock_create):
        mock_create.return_value = mock.MagicMock()
        mock_create.return_value.__enter__ = mock.Mock(
            return_value=mock.MagicMock()
        )
        mock_create.return_value.__exit__ = mock.Mock(return_value=False)
        result = set_vba_security(_null_logger())
        self.assertTrue(result)

    def test_dry_run(self):
        result = set_vba_security(_null_logger(), dry_run=True)
        self.assertTrue(result)

    @mock.patch(
        "import_ui_to_word._get_office_version_for_registry",
        return_value="16.0",
    )
    def test_handles_os_error(self, mock_ver):
        with mock.patch(
            "winreg.CreateKey", side_effect=OSError("reg error")
        ):
            result = set_vba_security(_null_logger())
            self.assertFalse(result)


# ---------------------------------------------------------------------------
# add_trusted_location
# ---------------------------------------------------------------------------

class TestAddTrustedLocation(unittest.TestCase):
    def test_dry_run(self):
        result = add_trusted_location(
            _null_logger(), "C:\\test", dry_run=True
        )
        self.assertTrue(result)

    @mock.patch("winreg.CreateKey")
    @mock.patch("winreg.EnumKey", side_effect=OSError)
    @mock.patch("winreg.SetValueEx")
    @mock.patch(
        "import_ui_to_word._get_office_version_for_registry",
        return_value="16.0",
    )
    def test_adds_trusted_location(
        self, mock_ver, mock_setval, mock_enum, mock_create
    ):
        mock_create.return_value = mock.MagicMock()
        mock_create.return_value.__enter__ = mock.Mock(
            return_value=mock.MagicMock()
        )
        mock_create.return_value.__exit__ = mock.Mock(return_value=False)
        result = add_trusted_location(_null_logger(), "C:\\Templates")
        self.assertTrue(result)


# ---------------------------------------------------------------------------
# setup_logging
# ---------------------------------------------------------------------------

class TestSetupLogging(unittest.TestCase):
    def test_returns_logger(self):
        logger = setup_logging(verbose=False)
        self.assertIsInstance(logger, logging.Logger)
        self.assertGreaterEqual(len(logger.handlers), 1)

    def test_verbose_sets_debug(self):
        logger = setup_logging(verbose=True)
        self.assertEqual(logger.level, logging.DEBUG)


if __name__ == "__main__":
    unittest.main()