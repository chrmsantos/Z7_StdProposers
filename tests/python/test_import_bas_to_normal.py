"""test_import_bas_to_normal.py
Unit tests for scripts/import_bas_to_normal.py
"""
import logging
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from import_bas_to_normal import (  # noqa: E402
    BACKUP_NAME,
    Z7_MODULE_PREFIX,
    _get_logs_dir,
    create_backup,
    discover_bas_files,
    parse_args,
    resolve_normal_dotm,
    resolve_project_root,
    setup_logging,
)


class TestParseArgs(unittest.TestCase):
    def test_defaults(self):
        args = parse_args([])
        self.assertFalse(args.dry_run)
        self.assertFalse(args.no_backup)
        self.assertFalse(args.verbose)

    def test_dry_run(self):
        args = parse_args(["--dry-run"])
        self.assertTrue(args.dry_run)

    def test_no_backup(self):
        args = parse_args(["--no-backup"])
        self.assertTrue(args.no_backup)

    def test_verbose(self):
        args = parse_args(["--verbose"])
        self.assertTrue(args.verbose)

    def test_combined(self):
        args = parse_args(["--dry-run", "--verbose"])
        self.assertTrue(args.dry_run)
        self.assertTrue(args.verbose)


class TestResolveProjectRoot(unittest.TestCase):
    def test_script_mode(self):
        root = resolve_project_root()
        self.assertTrue((root / "source" / "main").is_dir())
        self.assertTrue((root / "scripts").is_dir())

    def test_frozen_mode(self):
        fake_exe = Path(tempfile.mkdtemp()) / "dist" / "import_bas_to_normal.exe"
        fake_exe.parent.mkdir(parents=True, exist_ok=True)
        fake_exe.touch()
        try:
            with mock.patch("import_bas_to_normal.sys") as mock_sys:
                mock_sys.frozen = True
                mock_sys.executable = str(fake_exe)
                result = resolve_project_root()
                self.assertEqual(result, fake_exe.parent.parent)
        finally:
            shutil.rmtree(fake_exe.parent.parent, ignore_errors=True)

# PLACEHOLDER_TEST2


class TestResolveNormalDotm(unittest.TestCase):
    def test_uses_appdata(self):
        with mock.patch.dict(os.environ, {"APPDATA": "C:\\FakeAppData"}):
            result = resolve_normal_dotm()
            self.assertEqual(str(result), "C:\\FakeAppData\\Microsoft\\Templates\\Normal.dotm")

    def test_falls_back_to_userprofile(self):
        env = {"USERPROFILE": "C:\\Users\\test", "APPDATA": ""}
        with mock.patch.dict(os.environ, env, clear=False):
            os.environ.pop("APPDATA", None)
            result = resolve_normal_dotm()
            self.assertIn("Microsoft", str(result))
            self.assertTrue(str(result).endswith("Normal.dotm"))

    def test_raises_when_no_env(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(FileNotFoundError):
                resolve_normal_dotm()


class TestDiscoverBasFiles(unittest.TestCase):
    def test_finds_bas_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            (d / "Mod_01.bas").write_text("Attribute VB_Name = 'Mod_01'")
            (d / "Mod_02.bas").write_text("Attribute VB_Name = 'Mod_02'")
            (d / "readme.txt").write_text("not a bas")
            result = discover_bas_files(d)
            self.assertEqual(len(result), 2)
            self.assertTrue(all(f.endswith(".bas") for f in result))

    def test_empty_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = discover_bas_files(Path(tmp))
            self.assertEqual(result, [])

    def test_nonexistent_dir(self):
        with self.assertRaises(FileNotFoundError):
            discover_bas_files(Path("/nonexistent/path"))

# PLACEHOLDER_TEST3


class TestCreateBackup(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.normal = Path(self.tmp) / "Normal.dotm"
        self.normal.write_bytes(b"fake dotm content " + b"x" * 100)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_creates_backup(self):
        logger = logging.getLogger("test")
        backup = create_backup(self.normal, logger)
        self.assertTrue(backup.exists())
        self.assertEqual(backup.name, BACKUP_NAME)
        self.assertEqual(backup.read_bytes(), self.normal.read_bytes())

    def test_rotates_existing_backup(self):
        logger = logging.getLogger("test")
        create_backup(self.normal, logger)
        self.normal.write_bytes(b"modified content " + b"y" * 100)
        create_backup(self.normal, logger)
        backups = list(Path(self.tmp).glob("Normal_backup_*.dotm"))
        self.assertEqual(len(backups), 1)
        current = Path(self.tmp) / BACKUP_NAME
        self.assertEqual(current.read_bytes(), self.normal.read_bytes())

    def test_integrity_check(self):
        logger = logging.getLogger("test")
        backup = create_backup(self.normal, logger)
        self.assertEqual(backup.stat().st_size, self.normal.stat().st_size)

    def test_raises_when_normal_missing(self):
        logger = logging.getLogger("test")
        missing = Path(self.tmp) / "missing.dotm"
        with self.assertRaises(FileNotFoundError):
            create_backup(missing, logger)


class TestSetupLogging(unittest.TestCase):
    def test_creates_logger(self):
        logger = setup_logging(verbose=False)
        self.assertIsInstance(logger, logging.Logger)
        self.assertEqual(logger.name, "import_bas")
        self.assertGreaterEqual(len(logger.handlers), 1)

    def test_verbose_sets_debug(self):
        logger = setup_logging(verbose=True)
        self.assertEqual(logger.level, logging.DEBUG)

    def test_console_handler_exists(self):
        logger = setup_logging()
        handler_types = [type(h).__name__ for h in logger.handlers]
        self.assertIn("StreamHandler", handler_types)


class TestGetLogsDir(unittest.TestCase):
    def test_uses_localappdata(self):
        with mock.patch.dict(os.environ, {"LOCALAPPDATA": "C:\\Local"}):
            result = _get_logs_dir()
            self.assertEqual(str(result), "C:\\Local\\Z7\\Apps\\StdProposers\\LocalConfigs\\logs")

    def test_falls_back_to_userprofile(self):
        env = {"USERPROFILE": "C:\\Users\\test", "LOCALAPPDATA": ""}
        with mock.patch.dict(os.environ, env, clear=False):
            os.environ.pop("LOCALAPPDATA", None)
            result = _get_logs_dir()
            self.assertIn("Z7", str(result))
            self.assertIn("logs", str(result))


class TestZ7ModulePrefix(unittest.TestCase):
    def test_prefix_value(self):
        self.assertEqual(Z7_MODULE_PREFIX, "Mod_")

    def test_prefix_matches_known_modules(self):
        modules = [
            "Mod_01_Infrastructure", "Mod_02_Engine", "Mod_03_Pipeline",
            "Mod_04_Main", "Mod_05_Logging", "Mod_06_WordMacro",
            "Mod_07_Formatting", "Mod_08_Ementa", "Mod_09_SpecialParagraphs",
            "Mod_10_Validation", "Mod_11_RevisionText", "Mod_12_AIStructure",
        ]
        for name in modules:
            self.assertTrue(
                name.startswith(Z7_MODULE_PREFIX),
                f"{name} should start with {Z7_MODULE_PREFIX}",
            )

    def test_prefix_does_not_match_non_z7(self):
        non_z7 = ["Normal", "ThisDocument", "UserForm1", "Module1"]
        for name in non_z7:
            self.assertFalse(name.startswith(Z7_MODULE_PREFIX))


if __name__ == "__main__":
    unittest.main()