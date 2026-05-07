import logging
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
PY_ROOT = ROOT / "ai"
if str(PY_ROOT) not in os.sys.path:
    os.sys.path.insert(0, str(PY_ROOT))

from z7_logging import (  # noqa: E402
    build_log_path,
    configure_component_logger,
    get_component_log_path,
    get_data_dir,
    get_logs_dir,
    is_frozen,
    log_exception,
)


class TestBuildLogPath(unittest.TestCase):
    def test_build_log_path_uses_component_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                path = build_log_path("component.test")
                self.assertEqual(path.name, "component_test.log")
                self.assertTrue(path.suffix == ".log")

    def test_get_component_log_path_is_alias(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                self.assertEqual(
                    build_log_path("mycomp"),
                    get_component_log_path("mycomp"),
                )


class TestConfigureComponentLogger(unittest.TestCase):
    def test_configure_component_logger_creates_log_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                logger_name = "z7.test_logger"
                logger = logging.getLogger(logger_name)
                logger.handlers = []
                logger = configure_component_logger("test_logger")
                logger.info("log-entry")

                for handler in logger.handlers:
                    handler.flush()

                files = list(Path(tmp).glob("test_logger*.log"))
                self.assertEqual(len(files), 1)
                content = files[0].read_text(encoding="utf-8")
                self.assertIn("log-entry", content)

                handlers = list(logger.handlers)
                for handler in handlers:
                    handler.close()
                    logger.removeHandler(handler)


class TestGetLogsDir(unittest.TestCase):
    def test_get_logs_dir_creates_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake_user = Path(tmp)
            with mock.patch.dict(os.environ, {"USERPROFILE": str(fake_user)}, clear=False):
                logs_dir = get_logs_dir()
                self.assertTrue(logs_dir.exists())
                self.assertTrue(logs_dir.is_dir())


class TestGetDataDir(unittest.TestCase):
    def test_get_data_dir_creates_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake_user = Path(tmp)
            with mock.patch.dict(os.environ, {"USERPROFILE": str(fake_user)}, clear=False):
                data_dir = get_data_dir()
                self.assertTrue(data_dir.exists())
                self.assertTrue(data_dir.is_dir())

    def test_get_data_dir_is_under_appdata_local(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake_user = Path(tmp)
            with mock.patch.dict(os.environ, {"USERPROFILE": str(fake_user)}, clear=False):
                data_dir = get_data_dir()
                self.assertTrue(str(data_dir).startswith(str(fake_user)))


class TestIsFrozen(unittest.TestCase):
    def test_returns_false_when_not_compiled(self):
        with mock.patch.object(os.sys, "frozen", False, create=True):
            self.assertFalse(is_frozen())

    def test_returns_true_when_frozen(self):
        with mock.patch.object(os.sys, "frozen", True, create=True):
            with mock.patch.object(os.sys, "_MEIPASS", "/fake/path", create=True):
                self.assertTrue(is_frozen())


class TestLogException(unittest.TestCase):
    def test_logs_exception_without_reraise(self):
        logger = logging.getLogger("z7.test_exc")
        logger.handlers = []
        exc = ValueError("erro de teste")
        # Não deve lançar exceção
        log_exception(logger, "contexto", exc, reraise=False)

    def test_reraise_true_re_raises(self):
        logger = logging.getLogger("z7.test_exc_reraise")
        logger.handlers = []
        exc = ValueError("erro para relançar")
        with self.assertRaises(ValueError):
            log_exception(logger, "ctx", exc, reraise=True)

    def test_reraise_default_does_not_raise(self):
        logger = logging.getLogger("z7.test_exc_default")
        logger.handlers = []
        exc = RuntimeError("erro sem reraise")
        log_exception(logger, "ctx", exc)  # reraise padrão = False


if __name__ == "__main__":
    unittest.main()

