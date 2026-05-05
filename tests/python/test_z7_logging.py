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

from z7_logging import build_log_path, configure_component_logger, get_logs_dir  # noqa: E402


class TestZ7Logging(unittest.TestCase):
    def test_build_log_path_uses_component_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                path = build_log_path("component.test")
                self.assertTrue(path.name.startswith("component_test_"))
                self.assertTrue(path.suffix == ".log")

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

                files = list(Path(tmp).glob("test_logger_*.log"))
                self.assertEqual(len(files), 1)
                content = files[0].read_text(encoding="utf-8")
                self.assertIn("log-entry", content)

                handlers = list(logger.handlers)
                for handler in handlers:
                    handler.close()
                    logger.removeHandler(handler)

    def test_get_logs_dir_creates_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake_user = Path(tmp)
            with mock.patch.dict(os.environ, {"USERPROFILE": str(fake_user)}, clear=False):
                logs_dir = get_logs_dir()
                self.assertTrue(logs_dir.exists())
                self.assertTrue(logs_dir.is_dir())


if __name__ == "__main__":
    unittest.main()

