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


# ---------------------------------------------------------------------------
# Helper interno para capturar LogRecords sem I/O
# ---------------------------------------------------------------------------

class _RecordingHandler(logging.Handler):
    """Handler que coleta registros em memória para inspeção nos testes."""
    def __init__(self):
        super().__init__()
        self.records: list = []

    def emit(self, record: logging.LogRecord) -> None:
        self.records.append(record)

    def messages(self) -> list:
        return [self.format(record) for record in self.records]


# ---------------------------------------------------------------------------
# Idempotência do logger
# ---------------------------------------------------------------------------

class TestLoggerIdempotency(unittest.TestCase):
    """configure_component_logger não deve adicionar handlers duplicados."""

    def _cleanup(self, logger: logging.Logger) -> None:
        for h in list(logger.handlers):
            h.close()
            logger.removeHandler(h)

    def test_no_duplicate_handlers_on_repeated_calls(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.idem_test_1"
                logger = logging.getLogger(name)
                logger.handlers.clear()

                configure_component_logger("idem_test_1")
                count_first = len(logger.handlers)
                configure_component_logger("idem_test_1")
                count_second = len(logger.handlers)

                self.assertEqual(count_first, count_second,
                                 "Handlers duplicados ao chamar configure duas vezes")
                self._cleanup(logger)

    def test_returns_same_logger_instance(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.idem_test_2"
                logging.getLogger(name).handlers.clear()

                l1 = configure_component_logger("idem_test_2")
                l2 = configure_component_logger("idem_test_2")
                self.assertIs(l1, l2)
                self._cleanup(l1)

    def test_debug_mode_adds_stream_handler(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.idem_debug_test"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("idem_debug_test", debug=True)
                handler_types = [type(h).__name__ for h in logger.handlers]
                self.assertIn("StreamHandler", handler_types,
                              "Modo debug deve adicionar StreamHandler")
                self._cleanup(logger)


# ---------------------------------------------------------------------------
# Formato dos registros de log
# ---------------------------------------------------------------------------

class TestLogFormat(unittest.TestCase):
    """Verifica que o formato gravado em arquivo contém os campos obrigatórios."""

    def _cleanup(self, logger: logging.Logger) -> None:
        for h in list(logger.handlers):
            h.close()
            logger.removeHandler(h)

    def test_log_entry_contains_level_name_and_message(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.fmt_test_1"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("fmt_test_1")
                logger.info("mensagem_de_formato_unica")

                for h in logger.handlers:
                    h.flush()

                files = list(Path(tmp).glob("fmt_test_1.log"))
                self.assertEqual(len(files), 1)
                content = files[0].read_text(encoding="utf-8")

                self.assertIn("INFO", content)
                self.assertIn("z7.fmt_test_1", content)
                self.assertIn("mensagem_de_formato_unica", content)
                # Separador pipe do formatter
                self.assertIn("|", content)
                self._cleanup(logger)

    def test_warning_level_recorded_correctly(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.fmt_test_2"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("fmt_test_2")
                logger.warning("aviso_de_teste_fmt")

                for h in logger.handlers:
                    h.flush()

                files = list(Path(tmp).glob("fmt_test_2.log"))
                content = files[0].read_text(encoding="utf-8")
                self.assertIn("WARNING", content)
                self.assertIn("aviso_de_teste_fmt", content)
                self._cleanup(logger)

    def test_error_level_recorded_correctly(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.fmt_test_3"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("fmt_test_3")
                logger.error("erro_de_teste_fmt")

                for h in logger.handlers:
                    h.flush()

                files = list(Path(tmp).glob("fmt_test_3.log"))
                content = files[0].read_text(encoding="utf-8")
                self.assertIn("ERROR", content)
                self._cleanup(logger)


# ---------------------------------------------------------------------------
# log_exception grava traceback no arquivo
# ---------------------------------------------------------------------------

class TestLogExceptionWritesTraceback(unittest.TestCase):
    """log_exception deve gravar o traceback completo no arquivo de log."""

    def _cleanup(self, logger: logging.Logger) -> None:
        for h in list(logger.handlers):
            h.close()
            logger.removeHandler(h)

    def test_traceback_written_to_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.exc_file_1"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("exc_file_1")
                try:
                    raise ValueError("erro_simulado_traceback")
                except ValueError as exc:
                    log_exception(logger, "contexto_do_teste", exc)

                for h in logger.handlers:
                    h.flush()

                files = list(Path(tmp).glob("exc_file_1.log"))
                content = files[0].read_text(encoding="utf-8")

                self.assertIn("erro_simulado_traceback", content)
                self.assertIn("Traceback", content)
                self.assertIn("ValueError", content)
                self._cleanup(logger)

    def test_context_message_appears_in_log(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.exc_file_2"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("exc_file_2")
                try:
                    raise RuntimeError("runtime_failure_test")
                except RuntimeError as exc:
                    log_exception(logger, "contexto_esperado_no_log", exc)

                for h in logger.handlers:
                    h.flush()

                files = list(Path(tmp).glob("exc_file_2.log"))
                content = files[0].read_text(encoding="utf-8")
                self.assertIn("contexto_esperado_no_log", content)
                self._cleanup(logger)

    def test_multiple_exception_types_recorded(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_logs_dir", return_value=Path(tmp)):
                name = "z7.exc_file_3"
                logging.getLogger(name).handlers.clear()

                logger = configure_component_logger("exc_file_3")
                for exc_type, msg in [(KeyError, "key_error_msg"), (OSError, "os_error_msg")]:
                    try:
                        raise exc_type(msg)
                    except Exception as exc:
                        log_exception(logger, f"ctx_{msg}", exc)

                for h in logger.handlers:
                    h.flush()

                files = list(Path(tmp).glob("exc_file_3.log"))
                content = files[0].read_text(encoding="utf-8")
                self.assertIn("KeyError", content)
                self.assertIn("OSError", content)
                self._cleanup(logger)


# ---------------------------------------------------------------------------
# get_session_id
# ---------------------------------------------------------------------------

class TestGetSessionId(unittest.TestCase):
    """Verifica formato e unicidade de get_session_id."""

    def test_session_id_format_yyyymmdd_hhmmss(self):
        from z7_logging import get_session_id
        sid = get_session_id()
        self.assertEqual(len(sid), 15, f"Comprimento inesperado: {sid!r}")
        self.assertEqual(sid[8], "_", f"Separador ausente em posição 8: {sid!r}")
        self.assertTrue(sid[:8].isdigit(), f"Prefixo não numérico: {sid!r}")
        self.assertTrue(sid[9:].isdigit(), f"Sufixo não numérico: {sid!r}")

    def test_session_id_date_component_matches_today(self):
        from z7_logging import get_session_id
        import datetime as _dt
        sid = get_session_id()
        today = _dt.date.today().strftime("%Y%m%d")
        self.assertEqual(sid[:8], today,
                         f"Componente de data incorreto: {sid!r}")

    def test_session_id_exported_in_all(self):
        import z7_logging
        self.assertIn("get_session_id", z7_logging.__all__)


# ---------------------------------------------------------------------------
# log_context
# ---------------------------------------------------------------------------

class TestLogContext(unittest.TestCase):
    """Verifica o context manager log_context."""

    def test_start_and_end_messages_logged(self):
        from z7_logging import log_context
        logger = logging.getLogger("z7.ctx_ok_test")
        logger.handlers.clear()
        rec = _RecordingHandler()
        rec.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(rec)
        logger.setLevel(logging.DEBUG)

        with log_context(logger, "operacao_de_teste"):
            pass

        messages = rec.messages()
        combined = " ".join(messages)
        self.assertIn("START", combined)
        self.assertIn("END", combined)
        self.assertIn("operacao_de_teste", combined)
        logger.removeHandler(rec)

    def test_elapsed_time_recorded_on_success(self):
        from z7_logging import log_context
        logger = logging.getLogger("z7.ctx_elapsed_test")
        logger.handlers.clear()
        rec = _RecordingHandler()
        rec.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(rec)
        logger.setLevel(logging.DEBUG)

        with log_context(logger, "bloco_temporizado"):
            pass

        combined = " ".join(rec.messages())
        # Formato esperado: "[END]   bloco_temporizado | 0.000s"
        self.assertRegex(combined, r"\d+\.\d+s")
        logger.removeHandler(rec)

    def test_exception_reraised_and_fail_logged(self):
        from z7_logging import log_context
        logger = logging.getLogger("z7.ctx_fail_test")
        logger.handlers.clear()
        rec = _RecordingHandler()
        rec.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(rec)
        logger.setLevel(logging.DEBUG)

        with self.assertRaises(ValueError):
            with log_context(logger, "operacao_com_falha"):
                raise ValueError("falha_proposital")

        combined = " ".join(rec.messages())
        self.assertIn("FAIL", combined)
        self.assertIn("operacao_com_falha", combined)
        self.assertIn("falha_proposital", combined)
        logger.removeHandler(rec)

    def test_log_context_exported_in_all(self):
        import z7_logging
        self.assertIn("log_context", z7_logging.__all__)


if __name__ == "__main__":
    unittest.main()

