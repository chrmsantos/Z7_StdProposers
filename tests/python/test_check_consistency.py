import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
PY_ROOT = ROOT / "ai"
if str(PY_ROOT) not in sys.path:
    sys.path.insert(0, str(PY_ROOT))

# ---------------------------------------------------------------------------
# Stubs de dependências externas
# ---------------------------------------------------------------------------
_win32com_stub = mock.MagicMock()
_doc_stub = mock.MagicMock()
_doc_stub.Range.return_value.Text = "Texto completo da propositura para teste."
_word_stub = mock.MagicMock()
_word_stub.ActiveDocument = _doc_stub
_win32com_stub.client.GetActiveObject.return_value = _word_stub

_genai_stub = mock.MagicMock()
_response_stub = mock.MagicMock()
_response_stub.text = "Sem inconsistências graves detectadas."
_genai_stub.Client.return_value.models.generate_content.return_value = _response_stub

_STUBS = {
    "win32com": _win32com_stub,
    "win32com.client": _win32com_stub.client,
    "google": mock.MagicMock(),
    "google.genai": _genai_stub,
    "win32crypt": mock.MagicMock(),
}


def _import_check_consistency():
    import importlib
    with mock.patch.dict(sys.modules, _STUBS):
        import check_consistency
        importlib.reload(check_consistency)
        return check_consistency


class TestCheckConsistencyWordConnection(unittest.TestCase):
    def test_get_active_object_called_on_main(self):
        """Verifica que main() invoca GetActiveObject ao tentar conectar ao Word."""
        win32_spy = mock.MagicMock()
        win32_spy.client.GetActiveObject.return_value = _word_stub

        stubs = dict(_STUBS)
        stubs["win32com"] = win32_spy
        stubs["win32com.client"] = win32_spy.client

        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import check_consistency
            importlib.reload(check_consistency)
            # Suprime o diálogo de privacidade para evitar bloqueio da UI no teste
            with mock.patch("z7_theme.ask_privacy_warning", return_value=False):
                try:
                    check_consistency.main()
                except Exception:
                    pass

        win32_spy.client.GetActiveObject.assert_called_with("Word.Application")

    def test_get_active_object_is_primary_connection(self):
        """check_consistency.py deve tentar GetActiveObject antes de GetObject."""
        source = (PY_ROOT / "check_consistency.py").read_text(encoding="utf-8")
        lines = source.splitlines()
        get_active_idx = next(
            (i for i, l in enumerate(lines) if "GetActiveObject" in l), None
        )
        get_object_idx = next(
            (i for i, l in enumerate(lines) if 'GetObject(Class=' in l), None
        )
        self.assertIsNotNone(get_active_idx, "GetActiveObject não encontrado")
        self.assertIsNotNone(get_object_idx, "GetObject fallback não encontrado")
        self.assertLess(get_active_idx, get_object_idx,
                        "GetActiveObject deve vir antes de GetObject fallback")


class TestCheckConsistencyNoIssueMarker(unittest.TestCase):
    def test_no_issue_marker_constant_defined(self):
        """_NO_ISSUE_MARKER deve estar definido no módulo."""
        mod = _import_check_consistency()
        self.assertTrue(hasattr(mod, '_NO_ISSUE_MARKER'))
        self.assertIsInstance(mod._NO_ISSUE_MARKER, str)
        self.assertTrue(len(mod._NO_ISSUE_MARKER) > 0)

    def test_has_issues_false_when_marker_present(self):
        """has_issues deve ser False quando a resposta contém o marcador de sem problemas."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        analysis = f"Análise concluída. {marker.capitalize()}."
        has_issues = marker not in analysis.lower()
        self.assertFalse(has_issues)

    def test_has_issues_true_when_marker_absent(self):
        """has_issues deve ser True quando a resposta não contém o marcador."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        analysis = "1. A ementa menciona 'Rua X' mas o corpo do texto cita 'Rua Y'."
        has_issues = marker not in analysis.lower()
        self.assertTrue(has_issues)


class TestCheckConsistencyCodeConventions(unittest.TestCase):
    def test_uses_shared_logger(self):
        """check_consistency.py deve usar configure_component_logger de z7_logging."""
        source = (PY_ROOT / "check_consistency.py").read_text(encoding="utf-8")
        self.assertIn("from z7_logging import configure_component_logger", source)
        self.assertIn("LOGGER = configure_component_logger", source)

    def test_imports_from_z7_gemini_key(self):
        """check_consistency.py deve importar de z7_gemini_key."""
        source = (PY_ROOT / "check_consistency.py").read_text(encoding="utf-8")
        self.assertIn("from z7_gemini_key import", source)

    def test_loads_consistency_prompt_from_config(self):
        """check_consistency.py deve carregar o prompt de consistência via config_prompt."""
        source = (PY_ROOT / "check_consistency.py").read_text(encoding="utf-8")
        self.assertIn("load_consistency_prompt", source)

    def test_loads_ai_model_from_config(self):
        """check_consistency.py deve carregar o modelo via config_prompt.load_ai_model."""
        source = (PY_ROOT / "check_consistency.py").read_text(encoding="utf-8")
        self.assertIn("load_ai_model", source)


class TestCheckConsistencyConfigPromptIntegration(unittest.TestCase):
    def test_load_consistency_prompt_returns_default_when_no_file(self):
        """load_consistency_prompt deve retornar DEFAULT_CONSISTENCY_PROMPT quando sem arquivo."""
        import importlib
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                import config_prompt
                importlib.reload(config_prompt)
                result = config_prompt.load_consistency_prompt()
                self.assertEqual(result, config_prompt.DEFAULT_CONSISTENCY_PROMPT)

    def test_load_consistency_prompt_returns_custom_when_file_exists(self):
        """load_consistency_prompt deve retornar conteúdo do arquivo quando existir."""
        import importlib
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "consistency_prompt.txt").write_text("prompt personalizado", encoding="utf-8")
            with mock.patch("z7_logging.get_data_dir", return_value=tmp_path):
                import config_prompt
                importlib.reload(config_prompt)
                result = config_prompt.load_consistency_prompt()
                self.assertEqual(result, "prompt personalizado")

    def test_default_consistency_prompt_not_empty(self):
        """DEFAULT_CONSISTENCY_PROMPT deve estar definido e não estar vazio."""
        import importlib
        import config_prompt
        importlib.reload(config_prompt)
        self.assertTrue(hasattr(config_prompt, 'DEFAULT_CONSISTENCY_PROMPT'))
        self.assertGreater(len(config_prompt.DEFAULT_CONSISTENCY_PROMPT.strip()), 0)

    def test_get_consistency_prompt_file_path_uses_data_dir(self):
        """get_consistency_prompt_file_path deve usar o data_dir e nomear o arquivo corretamente."""
        import importlib
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch("z7_logging.get_data_dir", return_value=Path(tmp)):
                import config_prompt
                importlib.reload(config_prompt)
                path = config_prompt.get_consistency_prompt_file_path()
                self.assertEqual(path.parent, Path(tmp))
                self.assertEqual(path.name, "consistency_prompt.txt")


if __name__ == "__main__":
    unittest.main()
