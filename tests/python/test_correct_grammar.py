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
_word_stub = mock.MagicMock()
_word_stub.ActiveDocument.Content.Text = "Texto completo do documento para teste em producao."
_win32com_stub.client.GetActiveObject.return_value = _word_stub

_genai_stub = mock.MagicMock()
_response_stub = mock.MagicMock()
_response_stub.text = "Texto corrigido gramaticalmente."
_genai_stub.Client.return_value.models.generate_content.return_value = _response_stub

_STUBS = {
    "win32com": _win32com_stub,
    "win32com.client": _win32com_stub.client,
    "google": mock.MagicMock(),
    "google.genai": _genai_stub,
    "win32crypt": mock.MagicMock(),
    "tkinter": mock.MagicMock(),
}


def _import_correct_grammar():
    import importlib
    with mock.patch.dict(sys.modules, _STUBS):
        import correct_grammar
        importlib.reload(correct_grammar)
        return correct_grammar


class TestCorrectGrammarWordConnection(unittest.TestCase):
    def test_get_active_object_called_on_main(self):
        """Verifica que main() invoca GetActiveObject ao tentar conectar ao Word."""
        win32_spy = mock.MagicMock()
        win32_spy.client.GetActiveObject.return_value = _word_stub

        stubs = dict(_STUBS)
        stubs["win32com"] = win32_spy
        stubs["win32com.client"] = win32_spy.client

        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            # Cancela no aviso de privacidade para não bloquear o teste;
            # GetActiveObject já foi chamado antes dessa etapa.
            with mock.patch("z7_theme.ask_privacy_warning", return_value=False):
                try:
                    correct_grammar.main()
                except Exception:
                    pass

        win32_spy.client.GetActiveObject.assert_called_with("Word.Application")

    def test_get_active_object_is_primary_connection(self):
        """correct_grammar.py deve tentar GetActiveObject antes de GetObject."""
        source = (PY_ROOT / "correct_grammar.py").read_text(encoding="utf-8")
        lines = source.splitlines()
        get_active_idx = next(
            (i for i, l in enumerate(lines) if "GetActiveObject" in l), None
        )
        get_object_idx = next(
            (i for i, l in enumerate(lines) if 'GetObject(Class=' in l), None
        )
        self.assertIsNotNone(get_active_idx, "GetActiveObject não encontrado")
        self.assertIsNotNone(get_object_idx, "GetObject fallback não encontrado")
        self.assertLess(
            get_active_idx, get_object_idx,
            "GetActiveObject deve aparecer antes do fallback GetObject",
        )


class TestCorrectGrammarDocumentGuards(unittest.TestCase):
    def test_no_active_document_is_handled(self):
        """ActiveDocument None não deve causar AttributeError."""
        word_no_doc = mock.MagicMock()
        word_no_doc.ActiveDocument = None

        win32com_no_doc = mock.MagicMock()
        win32com_no_doc.client.GetActiveObject.return_value = word_no_doc

        stubs = dict(_STUBS)
        stubs["win32com"] = win32com_no_doc
        stubs["win32com.client"] = win32com_no_doc.client

        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            try:
                correct_grammar.main()
            except SystemExit:
                pass
            except Exception as exc:
                self.fail(f"main() lançou exceção inesperada: {exc}")

    def test_empty_document_exits_early(self):
        """Documento vazio ou muito curto não deve prosseguir para a API."""
        word_empty = mock.MagicMock()
        word_empty.ActiveDocument.Content.Text = "curto"

        win32com_empty = mock.MagicMock()
        win32com_empty.client.GetActiveObject.return_value = word_empty

        stubs = dict(_STUBS)
        stubs["win32com"] = win32com_empty
        stubs["win32com.client"] = win32com_empty.client

        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            _genai_stub.reset_mock()
            try:
                correct_grammar.main()
            except SystemExit:
                pass
            _genai_stub.Client.return_value.models.generate_content.assert_not_called()


class TestCorrectGrammarPrivacyCheck(unittest.TestCase):
    def test_cancelled_privacy_warning_aborts(self):
        """Se o usuário cancelar o aviso de privacidade, a API não é chamada."""
        stubs = dict(_STUBS)
        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            with mock.patch("z7_theme.ask_privacy_warning", return_value=False):
                with mock.patch("z7_gemini_key.get_api_key") as mock_key:
                    try:
                        correct_grammar.main()
                    except SystemExit:
                        pass
                    mock_key.assert_not_called()


class TestCorrectGrammarApiKeyMissing(unittest.TestCase):
    def test_missing_api_key_aborts_gracefully(self):
        stubs = dict(_STUBS)
        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            with mock.patch("z7_theme.ask_privacy_warning", return_value=True):
                # correct_grammar importa get_api_key via 'from', então o patch
                # deve ser feito no namespace do módulo, não no de z7_gemini_key.
                with mock.patch.object(correct_grammar, 'get_api_key', return_value=None):
                    _genai_stub.reset_mock()
                    try:
                        correct_grammar.main()
                    except (SystemExit, Exception):
                        pass
                    _genai_stub.Client.assert_not_called()


class TestCorrectGrammarDocumentApplication(unittest.TestCase):
    def test_corrected_text_replaces_document_content(self):
        """Quando o usuário confirma, doc.Content.Text deve receber o texto corrigido."""
        doc_mock = mock.MagicMock()
        doc_mock.Content.Text = "Texto completo do documento para ser corrigido pela IA."

        word_mock = mock.MagicMock()
        word_mock.ActiveDocument = doc_mock

        win32com_mock = mock.MagicMock()
        win32com_mock.client.GetActiveObject.return_value = word_mock

        genai_mock = mock.MagicMock()
        genai_mock.Client.return_value.models.generate_content.return_value.text = (
            "Texto completo do documento corrigido pela IA."
        )
        google_mock = mock.MagicMock()
        google_mock.genai = genai_mock

        stubs = {
            "win32com": win32com_mock,
            "win32com.client": win32com_mock.client,
            "google": google_mock,
            "google.genai": genai_mock,
            "win32crypt": mock.MagicMock(),
            "tkinter": mock.MagicMock(),
        }

        import importlib
        with mock.patch.dict(sys.modules, stubs):
            import correct_grammar
            importlib.reload(correct_grammar)
            with mock.patch("z7_theme.ask_privacy_warning", return_value=True):
                with mock.patch.object(correct_grammar, 'get_api_key', return_value="fake-key"):
                    with mock.patch.object(correct_grammar, '_show_result_window',
                                           side_effect=lambda text, **kw: text):
                        try:
                            correct_grammar.main()
                        except Exception:
                            pass

        self.assertEqual(
            doc_mock.Content.Text,
            "Texto completo do documento corrigido pela IA."
        )


if __name__ == "__main__":
    unittest.main()
