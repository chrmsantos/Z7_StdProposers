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
_selection_stub = mock.MagicMock()
_selection_stub.Text = "Texto selecionado para teste."
_selection_stub.Font.Name = "Arial"
_selection_stub.Font.Size = 12
_selection_stub.Font.Bold = False
_selection_stub.Font.Italic = False
_word_stub.Selection = _selection_stub
_word_stub.ActiveDocument = mock.MagicMock()
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


class TestCorrectGrammarSelectionGuards(unittest.TestCase):
    def test_none_selection_is_handled(self):
        """word.Selection retornando None não deve causar AttributeError."""
        word_none_sel = mock.MagicMock()
        word_none_sel.Selection = None

        win32com_none = mock.MagicMock()
        win32com_none.client.GetActiveObject.return_value = word_none_sel

        stubs = dict(_STUBS)
        stubs["win32com"] = win32com_none
        stubs["win32com.client"] = win32com_none.client

        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            # Não deve lançar exceção
            try:
                correct_grammar.main()
            except SystemExit:
                pass
            except Exception as exc:
                self.fail(f"main() lançou exceção inesperada: {exc}")

    def test_short_selection_exits_early(self):
        """Seleção com menos de 2 caracteres não deve prosseguir para a API."""
        short_sel = mock.MagicMock()
        short_sel.Text = "a"
        word_short = mock.MagicMock()
        word_short.Selection = short_sel

        win32com_short = mock.MagicMock()
        win32com_short.client.GetActiveObject.return_value = word_short

        stubs = dict(_STUBS)
        stubs["win32com"] = win32com_short
        stubs["win32com.client"] = win32com_short.client

        with mock.patch.dict(sys.modules, stubs):
            import importlib
            import correct_grammar
            importlib.reload(correct_grammar)
            try:
                correct_grammar.main()
            except SystemExit:
                pass
            # API Gemini não deve ter sido chamada
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


class TestCorrectGrammarFormattingPreservation(unittest.TestCase):
    """Testa que a formatação do texto selecionado é salva e reaplicada corretamente."""

    def _run_main_with_fmt(self, font_attrs: dict):
        """Executa main() com seleção formatada e retorna (sel, mod) para inspeção."""
        first_char_font = mock.MagicMock()
        for attr, value in font_attrs.items():
            setattr(first_char_font, attr, value)

        first_char = mock.MagicMock()
        first_char.Font = first_char_font

        sel = mock.MagicMock()
        sel.Text = "Texto com formatação específica."
        sel.Characters = mock.MagicMock(return_value=first_char)

        word = mock.MagicMock()
        word.Selection = sel

        win32com_mock = mock.MagicMock()
        win32com_mock.client.GetActiveObject.return_value = word

        genai_mock = mock.MagicMock()
        genai_mock.Client.return_value.models.generate_content.return_value.text = (
            "Texto corrigido gramaticalmente."
        )
        google_mock = mock.MagicMock()
        google_mock.genai = genai_mock

        stubs = {
            "win32com": win32com_mock,
            "win32com.client": win32com_mock.client,
            "google": google_mock,
            "google.genai": genai_mock,
            "win32crypt": mock.MagicMock(),
        }

        import importlib
        with mock.patch.dict(sys.modules, stubs):
            import correct_grammar
            importlib.reload(correct_grammar)
            with mock.patch("z7_theme.ask_privacy_warning", return_value=True):
                with mock.patch.object(correct_grammar, 'get_api_key', return_value="fake-key"):
                    try:
                        correct_grammar.main()
                    except Exception:
                        pass
        return sel

    def test_font_name_is_read_from_first_character(self):
        """Deve usar Characters(1).Font em vez do agregado selection.Font."""
        sel = self._run_main_with_fmt({"Name": "Times New Roman", "Size": 14, "Bold": True})
        sel.Characters.assert_called_with(1)

    def test_bold_formatting_is_reapplied(self):
        """Após substituição, Bold=True deve ser reaplicado à seleção."""
        attrs = {
            "Name": "Arial", "Size": 12, "Bold": True, "Italic": False,
            "Underline": 0, "Color": 0, "StrikeThrough": False,
            "DoubleStrikeThrough": False, "Subscript": False, "Superscript": False,
            "SmallCaps": False, "AllCaps": False, "HighlightColorIndex": 0,
        }
        sel = self._run_main_with_fmt(attrs)
        self.assertEqual(sel.Font.Bold, True)

    def test_all_font_attrs_are_reapplied(self):
        """Todos os 13 atributos de fonte devem ser reaplicados após a substituição."""
        attrs = {
            "Name": "Calibri", "Size": 11, "Bold": False, "Italic": True,
            "Underline": 1, "Color": 255, "StrikeThrough": False,
            "DoubleStrikeThrough": False, "Subscript": False, "Superscript": False,
            "SmallCaps": False, "AllCaps": False, "HighlightColorIndex": 0,
        }
        sel = self._run_main_with_fmt(attrs)
        for attr, expected in attrs.items():
            actual = getattr(sel.Font, attr)
            self.assertEqual(actual, expected, f"Font.{attr} não foi reaplicado")


if __name__ == "__main__":
    unittest.main()
