import sys
import threading
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
_word_stub = mock.MagicMock()
_word_stub.ActiveDocument.Content.Text = "Texto do documento ativo."
_word_stub.ActiveDocument.Name = "test_doc.docx"

_win32com_stub = mock.MagicMock()
_win32com_stub.client.GetActiveObject.return_value = _word_stub

_chat_session_stub = mock.MagicMock()
_chat_session_stub.send_message.return_value.text = "Olá! Posso ajudar."

_genai_stub = mock.MagicMock()
_genai_stub.Client.return_value.chats.create.return_value = _chat_session_stub

_STUBS = {
    "win32com": _win32com_stub,
    "win32com.client": _win32com_stub.client,
    "pythoncom": mock.MagicMock(),
    "google": mock.MagicMock(),
    "google.genai": _genai_stub,
    "google.genai.types": mock.MagicMock(),
    "win32crypt": mock.MagicMock(),
    "PIL": mock.MagicMock(),
    "PIL.Image": mock.MagicMock(),
}


def _import_chat_ia():
    import importlib
    with mock.patch.dict(sys.modules, _STUBS):
        import chat_ia
        importlib.reload(chat_ia)
        return chat_ia


class TestChatIaDocTextLoading(unittest.TestCase):
    """Testa que o texto do documento é carregado na thread principal."""

    def test_load_doc_text_main_thread_sets_doc_text(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = ""
            app._doc_truncated = False
            app.status_lbl = mock.MagicMock()

            # Chama o método diretamente (sem instância real para não abrir Tk)
            mod.ChatApp._load_doc_text_main_thread(app)

            self.assertEqual(app.doc_text, "Texto do documento ativo.")
            self.assertFalse(app._doc_truncated)

    def test_load_doc_text_sets_fallback_on_error(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = ""
            app._doc_truncated = False

            # Simula falha no GetActiveObject
            broken_win32com = mock.MagicMock()
            broken_win32com.client.GetActiveObject.side_effect = Exception("Word not running")
            broken_win32com.client.GetObject.side_effect = Exception("Word not running")

            with mock.patch.dict(sys.modules, {
                **_STUBS,
                "win32com": broken_win32com,
                "win32com.client": broken_win32com.client,
            }):
                import importlib
                import chat_ia
                importlib.reload(chat_ia)
                chat_ia.ChatApp._load_doc_text_main_thread(app)
                self.assertIn("nenhum documento", app.doc_text.lower())

    def test_load_doc_text_truncates_long_document(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            long_text = "A " * 100_000  # 200k chars > _MAX_CONTEXT_CHARS (150k)
            _word_stub.ActiveDocument.Content.Text = long_text

            app = mock.MagicMock()
            app.doc_text = ""
            app._doc_truncated = False

            mod.ChatApp._load_doc_text_main_thread(app)

            self.assertLessEqual(len(app.doc_text), mod._MAX_CONTEXT_CHARS + 1)
            self.assertTrue(app._doc_truncated)

            # Restaura texto original
            _word_stub.ActiveDocument.Content.Text = "Texto do documento ativo."


class TestChatIaGetActiveObjectIsPrimary(unittest.TestCase):
    def test_get_active_object_before_get_object(self):
        source = (PY_ROOT / "chat_ia.py").read_text(encoding="utf-8")
        lines = source.splitlines()
        # Verifica dentro de _load_doc_text_main_thread
        in_method = False
        get_active_idx = get_object_idx = None
        for i, line in enumerate(lines):
            if "_load_doc_text_main_thread" in line:
                in_method = True
            if in_method and get_active_idx is None and "GetActiveObject" in line:
                get_active_idx = i
            if in_method and get_active_idx is not None and 'GetObject(Class=' in line:
                get_object_idx = i
                break
        self.assertIsNotNone(get_active_idx, "GetActiveObject não encontrado em _load_doc_text_main_thread")
        self.assertIsNotNone(get_object_idx, "GetObject fallback não encontrado")
        self.assertLess(get_active_idx, get_object_idx)


class TestChatIaNewConversation(unittest.TestCase):
    def test_new_conversation_sends_context_when_doc_available(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()

            app = mock.MagicMock()
            app.doc_text = "Conteúdo do documento ativo."
            app.client = _genai_stub.Client()
            app._model = "gemini-2.5-flash"
            app.root = mock.MagicMock()

            # Executa a thread de nova conversa de forma síncrona
            mod.ChatApp._new_conversation_thread(app)

            _chat_session_stub.send_message.assert_called()
            call_args = _chat_session_stub.send_message.call_args[0][0]
            self.assertIn("Conteúdo do documento ativo.", call_args)

    def test_new_conversation_skips_context_when_no_doc(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            _chat_session_stub.reset_mock()

            app = mock.MagicMock()
            app.doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."
            app.client = _genai_stub.Client()
            app._model = "gemini-2.5-flash"
            app.root = mock.MagicMock()

            mod.ChatApp._new_conversation_thread(app)

            # Com doc_text indicando "nenhum documento", não deve enviar contexto
            _chat_session_stub.send_message.assert_not_called()


class TestChatIaPrivacyAndApiKey(unittest.TestCase):
    def test_init_ai_thread_aborts_without_api_key(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = "doc"
            app._doc_truncated = False
            app.root = mock.MagicMock()
            app.root.after = mock.MagicMock()
            app.status_lbl = mock.MagicMock()

            with mock.patch("z7_gemini_key.get_api_key", return_value=None):
                mod.ChatApp._init_ai_thread(app)
                # Verifica que root.after foi chamado com mensagem de erro
                app.root.after.assert_called()


if __name__ == "__main__":
    unittest.main()
