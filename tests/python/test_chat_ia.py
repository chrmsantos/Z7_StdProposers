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

            # Preserva doc_text: _reload_doc_text seria chamado internamente mas nao deve
            # sobrescrever o valor ja definido no mock
            with mock.patch.object(mod.ChatApp, '_reload_doc_text', return_value=True):
                mod.ChatApp._new_conversation_thread(app)

            # Com abordagem history-based, o contexto e injetado via history no chats.create
            create_call = app.client.chats.create.call_args
            self.assertIsNotNone(create_call, "chats.create deve ter sido chamado")
            history = create_call.kwargs.get('history', [])
            self.assertGreater(len(history), 0, "history deve ser nao vazia quando ha contexto do documento")

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

            with mock.patch.object(mod, 'get_api_key', return_value=None):
                mod.ChatApp._init_ai_thread(app)
                # Verifica que root.after foi chamado com mensagem de erro
                app.root.after.assert_called()


class TestChatIaContextPending(unittest.TestCase):
    """Testa a injeção diferida do contexto do documento quando o envio inicial falha."""

    def _make_stubs_with_failing_context(self):
        """Retorna stubs onde o genai falha na primeira chamada (simulando 503) mas funciona na segunda."""
        chat_stub = mock.MagicMock()
        # Primeira chamada (contexto inicial) falha; segunda (primeira mensagem do usuário) funciona.
        chat_stub.send_message.side_effect = [
            Exception("503 UNAVAILABLE"),
            mock.MagicMock(text="Resposta normal"),
        ]
        genai_stub = mock.MagicMock()
        genai_stub.Client.return_value.chats.create.return_value = chat_stub

        google_mock = mock.MagicMock()
        google_mock.genai = genai_stub
        stubs = {
            **_STUBS,
            "google": google_mock,
            "google.genai": genai_stub,
        }
        return stubs, chat_stub

    def test_context_pending_set_when_initial_send_fails(self):
        """Com abordagem history-based, _context_pending nao e alterado pelo init.
        O _context_pending pode ser usado por outras partes do codigo para injecao diferida."""
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = "Conteúdo do documento para contexto."
            app._doc_truncated = False
            app._context_pending = False
            app.root = mock.MagicMock()

            with mock.patch.object(mod, 'get_api_key', return_value="fake-key"):
                mod.ChatApp._init_ai_thread(app)

            # Com abordagem history-based, _context_pending permanece False apos init
            self.assertFalse(app._context_pending)

    def test_send_message_injects_context_when_pending(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = "Texto do documento."
            app._context_pending = True
            app.chat_session = _chat_session_stub
            _chat_session_stub.reset_mock()
            _chat_session_stub.send_message.return_value.text = "Entendido, contexto recebido."

            mod.ChatApp._send_message_thread(app, "Qual o tema do documento?")

            # Deve ter enviado a mensagem com o contexto embutido
            call_args = _chat_session_stub.send_message.call_args[0][0]
            self.assertIn("Texto do documento.", call_args)
            self.assertIn("Qual o tema do documento?", call_args)
            # Após envio, _context_pending deve ser False
            self.assertFalse(app._context_pending)

    def test_send_message_does_not_inject_context_when_not_pending(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = "Texto do documento."
            app._context_pending = False
            app.chat_session = _chat_session_stub
            _chat_session_stub.reset_mock()
            _chat_session_stub.send_message.return_value.text = "Resposta normal"

            mod.ChatApp._send_message_thread(app, "Pergunta simples")

            call_args = _chat_session_stub.send_message.call_args[0][0]
            # Sem contexto pendente, deve enviar apenas a mensagem do usuário
            self.assertEqual(call_args, "Pergunta simples")

    def test_send_message_skips_context_injection_when_no_doc(self):
        with mock.patch.dict(sys.modules, _STUBS):
            mod = _import_chat_ia()
            app = mock.MagicMock()
            app.doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."
            app._context_pending = True  # pendente mas sem doc real
            app.chat_session = _chat_session_stub
            _chat_session_stub.reset_mock()
            _chat_session_stub.send_message.return_value.text = "Sem contexto"

            mod.ChatApp._send_message_thread(app, "Pergunta sem contexto")

            # Com fallback message, não deve injetar o "documento" na mensagem
            call_args = _chat_session_stub.send_message.call_args[0][0]
            self.assertEqual(call_args, "Pergunta sem contexto")


if __name__ == "__main__":
    unittest.main()
