"""Testes unitários para chat_ia.py — foco na leitura robusta do documento do Word.

Cobre o cenário do bug recorrente: documento aberto a partir de
%LocalAppData%\\Temp entra em Protected View e NÃO aparece em
ActiveDocument/Documents, causando o erro COM
'Este comando não está disponível porque nenhum documento foi aberto'
(hresult interno -2146824040).
"""
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[2]
PY_ROOT = ROOT / "ai"
if str(PY_ROOT) not in sys.path:
    sys.path.insert(0, str(PY_ROOT))

import chat_ia  # noqa: E402

# Mensagem COM exata observada em produção (Word PT-BR, Protected View)
NO_DOC_COM_MSG = (
    "(-2147352567, 'Exceção.', (0, 'Microsoft Word', "
    "'Este comando não está disponível porque nenhum documento foi aberto.', "
    "'wdmain11.chm', 37016, -2146824040), None)"
)
NO_DOC_COM_MSG_EN = "This command is not available because no document is open."
RPC_E_CALL_REJECTED = -2147418111  # 0x80010001


# ---------------------------------------------------------------------------
# Fakes COM
# ---------------------------------------------------------------------------
class FakeComError(Exception):
    """Simula pywintypes.com_error (possui atributo hresult)."""

    def __init__(self, message, hresult=None):
        super().__init__(message)
        self.hresult = hresult


class _Content:
    def __init__(self, text):
        self.Text = text


class _Document:
    def __init__(self, text):
        self.Content = _Content(text)


class _Collection:
    """Simula coleção COM 1-based com propriedade Count (Documents, Windows, etc.)."""

    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


class _Window:
    def __init__(self, doc):
        self.Document = doc


class FakeWordApp:
    """Simula um Word.Application COM.

    normal_texts:    documentos normais (aparecem em Documents/ActiveDocument).
    protected_texts: documentos em Protected View (aparecem APENAS em
                     ProtectedViewWindows) — cenário de arquivos do Temp/Internet.
    """

    def __init__(self, normal_texts=(), protected_texts=()):
        self._docs = [_Document(t) for t in normal_texts]
        self.Documents = _Collection(self._docs)
        self.ProtectedViewWindows = _Collection(
            [_Window(_Document(t)) for t in protected_texts]
        )
        self.Windows = _Collection([_Window(d) for d in self._docs])
        self.Application = mock.MagicMock()
        self.StatusBar = ""

    def _no_doc(self):
        raise FakeComError(NO_DOC_COM_MSG, hresult=-2147352567)

    @property
    def ActiveDocument(self):
        if not self._docs:
            self._no_doc()
        return self._docs[0]

    @property
    def ActiveWindow(self):
        if not self._docs:
            self._no_doc()
        return _Window(self._docs[0])

    @property
    def Selection(self):
        if not self._docs:
            self._no_doc()
        return _Window(self._docs[0])


class ExplodingWord:
    """Word em que TODO acesso a propriedade lança o erro configurado."""

    def __init__(self, err):
        self._err = err

    def __getattr__(self, name):
        raise self._err


class RpcBusyWord:
    """Word que rejeita chamadas com RPC_E_CALL_REJECTED nas primeiras N chamadas."""

    def __init__(self, text, busy_calls):
        self._text = text
        self._busy = busy_calls
        self.total_calls = 0
        self.Application = mock.MagicMock()
        self.StatusBar = ""

    def _gate(self):
        self.total_calls += 1
        if self._busy > 0:
            self._busy -= 1
            raise FakeComError("Call was rejected by callee.", hresult=RPC_E_CALL_REJECTED)

    @property
    def ActiveDocument(self):
        self._gate()
        return _Document(self._text)

    @property
    def Documents(self):
        self._gate()
        return _Collection([_Document(self._text)])

    @property
    def ActiveWindow(self):
        self._gate()
        return _Window(_Document(self._text))

    @property
    def Selection(self):
        self._gate()
        return _Window(_Document(self._text))

    @property
    def ProtectedViewWindows(self):
        self._gate()
        return _Collection([])

    @property
    def Windows(self):
        self._gate()
        return _Collection([_Window(_Document(self._text))])


def _new_app():
    """Cria instância de ChatApp sem rodar __init__ (sem Tk/COM real)."""
    app = chat_ia.ChatApp.__new__(chat_ia.ChatApp)
    app.word_app = None
    app.doc_text = ""
    app._doc_truncated = False
    app._doc_load_error = ""
    return app


def _make_rot_stubs(word_instances):
    """Cria stubs de pythoncom/win32com simulando a ROT com as instâncias dadas."""
    pythoncom_stub = mock.MagicMock()
    win32com_client_stub = mock.MagicMock()

    monikers = []
    unks = []
    for i, _w in enumerate(word_instances):
        mk = mock.MagicMock()
        mk.GetDisplayName.return_value = f"!Word.Application.{i}"
        monikers.append(mk)
        unk = mock.MagicMock()
        unk.QueryInterface.return_value = f"disp-{i}"
        unks.append(unk)

    enum = mock.MagicMock()
    enum.Next.side_effect = [[mk] for mk in monikers] + [[]]
    rot = mock.MagicMock()
    rot.EnumRunning.return_value = enum
    rot.GetObject.side_effect = unks

    pythoncom_stub.GetRunningObjectTable.return_value = rot
    pythoncom_stub.CreateBindCtx.return_value = None

    win32com_client_stub.Dispatch.side_effect = list(word_instances)

    win32com_stub = mock.MagicMock()
    win32com_stub.client = win32com_client_stub
    return {
        "pythoncom": pythoncom_stub,
        "win32com": win32com_stub,
        "win32com.client": win32com_client_stub,
    }


# ---------------------------------------------------------------------------
# Helpers de classificação de erros e contagem de documentos
# ---------------------------------------------------------------------------
class TestWordDocCounts(unittest.TestCase):
    def test_counts_normal_documents(self):
        word = FakeWordApp(normal_texts=["a", "b"])
        self.assertEqual(chat_ia.ChatApp._word_doc_counts(word), (2, 0))

    def test_counts_protected_view_documents(self):
        word = FakeWordApp(protected_texts=["temp doc"])
        self.assertEqual(chat_ia.ChatApp._word_doc_counts(word), (0, 1))

    def test_defensive_when_collections_raise(self):
        word = ExplodingWord(FakeComError("COM dead", hresult=-1))
        self.assertEqual(chat_ia.ChatApp._word_doc_counts(word), (0, 0))


class TestErrorClassification(unittest.TestCase):
    def test_is_no_doc_error_pt(self):
        err = FakeComError(NO_DOC_COM_MSG, hresult=-2147352567)
        self.assertTrue(chat_ia.ChatApp._is_no_doc_error(err))

    def test_is_no_doc_error_en(self):
        err = FakeComError(NO_DOC_COM_MSG_EN, hresult=-2147352567)
        self.assertTrue(chat_ia.ChatApp._is_no_doc_error(err))

    def test_is_no_doc_error_empty_collection(self):
        self.assertTrue(chat_ia.ChatApp._is_no_doc_error(RuntimeError("Documents.Count == 0")))

    def test_is_no_doc_error_negative(self):
        self.assertFalse(chat_ia.ChatApp._is_no_doc_error(ValueError("boom")))

    def test_is_rpc_busy_error_call_rejected(self):
        err = FakeComError("rejected", hresult=RPC_E_CALL_REJECTED)
        self.assertTrue(chat_ia.ChatApp._is_rpc_busy_error(err))

    def test_is_rpc_busy_error_retry_later(self):
        err = FakeComError("retry later", hresult=-2147417846)  # 0x8001010A
        self.assertTrue(chat_ia.ChatApp._is_rpc_busy_error(err))

    def test_is_rpc_busy_error_negative(self):
        self.assertFalse(chat_ia.ChatApp._is_rpc_busy_error(FakeComError("x", hresult=-2147352567)))
        self.assertFalse(chat_ia.ChatApp._is_rpc_busy_error(ValueError("sem hresult")))
        self.assertFalse(chat_ia.ChatApp._is_rpc_busy_error(FakeComError("x", hresult="abc")))


# ---------------------------------------------------------------------------
# _read_word_doc_text
# ---------------------------------------------------------------------------
class TestReadWordDocText(unittest.TestCase):
    def setUp(self):
        self.app = _new_app()

    def test_reads_via_active_document(self):
        word = FakeWordApp(normal_texts=["Texto do documento ativo."])
        result = self.app._read_word_doc_text(word)
        self.assertEqual(result, "Texto do documento ativo.")

    def test_reads_via_protected_view_when_no_normal_document(self):
        """CENÁRIO DO BUG: documento aberto do Temp está em Protected View.

        ActiveDocument/Documents/ActiveWindow/Selection lançam o erro COM
        'nenhum documento foi aberto', mas o texto DEVE ser obtido via
        ProtectedViewWindows.
        """
        word = FakeWordApp(protected_texts=["Conteúdo do documento em Protected View."])
        result = self.app._read_word_doc_text(word)
        self.assertEqual(result, "Conteúdo do documento em Protected View.")

    def test_none_text_treated_as_empty(self):
        word = FakeWordApp(normal_texts=[None])
        result = self.app._read_word_doc_text(word)
        self.assertEqual(result, "")

    @mock.patch("time.sleep")
    def test_raises_friendly_message_when_no_document(self, _sleep):
        word = FakeWordApp()  # sem documento algum
        with self.assertRaises(Exception) as ctx:
            self.app._read_word_doc_text(word)
        self.assertIn("nenhum documento", str(ctx.exception).lower())

    @mock.patch("time.sleep")
    def test_retries_when_word_busy_rpc(self, sleep_mock):
        # 9 chamadas COM por rodada (1+2+1+1+2+2); 9 busy = 1ª rodada inteira falha
        word = RpcBusyWord("Texto após retry RPC.", busy_calls=9)
        result = self.app._read_word_doc_text(word)
        self.assertEqual(result, "Texto após retry RPC.")
        self.assertGreater(word.total_calls, 9, "Deve ter tentado mais de uma rodada")
        sleep_mock.assert_called()

    @mock.patch("time.sleep")
    def test_reraises_last_error_when_not_no_doc(self, _sleep):
        err = ValueError("falha genérica de leitura")
        word = ExplodingWord(err)
        with self.assertRaises(ValueError) as ctx:
            self.app._read_word_doc_text(word)
        self.assertIs(ctx.exception, err)


# ---------------------------------------------------------------------------
# _find_word_with_documents (Running Object Table)
# ---------------------------------------------------------------------------
class TestFindWordWithDocuments(unittest.TestCase):
    def setUp(self):
        self.app = _new_app()

    def test_prefers_instance_with_protected_view_doc(self):
        """Instância vazia (tela inicial) deve ser ignorada em favor da que
        possui documento em Protected View."""
        empty_word = FakeWordApp()
        temp_word = FakeWordApp(protected_texts=["doc do temp"])
        stubs = _make_rot_stubs([empty_word, temp_word])
        with mock.patch.dict(sys.modules, stubs):
            result = self.app._find_word_with_documents()
        self.assertIs(result, temp_word)

    def test_prefers_instance_with_normal_doc(self):
        empty_word = FakeWordApp()
        normal_word = FakeWordApp(normal_texts=["doc normal"])
        stubs = _make_rot_stubs([empty_word, normal_word])
        with mock.patch.dict(sys.modules, stubs):
            result = self.app._find_word_with_documents()
        self.assertIs(result, normal_word)

    def test_returns_none_when_no_word_in_rot(self):
        stubs = _make_rot_stubs([])
        with mock.patch.dict(sys.modules, stubs):
            result = self.app._find_word_with_documents()
        self.assertIsNone(result)

    def test_returns_first_instance_when_none_has_docs(self):
        word_a = FakeWordApp()
        word_b = FakeWordApp()
        stubs = _make_rot_stubs([word_a, word_b])
        with mock.patch.dict(sys.modules, stubs):
            result = self.app._find_word_with_documents()
        self.assertIs(result, word_a)


# ---------------------------------------------------------------------------
# _load_doc_text_main_thread (integração das camadas defensivas)
# ---------------------------------------------------------------------------
class TestLoadDocTextMainThread(unittest.TestCase):
    def setUp(self):
        self.app = _new_app()
        self._pythoncom_stub = mock.MagicMock()

    def _run(self):
        with mock.patch.dict(sys.modules, {"pythoncom": self._pythoncom_stub}):
            self.app._load_doc_text_main_thread()

    def test_sets_doc_text_from_active_document(self):
        word = FakeWordApp(normal_texts=["Texto do documento ativo."])
        with mock.patch.object(chat_ia.ChatApp, "_get_word_app", return_value=word):
            self._run()
        self.assertEqual(self.app.doc_text, "Texto do documento ativo.")
        self.assertFalse(self.app._doc_truncated)
        self.assertEqual(self.app._doc_load_error, "")
        # Backup deve ter sido disparado (há ActiveDocument)
        word.Application.Run.assert_called_once()
        self.assertEqual(word.Application.Run.call_args[0][0], "CreateDocumentBackup")

    def test_sets_doc_text_from_protected_view(self):
        """CENÁRIO DO BUG: inteiro teor do documento em Protected View (aberto
        do Temp) deve ser carregado como contexto, sem erro."""
        full_text = "A " * 100 + "documento em protected view vindo da pasta Temp."
        word = FakeWordApp(protected_texts=[full_text])
        with mock.patch.object(chat_ia.ChatApp, "_get_word_app", return_value=word):
            self._run()
        self.assertEqual(self.app.doc_text, full_text)
        self.assertEqual(self.app._doc_load_error, "")
        # Sem ActiveDocument → backup é pulado, sem quebrar o fluxo
        word.Application.Run.assert_not_called()

    @mock.patch("time.sleep")
    def test_reconnects_and_succeeds_on_second_cycle(self, _sleep):
        """1º ciclo falha (instância errada/sem doc); 2º ciclo reconecta e lê."""
        empty_word = FakeWordApp()
        good_word = FakeWordApp(normal_texts=["Texto após reconexão."])
        with mock.patch.object(
            chat_ia.ChatApp, "_get_word_app", side_effect=[empty_word, good_word]
        ) as get_app:
            self._run()
        self.assertEqual(self.app.doc_text, "Texto após reconexão.")
        self.assertEqual(get_app.call_count, 2)
        self.assertIs(self.app.word_app, good_word)

    @mock.patch("time.sleep")
    def test_sets_fallback_on_total_failure(self, _sleep):
        with mock.patch.object(
            chat_ia.ChatApp, "_get_word_app", side_effect=Exception("COM dead")
        ):
            self._run()
        self.assertIn("nenhum documento", self.app.doc_text.lower())
        self.assertIn("Não foi possível conectar", self.app._doc_load_error)

    def test_truncates_long_document(self):
        long_text = "A " * 100_000  # 200k chars > _MAX_CONTEXT_CHARS (150k)
        word = FakeWordApp(normal_texts=[long_text])
        with mock.patch.object(chat_ia.ChatApp, "_get_word_app", return_value=word):
            self._run()
        self.assertLessEqual(len(self.app.doc_text), chat_ia._MAX_CONTEXT_CHARS + 1)
        self.assertTrue(self.app._doc_truncated)


# ---------------------------------------------------------------------------
# _reload_doc_text
# ---------------------------------------------------------------------------
class TestReloadDocText(unittest.TestCase):
    def setUp(self):
        self.app = _new_app()

    def test_reload_success_with_existing_word_app(self):
        self.app.word_app = FakeWordApp(normal_texts=["Texto recarregado."])
        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}):
            ok = self.app._reload_doc_text()
        self.assertTrue(ok)
        self.assertEqual(self.app.doc_text, "Texto recarregado.")

    def test_reload_success_from_protected_view(self):
        self.app.word_app = FakeWordApp(protected_texts=["Temp doc protegido."])
        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}):
            ok = self.app._reload_doc_text()
        self.assertTrue(ok)
        self.assertEqual(self.app.doc_text, "Temp doc protegido.")

    @mock.patch("time.sleep")
    def test_reload_returns_false_on_persistent_failure(self, _sleep):
        self.app.word_app = ExplodingWord(ValueError("stale ref"))
        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}):
            with mock.patch.object(
                chat_ia.ChatApp, "_get_word_app",
                side_effect=Exception("COM dead"),
            ):
                ok = self.app._reload_doc_text()
        self.assertFalse(ok)


# ---------------------------------------------------------------------------
# Convenções de código (guardas de regressão)
# ---------------------------------------------------------------------------
class TestSourceConventions(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (PY_ROOT / "chat_ia.py").read_text(encoding="utf-8")

    def test_get_active_object_before_get_object(self):
        lines = self.source.splitlines()
        in_method = False
        get_active_idx = get_object_idx = None
        for i, line in enumerate(lines):
            if "_load_doc_text_main_thread" in line:
                in_method = True
            if in_method and get_active_idx is None and "GetActiveObject" in line:
                get_active_idx = i
            if in_method and get_active_idx is not None and "GetObject(Class=" in line:
                get_object_idx = i
                break
        self.assertIsNotNone(get_active_idx, "GetActiveObject não encontrado")
        self.assertIsNotNone(get_object_idx, "GetObject fallback não encontrado")
        self.assertLess(get_active_idx, get_object_idx)

    def test_protected_view_is_supported(self):
        """Guarda de regressão do bug: leitura DEVE cobrir ProtectedViewWindows."""
        self.assertIn("ProtectedViewWindows", self.source)

    def test_read_methods_cover_six_strategies(self):
        for method_name in (
            "ActiveDocument",
            "Documents(1)",
            "ActiveWindow.Document",
            "Selection.Document",
            "ProtectedViewWindows(1).Document",
            "Windows(1).Document",
        ):
            self.assertIn(method_name, self.source)


# ---------------------------------------------------------------------------
# Data atual
# ---------------------------------------------------------------------------
class TestDateText(unittest.TestCase):
    def test_get_today_date_text_format(self):
        date_text = chat_ia.get_today_date_text()
        self.assertTrue(date_text.startswith("Hoje é "))
        self.assertTrue(date_text.endswith("."))
        parts = date_text.split(" ")
        self.assertEqual(len(parts), 7)  # ['Hoje', 'é', 'DD', 'de', 'MM', 'de', 'AAAA.']
        self.assertEqual(parts[0], "Hoje")
        self.assertEqual(parts[1], "é")
        self.assertEqual(parts[3], "de")
        self.assertEqual(parts[5], "de")


if __name__ == "__main__":
    unittest.main()
