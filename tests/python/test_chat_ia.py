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

    def test_logs_warning_with_hresult_when_count_read_fails(self):
        """Erro de leitura da contagem NÃO deve ser silencioso: WARNING + hresult."""
        word = ExplodingWord(FakeComError("COM dead", hresult=-2147418111))
        with self.assertLogs("z7.chat_ia", level="WARNING") as cm:
            result = chat_ia.ChatApp._word_doc_counts(word)
        self.assertEqual(result, (0, 0))
        joined = "\n".join(cm.output)
        self.assertIn("Could not read Documents.Count", joined)
        self.assertIn("Could not read ProtectedViewWindows.Count", joined)
        self.assertIn("hresult=-2147418111", joined)


class TestWordDocSnapshot(unittest.TestCase):
    """Cobre _word_doc_snapshot/_fmt_count/_log_word_snapshot — diagnóstico do Word."""

    def setUp(self):
        self.app = _new_app()

    def test_snapshot_ok_with_documents(self):
        word = FakeWordApp(normal_texts=["doc"], protected_texts=[])
        snap = self.app._word_doc_snapshot(word)
        self.assertEqual(snap["documents"]["status"], "ok")
        self.assertEqual(snap["documents"]["value"], 1)
        self.assertEqual(snap["protected_view"]["status"], "ok")
        self.assertEqual(snap["protected_view"]["value"], 0)
        self.assertEqual(snap["windows"]["status"], "ok")
        self.assertEqual(snap["windows"]["value"], 1)
        self.assertEqual(snap["active_document"]["status"], "ok")
        self.assertTrue(snap["active_document"]["present"])

    def test_snapshot_reports_error_when_count_read_fails(self):
        """Contagem falhou ≠ instância vazia: status 'error' com hresult."""
        word = ExplodingWord(FakeComError("COM dead", hresult=-1))
        snap = self.app._word_doc_snapshot(word)
        for key in ("documents", "protected_view", "windows", "active_document"):
            self.assertEqual(snap[key]["status"], "error", f"key {key} deveria ser 'error'")
        self.assertEqual(snap["documents"]["value"], None)
        self.assertEqual(snap["documents"]["hresult"], -1)

    def test_snapshot_empty_word_reports_zero_ok(self):
        """Word de fato vazio: contagens 'ok' com valor 0 (não é erro)."""
        word = FakeWordApp()
        snap = self.app._word_doc_snapshot(word)
        self.assertEqual(snap["documents"]["status"], "ok")
        self.assertEqual(snap["documents"]["value"], 0)
        # FakeWordApp sem docs levanta 'no doc' ao acessar ActiveDocument
        self.assertEqual(snap["active_document"]["status"], "error")

    def test_snapshot_protected_view_only(self):
        word = FakeWordApp(protected_texts=["temp"])
        snap = self.app._word_doc_snapshot(word)
        self.assertEqual(snap["documents"]["value"], 0)
        self.assertEqual(snap["protected_view"]["value"], 1)
        self.assertEqual(snap["windows"]["value"], 0)

    def test_fmt_count(self):
        self.assertEqual(chat_ia.ChatApp._fmt_count({"status": "ok", "value": 3}), "3")
        self.assertEqual(
            chat_ia.ChatApp._fmt_count({"status": "error", "hresult": -1}),
            "ERR(hresult=-1)",
        )

    def test_log_word_snapshot_logs_info_and_returns_snapshot(self):
        word = FakeWordApp(normal_texts=["doc"])
        with self.assertLogs("z7.chat_ia", level="INFO") as cm:
            snap = self.app._log_word_snapshot(word, "test-label")
        self.assertEqual(snap["documents"]["value"], 1)
        self.assertTrue(any("Word snapshot [test-label]" in m for m in cm.output))


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

    @mock.patch("time.sleep")
    def test_early_stop_when_consistently_no_doc(self, sleep_mock):
        """Word vazio: todos os métodos reportam 'no document' → retry cedo.

        Sem o early-break seriam 4 sleeps (0.6s+1.2s+1.8s+2.4s); com ele,
        apenas 1 (estado determinístico não merece backoff completo).
        """
        word = FakeWordApp()  # sem documento algum
        with self.assertRaises(Exception) as ctx:
            self.app._read_word_doc_text(word)
        self.assertIn("nenhum documento", str(ctx.exception).lower())
        self.assertEqual(sleep_mock.call_count, 1)

    @mock.patch("time.sleep")
    def test_logs_failure_summary_with_categories(self, _sleep):
        """Resumo de falha deve registrar métodos e categorias de erro."""
        word = FakeWordApp()
        with self.assertLogs("z7.chat_ia", level="WARNING") as cm:
            with self.assertRaises(Exception):
                self.app._read_word_doc_text(word)
        joined = "\n".join(cm.output)
        self.assertTrue(any("Document read FAILED" in m for m in cm.output))
        self.assertIn("no_doc", joined)
        self.assertIn("ActiveDocument", joined)

    @mock.patch("time.sleep")
    def test_logs_method_failures_at_info(self, _sleep):
        """Falhas individuais de método devem ser visíveis (INFO) com categoria."""
        word = FakeWordApp()
        with self.assertLogs("z7.chat_ia", level="INFO") as cm:
            with self.assertRaises(Exception):
                self.app._read_word_doc_text(word)
        self.assertTrue(any("Read method ActiveDocument failed [no_doc]" in m for m in cm.output))


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
# _get_word_app (fallback chain com verificação de documentos)
# ---------------------------------------------------------------------------
class TestGetWordApp(unittest.TestCase):
    """Valida que _get_word_app verifica documentos em cada fallback."""

    def setUp(self):
        self.app = _new_app()

    def test_get_active_object_skipped_when_empty_then_get_object_succeeds(self):
        """GetActiveObject retorna instância vazia → código tenta GetObject."""
        empty_word = FakeWordApp()
        good_word = FakeWordApp(normal_texts=["Documento via GetObject."])

        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}), \
             mock.patch.object(chat_ia.ChatApp, "_find_word_with_documents", return_value=None), \
             mock.patch("win32com.client.GetActiveObject", return_value=empty_word) as gao, \
             mock.patch("win32com.client.GetObject", return_value=good_word) as go:
            result = self.app._get_word_app()
        self.assertIs(result, good_word)
        gao.assert_called_once()
        go.assert_called_once()

    def test_rot_returns_empty_then_get_active_object_with_docs(self):
        """ROT encontra instância vazia → GetActiveObject com docs é usada."""
        empty_rot = FakeWordApp()
        good_gao = FakeWordApp(normal_texts=["Doc via fallback."])

        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}), \
             mock.patch.object(chat_ia.ChatApp, "_find_word_with_documents", return_value=empty_rot), \
             mock.patch("win32com.client.GetActiveObject", return_value=good_gao):
            result = self.app._get_word_app()
        self.assertIs(result, good_gao)

    def test_all_fallbacks_empty_returns_last_resort(self):
        """Todos os fallbacks retornam instância vazia → último recurso retorna instância."""
        empty = FakeWordApp()
        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}), \
             mock.patch.object(chat_ia.ChatApp, "_find_word_with_documents", return_value=None), \
             mock.patch("win32com.client.GetActiveObject", return_value=empty), \
             mock.patch("win32com.client.GetObject", return_value=empty):
            result = self.app._get_word_app()
        # Último recurso: GetActiveObject retorna a instância (vazia, mas existente)
        self.assertIsNotNone(result)

    def test_word_not_running_raises(self):
        """Word não está instalado/rodando → lança exceção."""
        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}), \
             mock.patch.object(chat_ia.ChatApp, "_find_word_with_documents", return_value=None), \
             mock.patch("win32com.client.GetActiveObject", side_effect=Exception("not running")), \
             mock.patch("win32com.client.GetObject", side_effect=Exception("not running")):
            with self.assertRaises(Exception) as ctx:
                self.app._get_word_app()
        self.assertIn("Não foi possível", str(ctx.exception))




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
        # Nenhum backup deve ser disparado na inicialização
        word.Application.Run.assert_not_called()

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

    def test_logs_success_summary(self):
        """Carga bem-sucedida deve emitir resumo estruturado [doc-load] SUCCESS."""
        word = FakeWordApp(normal_texts=["Texto do documento."])
        with mock.patch.object(chat_ia.ChatApp, "_get_word_app", return_value=word):
            with self.assertLogs("z7.chat_ia", level="INFO") as cm:
                self._run()
        self.assertTrue(any("[doc-load] SUCCESS" in m for m in cm.output))
        self.assertTrue(any("[doc-load] START" in m for m in cm.output))

    @mock.patch("time.sleep")
    def test_logs_failure_summary(self, _sleep):
        """Falha deve emitir [doc-load] FAILURE com duração total."""
        word = FakeWordApp()  # vazio → 'nenhum documento'
        with mock.patch.object(chat_ia.ChatApp, "_get_word_app", return_value=word):
            with self.assertLogs("z7.chat_ia", level="ERROR") as cm:
                self._run()
        self.assertTrue(any("[doc-load] FAILURE" in m for m in cm.output))

    @mock.patch("time.sleep")
    def test_logs_word_snapshot_on_connection(self, _sleep):
        """Cada conexão bem-sucedida deve registrar um Word snapshot."""
        word = FakeWordApp(normal_texts=["Texto do documento."])
        with mock.patch.object(chat_ia.ChatApp, "_get_word_app", return_value=word):
            with self.assertLogs("z7.chat_ia", level="INFO") as cm:
                self._run()
        self.assertTrue(any("Word snapshot [cycle 1 connected]" in m for m in cm.output))


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
# _init_ai_thread — decisão de pré-popular o contexto do documento
# ---------------------------------------------------------------------------
class TestInitAiThreadDocContext(unittest.TestCase):
    """Cobre a decisão de enviar (ou não) o contexto do documento à IA."""

    def setUp(self):
        self.app = _new_app()
        self.app.root = mock.MagicMock()
        self.app.messages = []
        self.app._set_word_status = mock.MagicMock()
        self.app.update_status = mock.MagicMock()
        self.app.append_message = mock.MagicMock()

    def _run_init(self):
        with mock.patch.dict(sys.modules, {"pythoncom": mock.MagicMock()}), \
             mock.patch.object(chat_ia, "_configure_ssl_certifi"), \
             mock.patch.object(chat_ia, "get_api_key", return_value="fake-key"), \
             mock.patch("openai.OpenAI"), \
             mock.patch("config_prompt.load_ai_model", return_value="deepseek/deepseek-v4-pro"), \
             mock.patch("config_prompt.load_chat_system_prompt", return_value="system prompt"):
            self.app._init_ai_thread()

    def test_pres_seeds_context_when_document_loaded(self):
        self.app.doc_text = "Texto da propositura."
        self.app._doc_truncated = False
        self.app._doc_load_error = ""
        self._run_init()
        self.assertEqual(len(self.app.messages), 2)
        self.assertEqual(self.app.messages[0]["role"], "user")
        self.assertIn("Texto da propositura.", self.app.messages[0]["content"])
        self.assertIn("Contexto do documento carregado", self.app.initial_greeting)

    def test_no_context_when_load_error(self):
        self.app.doc_text = "Nenhum documento ativo ou erro ao obter texto do Word."
        self.app._doc_load_error = "O Microsoft Word está aberto, mas nenhum documento foi encontrado."
        self._run_init()
        self.assertEqual(self.app.messages, [])
        self.assertIn("Não consegui acessar o documento", self.app.initial_greeting)

    def test_blank_document_greeting(self):
        self.app.doc_text = ""
        self.app._doc_load_error = ""
        self._run_init()
        self.assertEqual(self.app.messages, [])
        self.assertIn("em branco", self.app.initial_greeting)

    def test_truncated_document_marks_truncation_notice(self):
        self.app.doc_text = "Texto longo."
        self.app._doc_truncated = True
        self.app._doc_load_error = ""
        self._run_init()
        self.assertEqual(len(self.app.messages), 2)


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
