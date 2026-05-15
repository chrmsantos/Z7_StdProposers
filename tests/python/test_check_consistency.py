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


# ===========================================================================
# Infraestrutura compartilhada para testes de fluxo completo
# ===========================================================================

class _SyncThread:
    """Substituto de threading.Thread que executa o target de forma síncrona."""
    def __init__(self, target=None, daemon=False):
        self._target = target

    def start(self):
        if self._target:
            self._target()

    def is_alive(self):
        return False


_FAKE_THEME_COLORS = {
    "bg": "#1e1e2e", "fg": "#cdd6f4", "text_bg": "#313244",
    "border": "#45475a", "btn_primary_bg": "#2563eb", "btn_primary_hover": "#1d4ed8",
}


def _make_word_stubs(doc_text: str):
    """Devolve (stubs_dict, word_mock) configurados para o texto informado."""
    doc_mock = mock.MagicMock()
    doc_mock.Range.return_value.Text = doc_text
    word_mock = mock.MagicMock()
    word_mock.ActiveDocument = doc_mock
    win32_mock = mock.MagicMock()
    win32_mock.client.GetActiveObject.return_value = word_mock
    stubs = dict(_STUBS)
    stubs["win32com"] = win32_mock
    stubs["win32com.client"] = win32_mock.client
    return stubs, word_mock


# ===========================================================================
# 1. Tratamento de documentos inválidos / muito curtos
# ===========================================================================

class TestCheckConsistencyDocumentHandling(unittest.TestCase):
    """Testa saídas antecipadas provocadas por documentos inválidos ou curtos."""

    def _run(self, doc_text, privacy_mock=None):
        """Executa main() e devolve o mock de ask_privacy_warning."""
        import importlib
        stubs, _ = _make_word_stubs(doc_text)
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            privacy = privacy_mock or mock.MagicMock(return_value=True)
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", privacy), \
                 mock.patch("z7_theme.show_error"):
                check_consistency.main()
        return privacy

    def test_empty_document_skips_privacy_check(self):
        """Documento vazio deve sair antes do aviso de privacidade."""
        privacy = self._run("")
        privacy.assert_not_called()

    def test_whitespace_only_document_skips_privacy_check(self):
        """Documento com apenas espaços deve sair antes do aviso de privacidade."""
        privacy = self._run("   \t\n")
        privacy.assert_not_called()

    def test_nine_char_document_skips_privacy_check(self):
        """Documento com 9 caracteres (< 10) deve ser considerado muito curto."""
        privacy = self._run("123456789")
        privacy.assert_not_called()

    def test_ten_char_document_reaches_privacy_check(self):
        """Documento com exatamente 10 caracteres deve chegar ao aviso de privacidade."""
        privacy = self._run("1234567890", privacy_mock=mock.MagicMock(return_value=False))
        privacy.assert_called_once()

    def test_no_active_document_skips_privacy_check(self):
        """doc=None deve causar saída antes do aviso de privacidade."""
        import importlib
        word_mock = mock.MagicMock()
        word_mock.ActiveDocument = None
        win32_mock = mock.MagicMock()
        win32_mock.client.GetActiveObject.return_value = word_mock
        stubs = dict(_STUBS)
        stubs["win32com"] = win32_mock
        stubs["win32com.client"] = win32_mock.client
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            privacy = mock.MagicMock(return_value=True)
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", privacy), \
                 mock.patch("z7_theme.show_error"):
                check_consistency.main()
        privacy.assert_not_called()

    def test_active_document_com_error_shows_dialog(self):
        """word.ActiveDocument lançando com_error deve exibir diálogo, não silenciar.

        Regressão: 'Este comando não está disponível porque nenhum documento foi aberto.'
        é lançado como com_error — não retorna None — quando o Word está aberto mas
        sem documento ativo. O guard 'if doc is None' não captura este caso; a
        exceção deve ser convertida em None e resultar em z7_theme.show_error.
        """
        import importlib
        word_mock = mock.MagicMock()
        type(word_mock).ActiveDocument = mock.PropertyMock(
            side_effect=Exception(
                "Este comando não está disponível porque nenhum documento foi aberto."
            )
        )
        win32_mock = mock.MagicMock()
        win32_mock.client.GetActiveObject.return_value = word_mock
        stubs = dict(_STUBS)
        stubs["win32com"] = win32_mock
        stubs["win32com.client"] = win32_mock.client
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            privacy = mock.MagicMock(return_value=True)
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", privacy), \
                 mock.patch("z7_theme.show_error") as mock_err:
                check_consistency.main()
        privacy.assert_not_called()
        mock_err.assert_called_once()

    def test_short_document_sets_word_status_bar(self):
        """StatusBar do Word deve ser atualizado quando o documento é muito curto."""
        import importlib
        stubs, word_mock = _make_word_stubs("Curto")
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning"), \
                 mock.patch("z7_theme.show_error"):
                check_consistency.main()
        status = str(word_mock.StatusBar)
        self.assertTrue(
            "vazio" in status.lower() or "curto" in status.lower(),
            f"StatusBar não refletiu documento curto. Valor: {status}",
        )


# ===========================================================================
# 2. Fluxo de privacidade e chave de API
# ===========================================================================

class TestCheckConsistencyPrivacyAndKeyFlow(unittest.TestCase):
    """Testa aborts causados por recusa de privacidade ou chave de API ausente."""

    _GOOD_TEXT = "Texto suficientemente longo para análise completa da propositura."

    def _run_until_api(self, privacy_ok: bool, api_key):
        """Executa main() até o ponto de chamar (ou não) a API Gemini."""
        import importlib
        stubs, _ = _make_word_stubs(self._GOOD_TEXT)
        genai_mock = mock.MagicMock()
        stubs["google.genai"] = genai_mock
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", return_value=privacy_ok), \
                 mock.patch.object(check_consistency, "get_api_key", return_value=api_key), \
                 mock.patch("z7_theme.show_error"):
                check_consistency.main()
        return genai_mock

    def test_privacy_declined_does_not_call_gemini_client(self):
        """Recusar aviso de privacidade deve impedir criação do cliente Gemini."""
        genai_mock = self._run_until_api(privacy_ok=False, api_key="qualquer")
        genai_mock.Client.assert_not_called()

    def test_privacy_declined_does_not_call_get_api_key(self):
        """Recusar aviso de privacidade deve impedir chamada a get_api_key."""
        import importlib
        stubs, _ = _make_word_stubs(self._GOOD_TEXT)
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", return_value=False), \
                 mock.patch.object(check_consistency, "get_api_key") as mock_get_key, \
                 mock.patch("z7_theme.show_error"):
                check_consistency.main()
        mock_get_key.assert_not_called()

    def test_missing_api_key_does_not_call_gemini_client(self):
        """get_api_key retornando None deve impedir criação do cliente Gemini."""
        genai_mock = self._run_until_api(privacy_ok=True, api_key=None)
        genai_mock.Client.assert_not_called()

    def test_empty_api_key_does_not_call_gemini_client(self):
        """get_api_key retornando string vazia deve impedir criação do cliente Gemini."""
        genai_mock = self._run_until_api(privacy_ok=True, api_key="")
        genai_mock.Client.assert_not_called()


# ===========================================================================
# 3. Fallback de conexão com o Word
# ===========================================================================

class TestCheckConsistencyWordConnectionFallback(unittest.TestCase):
    """Testa o mecanismo de fallback de GetActiveObject → GetObject."""

    def test_fallback_to_get_object_when_get_active_object_fails(self):
        """Se GetActiveObject falhar, deve tentar GetObject como fallback."""
        import importlib
        doc_mock = mock.MagicMock()
        doc_mock.Range.return_value.Text = "Texto longo o suficiente para análise."
        word_mock = mock.MagicMock()
        word_mock.ActiveDocument = doc_mock
        win32_mock = mock.MagicMock()
        win32_mock.client.GetActiveObject.side_effect = Exception("COM error")
        win32_mock.client.GetObject.return_value = word_mock
        stubs = dict(_STUBS)
        stubs["win32com"] = win32_mock
        stubs["win32com.client"] = win32_mock.client
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", return_value=False), \
                 mock.patch("z7_theme.show_error"):
                check_consistency.main()
        win32_mock.client.GetActiveObject.assert_called_once_with("Word.Application")
        win32_mock.client.GetObject.assert_called_once()

    def test_both_connection_methods_fail_shows_error(self):
        """Se GetActiveObject e GetObject falharem, deve exibir show_error."""
        import importlib
        win32_mock = mock.MagicMock()
        win32_mock.client.GetActiveObject.side_effect = Exception("COM error 1")
        win32_mock.client.GetObject.side_effect = Exception("COM error 2")
        stubs = dict(_STUBS)
        stubs["win32com"] = win32_mock
        stubs["win32com.client"] = win32_mock.client
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.show_error") as mock_show_error, \
                 mock.patch("z7_theme.ask_privacy_warning"):
                check_consistency.main()
        mock_show_error.assert_called_once()


# ===========================================================================
# 4. Tratamento de erros da API e respostas da IA
# ===========================================================================

class TestCheckConsistencyApiErrorHandling(unittest.TestCase):
    """Testa o comportamento diante de erros da API Gemini e respostas especiais."""

    _GOOD_TEXT = "Texto suficientemente longo para análise completa da propositura legislativa."

    def _run_full_main(self, api_response=None, api_exception=None, ask_ok_cancel=True):
        """
        Executa main() com threading síncrono e tkinter mockado.
        Devolve dict com os mocks das funções de UI do z7_theme.
        """
        import importlib
        stubs, _ = _make_word_stubs(self._GOOD_TEXT)
        genai_mock = mock.MagicMock()
        if api_exception is not None:
            genai_mock.Client.return_value.models.generate_content.side_effect = api_exception
        else:
            resp = mock.MagicMock()
            resp.text = api_response if api_response is not None else ""
            genai_mock.Client.return_value.models.generate_content.return_value = resp
        google_mock = mock.MagicMock()
        google_mock.genai = genai_mock
        stubs["google"] = google_mock
        stubs["google.genai"] = genai_mock

        captured = {}
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            mock_tk.Toplevel.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", return_value=True), \
                 mock.patch("z7_theme.get_theme_colors", return_value=_FAKE_THEME_COLORS), \
                 mock.patch("z7_theme.show_error") as m_err, \
                 mock.patch("z7_theme.show_info") as m_info, \
                 mock.patch("z7_theme.show_warning") as m_warn, \
                 mock.patch("z7_theme.ask_ok_cancel", return_value=ask_ok_cancel) as m_ok_cancel, \
                 mock.patch.object(check_consistency, "get_api_key", return_value="test_key"), \
                 mock.patch.object(check_consistency, "delete_api_key") as m_delete, \
                 mock.patch("threading.Thread", _SyncThread):
                check_consistency.main()
            captured["show_error"] = m_err
            captured["show_info"] = m_info
            captured["show_warning"] = m_warn
            captured["ask_ok_cancel"] = m_ok_cancel
            captured["delete_api_key"] = m_delete
        return captured

    def test_empty_api_response_shows_warning(self):
        """Resposta vazia da API deve exibir show_warning."""
        captured = self._run_full_main(api_response="")
        captured["show_warning"].assert_called_once()

    def test_no_issues_response_shows_info(self):
        """Resposta sem inconsistências deve exibir show_info (não show_error)."""
        captured = self._run_full_main(
            api_response="Sem inconsistências graves detectadas. O documento está correto."
        )
        captured["show_info"].assert_called_once()
        captured["show_error"].assert_not_called()

    def test_issues_response_calls_show_issues_window(self):
        """Resposta com inconsistências deve chamar _show_issues_window."""
        import importlib
        stubs, _ = _make_word_stubs(self._GOOD_TEXT)
        genai_mock = mock.MagicMock()
        resp = mock.MagicMock()
        resp.text = "1. Ementa cita 'Rua X' mas artigo 2º cita 'Av. Y'."
        genai_mock.Client.return_value.models.generate_content.return_value = resp
        google_mock = mock.MagicMock()
        google_mock.genai = genai_mock
        stubs["google"] = google_mock
        stubs["google.genai"] = genai_mock
        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            mock_tk.Toplevel.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", return_value=True), \
                 mock.patch("z7_theme.get_theme_colors", return_value=_FAKE_THEME_COLORS), \
                 mock.patch("z7_theme.show_error"), \
                 mock.patch("z7_theme.show_info"), \
                 mock.patch.object(check_consistency, "get_api_key", return_value="key"), \
                 mock.patch.object(check_consistency, "_show_issues_window") as mock_show, \
                 mock.patch("threading.Thread", _SyncThread):
                check_consistency.main()
        mock_show.assert_called_once()
        _, call_kwargs = mock_show.call_args[0], mock_show.call_args[1] if mock_show.call_args[1] else {}
        args = mock_show.call_args[0]
        self.assertIn("1. Ementa cita", args[1])

    def test_generic_api_error_shows_error_dialog(self):
        """Erro genérico da API deve exibir show_error."""
        captured = self._run_full_main(api_exception=Exception("Connection timeout"))
        captured["show_error"].assert_called()

    def test_generic_api_error_does_not_prompt_key_deletion(self):
        """Erro genérico (sem 401/403) não deve perguntar sobre excluir a chave."""
        captured = self._run_full_main(api_exception=Exception("timeout"))
        captured["ask_ok_cancel"].assert_not_called()

    def test_403_api_error_prompts_key_deletion_dialog(self):
        """Erro 403 deve perguntar se o usuário deseja excluir a chave inválida."""
        captured = self._run_full_main(
            api_exception=Exception("403 PERMISSION_DENIED: API key not valid")
        )
        captured["ask_ok_cancel"].assert_called_once()

    def test_401_api_error_prompts_key_deletion_dialog(self):
        """Erro 401 deve perguntar se o usuário deseja excluir a chave inválida."""
        captured = self._run_full_main(
            api_exception=Exception("401 UNAUTHENTICATED: invalid api_key provided")
        )
        captured["ask_ok_cancel"].assert_called_once()

    def test_403_error_with_confirmation_deletes_api_key(self):
        """Confirmar exclusão após erro 403 deve chamar delete_api_key."""
        captured = self._run_full_main(
            api_exception=Exception("403 invalid api key"), ask_ok_cancel=True
        )
        captured["delete_api_key"].assert_called_once()

    def test_403_error_without_confirmation_keeps_api_key(self):
        """Recusar exclusão após erro 403 não deve chamar delete_api_key."""
        captured = self._run_full_main(
            api_exception=Exception("403 invalid api key"), ask_ok_cancel=False
        )
        captured["delete_api_key"].assert_not_called()


# ===========================================================================
# 5. Lógica de has_issues (testes estendidos)
# ===========================================================================

class TestCheckConsistencyHasIssuesExtended(unittest.TestCase):
    """Testes adicionais para a lógica de detecção do marcador de sem-problemas."""

    def test_marker_detection_is_case_insensitive(self):
        """_NO_ISSUE_MARKER deve ser detectado independente de maiúsculas/minúsculas."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        self.assertFalse(marker not in marker.upper().lower())

    def test_marker_in_mixed_case_not_triggers_has_issues(self):
        """Marcador escrito com capitalização mista não deve acionar has_issues."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        capitalized = marker.capitalize()
        analysis = f"Análise concluída. {capitalized}."
        has_issues = marker not in analysis.lower()
        self.assertFalse(has_issues)

    def test_partial_marker_still_triggers_has_issues(self):
        """Presença apenas de parte do marcador deve manter has_issues=True."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        partial = marker[:6]
        analysis = f"Análise: {partial} algo encontrado."
        has_issues = marker not in analysis.lower()
        self.assertTrue(has_issues)

    def test_marker_embedded_in_long_response_is_detected(self):
        """Marcador no meio de resposta longa deve ser detectado corretamente."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        analysis = (
            "Realizei a análise linha a linha da propositura. "
            f"{marker.capitalize()}. "
            "O documento está internamente consistente e bem estruturado."
        )
        has_issues = marker not in analysis.lower()
        self.assertFalse(has_issues)

    def test_numbered_list_of_issues_triggers_has_issues(self):
        """Lista numerada de inconsistências deve resultar em has_issues=True."""
        mod = _import_check_consistency()
        marker = mod._NO_ISSUE_MARKER
        analysis = (
            "Foram identificadas as seguintes inconsistências:\n"
            "1. A ementa menciona 'Rua das Flores' mas o artigo 2º cita 'Av. Central'.\n"
            "2. O valor de R$ 10.000 no título difere de R$ 15.000 no corpo do texto."
        )
        has_issues = marker not in analysis.lower()
        self.assertTrue(has_issues)

    def test_marker_not_empty_string(self):
        """_NO_ISSUE_MARKER não deve ser uma string vazia."""
        mod = _import_check_consistency()
        self.assertGreater(len(mod._NO_ISSUE_MARKER.strip()), 5)


# ===========================================================================
# 6. Formato do prompt enviado ao Gemini
# ===========================================================================

class TestCheckConsistencyPromptFormat(unittest.TestCase):
    """Verifica a estrutura do prompt construído em main()."""

    def _src(self):
        return (PY_ROOT / "check_consistency.py").read_text(encoding="utf-8")

    def test_prompt_contains_inicio_da_propositura_marker(self):
        """O prompt deve delimitar o início da propositura com um marcador textual."""
        self.assertIn("INICIO DA PROPOSITURA", self._src())

    def test_prompt_contains_fim_da_propositura_marker(self):
        """O prompt deve delimitar o fim da propositura com um marcador textual."""
        self.assertIn("FIM DA PROPOSITURA", self._src())

    def test_prompt_includes_full_text_variable(self):
        """O prompt deve incorporar a variável full_text (conteúdo do documento)."""
        src = self._src()
        self.assertIn("full_text", src)

    def test_prompt_combines_base_prompt_and_document(self):
        """O prompt final deve combinar base_prompt com o texto do documento."""
        src = self._src()
        self.assertIn("base_prompt", src)
        # Deve existir uma concatenação ou f-string que una ambos
        self.assertTrue(
            "f\"{base_prompt}" in src or 'base_prompt + ' in src or "base_prompt\\n" in src,
            "base_prompt não parece ser combinado com o texto do documento.",
        )

    def test_gemini_client_receives_prompt_with_full_text(self):
        """generate_content deve ser chamado com o texto do documento incluído."""
        import importlib
        good_text = "Propositura de teste com conteúdo suficiente para validação."
        stubs, _ = _make_word_stubs(good_text)
        genai_mock = mock.MagicMock()
        resp = mock.MagicMock()
        resp.text = "Sem inconsistências graves detectadas."
        genai_mock.Client.return_value.models.generate_content.return_value = resp
        google_mock = mock.MagicMock()
        google_mock.genai = genai_mock
        stubs["google"] = google_mock
        stubs["google.genai"] = genai_mock

        with mock.patch.dict(sys.modules, stubs):
            import check_consistency
            importlib.reload(check_consistency)
            mock_tk = mock.MagicMock()
            mock_tk.Tk.return_value = mock.MagicMock()
            mock_tk.Toplevel.return_value = mock.MagicMock()
            with mock.patch.object(check_consistency, "tk", mock_tk), \
                 mock.patch("z7_theme.ask_privacy_warning", return_value=True), \
                 mock.patch("z7_theme.get_theme_colors", return_value=_FAKE_THEME_COLORS), \
                 mock.patch("z7_theme.show_info"), \
                 mock.patch.object(check_consistency, "get_api_key", return_value="key"), \
                 mock.patch("threading.Thread", _SyncThread):
                check_consistency.main()

        generate_call = genai_mock.Client.return_value.models.generate_content
        generate_call.assert_called_once()
        call_kwargs = generate_call.call_args
        contents_arg = call_kwargs[1].get("contents") or (call_kwargs[0][1] if len(call_kwargs[0]) > 1 else "")
        self.assertIn(good_text, str(contents_arg))


if __name__ == "__main__":
    unittest.main()
