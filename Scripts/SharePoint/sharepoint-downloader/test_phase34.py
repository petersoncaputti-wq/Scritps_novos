"""Fases 3 e 4: perfil persistente, deteccao de login e renovacao coordenada."""
import threading
import unittest
from unittest.mock import MagicMock, Mock, patch

import browser_bulk as bulk
import browser_test
import bulk_parallel
from test_phase2 import ParallelBase


class BrowserProfileTests(unittest.TestCase):
    def test_persistent_profile_and_fallback(self):
        pw = MagicMock()
        self.assertIs(browser_test.open_browser(pw), pw.chromium.launch_persistent_context.return_value)
        pw.chromium.launch.assert_not_called()
        pw.chromium.launch_persistent_context.side_effect = RuntimeError('perfil em uso')
        with patch('builtins.print'):
            context = browser_test.open_browser(pw)
        self.assertIs(context, pw.chromium.launch.return_value.new_context.return_value)
        pw = MagicMock()
        browser_test.open_browser(pw, fresh_login=True)
        pw.chromium.launch_persistent_context.assert_not_called()

    def test_close_browser_handles_persistent_context(self):
        context = Mock(browser=None)
        browser_test.close_browser(context)
        context.close.assert_called_once()


class LoginDetectionTests(unittest.TestCase):
    def tab(self, url, result=None, closed=False):
        tab = Mock(url=url)
        tab.is_closed.return_value = closed
        tab.evaluate.return_value = result or {}
        return tab

    def test_session_ready_rules(self):
        host = 'tenant.sharepoint.com'
        ok = self.tab('https://tenant.sharepoint.com/sites/x', {'ready': True, 'status': 200})
        denied = self.tab('https://tenant.sharepoint.com/', {'ready': True, 'status': 403})
        anonymous = self.tab('https://tenant.sharepoint.com/', {'ready': True, 'status': 403, 'forms': 'x'})
        login = self.tab('https://login.microsoftonline.com/', {'ready': True, 'status': 200})
        loading = self.tab('https://tenant.sharepoint.com/', {'ready': False})
        self.assertTrue(browser_test.session_ready(ok, host))
        self.assertTrue(browser_test.session_ready(denied, host), '403 de permissao ja e sessao valida')
        for tab in (anonymous, login, loading):
            self.assertFalse(browser_test.session_ready(tab, host))
        broken = self.tab('https://tenant.sharepoint.com/')
        broken.evaluate.side_effect = RuntimeError('Execution context was destroyed')
        self.assertFalse(browser_test.session_ready(broken, host))

    def test_wait_detects_login_without_enter(self):
        login = self.tab('https://login.microsoftonline.com/')
        done = self.tab('https://tenant.sharepoint.com/sites/x', {'ready': True, 'status': 200})
        context = Mock()
        states = iter([[login], [login], [login, done]])
        type(context).pages = property(lambda self: next(states))
        with patch('builtins.print'), patch.object(browser_test, 'read_key', return_value=None):
            self.assertIs(browser_test.wait_for_sharepoint(context, 'tenant.sharepoint.com', poll_ms=1), done)

    def test_q_cancels(self):
        context = Mock(pages=[self.tab('https://login.microsoftonline.com/')])
        with patch('builtins.print'), patch.object(browser_test, 'read_key', return_value='q'):
            with self.assertRaises(KeyboardInterrupt):
                browser_test.wait_for_sharepoint(context, 'tenant.sharepoint.com', poll_ms=1)


class TypedErrorTests(unittest.TestCase):
    def test_auth_and_not_found_classification(self):
        self.assertTrue(browser_test.is_auth_failure(browser_test.SharePointHTTPError('x', 401)))
        self.assertTrue(browser_test.is_auth_failure(browser_test.SharePointHTTPError('x', 200)))
        self.assertFalse(browser_test.is_auth_failure(RuntimeError('HTTP 401')))
        self.assertTrue(browser_test.is_not_found(browser_test.SharePointHTTPError('x', 404)))
        self.assertTrue(bulk.DownloadError('x', 403).auth)
        self.assertTrue(bulk.DownloadError('html', auth=True).auth)
        self.assertFalse(bulk.DownloadError('x', 404).auth)


class CoordinatedRenewTests(ParallelBase):
    def test_expired_session_renews_once_without_consuming_attempts(self):
        self.add_files(6)
        state = {'renewed': False}
        lock = threading.Lock()
        page = Mock()
        cookies = [{'name': 'FedAuth', 'value': 'old', 'domain': 'tenant.sharepoint.com', 'path': '/'}]
        page.context.cookies.side_effect = lambda urls: list(cookies)

        def renew():
            state['renewed'] = True
            cookies[0] = dict(cookies[0], value='new')
            return page

        renew_mock = Mock(side_effect=renew)

        def transfer(session, url, target, size, progress=None, **kwargs):
            if session.cookies.get('FedAuth') != 'new':
                raise bulk.DownloadError('Download retornou HTML; a sessao pode ter expirado.', auth=True)
            target.write_bytes(b'abc')
            return 3

        with patch('builtins.print'):
            result = bulk_parallel.run_parallel(self.db, page, 'https://tenant.sharepoint.com', self.root,
                                                bulk.Progress(6, 18), renew_mock, 3, transfer, probe=Mock())
        self.assertEqual(result, (6, 0, 0))
        self.assertEqual(renew_mock.call_count, 1)

    def test_permission_denied_counts_attempts(self):
        self.add_files(1)
        transfer = Mock(side_effect=bulk.DownloadError('HTTP 403', 403, 0))
        renew = Mock(side_effect=lambda: page)
        page = Mock()
        page.context.cookies.return_value = []
        with patch('builtins.print'):
            result = bulk_parallel.run_parallel(self.db, page, 'https://tenant.sharepoint.com', self.root,
                                                bulk.Progress(1, 3), renew, 1, transfer, probe=Mock())
        self.assertEqual(result, (0, 0, 1))
        self.assertEqual(transfer.call_count, 3)


if __name__ == '__main__':
    unittest.main()
