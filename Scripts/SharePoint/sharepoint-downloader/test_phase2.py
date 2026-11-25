"""Fase 2: disjuntor de rede, sessao por thread e recuperacao da concorrencia."""
import sqlite3
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import requests

import bulk_parallel
import browser_bulk as bulk


class ParallelBase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.db = sqlite3.connect(':memory:')
        self.addCleanup(self.db.close)
        self.db.execute('CREATE TABLE files (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT, status TEXT)')
        self.db.execute('CREATE TABLE inventory (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT)')

    def add_files(self, count):
        for index in range(count):
            self.db.execute('INSERT INTO inventory VALUES (?,?,?,?)',
                            (f'/r/{index:03}.bin', str(self.root / f'{index:03}.bin'), 3, 'v1'))

    def run_parallel(self, transfer, workers=1, **kwargs):
        page = Mock()
        page.context.cookies.return_value = [{'name': 'FedAuth', 'value': 'x', 'domain': 'tenant.sharepoint.com', 'path': '/'}]
        total = self.db.execute('SELECT COUNT(*) FROM inventory').fetchone()[0]
        with patch('builtins.print'):
            return bulk_parallel.run_parallel(self.db, page, 'https://tenant.sharepoint.com', self.root,
                                              bulk.Progress(total, total * 3), Mock(), workers, transfer, **kwargs)


class BreakerTests(ParallelBase):
    def test_outage_does_not_consume_attempts(self):
        self.add_files(3)
        calls = {'n': 0}
        lock = threading.Lock()

        def transfer(session, url, target, size, progress=None, **kwargs):
            with lock:
                calls['n'] += 1
                failing = calls['n'] <= 6  # mais que 3 tentativas do mesmo arquivo
            if failing:
                raise requests.ConnectionError('rede caiu')
            target.write_bytes(b'abc')
            return 3

        probe = Mock(side_effect=[False, True] + [True] * 10)
        breaker = bulk_parallel.NetworkBreaker(threshold=1, first_wait=0, max_wait=0)
        self.assertEqual(self.run_parallel(transfer, probe=probe, breaker=breaker), (3, 0, 0))
        self.assertGreaterEqual(probe.call_count, 2)
        self.assertFalse(breaker.open)

    def test_breaker_state_machine(self):
        now = [100.0]
        breaker = bulk_parallel.NetworkBreaker(threshold=3, first_wait=10, max_wait=40, clock=lambda: now[0])
        breaker.failure(); breaker.failure()
        self.assertFalse(breaker.open)
        breaker.success()
        breaker.failure(); breaker.failure()
        self.assertFalse(breaker.open, 'sucesso zera a contagem')
        breaker.failure()
        self.assertTrue(breaker.open)
        self.assertFalse(breaker.due())
        now[0] += 10
        self.assertTrue(breaker.due())
        breaker.probe_result(False)
        self.assertEqual(breaker.wait, 20)
        for _ in range(3):
            breaker.probe_result(False)
        self.assertEqual(breaker.wait, 40)
        breaker.probe_result(True)
        self.assertFalse(breaker.open)

    def test_isolated_network_errors_still_fail_after_three_attempts(self):
        self.add_files(1)
        transfer = Mock(side_effect=requests.ConnectionError('reset'))
        breaker = bulk_parallel.NetworkBreaker(threshold=99)
        with patch.object(bulk_parallel.time, 'monotonic', side_effect=_fast_clock()), \
                patch.object(bulk_parallel, 'KEEPALIVE_SECONDS', 10 ** 9):
            self.assertEqual(self.run_parallel(transfer, breaker=breaker, probe=Mock()), (0, 0, 1))
        self.assertEqual(transfer.call_count, 3)


class SessionAndConcurrencyTests(ParallelBase):
    def test_one_session_per_thread_with_cookies(self):
        self.add_files(5)
        seen = []

        def transfer(session, url, target, size, progress=None, **kwargs):
            seen.append(session)
            self.assertEqual(session.cookies.get('FedAuth'), 'x')
            self.assertIn('NONISV', session.headers['User-Agent'])
            target.write_bytes(b'abc')
            return 3

        self.assertEqual(self.run_parallel(transfer, workers=1, probe=Mock()), (5, 0, 0))
        self.assertEqual(len({id(s) for s in seen}), 1)

    def test_concurrency_recovers_after_throttling(self):
        self.add_files(8)
        state = {'throttled': False}
        lock = threading.Lock()

        def transfer(session, url, target, size, progress=None, **kwargs):
            with lock:
                first = not state['throttled']
                state['throttled'] = True
            if first:
                raise bulk.DownloadError('HTTP 429', 429, 0)
            target.write_bytes(b'abc')
            return 3

        with patch.object(bulk_parallel, 'RECOVERY_SUCCESSES', 2), \
                self.assertLogs(level='INFO') as logs:
            self.assertEqual(self.run_parallel(transfer, workers=3, probe=Mock()), (8, 0, 0))
        text = '\n'.join(logs.output)
        self.assertIn('LOTE_LIMITADO', text)
        self.assertIn('LOTE_CONCORRENCIA_AMPLIADA concorrencia=2', text)


def _fast_clock():
    """Relogio que avanca 5s por leitura, para os atrasos de nova tentativa."""
    value = 0.0
    while True:
        value += 5
        yield value


if __name__ == '__main__':
    unittest.main()
