"""Fase 1: tamanho Office ajustado e reaproveitamento sem releitura completa."""
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import browser_bulk as bulk
import bulk_transfer


def fake_session(status, headers, chunks, url='https://tenant.sharepoint.com/file'):
    session = Mock()
    response = Mock(status_code=status, headers=headers, url=url)
    response.iter_content.return_value = chunks
    session.get.return_value.__enter__ = Mock(return_value=response)
    session.get.return_value.__exit__ = Mock(return_value=False)
    return session, url


class OfficeSizeTests(unittest.TestCase):
    def test_office_accepts_content_length_with_properties(self):
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / 'planilha.xlsx'
            session, url = fake_session(200, {'ETag': '"v1"', 'Content-Length': '7'}, [b'abcdefg'])
            self.assertEqual(bulk.transfer(session, url, target, 4, resume_key='k'), 7)
            self.assertEqual(target.read_bytes(), b'abcdefg')

    def test_office_accepts_growth_without_content_length(self):
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / 'texto.docx'
            session, url = fake_session(200, {'ETag': '"v1"'}, [b'abcdef'])
            self.assertEqual(bulk.transfer(session, url, target, 4, resume_key='k'), 6)

    def test_non_office_and_shrink_still_fail(self):
        cases = (('dados.bin', [b'abcdef']), ('texto.docx', [b'ab']))
        for name, chunks in cases:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as temp:
                target = Path(temp) / name
                session, url = fake_session(200, {'ETag': '"v1"'}, chunks)
                with self.assertRaises(bulk.DownloadError):
                    bulk.transfer(session, url, target, 4, resume_key='k')
                self.assertFalse(target.exists())

    def test_office_growth_above_limit_fails(self):
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / 'planilha.xlsx'
            big = 4 + bulk_transfer.OFFICE_MAX_DELTA + 1
            session, url = fake_session(200, {'ETag': '"v1"', 'Content-Length': str(big)}, [b'x' * big])
            with self.assertRaises(bulk.DownloadError):
                bulk.transfer(session, url, target, 4, resume_key='k')

    def test_office_resume_uses_adjusted_total(self):
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / 'planilha.xlsx'
            partial = target.with_name(target.name + '.spdownload.part')
            partial.write_bytes(b'abc')
            partial.with_name(partial.name + '.json').write_text(json.dumps(
                {'identity': 'k', 'size': 4, 'expected': 6, 'etag': '"v1"'}))
            session, url = fake_session(206, {'ETag': '"v1"', 'Content-Range': 'bytes 3-5/6'}, [b'def'])
            self.assertEqual(bulk.transfer(session, url, target, 4, resume_key='k'), 6)
            self.assertEqual(session.get.call_args.kwargs['headers']['Range'], 'bytes=3-')
            self.assertEqual(target.read_bytes(), b'abcdef')


class SkipExistingTests(unittest.TestCase):
    def run_once(self, root, db, transfer=None, verify_existing=False):
        from bulk_parallel import run_parallel
        page = Mock()
        page.context.cookies.return_value = []
        transfer = transfer or Mock(side_effect=AssertionError('nao deve baixar'))
        with patch('builtins.print'):
            return run_parallel(db, page, 'https://tenant.sharepoint.com', root, bulk.Progress(1, 4),
                                Mock(), 2, transfer, verify_existing)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.db = sqlite3.connect(':memory:')
        self.addCleanup(self.db.close)
        self.db.execute('CREATE TABLE files (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT, status TEXT)')
        self.db.execute('CREATE TABLE inventory (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT)')
        self.target = self.root / 'planilha.xlsx'
        self.db.execute('INSERT INTO inventory VALUES (?,?,?,?)', ('/r/planilha.xlsx', str(self.target), 4, 'v1'))

    def test_downloaded_office_file_is_reused_without_full_read(self):
        def transfer(session, url, target, size, progress=None, **kwargs):
            target.write_bytes(b'abcdef')
            return 6
        self.assertEqual(self.run_once(self.root, self.db, transfer), (1, 0, 0))
        row = self.db.execute('SELECT status,local_size,mtime_ns FROM files').fetchone()
        self.assertEqual(row[:2], ('done', 6))
        self.assertIsNotNone(row[2])
        with patch.object(bulk_transfer, 'verify_local_file', side_effect=AssertionError('releu')):
            self.assertEqual(self.run_once(self.root, self.db), (0, 1, 0))

    def test_legacy_record_trusted_then_mtime_change_forces_read(self):
        self.target.write_bytes(b'abcd')
        self.db.execute("INSERT INTO files VALUES ('/r/planilha.xlsx', ?, 4, 'v1', 'done')", (str(self.target),))
        with patch.object(bulk_transfer, 'verify_local_file', side_effect=AssertionError('releu')):
            self.assertEqual(self.run_once(self.root, self.db), (0, 1, 0))
        self.db.execute('UPDATE files SET mtime_ns = mtime_ns - 1')
        with patch.object(bulk_transfer, 'verify_local_file') as verify:
            self.assertEqual(self.run_once(self.root, self.db), (0, 1, 0))
            verify.assert_called_once()

    def test_verify_existing_flag_forces_full_read(self):
        self.target.write_bytes(b'abcd')
        self.db.execute("INSERT INTO files VALUES ('/r/planilha.xlsx', ?, 4, 'v1', 'done')", (str(self.target),))
        with patch.object(bulk_transfer, 'verify_local_file') as verify:
            self.assertEqual(self.run_once(self.root, self.db, verify_existing=True), (0, 1, 0))
            verify.assert_called_once()

    def test_other_version_is_still_conflict(self):
        self.target.write_bytes(b'abcd')
        self.db.execute("INSERT INTO files VALUES ('/r/planilha.xlsx', ?, 4, 'v0', 'done')", (str(self.target),))
        self.assertEqual(self.run_once(self.root, self.db), (0, 0, 1))
        self.assertEqual(self.target.read_bytes(), b'abcd')


if __name__ == '__main__':
    unittest.main()
