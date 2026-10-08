import tempfile
import json
import sqlite3
import threading
import logging
import os
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import browser_bulk as bulk


class BulkTests(unittest.TestCase):
    def test_http400_uses_direct_route_but_auth_failure_does_not(self):
        from bulk_parallel import transfer_with_fallback
        transfer = Mock(side_effect=[bulk.DownloadError('HTTP 400', 400), None])
        transfer_with_fallback(Mock(), 'https://tenant.sharepoint.com/sites/project',
                               '/sites/project/library/a (1).xlsx', Path('a.xlsx'), 3,
                               None, 'v1', threading.Event(), transfer)
        self.assertEqual(transfer.call_count, 2)
        direct_url = transfer.call_args.args[1]
        self.assertEqual(direct_url, 'https://tenant.sharepoint.com/sites/project/library/a%20%281%29.xlsx')
        transfer = Mock(side_effect=bulk.DownloadError('HTTP 403', 403))
        with self.assertRaises(bulk.DownloadError):
            transfer_with_fallback(Mock(), 'https://tenant.sharepoint.com/sites/project',
                                   '/sites/project/library/a.xlsx', Path('a.xlsx'), 3,
                                   None, 'v1', threading.Event(), transfer)
        self.assertEqual(transfer.call_count, 1)

    def test_progress_stays_on_one_line_and_fits_terminal(self):
        progress = bulk.Progress(2, 10)
        with patch.object(bulk.shutil, 'get_terminal_size', return_value=os.terminal_size((80, 24))), patch('builtins.print') as output:
            progress.message('Arquivo ' + 'x' * 200)
            progress.show(force=True)
            progress.complete(5)
            progress.show(force=True)
            for call in output.call_args_list:
                self.assertEqual(call.kwargs.get('end'), '')
                self.assertTrue(call.args[0].startswith('\r'))
                self.assertNotIn('\n', call.args[0])
                self.assertLessEqual(len(call.args[0]), 80)
            progress.close()
            self.assertEqual(output.call_args.args, ())

    def test_interruption_preserves_partial_and_invalid_range_does_not_append(self):
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / 'file.bin'
            session = Mock()
            url = 'https://tenant.sharepoint.com/file'
            response = Mock(status_code=200, headers={'ETag': '"v1"'}, url=url)
            def chunks():
                yield b'ab'
                raise bulk.requests.ConnectionError('connection lost')
            response.iter_content.return_value = chunks()
            session.get.return_value.__enter__ = Mock(return_value=response)
            session.get.return_value.__exit__ = Mock(return_value=False)
            with self.assertRaises(bulk.requests.ConnectionError):
                bulk.transfer(session, url, target, 4, resume_key='key')
            partial = target.with_name(target.name + '.spdownload.part')
            self.assertEqual(partial.read_bytes(), b'ab')
            response.status_code = 206
            response.headers = {'ETag': '"v2"', 'Content-Range': 'bytes 2-3/4'}
            with self.assertRaises(bulk.DownloadError):
                bulk.transfer(session, url, target, 4, resume_key='key')
            self.assertEqual(partial.read_bytes(), b'ab')
            self.assertFalse(target.exists())

    def test_range_resume_and_server_restart(self):
        for status in (206, 200):
            with self.subTest(status=status), tempfile.TemporaryDirectory() as temp:
                target = Path(temp) / 'file.bin'
                partial = target.with_name(target.name + '.spdownload.part')
                partial.write_bytes(b'ab')
                partial.with_name(partial.name + '.json').write_text(json.dumps({'identity': 'key', 'size': 4, 'etag': '"v1"'}))
                session = Mock()
                response = Mock(status_code=status, headers={'ETag': '"v1"', 'Content-Range': 'bytes 2-3/4'}, url='https://tenant.sharepoint.com/file')
                response.iter_content.return_value = [b'cd'] if status == 206 else [b'abcd']
                session.get.return_value.__enter__ = Mock(return_value=response)
                session.get.return_value.__exit__ = Mock(return_value=False)
                bulk.transfer(session, response.url, target, 4, resume_key='key')
                self.assertEqual(target.read_bytes(), b'abcd')
                self.assertEqual(session.get.call_args.kwargs['headers']['Range'], 'bytes=2-')
                self.assertFalse(partial.exists())

    def test_parallel_and_existing_without_manifest(self):
        from bulk_parallel import run_parallel
        from log_setup import configure_logging
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            db = sqlite3.connect(':memory:')
            self.addCleanup(db.close)
            db.execute('CREATE TABLE files (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT, status TEXT)')
            db.execute('CREATE TABLE inventory (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT)')
            for index in range(4):
                db.execute('INSERT INTO inventory VALUES (?,?,?,?)', (f'/root/{index}', str(root / f'{index}.bin'), 3, 'v1'))
            (root / '0.bin').write_bytes(b'abc')
            barrier = threading.Barrier(3, timeout=5)
            threads = set()
            def transfer(session, url, target, size, progress=None, **kwargs):
                threads.add(threading.get_ident())
                barrier.wait()
                target.write_bytes(b'abc')
            page = Mock()
            page.context.cookies.return_value = []
            old_level = logging.getLogger().level
            run_log, handlers = configure_logging(root / 'logs')
            try:
                with patch('builtins.print'):
                    result = run_parallel(db, page, 'https://tenant.sharepoint.com', root, bulk.Progress(4, 12), Mock(), 3, transfer)
            finally:
                for handler in handlers:
                    logging.getLogger().removeHandler(handler)
                    handler.close()
                logging.getLogger().setLevel(old_level)
            contents = run_log.read_text(encoding='utf-8')
            self.assertEqual(contents.count('DOWNLOAD_INICIO'), 3)
            self.assertEqual(contents.count('DOWNLOAD_CONCLUIDO'), 3)
            self.assertEqual(contents.count('ARQUIVO_EXISTENTE'), 1)
            self.assertIn('execucao=', contents)
            self.assertIn('-03:00', contents)
            self.assertEqual(result, (3, 1, 0))
            self.assertEqual(len(threads), 3)

    def test_transfer_reports_status_and_retry_after(self):
        session = Mock()
        response = Mock(status_code=429, headers={'Retry-After': '17'})
        session.get.return_value.__enter__ = Mock(return_value=response)
        session.get.return_value.__exit__ = Mock(return_value=False)
        with tempfile.TemporaryDirectory() as temp:
            with self.assertRaises(bulk.DownloadError) as caught:
                bulk.transfer(session, 'https://tenant.sharepoint.com/file', Path(temp) / 'file.xlsx', 3)
        self.assertEqual(caught.exception.status, 429)
        self.assertEqual(caught.exception.retry_after, 17)
        self.assertIn('HTTP 429', bulk.safe_error(caught.exception))

    def test_transfer_size_failure_preserves_target(self):
        session = Mock()
        response = Mock(status_code=200, headers={}, url='https://tenant.sharepoint.com/file')
        response.iter_content.return_value = [b'ab']
        session.get.return_value.__enter__ = Mock(return_value=response)
        session.get.return_value.__exit__ = Mock(return_value=False)
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / 'file.xlsx'
            with self.assertRaises(bulk.DownloadError) as caught:
                bulk.transfer(session, response.url, target, 3)
            self.assertIn('esperado 3 bytes; recebido 2 bytes', str(caught.exception))
            self.assertFalse(target.exists())
            self.assertFalse(target.exists())

    def test_progress_retry_reset_and_completion(self):
        progress = bulk.Progress(2, 10)
        with patch('builtins.print'):
            progress.show(4, force=True)
            self.assertEqual(progress.completed_bytes, 0)
            progress.show(force=True)
            self.assertEqual(progress.active_bytes, 0)
            progress.complete(5)
            progress.complete(5)
        self.assertEqual(progress.completed_files, 2)
        self.assertEqual(progress.completed_bytes, 10)

    def test_pagination_and_foreign_host(self):
        endpoint = 'https://tenant.sharepoint.com/_api/files'
        with patch.object(bulk, 'read_json', side_effect=[
            {'value': [1], 'odata.nextLink': endpoint + '?page=2'}, {'value': [2]}
        ]):
            self.assertEqual(list(bulk.pages(Mock(), endpoint)), [1, 2])
        with patch.object(bulk, 'read_json', return_value={'value': [], 'odata.nextLink': 'https://foreign.com/files'}):
            with self.assertRaises(RuntimeError):
                list(bulk.pages(Mock(), endpoint))

    def test_tree_and_resume_and_conflict(self):
        root = '/sites/project/library/root'
        page = Mock()
        page.context.cookies.return_value = []
        item = {'Name': 'file.bin', 'ServerRelativeUrl': root + '/child/file.bin', 'Length': '3', 'TimeLastModified': 'v1'}
        scanned = set()

        def listing(page, endpoint):
            scanned.add(endpoint)
            if '/Folders?' in endpoint:
                return iter([] if "root/child'" in endpoint else [{'Name': 'child', 'ServerRelativeUrl': root + '/child'}])
            return iter([item] if "root/child'" in endpoint else [])

        def transfer(session, url, target, size, progress=None, **kwargs):
            self.assertEqual(len(scanned), 4, 'Todas as pastas devem ser lidas antes do download')
            if progress:
                progress(size)
            target.write_bytes(b'abc')

        with tempfile.TemporaryDirectory() as temp, patch.object(bulk, 'read_json', return_value={'Name': 'root', 'ServerRelativeUrl': root}), patch.object(bulk, 'pages', side_effect=listing), patch.object(bulk, 'transfer', side_effect=transfer) as download:
            destination = Path(temp)
            self.assertEqual(bulk.download_tree(page, 'https://tenant.sharepoint.com/sites/project', root, destination, 'url'), 0)
            self.assertEqual((destination / 'root/child/file.bin').read_bytes(), b'abc')
            with patch.object(bulk, 'pages', side_effect=AssertionError('Retomada nao deve repetir o inventario')):
                self.assertEqual(bulk.download_tree(page, 'https://tenant.sharepoint.com/sites/project', root, destination, 'url'), 0)
            self.assertEqual(download.call_count, 1)
            item['TimeLastModified'] = 'v2'
            self.assertEqual(bulk.download_tree(page, 'https://tenant.sharepoint.com/sites/project', root, destination, 'url', refresh_inventory=True), 1)
            self.assertEqual(download.call_count, 1)

    def test_safe_names(self):
        for name in ('..', 'CON', 'bad/name', 'x' * 200):
            safe = bulk.local_name(name)
            self.assertNotIn('/', safe)
            self.assertNotIn(safe, ('.', '..', 'CON'))
            self.assertLessEqual(len(safe), 120)


if __name__ == '__main__':
    unittest.main()
