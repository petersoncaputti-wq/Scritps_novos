import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
import download_diagnostic as diagnostic


class DiagnosticTests(unittest.TestCase):
    def test_chooses_last_failed_file_inside_folder(self):
        with tempfile.TemporaryDirectory() as temp:
            base = Path(temp)
            (base / 'logs').mkdir()
            (base / 'logs/execucao_test.log').write_text('ORIGEM hostname=tenant.sharepoint.com pasta=/sites/project/library\nDOWNLOAD_FALHA_FINAL remoto=/sites/project/library/file.xlsx local=C:\\file.xlsx\n')
            with patch.object(diagnostic, 'BASE', base), patch('builtins.print'):
                url, remote = diagnostic.choose_target()
            self.assertEqual(url, 'https://tenant.sharepoint.com/sites/project/library')
            self.assertEqual(remote, '/sites/project/library/file.xlsx')

    def test_probe_reads_only_sample_and_never_follows_redirects(self):
        session = Mock()
        response = Mock(status_code=200, headers={'Content-Type': 'application/octet-stream'})
        chunks = []
        def content(*args):
            for _ in range(100):
                chunks.append(1)
                yield b'PK\x03\x04' + b'x' * 4092
        response.iter_content.side_effect = content
        session.get.return_value.__enter__ = Mock(return_value=response)
        session.get.return_value.__exit__ = Mock(return_value=False)
        with patch('builtins.print'):
            diagnostic.probe(session, 'test', 'https://tenant.sharepoint.com/file')
        self.assertEqual(len(chunks), 8)
        self.assertFalse(session.get.call_args.kwargs['allow_redirects'])


if __name__ == '__main__':
    unittest.main()
