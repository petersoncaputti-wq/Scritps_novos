import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from browser_bulk import DownloadError
from bulk_transfer import verify_local_file


class LocalVerificationTests(unittest.TestCase):
    def test_readable_files_including_empty(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'arquivo.bin'
            for content in (b'', b'abc' * 500000):
                target.write_bytes(content)
                verify_local_file(target, len(content))

    def test_missing_directory_and_wrong_size_fail(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'arquivo.bin'
            with self.assertRaises(DownloadError):
                verify_local_file(target, 3)
            with self.assertRaises(DownloadError):
                verify_local_file(Path(directory), 3)
            target.write_bytes(b'ab')
            with self.assertRaises(DownloadError):
                verify_local_file(target, 3)

    def test_unreadable_file_fails_and_is_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'arquivo.bin'
            target.write_bytes(b'abc')
            with patch.object(Path, 'open', side_effect=PermissionError('acesso negado')):
                with self.assertRaises(DownloadError):
                    verify_local_file(target, 3)
            self.assertEqual(target.read_bytes(), b'abc')


if __name__ == '__main__':
    unittest.main()
