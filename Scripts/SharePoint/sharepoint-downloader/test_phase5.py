"""Fase 5: inventario pela lista da biblioteca, com retorno a leitura por pasta."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import browser_bulk as bulk
from browser_test import SharePointHTTPError

WEB = 'https://tenant.sharepoint.com/sites/project'
ROOT = '/sites/project/Shared Documents/Lote 21'
LIBRARIES = [
    {'Id': 'aaa', 'ItemCount': 10, 'RootFolder': {'ServerRelativeUrl': '/sites/project/Shared Documents'}},
    {'Id': 'bbb', 'ItemCount': 10, 'RootFolder': {'ServerRelativeUrl': '/sites/project/Shared'}},
]


def file_item(path, size=3, etag='"{1},1"'):
    return {'FileSystemObjectType': 0, 'FileRef': path,
            'File': {'Name': path.rsplit('/', 1)[1], 'Length': str(size), 'ServerRelativeUrl': path, 'ETag': etag}}


ITEMS = [
    {'FileSystemObjectType': 1, 'FileRef': ROOT},
    {'FileSystemObjectType': 1, 'FileRef': ROOT + '/Vazia'},
    {'FileSystemObjectType': 1, 'FileRef': ROOT + '/A'},
    file_item(ROOT + '/A/planilha.xlsx'),
    file_item(ROOT.upper() + '/raiz.pdf'),  # caixa diferente continua dentro da raiz
    file_item('/sites/project/Shared Documents/Outro/fora.pdf'),
    file_item('/sites/project/Shared Documents/Lote 210/vizinho.pdf'),
]


def fake_listing(endpoint):
    if '/_api/web/lists?' in endpoint:
        return LIBRARIES
    assert "lists(guid'aaa')/items" in endpoint, endpoint
    assert '$expand=File' in endpoint and '$top=5000' in endpoint
    return ITEMS


class ListInventoryTests(unittest.TestCase):
    def test_longest_library_and_filtering(self):
        with patch('builtins.print'):
            folders, files = bulk.list_inventory(fake_listing, WEB, ROOT)
        self.assertEqual(sorted(folders), [ROOT + '/A', ROOT + '/Vazia'])
        self.assertEqual(sorted(f['Name'] for f in files), ['planilha.xlsx', 'raiz.pdf'])

    def test_unknown_or_large_library_returns_none(self):
        self.assertIsNone(bulk.list_inventory(lambda e: [], WEB, ROOT))
        big = [dict(LIBRARIES[0], ItemCount=bulk.LIST_SCAN_MAX_ITEMS + 1)]
        self.assertIsNone(bulk.list_inventory(lambda e: big, WEB, ROOT))


class DownloadTreeInventoryTests(unittest.TestCase):
    def run_tree(self, pages_side_effect, root_count=3):
        page = Mock()
        page.context.cookies.return_value = []
        transfer = Mock(side_effect=lambda session, url, target, size, *a, **k: target.write_bytes(b'abc') or 3)
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        destination = Path(temp.name)
        info = {'Name': 'Lote 21', 'ServerRelativeUrl': ROOT, 'ItemCount': root_count}
        with patch.object(bulk, 'read_json', return_value=info), \
                patch.object(bulk, 'pages', side_effect=pages_side_effect), \
                patch.object(bulk, 'transfer', transfer), patch('builtins.print'):
            result = bulk.download_tree(page, WEB, ROOT, destination, 'url', workers=2)
        return result, destination / 'Lote 21', transfer

    def test_list_scan_builds_same_layout_with_empty_folders(self):
        result, root, transfer = self.run_tree(lambda page, endpoint: iter(fake_listing(endpoint)))
        self.assertEqual(result, 0)
        self.assertEqual(transfer.call_count, 2)
        self.assertTrue((root / 'A' / 'planilha.xlsx').is_file())
        self.assertTrue((root / 'raiz.pdf').is_file())
        self.assertTrue((root / 'Vazia').is_dir())
        self.assertFalse((root / 'Outro').exists())

    def walk_listing(self, failing):
        reads = []

        def listing(page, endpoint):
            reads.append(endpoint)
            if '/_api/web/lists' in endpoint:
                return failing(endpoint)
            if '/Folders?' in endpoint:
                return iter([])
            return iter([{'Name': 'a.pdf', 'ServerRelativeUrl': ROOT + '/a.pdf', 'Length': '3', 'ETag': 'v1'}])
        return listing, reads

    def test_list_errors_fall_back_to_folder_walk(self):
        def threshold(endpoint):
            raise SharePointHTTPError('HTTP 500 limite da lista', 500)
        listing, reads = self.walk_listing(threshold)
        result, root, transfer = self.run_tree(listing)
        self.assertEqual(result, 0)
        self.assertTrue((root / 'a.pdf').is_file())
        self.assertTrue(any('/Files?' in r for r in reads))

    def test_empty_scan_with_items_in_root_falls_back(self):
        listing, reads = self.walk_listing(lambda endpoint: iter(LIBRARIES if '/lists?' in endpoint else []))
        result, root, transfer = self.run_tree(listing)
        self.assertTrue((root / 'a.pdf').is_file())

    def test_auth_failure_during_scan_is_not_hidden(self):
        def expired(endpoint):
            raise SharePointHTTPError('HTTP 401', 401)
        listing, _ = self.walk_listing(expired)
        page_renew = patch.object(bulk, 'wait_for_sharepoint', side_effect=KeyboardInterrupt)
        with page_renew, self.assertRaises(KeyboardInterrupt):
            self.run_tree(listing)


if __name__ == '__main__':
    unittest.main()
