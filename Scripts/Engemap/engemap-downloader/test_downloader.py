import json
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import downloader as d

DATA = b'arquivo-de-teste' * 100


class Handler(BaseHTTPRequestHandler):
    ignore_range = False
    invalid_range = False

    def log_message(self, *args):
        pass

    def do_GET(self):
        offset = int(self.headers.get('Range', 'bytes=0-')[6:-1])
        if self.ignore_range:
            offset = 0
        self.send_response(206 if offset else 200)
        self.send_header('ETag', '"v1"')
        if offset:
            self.send_header('Content-Range', 'bytes 0-1/2' if self.invalid_range else
                             f'bytes {offset}-{len(DATA)-1}/{len(DATA)}')
        self.send_header('Content-Length', str(len(DATA) - offset))
        self.end_headers()
        self.wfile.write(DATA[offset:])

    def do_PROPFIND(self):
        self.rfile.read(int(self.headers.get('Content-Length', 0)))
        self.send_response(207)
        self.end_headers()
        self.wfile.write(b'''<d:multistatus xmlns:d="DAV:">
        <d:response><d:href>/root/</d:href></d:response>
        <d:response><d:href>/root/arquivo%20teste.pdf</d:href><d:propstat>
        <d:status>HTTP/1.1 200 OK</d:status><d:prop><d:resourcetype/>
        <d:getcontentlength>1500</d:getcontentlength><d:getetag>"v1"</d:getetag>
        </d:prop></d:propstat></d:response></d:multistatus>''')


class Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.target = Path(self.temp.name) / 'teste.pdf'
        self.client = d.Client(d.DEFAULT_URL, 'senha-teste')
        self.client.origin = f'http://127.0.0.1:{self.server.server_port}'
        self.item = {'path': '/root/arquivo', 'size': len(DATA), 'etag': '"v1"'}
        Handler.ignore_range = Handler.invalid_range = False

    def partial(self):
        partial = self.target.with_name('teste.pdf.engemap.part')
        partial.write_bytes(DATA[:100])
        d.save_json(partial.with_name(partial.name + '.json'),
                    {'source': self.client.origin + self.item['path'],
                     'size': len(DATA), 'etag': '"v1"'})
        return partial

    def test_full_download_and_hash(self):
        result = d.transfer(self.client, self.item, self.target)
        self.assertEqual(self.target.read_bytes(), DATA)
        self.assertEqual(result['sha256'], d.file_hash(self.target))

    def test_resume(self):
        self.partial()
        d.transfer(self.client, self.item, self.target)
        self.assertEqual(self.target.read_bytes(), DATA)

    def test_ignored_range_restarts(self):
        self.partial()
        Handler.ignore_range = True
        d.transfer(self.client, self.item, self.target)
        self.assertEqual(self.target.read_bytes(), DATA)

    def test_invalid_range_preserves_partial(self):
        partial = self.partial()
        Handler.invalid_range = True
        with self.assertRaises(d.DownloadError):
            d.transfer(self.client, self.item, self.target)
        self.assertEqual(partial.read_bytes(), DATA[:100])
        self.assertFalse(self.target.exists())

    def test_existing_file_preserved(self):
        self.target.write_bytes(b'original')
        with self.assertRaises(d.DownloadError):
            d.transfer(self.client, self.item, self.target)
        self.assertEqual(self.target.read_bytes(), b'original')

    def test_listing_decodes_names(self):
        items = self.client.listing('/root/')
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]['name'], 'arquivo teste.pdf')
        self.assertEqual(items[0]['size'], 1500)

    def test_names_and_traversal(self):
        self.assertEqual(d.local_name('Ortomosaico_SP310_001.tif'), 'Ortomosaico_SP310_001.tif')
        self.assertEqual(d.local_name('Ortomosaico_SP310_001.tfw'), 'Ortomosaico_SP310_001.tfw')
        self.assertNotEqual(d.local_name('a:b'), d.local_name('a?b'))
        with self.assertRaises(d.DownloadError):
            d.safe_path(Path(self.temp.name), '../escape')

    def test_reject_invalid_share(self):
        for url in ['http://servidor/s/token', 'https://usuario:senha@servidor/s/token',
                    'https://servidor/outro']:
            with self.assertRaises(d.DownloadError):
                d.Client(url, '')

    def test_reject_case_collision(self):
        with self.assertRaises(d.DownloadError):
            d.validate_names([{'relative': 'pasta/A.pdf'}, {'relative': 'Pasta/b.pdf'}], [])

    def test_associated_files_preserve_stem(self):
        d.validate_names([{'relative': 'mapa' + ext} for ext in ['.shp', '.shx', '.dbf', '.prj']], [])
        for ext in ['.shp', '.shx', '.dbf', '.prj']:
            self.assertEqual(Path(d.local_name('mapa' + ext)).stem, 'mapa')


if __name__ == '__main__':
    unittest.main()
