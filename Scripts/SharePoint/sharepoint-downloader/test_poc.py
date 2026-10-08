import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from graph_client import GraphClient, GraphError
from sharepoint import parse_sharepoint_url, resolve_sharepoint_folder, get_site
from downloader import download_file
from msal_extensions import FilePersistenceWithDataProtection

class POCTests(unittest.TestCase):
    def test_view_url_keeps_encoded_percent(self):
        parsed = parse_sharepoint_url("https://empresa.sharepoint.com/sites/P/Docs/Forms/AllItems.aspx?id=%2Fsites%2FP%2FDocs%2F100%25")
        self.assertEqual(parsed.path, "/sites/P/Docs/100%")

    def test_untrusted_hosts_and_share_links(self):
        for url in ["http://empresa.sharepoint.com/a", "https://empresa.sharepoint.com.evil/a", "https://empresa.sharepoint.com/:f:/s/P/a"]:
            with self.assertRaises(ValueError):
                parse_sharepoint_url(url)

    def test_arbitrary_library_and_special_characters(self):
        graph = Mock()
        graph.get.return_value = {"id": "folder", "folder": {}}
        drive = {"id": "drive", "name": "Engenharia", "webUrl": "https://empresa.sharepoint.com/sites/P/Acervo%20Tecnico"}
        parsed = parse_sharepoint_url("https://empresa.sharepoint.com/sites/P/Acervo%20Tecnico/A%23B")
        self.assertEqual(resolve_sharepoint_folder(graph, parsed, [drive])[0], drive)
        graph.get.assert_called_with("drives/drive/root:/A%23B")

    def test_subsite_resolution(self):
        graph = Mock()
        graph.get.side_effect = [GraphError(404, "itemNotFound"), {"id": "site", "displayName": "Subsite"}]
        parsed = parse_sharepoint_url("https://empresa.sharepoint.com/sites/P/Sub/Docs")
        self.assertEqual(get_site(graph, parsed)["id"], "site")
        self.assertEqual(graph.get.call_args.args[0], "sites/empresa.sharepoint.com:/sites/P/Sub")

    def test_retry_and_pagination(self):
        graph = GraphClient("dummy-not-a-real-token")
        limited = Mock(status_code=429, ok=False, headers={"Retry-After": "7"})
        limited.json.return_value = {"error": {"code": "tooManyRequests"}}
        first = Mock(ok=True)
        first.json.return_value = {"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"}
        last = Mock(ok=True)
        last.json.return_value = {"value": [{"id": "b"}]}
        graph.session.get = Mock(side_effect=[limited, first, last])
        with patch("graph_client.time.sleep") as sleep:
            self.assertEqual(list(graph.pages("start")), [{"id": "a"}, {"id": "b"}])
            sleep.assert_called_once_with(7)
        with self.assertRaises(ValueError):
            graph.get("https://evil.example/next")

    def test_download_stream_and_no_bearer_forwarding(self):
        graph = Mock()
        graph.get.return_value = {"@microsoft.graph.downloadUrl": "https://download.example/file", "size": 3}
        response = Mock(status_code=200)
        response.iter_content.return_value = [b"abc"]
        context = Mock()
        context.__enter__ = Mock(return_value=response)
        context.__exit__ = Mock(return_value=False)
        with tempfile.TemporaryDirectory() as directory, patch("downloader.requests.get", return_value=context) as request:
            target = download_file(graph, "drive", {"id": "item", "name": "teste.txt"}, Path(directory))
            self.assertEqual(target.read_bytes(), b"abc")
            self.assertNotIn("headers", request.call_args.kwargs)
            with self.assertRaises(ValueError):
                download_file(graph, "drive", {"id": "item", "name": "CON.txt"}, Path(directory))

    def test_windows_cache_encryption(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cache"
            persistence = FilePersistenceWithDataProtection(str(path))
            persistence.save("test-only-no-real-token")
            self.assertEqual(persistence.load(), "test-only-no-real-token")
            self.assertNotIn(b"test-only-no-real-token", path.read_bytes())

if __name__ == "__main__":
    unittest.main()
