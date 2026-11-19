import unittest
from unittest.mock import Mock
from urllib.parse import parse_qs, urlsplit

from browser_test import folder_endpoint, read_json, alias_folder_path


class FolderQueryTests(unittest.TestCase):
    def test_long_path_preserves_odata_literal_and_query(self):
        path = '/sites/P/Docs/' + 'Relatórios Técnicos/' * 25 + "DNIT (categorias) #100% d'água"
        endpoint = folder_endpoint('https://empresa.sharepoint.com/sites/P', path)
        page = Mock()
        page.evaluate.return_value = {'status': 200, 'data': {'value': []}}
        self.assertEqual(read_json(page, endpoint + '/Files?$select=Name&$top=200'), {'value': []})
        sent = urlsplit(page.evaluate.call_args.args[1])
        query = parse_qs(sent.query)
        self.assertNotIn('/Docs/', sent.path)
        self.assertTrue(sent.path.endswith('/Files'))
        self.assertEqual(query['@folderPath'], ["'" + path.replace("'", "''") + "'"])
        self.assertEqual(query['$select'], ['Name'])
        self.assertEqual(query['$top'], ['200'])

    def test_throttling_406_is_retried(self):
        page = Mock()
        page.evaluate.side_effect = [{'status': 406, 'retryAfter': '2', 'detail': ''},
                                     {'status': 200, 'data': {'value': [1]}}]
        self.assertEqual(read_json(page, 'https://empresa.sharepoint.com/_api/web'), {'value': [1]})
        page.wait_for_timeout.assert_called_once_with(5000)

    def test_permanent_error_not_retried_and_shows_detail(self):
        page = Mock()
        page.evaluate.return_value = {'status': 403, 'detail': '{"odata.error":{"message":"Acesso negado"}}'}
        with self.assertRaisesRegex(RuntimeError, 'HTTP 403.*Acesso negado'):
            read_json(page, 'https://empresa.sharepoint.com/_api/web')
        page.wait_for_timeout.assert_not_called()

    def test_other_endpoints_and_aliased_pagination_unchanged(self):
        for endpoint in ('https://empresa.sharepoint.com/_api/web?$select=Url',
                         'https://empresa.sharepoint.com/_api/web/GetFolderByServerRelativePath(decodedUrl=@folderPath)/Files?@folderPath=%27Docs%27&$skiptoken=abc'):
            self.assertEqual(alias_folder_path(endpoint), endpoint)


if __name__ == '__main__':
    unittest.main()
