import json
import unittest
from unittest.mock import Mock
from http_diagnostics import error_details


class DiagnosticsTests(unittest.TestCase):
    def response(self, body, media='application/json'):
        response = Mock(headers={'Content-Type': media, 'SPRequestGuid': '12345678-1234-1234-1234-123456789abc', 'Set-Cookie': 'secret-cookie'})
        response.iter_content.return_value = [body]
        return response

    def test_odata_and_redaction(self):
        response = self.response(json.dumps({'error': {'code': '-1, Microsoft.SharePoint.SPException', 'message': {'value': 'Invalid parameter. https://example.com/?sig=secret-url Bearer secret-bearer token="secret with spaces"\nuser@example.com'}, 'innererror': {'password': 'secret-inner'}}}).encode())
        detail = error_details(response)
        self.assertIn('Invalid parameter.', detail['mensagem'])
        self.assertIn('SPException', detail['codigo'])
        self.assertEqual(detail['SPRequestGuid'], '12345678-1234-1234-1234-123456789abc')
        self.assertNotIn('secret', json.dumps(detail))
        self.assertNotIn('user@example.com', json.dumps(detail))
        self.assertNotIn('\n', detail['mensagem'])

    def test_body_bound_and_html_title_only(self):
        response = self.response(b'x' * 20000, 'text/plain')
        self.assertTrue(error_details(response)['corpo_truncado'])
        response = self.response(b'<title>Bad Request</title><input value="secret"><script>secret</script>', 'text/html')
        detail = error_details(response)
        self.assertEqual(detail['mensagem'], 'Bad Request')
        self.assertNotIn('secret', json.dumps(detail))


if __name__ == '__main__':
    unittest.main()
