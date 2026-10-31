import importlib.util
import http.client
from pathlib import Path
import threading
from urllib.parse import urlencode
import unittest
from unittest.mock import MagicMock, patch

spec = importlib.util.spec_from_file_location('setup', Path(__file__).with_name('google-recovery-setup.py'))
setup = importlib.util.module_from_spec(spec)
spec.loader.exec_module(setup)


class SetupTests(unittest.TestCase):
    def setUp(self):
        self.operation = MagicMock(side_effect=RuntimeError('fixture error with fixture-secret'))
        self.server, self.nonce = setup.make_server('fixture-token', activate=self.operation)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.origin = f'http://127.0.0.1:{self.server.server_port}'

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def request(self, method='GET', body=None, origin=None, csrf=None):
        connection = http.client.HTTPConnection('127.0.0.1', self.server.server_port)
        headers = {'Content-Type': 'application/x-www-form-urlencoded', 'Origin': self.origin if origin is None else origin}
        connection.request(method, '/' + self.nonce, body=None if body is None else urlencode({'csrf': self.nonce if csrf is None else csrf, 'password': body}), headers=headers)
        response = connection.getresponse()
        result = response.status, response.read().decode(), dict(response.getheaders())
        connection.close()
        return result

    def test_form_is_loopback_only_and_has_no_external_resources_or_cache(self):
        status, body, headers = self.request()
        self.assertEqual(status, 200)
        self.assertEqual(self.server.server_address[0], '127.0.0.1')
        self.assertEqual(headers['Cache-Control'], 'no-store')
        self.assertIn("frame-ancestors 'none'", headers['Content-Security-Policy'])
        self.assertIn('type="password"', body)
        self.operation.assert_not_called()

    def test_cross_origin_and_missing_csrf_cannot_submit_credentials(self):
        self.assertEqual(self.request('POST', 'abcdefghijklmnop', origin='https://example.invalid')[0], 403)
        self.assertEqual(self.request('POST', 'abcdefghijklmnop', csrf='wrong')[0], 400)
        self.operation.assert_not_called()

    def test_failure_does_not_echo_password_or_internal_error(self):
        status, body, _ = self.request('POST', 'abcdefghijklmnop')
        self.assertEqual(status, 400)
        self.assertNotIn('abcdefghijklmnop', body)
        self.assertNotIn('fixture-secret', body)
        self.operation.assert_called_once_with('fixture-token', 'abcdefghijklmnop')

    def test_success_activates_once_and_returns_confirmation(self):
        self.operation.side_effect = None
        with patch('builtins.print'):
            status, body, _ = self.request('POST', 'abcdefghijklmnop')
        self.assertEqual(status, 200)
        self.assertIn('Envio ativado', body)
        self.operation.assert_called_once()


if __name__ == '__main__':
    unittest.main()
