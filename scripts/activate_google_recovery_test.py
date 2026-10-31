import hashlib
import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
import unittest
import urllib.error
from unittest.mock import MagicMock, patch

spec = importlib.util.spec_from_file_location('activation', Path(__file__).with_name('activate-google-recovery-email.py'))
activation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(activation)


class ActivationTests(unittest.TestCase):
    def configure(self, stored=None, secret_present=True):
        expected = {'smtp_host': 'smtp.gmail.com', 'smtp_port': 465, 'smtp_user': activation.SENDER, 'smtp_admin_email': activation.SENDER, 'smtp_sender_name': 'Verifica Placa'}
        smtp = MagicMock()
        api = MagicMock(side_effect=[{}] + [expected if stored is None else stored] * 8)
        entries = [{'name': 'CONSOLE_RECOVERY_EMAIL_ENABLED', 'digest': hashlib.sha256(b'true').hexdigest()}] if secret_present else []
        cli = MagicMock(side_effect=[SimpleNamespace(returncode=0), SimpleNamespace(returncode=0, stdout=json.dumps(entries))])
        return smtp, api, cli

    def test_unpersisted_smtp_cannot_activate_even_if_patch_returns_success(self):
        smtp, api, cli = self.configure(stored={'smtp_host': None})
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli), patch.object(activation.time, 'sleep'):
            with self.assertRaisesRegex(RuntimeError, 'não confirmou'):
                activation.activate('fixture-token', 'fixture-password')
        cli.assert_not_called()

    def test_success_exit_from_cli_is_insufficient_without_enabled_secret(self):
        smtp, api, cli = self.configure(secret_present=False)
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli), patch.object(activation.time, 'sleep'):
            with self.assertRaisesRegex(RuntimeError, 'ativação não foi confirmada'):
                activation.activate('fixture-token', 'fixture-password')

    def test_verifies_google_smtp_persistence_and_exact_enabled_value(self):
        smtp, api, cli = self.configure()
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli), patch.object(activation.time, 'sleep'):
            activation.activate('fixture-token', 'fixture-password')
        smtp.return_value.__enter__.return_value.login.assert_called_once_with(activation.SENDER, 'fixture-password')
        self.assertEqual(api.call_count, 2)
        self.assertIsInstance(api.call_args_list[0].args[1]['smtp_port'], str)
        self.assertEqual(cli.call_count, 2)

    def test_google_login_failure_never_changes_server(self):
        smtp, api, cli = self.configure()
        smtp.return_value.__enter__.return_value.login.side_effect = RuntimeError('fixture failure')
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli), patch.object(activation.time, 'sleep'):
            with self.assertRaises(RuntimeError):
                activation.activate('fixture-token', 'fixture-password')
        api.assert_not_called()
        cli.assert_not_called()

    def test_management_error_is_reported_as_server_failure_after_google_auth(self):
        smtp, api, cli = self.configure()
        api.side_effect = urllib.error.HTTPError('https://api.supabase.com', 400, 'fixture', None, None)
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli):
            with self.assertRaises(activation.ActivationError) as raised:
                activation.activate('fixture-token', 'fixture-password')
        self.assertEqual(raised.exception.stage, 'server_config')
        self.assertIn('HTTP 400', str(raised.exception))
        self.assertNotIn('fixture-password', str(raised.exception))
        cli.assert_not_called()

    def test_configuration_propagation_is_rechecked_before_enabling(self):
        smtp, api, cli = self.configure()
        expected = {'smtp_host': 'smtp.gmail.com', 'smtp_port': '465', 'smtp_user': activation.SENDER, 'smtp_admin_email': activation.SENDER, 'smtp_sender_name': 'Verifica Placa'}
        api.side_effect = [{}, {'smtp_host': None}, expected]
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli), patch.object(activation.time, 'sleep') as wait:
            activation.activate('fixture-token', 'fixture-password')
        wait.assert_called_once_with(2)

    def test_google_auth_rejection_has_specific_safe_stage(self):
        smtp, api, cli = self.configure()
        smtp.return_value.__enter__.return_value.login.side_effect = activation.smtplib.SMTPAuthenticationError(535, b'fixture-secret')
        with patch.object(activation.smtplib, 'SMTP_SSL', smtp), patch.object(activation, 'auth_config', api), patch.object(activation.subprocess, 'run', cli):
            with self.assertRaises(activation.ActivationError) as raised:
                activation.activate('fixture-token', 'fixture-password')
        self.assertEqual(raised.exception.stage, 'google_login')
        self.assertNotIn('fixture-secret', str(raised.exception))
        api.assert_not_called()


if __name__ == '__main__':
    unittest.main()
