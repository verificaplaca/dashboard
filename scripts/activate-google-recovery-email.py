#!/usr/bin/env python3
"""Run locally by the account owner; never put an app password in chat or arguments."""
import base64
import getpass
import hashlib
import json
import os
import smtplib
import ssl
import subprocess
import sys
import time
import urllib.error
import urllib.request

PROJECT = 'ftmgmfdqdqxboiktxcoj'
SENDER = 'suporte@verificaplaca.com.br'
API = 'https://api.supabase.com/v1/projects/' + PROJECT + '/config/auth'


def management_token():
    token = os.environ.get('SUPABASE_ACCESS_TOKEN', '')
    if not token and sys.platform == 'darwin':
        result = subprocess.run(['security', 'find-generic-password', '-s', 'Supabase CLI', '-a', 'supabase', '-w'], capture_output=True, text=True)
        if result.returncode == 0:
            token = result.stdout.strip()
            if token.startswith('go-keyring-base64:'):
                token = base64.b64decode(token.split(':', 1)[1]).decode()
    if not token.startswith('sbp_'):
        raise RuntimeError('Entre no Supabase CLI com supabase login antes de continuar.')
    return token


def auth_config(token, values=None):
    request = urllib.request.Request(API, data=None if values is None else json.dumps(values).encode(), method='GET' if values is None else 'PATCH', headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


class ActivationError(RuntimeError):
    def __init__(self, stage, message):
        super().__init__(message)
        self.stage = stage


def activation_confirmed(token):
    stored = auth_config(token)
    expected = {'smtp_host': 'smtp.gmail.com', 'smtp_port': '465', 'smtp_user': SENDER, 'smtp_admin_email': SENDER, 'smtp_sender_name': 'Verifica Placa'}
    if any(str(stored.get(field)) != value for field, value in expected.items()):
        return False
    listed = subprocess.run(['supabase', 'secrets', 'list', '--project-ref', PROJECT, '--output', 'json'], capture_output=True, text=True)
    try:
        entries = json.loads(listed.stdout) if listed.returncode == 0 else []
    except (ValueError, TypeError):
        return False
    expected_digest = hashlib.sha256(b'true').hexdigest()
    return isinstance(entries, list) and any(entry.get('name') == 'CONSOLE_RECOVERY_EMAIL_ENABLED' and (entry.get('digest') or entry.get('value')) == expected_digest for entry in entries)


def activate(token, password):
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=25, context=ssl.create_default_context()) as smtp:
            smtp.login(SENDER, password)  # Validate only; no test message is sent.
    except smtplib.SMTPAuthenticationError:
        raise ActivationError('google_login', 'O Google recusou a autenticação. Confira a senha de app e a conta suporte@verificaplaca.com.br.') from None
    except Exception:
        raise ActivationError('google_connection', 'Não foi possível concluir a conexão com o Google. Tente novamente em instantes.') from None
    # The Management API requires smtp_port as a STRING (unlike smtplib).
    expected = {'smtp_host': 'smtp.gmail.com', 'smtp_port': '465', 'smtp_user': SENDER, 'smtp_pass': password, 'smtp_admin_email': SENDER, 'smtp_sender_name': 'Verifica Placa', 'rate_limit_email_sent': 30}
    try:
        auth_config(token, expected)
        public_fields = ['smtp_host', 'smtp_port', 'smtp_user', 'smtp_admin_email', 'smtp_sender_name']
        for attempt in range(8):
            stored = auth_config(token)
            if all(str(stored.get(field)) == str(expected[field]) for field in public_fields):
                break
            if attempt < 7:
                time.sleep(2)
        else:
            raise ActivationError('server_persistence', 'O Google validou a senha, mas o servidor ainda não confirmou a configuração. Aguarde e tente novamente.')
    except urllib.error.HTTPError as error:
        raise ActivationError('server_config', f'O Google validou a senha, mas houve uma falha ao salvar no servidor (HTTP {error.code}). A senha não precisa ser recriada.') from None
    except ActivationError:
        raise
    except Exception:
        raise ActivationError('server_connection', 'O Google validou a senha, mas não foi possível concluir a conexão com o servidor. Tente novamente em instantes.') from None
    result = subprocess.run(['supabase', 'secrets', 'set', 'CONSOLE_RECOVERY_EMAIL_ENABLED=true', '--project-ref', PROJECT], capture_output=True, text=True)
    if result.returncode:
        raise ActivationError('server_flag', 'Google configurado, mas faltou ativar a opção no servidor. Execute novamente ou revise sua sessão do Supabase CLI.')
    listed = subprocess.run(['supabase', 'secrets', 'list', '--project-ref', PROJECT, '--output', 'json'], capture_output=True, text=True)
    try:
        entries = json.loads(listed.stdout) if listed.returncode == 0 else []
    except (ValueError, TypeError):
        entries = []
    expected_digest = hashlib.sha256(b'true').hexdigest()
    if not any(entry.get('name') == 'CONSOLE_RECOVERY_EMAIL_ENABLED' and (entry.get('digest') or entry.get('value')) == expected_digest for entry in entries):
        raise ActivationError('server_flag', 'Google configurado, mas a ativação não foi confirmada no servidor. Execute novamente ou revise sua sessão do Supabase CLI.')


def main():
    print('Ativar emails de recuperação do Verifica Placa pelo Google')
    print('Conta remetente: ' + SENDER)
    print('Use uma senha de app dessa conta Google. Não use a senha principal.')
    token = management_token()
    if activation_confirmed(token):
        print('Envio já está ativado e confirmado no servidor. Nenhuma senha precisa ser informada.')
        return
    if not sys.stdin.isatty():
        raise RuntimeError('Execute em um terminal interativo para manter a senha oculta.')
    password = getpass.getpass('Cole a senha de app (a digitação fica oculta): ').replace(' ', '')
    if not password:
        raise RuntimeError('Nenhuma senha foi informada. Nada foi alterado.')
    activate(token, password)
    password = None
    print('Envio ativado. No painel, use Usuários → Enviar link por email.')
    print('Nenhum email foi enviado durante esta configuração.')


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('\nCancelado.')
        sys.exit(1)
    except Exception as error:
        # Transport errors can contain credentials/headers; never print raw objects.
        if isinstance(error, RuntimeError):
            print(str(error), file=sys.stderr)
        else:
            print('Não foi possível validar ou configurar o envio. Confira a senha de app, as regras da conta Google e o acesso ao Supabase.', file=sys.stderr)
        sys.exit(1)
