#!/usr/bin/env python3
"""One-use loopback-only setup form. The owner enters the credential personally."""
import html
import importlib.util
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
import secrets
import threading
from urllib.parse import parse_qs

spec = importlib.util.spec_from_file_location('activation', Path(__file__).with_name('activate-google-recovery-email.py'))
activation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(activation)


def page(message='', success=False):
    notice = '<p role="status">Envio ativado e conferido. Você já pode fechar esta página e usar “Enviar link por email” no painel.</p>' if success else '<p>Use a senha de app criada na conta <strong>suporte@verificaplaca.com.br</strong>.</p><p>Ela será usada para autorizar o envio pelo Google. Não use a senha principal da conta.</p>'
    form = '' if success else '<form method="post"><input type="hidden" name="csrf" value="TOKEN"><label for="password">Senha de app do Google</label><input id="password" name="password" type="password" autocomplete="new-password" required minlength="16" maxlength="32" autofocus><button>Autorizar e ativar envio</button></form><small>O campo fica oculto. Nenhum email será enviado durante a configuração.</small>'
    return ('<!doctype html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Ativar email · Verifica Placa</title><style>body{margin:0;background:#f6f8f6;color:#17231c;font:16px system-ui;min-height:100vh;display:grid;place-items:center}main{background:white;border:1px solid #dce5dd;border-radius:24px;padding:32px;width:min(430px,calc(100vw - 96px))}h1{font-size:26px;font-weight:600}p{line-height:1.6;color:#526059}label{display:block;margin:24px 0 10px}input[type=password]{box-sizing:border-box;width:100%;padding:14px;border:1px solid #8a9e8d;border-radius:10px;font:inherit}button{margin:20px 0;width:100%;background:#223d2b;color:white;border:0;border-radius:12px;padding:14px;font:inherit;cursor:pointer}small{color:#64736a;line-height:1.5;display:block}.error{color:#a42b35}strong{overflow-wrap:anywhere}</style><main><small>VERIFICA PLACA · CONFIGURAÇÃO LOCAL</small><h1>' + ('Envio ativado' if success else 'Ativar envio pelo Google') + '</h1>' + notice + ('<p class="error" role="alert">' + html.escape(message) + '</p>' if message else '') + form + '</main></html>')


def make_server(token, port=0, activate=None):
    nonce = secrets.token_urlsafe(32)
    operation = activation.activate if activate is None else activate
    state = {'completed': False, 'attempts': 0}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_args):
            pass  # Never log paths, POST bodies or credentials.

        def reply(self, status, body):
            self.send_response(status)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Referrer-Policy', 'no-referrer')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy', "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; frame-ancestors 'none'; base-uri 'none'")
            self.end_headers()
            self.wfile.write(body.replace('TOKEN', nonce).encode())

        def valid_target(self):
            return self.headers.get('Host') == f'127.0.0.1:{self.server.server_port}' and self.path == '/' + nonce

        def do_GET(self):
            if not self.valid_target():
                self.reply(404, 'Página indisponível.'); return
            self.reply(200, page(success=state['completed']))

        def do_POST(self):
            expected_origin = f'http://127.0.0.1:{self.server.server_port}'
            if not self.valid_target() or self.headers.get('Origin') != expected_origin or self.headers.get_content_type() != 'application/x-www-form-urlencoded':
                self.reply(403, 'Solicitação inválida.'); return
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if length < 1 or length > 1024:
                    raise ValueError()
                values = parse_qs(self.rfile.read(length).decode(), strict_parsing=True)
                csrf = values.get('csrf', [''])[0]
                password = values.get('password', [''])[0].replace(' ', '')
                if not secrets.compare_digest(csrf, nonce) or len(password) != 16 or not password.isascii() or not password.isalpha():
                    raise ValueError()
            except (ValueError, UnicodeError):
                self.reply(400, page('Confira a senha de app de 16 letras gerada pelo Google.')); return
            if state['completed'] or state['attempts'] >= 5:
                self.reply(409, 'Esta configuração foi encerrada.'); return
            state['attempts'] += 1
            try:
                operation(token, password)
                state['completed'] = True
                self.reply(200, page(success=True))
                print('Envio ativado e verificado no servidor.', flush=True)
                threading.Thread(target=self.server.shutdown, daemon=True).start()
            except Exception:
                self.reply(400, page('Não foi possível ativar. Confira se a senha de app pertence à conta suporte@verificaplaca.com.br.'))
            finally:
                password = None
                values.clear()

    server = HTTPServer(('127.0.0.1', port), Handler)
    return server, nonce


if __name__ == '__main__':
    token = activation.management_token()
    activation.auth_config(token)  # Check project access before opening the form.
    server, nonce = make_server(token)
    timer = threading.Timer(900, server.shutdown)
    timer.daemon = True
    timer.start()
    print(f'LOCAL_SETUP_URL=http://127.0.0.1:{server.server_port}/{nonce}', flush=True)
    try:
        server.serve_forever()
    finally:
        timer.cancel()
        server.server_close()
