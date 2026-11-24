#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Servidor do Painel de Ordens SAP (versão B: painel compartilhado).

- Serve o index.html e a biblioteca de Excel.
- Guarda no servidor as planilhas carregadas (GPI, operações, materiais)
  e as confirmações de entrega de material, para todos verem o mesmo painel.
- Só usa a biblioteca padrão do Python (3.7+). Nada para instalar.

Uso:   python servidor.py [--porta 8000] [--pin 1234]
       (ou defina a variável de ambiente PAINEL_PIN)
"""
import argparse
import json
import os
import socket
import sqlite3
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE, 'dados')
DB = os.path.join(DB_DIR, 'painel.db')
DATASETS = ('gpi', 'ops', 'mat')
MAX_BODY = 100 * 1024 * 1024
STATIC = {'/': ('index.html', 'text/html; charset=utf-8'),
          '/index.html': ('index.html', 'text/html; charset=utf-8'),
          '/lib/xlsx.full.min.js': ('lib/xlsx.full.min.js', 'application/javascript; charset=utf-8')}

WRITE_LOCK = threading.Lock()
PIN = ''


def db():
    con = sqlite3.connect(DB, timeout=15)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    os.makedirs(DB_DIR, exist_ok=True)
    with db() as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v INTEGER);
            INSERT OR IGNORE INTO meta VALUES('rev', 0);
            CREATE TABLE IF NOT EXISTS dataset(k TEXT PRIMARY KEY, name TEXT, quem TEXT, ts TEXT, rows TEXT);
            CREATE TABLE IF NOT EXISTS planilha(id INTEGER PRIMARY KEY AUTOINCREMENT, tipo TEXT, semana TEXT, name TEXT, quem TEXT, ts TEXT, nlinhas INTEGER, rows TEXT);
            CREATE TABLE IF NOT EXISTS deliv(key TEXT PRIMARY KEY, quem TEXT, ts TEXT);
            CREATE TABLE IF NOT EXISTS ajuste(key TEXT PRIMARY KEY, status TEXT, nota TEXT, quem TEXT, ts TEXT);
            CREATE TABLE IF NOT EXISTS log(id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, acao TEXT, chave TEXT, quem TEXT);
        """)


def migrate(con):
    # versão anterior guardava só a última planilha de cada tipo: vira o 1º registro da base (semana em branco)
    if con.execute('SELECT COUNT(*) FROM planilha').fetchone()[0] == 0:
        for r in con.execute('SELECT k,name,quem,ts,rows FROM dataset').fetchall():
            n = len(json.loads(r['rows']))
            con.execute('INSERT INTO planilha(tipo,semana,name,quem,ts,nlinhas,rows) VALUES(?,?,?,?,?,?,?)',
                        (r['k'], '', r['name'], r['quem'], r['ts'], n, r['rows']))


def now():
    return datetime.now().strftime('%d/%m/%Y %H:%M')


def bump(con):
    con.execute("UPDATE meta SET v = v + 1 WHERE k='rev'")


def get_rev(con):
    return con.execute("SELECT v FROM meta WHERE k='rev'").fetchone()[0]


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def log_message(self, fmt, *args):
        pass  # silencioso

    # ---------- utilidades ----------
    def send(self, code, body, ctype='application/json; charset=utf-8'):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def jerr(self, code, msg):
        self.send(code, json.dumps({'erro': msg}, ensure_ascii=False))

    def body_json(self):
        n = int(self.headers.get('Content-Length') or 0)
        if n > MAX_BODY:
            raise ValueError('arquivo grande demais')
        return json.loads(self.rfile.read(n).decode('utf-8') or '{}')

    def authorized(self):
        return (not PIN) or self.headers.get('X-PIN', '') == PIN

    # ---------- GET ----------
    def do_GET(self):
        path = urlparse(self.path).path
        if path in STATIC:
            fname, ctype = STATIC[path]
            try:
                with open(os.path.join(BASE, fname), 'rb') as f:
                    return self.send(200, f.read(), ctype)
            except OSError:
                return self.jerr(404, 'arquivo não encontrado')
        if path.startswith('/lib/'):
            name = os.path.basename(path)
            ext = os.path.splitext(name)[1].lower()
            types = {'.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg'}
            if ext in types:
                try:
                    with open(os.path.join(BASE, 'lib', name), 'rb') as f:
                        return self.send(200, f.read(), types[ext])
                except OSError:
                    return self.jerr(404, 'arquivo não encontrado')
        if not path.startswith('/api/'):
            return self.jerr(404, 'não encontrado')
        if not self.authorized():
            return self.jerr(401, 'PIN necessário')
        with db() as con:
            if path == '/api/rev':
                return self.send(200, json.dumps({'rev': get_rev(con)}))
            if path.startswith('/api/planilha/'):
                try:
                    pid = int(path.rsplit('/', 1)[1])
                except ValueError:
                    return self.jerr(400, 'id inválido')
                r = con.execute('SELECT rows FROM planilha WHERE id=?', (pid,)).fetchone()
                if not r:
                    return self.jerr(404, 'planilha não encontrada')
                return self.send(200, '{"rows":' + r['rows'] + '}')
            if path == '/api/state':
                planilhas = [dict(r) for r in con.execute(
                    'SELECT id,tipo,semana,name,quem,ts,nlinhas FROM planilha ORDER BY id')]
                deliv = {r['key']: {'quem': r['quem'], 'ts': r['ts']} for r in con.execute('SELECT * FROM deliv')}
                ajustes = {r['key']: {'status': r['status'], 'nota': r['nota'], 'quem': r['quem'], 'ts': r['ts']}
                           for r in con.execute('SELECT * FROM ajuste')}
                out = '{"rev":%d,"planilhas":%s,"deliv":%s,"ajustes":%s}' % (
                    get_rev(con), json.dumps(planilhas, ensure_ascii=False), json.dumps(deliv, ensure_ascii=False),
                    json.dumps(ajustes, ensure_ascii=False))
                return self.send(200, out)
        return self.jerr(404, 'não encontrado')

    # ---------- PUT / POST ----------
    def do_PUT(self):
        self.write_route()

    def do_POST(self):
        self.write_route()

    def write_route(self):
        path = urlparse(self.path).path
        if not self.authorized():
            return self.jerr(401, 'PIN necessário')
        try:
            data = self.body_json()
        except Exception as e:
            return self.jerr(400, 'corpo inválido: %s' % e)
        quem = str(data.get('quem') or '')[:80]
        with WRITE_LOCK, db() as con:
            if path == '/api/planilha' and self.command == 'POST':
                tipo = str(data.get('tipo') or '')
                rows = data.get('rows')
                if tipo not in DATASETS or not isinstance(rows, list):
                    return self.jerr(400, 'planilha inválida')
                cur = con.execute('INSERT INTO planilha(tipo,semana,name,quem,ts,nlinhas,rows) VALUES(?,?,?,?,?,?,?)',
                                  (tipo, str(data.get('semana') or '')[:20], str(data.get('name') or '')[:200], quem, now(),
                                   len(rows), json.dumps(rows, ensure_ascii=False)))
                con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'upload', '%s #%d' % (tipo, cur.lastrowid), quem))
                bump(con)
                return self.send(200, json.dumps({'id': cur.lastrowid, 'rev': get_rev(con)}))
            if path == '/api/planilha/delete' and self.command == 'POST':
                pid = int(data.get('id') or 0)
                con.execute('DELETE FROM planilha WHERE id=?', (pid,))
                con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'planilha excluída', str(pid), quem))
                bump(con)
                return self.send(200, json.dumps({'rev': get_rev(con)}))
            if path == '/api/planilha/semana' and self.command == 'POST':
                pid = int(data.get('id') or 0)
                con.execute('UPDATE planilha SET semana=? WHERE id=?', (str(data.get('semana') or '')[:20], pid))
                con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'semana alterada', '%d -> %s' % (pid, data.get('semana')), quem))
                bump(con)
                return self.send(200, json.dumps({'rev': get_rev(con)}))
            if path == '/api/deliv' and self.command == 'POST':
                keys = data.get('keys')
                if not isinstance(keys, list) or not quem:
                    return self.jerr(400, 'informe o responsável e as chaves')
                for key in keys:
                    con.execute('INSERT OR REPLACE INTO deliv VALUES(?,?,?)', (str(key), quem, now()))
                    con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'entregue', str(key), quem))
                bump(con)
                return self.send(200, json.dumps({'rev': get_rev(con)}))
            if path == '/api/deliv/remove' and self.command == 'POST':
                key = str(data.get('key') or '')
                con.execute('DELETE FROM deliv WHERE key=?', (key,))
                con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'desfez', key, quem))
                bump(con)
                return self.send(200, json.dumps({'rev': get_rev(con)}))
            if path == '/api/ajuste' and self.command == 'POST':
                key = str(data.get('key') or '')
                status = str(data.get('status') or '')
                if not key or status not in ('ok', 'warn', 'bad', 'na') or not quem:
                    return self.jerr(400, 'informe chave, status e responsável')
                con.execute('INSERT OR REPLACE INTO ajuste VALUES(?,?,?,?,?)',
                            (key, status, str(data.get('nota') or '')[:500], quem, now()))
                con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'ajuste:' + status, key, quem))
                bump(con)
                return self.send(200, json.dumps({'rev': get_rev(con)}))
            if path == '/api/ajuste/remove' and self.command == 'POST':
                key = str(data.get('key') or '')
                con.execute('DELETE FROM ajuste WHERE key=?', (key,))
                con.execute('INSERT INTO log(ts,acao,chave,quem) VALUES(?,?,?,?)', (now(), 'ajuste desfeito', key, quem))
                bump(con)
                return self.send(200, json.dumps({'rev': get_rev(con)}))
        return self.jerr(404, 'não encontrado')


def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return '127.0.0.1'


def main():
    global PIN
    ap = argparse.ArgumentParser()
    ap.add_argument('--porta', type=int, default=8000)
    ap.add_argument('--pin', default=os.environ.get('PAINEL_PIN', ''))
    a = ap.parse_args()
    PIN = a.pin
    init_db()
    with db() as con:
        migrate(con)
    srv = ThreadingHTTPServer(('0.0.0.0', a.porta), Handler)
    print('Painel de Ordens SAP no ar. Ctrl+C para parar.')
    print('  Neste computador : http://localhost:%d' % a.porta)
    print('  Para os colegas  : http://%s:%d' % (lan_ip(), a.porta))
    print('  PIN de acesso    : %s' % (PIN or '(sem PIN)'))
    print('  Dados guardados em: %s' % DB)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\nEncerrado.')


if __name__ == '__main__':
    main()
