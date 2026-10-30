"""Lớp đồng bộ GitHub kiểm thử với một GitHub GIẢ (HTTP mock tại chỗ) — không dùng token thật."""
import os, sys, json, base64, hashlib, threading, tempfile, time, logging, io, subprocess, atexit, re
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import storage, requests
TOKEN = 'github_pat_FAKE_' + 'z' * 40
class Fake:
    def __init__(s): s.objs = {}; s.refs = {}; s.lock = threading.Lock(); s.calls = []; s.race = 0; s.auth_bad = 0
    def put(s, kind, data):
        sha = hashlib.sha1(json.dumps([kind, data], sort_keys=True).encode()).hexdigest() if kind != 'blob' else storage.git_sha(data)
        s.objs[sha] = (kind, data); return sha
    def files(s, branch='main'):  # path -> bytes của commit đầu nhánh
        h = s.refs.get(branch)
        if not h: return {}
        return {p: s.objs[b][1] for p, b in s.objs[s.objs[h][1]['tree']][1].items()}
    def commit_files(s, files, msg='ngoài'):  # mô phỏng một máy khác đẩy commit (xung đột)
        tree = s.objs[s.objs[s.refs['main']][1]['tree']][1].copy() if s.refs.get('main') else {}
        for p, b in files.items(): tree[p] = s.put('blob', b)
        t = s.put('tree', tree); s.refs['main'] = s.put('commit', {'tree': t, 'parents': [s.refs.get('main')], 'm': msg + str(time.time())})
F = Fake()
class H(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def send(s, code, obj=None):
        b = json.dumps(obj or {}).encode(); s.send_response(code); s.send_header('Content-Type', 'application/json'); s.send_header('Content-Length', str(len(b))); s.end_headers(); s.wfile.write(b)
    def body(s): n = int(s.headers.get('Content-Length') or 0); return json.loads(s.rfile.read(n)) if n else {}
    def route(s, m):
        F.calls.append((m, s.path))
        if s.headers.get('Authorization') != 'Bearer ' + TOKEN: F.auth_bad += 1; return s.send(401, {'message': 'Bad credentials'})
        p = s.path.split('?')[0]; mm = re.match(r'/repos/me/bao-bai-data(/.*)', p)
        if not mm: return s.send(404)
        p = mm.group(1); b = s.body() if m in ('POST', 'PATCH') else None
        with F.lock:
            if m == 'GET' and p.startswith('/git/ref/heads/'):
                h = F.refs.get(p.rsplit('/', 1)[1]); return s.send(200, {'object': {'sha': h}}) if h else s.send(404)
            if m == 'GET' and p.startswith('/git/commits/'):
                return s.send(200, {'tree': {'sha': F.objs[p.rsplit('/', 1)[1]][1]['tree']}})
            if m == 'GET' and p.startswith('/git/trees/'):
                t = F.objs[p.rsplit('/', 1)[1]][1]; return s.send(200, {'tree': [{'path': k, 'type': 'blob', 'sha': v, 'mode': '100644'} for k, v in t.items()]})
            if m == 'GET' and p.startswith('/git/blobs/'):
                return s.send(200, {'content': base64.b64encode(F.objs[p.rsplit('/', 1)[1]][1]).decode(), 'encoding': 'base64'})
            if m == 'POST' and p == '/git/blobs': return s.send(201, {'sha': F.put('blob', base64.b64decode(b['content']))})
            if m == 'POST' and p == '/git/trees':
                t = F.objs[F.objs[b['base_tree']][1]['tree']][1].copy() if False else (F.objs[b['base_tree']][1].copy() if b.get('base_tree') else {})
                for e in b['tree']:
                    if e['sha'] is None: t.pop(e['path'], None)
                    else: t[e['path']] = e['sha']
                return s.send(201, {'sha': F.put('tree', t)})
            if m == 'POST' and p == '/git/commits': return s.send(201, {'sha': F.put('commit', {'tree': b['tree'], 'parents': b['parents'], 'm': b['message']})})
            if m == 'POST' and p == '/git/refs': F.refs[b['ref'].rsplit('/', 1)[1]] = b['sha']; return s.send(201, {})
        if m == 'PATCH' and p.startswith('/git/refs/heads/'):
            if F.race: F.race -= 1; F.commit_files({'data/khac.txt': b'may khac'})  # có người đẩy chen ngang
            with F.lock:
                br = p.rsplit('/', 1)[1]
                if F.objs[b['sha']][1]['parents'] != [F.refs.get(br)] and not b.get('force'): return s.send(422, {'message': 'Update is not a fast forward'})
                F.refs[br] = b['sha']; return s.send(200, {})
        s.send(404)
    def do_GET(s): s.route('GET')
    def do_POST(s): s.route('POST')
    def do_PATCH(s): s.route('PATCH')
srv = ThreadingHTTPServer(('127.0.0.1', 0), H); threading.Thread(target=srv.serve_forever, daemon=True).start()
API = f'http://127.0.0.1:{srv.server_port}'
logbuf = io.StringIO(); logging.basicConfig(stream=logbuf, level=logging.DEBUG)
F.commit_files({'README.md': b'# kho du lieu'})
# 1) kho chỉ có README -> kéo về 0 tệp; ghi và đẩy
d1 = tempfile.mkdtemp(); g = storage.GitSync(d1, 'me/bao-bai-data', TOKEN, api=API, debounce=0.3)
assert g.pull() == 0 and g.ready
os.makedirs(f'{d1}/users/gv-a/out'); os.makedirs(f'{d1}/users/gv-a/backup')
open(f'{d1}/accounts.json', 'w').write('{"x": 1}'); open(f'{d1}/users/gv-a/baobai_goc.docx', 'wb').write(b'\x00DOCX' * 1000)
open(f'{d1}/users/gv-a/tmp_abc.docx', 'w').write('tam'); open(f'{d1}/users/gv-a/out/ga.docx', 'w').write('ra')
open(f'{d1}/users/gv-a/backup/baobai_1.docx', 'w').write('bk')
assert g.push() == 3
fs = F.files(); assert set(fs) == {'README.md', 'data/accounts.json', 'data/users/gv-a/baobai_goc.docx', 'data/users/gv-a/backup/baobai_1.docx'}, set(fs)
assert fs['data/users/gv-a/baobai_goc.docx'] == b'\x00DOCX' * 1000 and fs['README.md'] == b'# kho du lieu'
assert g.push() == 0, 'không đổi gì thì không commit'
# 2) đĩa bị xoá sạch (Render khởi động lại) -> kéo về đầy đủ
d2 = tempfile.mkdtemp(); g2 = storage.GitSync(d2, 'me/bao-bai-data', TOKEN, api=API)
assert g2.pull() == 3 and open(f'{d2}/users/gv-a/baobai_goc.docx', 'rb').read() == b'\x00DOCX' * 1000
# 3) xung đột: máy khác đẩy chen giữa -> làm lại trên đầu mới, không mất tệp của ai
open(f'{d2}/accounts.json', 'w').write('{"x": 2}'); F.race = 1
assert g2.push() == 1
fs = F.files(); assert fs['data/accounts.json'] == b'{"x": 2}' and fs['data/khac.txt'] == b'may khac' and 'data/users/gv-a/baobai_goc.docx' in fs
# 4) xoá tệp cục bộ -> xoá trên kho
os.remove(f'{d2}/users/gv-a/backup/baobai_1.docx'); assert g2.push() == 1 and 'data/users/gv-a/backup/baobai_1.docx' not in F.files()
# 5) nền: gom nhiều lần ghi thành 1 commit
g2.debounce = 0.4; g2.start(); n0 = g2.pushes
for i in range(5): open(f'{d2}/accounts.json', 'w').write('{"x": %d}' % (10 + i)); g2.schedule(); time.sleep(0.05)
for _ in range(60):
    if g2.pushes > n0: break
    time.sleep(0.1)
time.sleep(0.6); assert g2.pushes == n0 + 1 and F.files()['data/accounts.json'] == b'{"x": 14}'
# 6) token sai -> lỗi được log, KHÔNG lộ token; thiếu biến -> LocalOnly
gbad = storage.GitSync(tempfile.mkdtemp(), 'me/bao-bai-data', 'sai-token-bimat', api=API)
try: gbad.pull(); assert False
except storage.HTTPErr as e: assert e.code == 401 and 'sai-token' not in str(e)
assert 'sai-token' not in repr(gbad) and TOKEN not in repr(g)
os.environ.pop('DATA_REPO', None); assert isinstance(storage.from_env('/tmp'), storage.LocalOnly)
# 7) tích hợp với ứng dụng: đăng ký -> tự đẩy; máy mới (đĩa trống) kéo về -> duyệt & đăng nhập được
here = os.path.dirname(os.path.abspath(__file__))
def run(port, d):
    env = dict(os.environ, DATA_DIR=d, SECRET_KEY='y' * 32, DATA_REPO='me/bao-bai-data', DATA_TOKEN=TOKEN, DATA_API=API, DATA_DEBOUNCE='0.5',
               TEACHER_PASSWORD_SHA256='feb9a081d082f4d8d31fa3c51cd960922c29d6fb9440d1ac1a00a8c59bde3cdb',
               GUEST_PASSWORD_SHA256='4fcacb897dec6fd766c0538c222a2f58cfa5e10f13d682facac89f81d02e7011')
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(port)], cwd=here, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    atexit.register(p.kill)
    for _ in range(60):
        try: requests.get(f'http://127.0.0.1:{port}/', timeout=1); return p
        except Exception: time.sleep(.2)
    raise SystemExit('server không chạy')
F.commit_files({'data/accounts.json': b'{}'})  # kho thật bắt đầu trống
p1 = run(8016, tempfile.mkdtemp())
r = requests.post('http://127.0.0.1:8016/api/register', data=dict(name='Cô Đồng Bộ', phone='0911222333', password='abcdef', password2='abcdef')); assert r.status_code == 200, r.text
for _ in range(50):
    if b'0911222333' in F.files().get('data/accounts.json', b''): break
    time.sleep(.2)
else: raise SystemExit('accounts.json chưa được đẩy')
p1.terminate(); out1 = p1.communicate(timeout=10)[0]
d3 = tempfile.mkdtemp(); p2 = run(8017, d3)
assert os.path.exists(f'{d3}/accounts.json') and os.path.exists(f'{d3}/users/gv-a/baobai_goc.docx')
s = requests.Session(); s.post('http://127.0.0.1:8017/api/login', data=dict(teacher='Anh Lê Đạt', password='test-teacher-pw'))
s.headers['Cookie'] = 'gv_session=' + s.cookies.get('gv_session')  # test-only: requests không gửi cookie Secure qua HTTP
aid = next(a['id'] for a in s.get('http://127.0.0.1:8017/api/admin/accounts').json()['accounts'] if a['phone'] == '0911222333')
assert s.post(f'http://127.0.0.1:8017/api/admin/accounts/{aid}', data={'action': 'approve'}).status_code == 200
t = requests.post('http://127.0.0.1:8017/api/login', data=dict(teacher='registered', phone='0911222333', password='abcdef')); assert t.status_code == 200, t.text
for _ in range(50):
    if b'approved' in F.files()['data/accounts.json']: break
    time.sleep(.2)
else: raise SystemExit('duyệt chưa được đẩy')
p2.terminate(); out2 = p2.communicate(timeout=10)[0]
assert TOKEN not in out1 + out2 + logbuf.getvalue(), 'token bị ghi vào log!'
assert F.auth_bad == 1  # chỉ lần thử token sai
print('SYNC MOCK OK', len(F.calls), 'lời gọi API giả')
