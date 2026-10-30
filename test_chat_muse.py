"""V76: cô bé chat dùng não Muse Spark (Meta Model API) — test toàn trình, máy chủ Meta GIẢ lập.
Kiểm chứng: ưu tiên Muse hơn Gemini dự phòng; đúng giao thức OpenAI (đường dẫn, Bearer, model);
system prompt là Hiền + luật bảo mật của anh; quản trị được nhận diện (anh/em), giáo viên thì không;
lỗi 401 từ Meta -> thông báo thân thiện; không key nào -> FAQ như cũ.
Chạy: ../venv/bin/python test_chat_muse.py
"""
import os, sys, json, time, hashlib, tempfile, threading, subprocess, atexit
import urllib.request, urllib.error, urllib.parse, http.cookiejar
from http.server import BaseHTTPRequestHandler, HTTPServer

HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
MKEY = 'test-muse-key-123'
SEEN = []       # các lần gọi vào máy chủ Meta giả
GEM_HITS = []   # nếu code lỡ gọi đường Gemini (/v1beta/...) sẽ ghi vào đây
REPLY = 'Chào anh, em là Hiền đây ạ. Anh cứ hỏi em bất cứ điều gì nhé!'

STRICT = {'on': False, 'dead': False}   # on: chỉ nhận REAL_MODEL; dead: Meta chết hẳn (mọi tên 404, /models rỗng)
REAL_MODEL = 'muse-spark-1.3-real'
MODEL_404 = []

class FakeMeta(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_GET(s):
        if s.path.startswith('/v1/models'):
            ids = [] if STRICT['dead'] else ([REAL_MODEL] if STRICT['on'] else ['muse-spark-1.3'])
            out = json.dumps({'data': [{'id': i} for i in ids]}).encode()
            s.send_response(200); s.end_headers(); s.wfile.write(out); return
        s.send_response(404); s.end_headers()
    def do_POST(s):
        b = json.loads(s.rfile.read(int(s.headers['Content-Length'])))
        if s.path.startswith('/v1beta/'):
            GEM_HITS.append(s.path)
            out = json.dumps({'candidates': [{'content': {'parts': [{'text': 'GEMINI'}]}}]}).encode()
            s.send_response(200); s.end_headers(); s.wfile.write(out); return
        SEEN.append((s.path, s.headers.get('Authorization'), b))
        if STRICT['dead'] or (STRICT['on'] and b.get('model') != REAL_MODEL):
            MODEL_404.append(b.get('model'))
            s.send_response(404); s.end_headers(); s.wfile.write(b'{"error":{"message":"model not found"}}'); return
        last = b['messages'][-1]['content']
        if 'loi-key' in last:
            s.send_response(401); s.end_headers(); s.wfile.write(b'{"error":"unauthorized"}'); return
        out = json.dumps({'choices': [{'message': {'role': 'assistant', 'content': REPLY}}]}).encode()
        s.send_response(200); s.end_headers(); s.wfile.write(out)

fake = HTTPServer(('127.0.0.1', 0), FakeMeta)
threading.Thread(target=fake.serve_forever, daemon=True).start()
FBASE = f'http://127.0.0.1:{fake.server_port}/v1'
GBASE = f'http://127.0.0.1:{fake.server_port}/v1beta'

def sha(x): return hashlib.sha256(x.encode()).hexdigest()

def boot(port, extra):
    D = tempfile.mkdtemp()
    env = dict(os.environ, DATA_DIR=D, SECRET_KEY='c' * 32,
               TEACHER_PASSWORD_SHA256=sha('1993'), GUEST_PASSWORD_SHA256=sha('0865291703'),
               ADMIN_PASSWORD_SHA256=sha('16122022'), ADMIN_PHONE='0865291703')
    for k in ('DATA_REPO', 'DATA_TOKEN', 'CHAT_DAILY_LIMIT', 'MODEL_API_KEY', 'META_API_BASE', 'MUSE_MODEL', 'GEMINI_API_KEY', 'GEMINI_API_URL'):
        env.pop(k, None)
    env.update(extra)
    lg = open(os.path.join(D, 'server.log'), 'w')
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(port)], cwd=HERE, env=env, stdout=lg, stderr=subprocess.STDOUT)
    atexit.register(p.terminate)
    for _ in range(100):
        try: urllib.request.urlopen(f'http://127.0.0.1:{port}/', timeout=1); break
        except Exception: time.sleep(.2)
    return p

class C:
    """Client test tự quản cookie: cookie phiên của app có cờ Secure (đúng chuẩn HTTPS)
    nên CookieJar thường không gửi lại qua HTTP nội bộ — ở đây tự bắt Set-Cookie và gắn tay."""
    def __init__(s, port):
        s.B = f'http://127.0.0.1:{port}'
        s.op = urllib.request.build_opener()
        s.cookies = {}
    def _grab(s, headers):
        for h in (headers.get_all('Set-Cookie') if headers else None) or []:
            kv = h.split(';', 1)[0].strip()
            if '=' in kv:
                k, v = kv.split('=', 1); s.cookies[k.strip()] = v.strip()
    def req(s, method, path, data=None, is_json=False):
        body = None; hdrs = {}
        if data is not None:
            if is_json: body = json.dumps(data).encode(); hdrs['Content-Type'] = 'application/json'
            else:
                body = urllib.parse.urlencode(data).encode(); hdrs['Content-Type'] = 'application/x-www-form-urlencoded'
        if s.cookies: hdrs['Cookie'] = '; '.join(f'{k}={v}' for k, v in s.cookies.items())
        r = urllib.request.Request(s.B + path, data=body, method=method, headers=hdrs)
        try:
            with s.op.open(r, timeout=30) as resp:
                s._grab(resp.headers); return resp.status, json.loads(resp.read() or b'{}')
        except urllib.error.HTTPError as e:
            s._grab(e.headers)
            try: return e.code, json.loads(e.read() or b'{}')
            except Exception: return e.code, {}
    def login_admin(s):
        st, r = s.req('POST', '/api/login', dict(teacher='registered', phone='0865291703', password='16122022'))
        assert st == 200, r

# ---------- A. có key Muse (+ cả key Gemini): Muse phải thắng ----------
p = boot(8097, dict(MODEL_API_KEY=MKEY, META_API_BASE=FBASE, GEMINI_API_KEY='gem-fake', GEMINI_API_URL=GBASE))
adm = C(8097); adm.login_admin()
st, info = adm.req('GET', '/api/chat')
assert st == 200 and info['ai'] is True and info['brain'] == 'muse' and info['name'] == 'Hiền', info
st, r = adm.req('POST', '/api/chat', dict(message='Chào em, em là ai?'), is_json=True)
assert st == 200 and r['src'] == 'ai' and r['reply'] == REPLY, r
path, auth, body = SEEN[-1]
assert path == '/v1/chat/completions' and auth == 'Bearer ' + MKEY, (path, auth)
assert body['model'] == 'muse-spark-1.3', body['model']
sysp = body['messages'][0]
assert sysp['role'] == 'system' and 'Hiền' in sysp['content'] and 'BẢO MẬT' in sysp['content'], 'system prompt thiếu Hiền/luật bảo mật'
assert 'anh Lê Đắc Đạt' in sysp['content'] and 'trợ lý riêng' in sysp['content'], 'quản trị phải được nhận diện là anh Đạt'
assert body['messages'][-1] == {'role': 'user', 'content': 'Chào em, em là ai?'}
assert not GEM_HITS, f'không được gọi Gemini khi đã có Muse: {GEM_HITS}'
st, r = adm.req('POST', '/api/chat', dict(message='Kể tiếp đi em'), is_json=True)
roles = [m['role'] for m in SEEN[-1][2]['messages']]
assert roles == ['system', 'user', 'assistant', 'user'], roles
print('A. Muse thắng Gemini, giao thức OpenAI đúng, quản trị được nhận diện, lịch sử assistant đúng')

# giáo viên thường: có tên trong ngữ cảnh, KHÔNG có đoạn "trợ lý riêng của anh"
st, rr = adm.req('POST', '/api/register', dict(name='Cô Kiểm V76', phone='0900000076', school='Test', password='kiemv76x', password2='kiemv76x'))
st, accs = adm.req('GET', '/api/admin/accounts')
aid = [a['id'] for a in accs['accounts'] if a['phone'] == '0900000076'][0]
adm.req('POST', f'/api/admin/accounts/{aid}', dict(action='approve'))
gv = C(8097)
st, r = gv.req('POST', '/api/login', dict(teacher='registered', phone='0900000076', password='kiemv76x')); assert st == 200, r
n0 = len(SEEN)
st, r = gv.req('POST', '/api/chat', dict(message='Em ơi, 2 + 2 bằng mấy?'), is_json=True)
assert st == 200 and r['src'] == 'ai', r
sysp = SEEN[-1][2]['messages'][0]['content']
assert 'Cô Kiểm V76' in sysp and 'trợ lý riêng' not in sysp, 'giáo viên: phải có tên, không có đoạn quản trị'
print('B. Giáo viên: nhận diện bằng tên, không lẫn đặc quyền quản trị')

# lỗi key từ Meta -> 502 thân thiện, không lộ chi tiết kỹ thuật
st, r = gv.req('POST', '/api/chat', dict(message='loi-key'), is_json=True)
assert st == 502 and 'lỗi cấu hình khóa' in r['detail'] and MKEY not in json.dumps(r), r
print('C. Meta trả 401 -> thông báo thân thiện, không lộ key')
p.terminate(); p.wait()

# ---------- D. không key nào: FAQ cũ vẫn chạy ----------
p = boot(8098, {})
c = C(8098); c.login_admin()
st, info = c.req('GET', '/api/chat')
assert info['ai'] is False and info['brain'] is None and info['name'] == 'Hiền', info
st, r = c.req('POST', '/api/chat', dict(message='Cách làm báo bài'), is_json=True)
assert r['src'] == 'faq' and 'tự gióng đúng tuần' in r['reply'], r
print('D. Không key: FAQ chạy, nội dung báo bài đã đúng luồng tự gióng tuần')
p.terminate(); p.wait()

# ---------- E. sai tên mô hình cấu hình: tự dò qua /models, nhớ luôn cho lần sau ----------
STRICT['on'] = True
p = boot(8096, dict(MODEL_API_KEY=MKEY, META_API_BASE=FBASE))
adm = C(8096); adm.login_admin()
st, r = adm.req('POST', '/api/chat', dict(message='Chào em'), is_json=True)
assert st == 200 and r['reply'] == REPLY, r
assert SEEN[-1][2]['model'] == REAL_MODEL, SEEN[-1][2]['model']
n404 = len(MODEL_404)
st, r = adm.req('POST', '/api/chat', dict(message='Câu nữa nhé'), is_json=True)
assert st == 200 and len(MODEL_404) == n404, 'lần 2 phải đi thẳng mô hình đã dò được, không 404 lại'
st, dg = adm.req('GET', '/api/chat/diag')
assert dg['models'] == [REAL_MODEL] and dg['probe']['http'] == 404, dg
print('E. Sai tên mô hình: tự dò đúng qua /models, nhớ cho lần sau; /api/chat/diag báo đúng sự thật')
p.terminate(); p.wait()

# ---------- F. Muse lỗi (404 mọi mô hình) + có Gemini: lùi về Gemini, không lỗi trắng ----------
STRICT['dead'] = True
p = boot(8095, dict(MODEL_API_KEY=MKEY, META_API_BASE=FBASE, GEMINI_API_KEY='gem-fake', GEMINI_API_URL=GBASE))
adm = C(8095); adm.login_admin()
st, r = adm.req('POST', '/api/chat', dict(message='Chào em'), is_json=True)
assert st == 200 and r['reply'] == 'GEMINI', r
assert GEM_HITS, 'phải có gọi Gemini dự phòng'
print('F. Muse hỏng -> tự lùi Gemini, thầy cô không thấy lỗi trắng')
p.terminate(); p.wait()
print('TẤT CẢ OK — test_chat_muse')
