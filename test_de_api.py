"""v16 API: tạo đề từ Word, sửa + dựng lại, tải Word, ảnh khi chưa có key, giới hạn lượt AI, thay đề mẫu (chỉ quản trị), không lộ key."""
import os, sys, json, tempfile, threading, subprocess, time, atexit, zipfile, io, re
from http.server import BaseHTTPRequestHandler, HTTPServer
import requests
HERE = os.path.dirname(os.path.abspath(__file__))
KEY = 'AIzaFAKE-api-test-key-987'
RESP = {'sections': [{'title': 'I. TRẮC NGHIỆM', 'questions': [{'so': 1, 'text': 'Số liền sau của 99 là:', 'muc': 'M1', 'diem': '0,5', 'blocks': [{'t': 'options', 'items': ['98', '100', '101', '90'], 'layout': 'row'}]}]}], 'figures': []}
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        s.rfile.read(int(s.headers['Content-Length']))
        ok = s.headers.get('x-goog-api-key') == KEY
        out = json.dumps({'candidates': [{'content': {'parts': [{'text': json.dumps(RESP)}]}}]}).encode()
        s.send_response(200 if ok else 403); s.end_headers(); s.wfile.write(out)
gs = HTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()

def server(port, extra):
    D = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=D, SECRET_KEY='y' * 32,
        TEACHER_PASSWORD_SHA256='feb9a081d082f4d8d31fa3c51cd960922c29d6fb9440d1ac1a00a8c59bde3cdb',
        GUEST_PASSWORD_SHA256='4fcacb897dec6fd766c0538c222a2f58cfa5e10f13d682facac89f81d02e7011', **extra)
    for k in ('DATA_REPO', 'DATA_TOKEN'): env.pop(k, None)
    if 'GEMINI_API_KEY' not in extra: env.pop('GEMINI_API_KEY', None)
    log = open(os.path.join(D, 'server.log'), 'w')
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(port)], cwd=HERE, env=env, stdout=log, stderr=subprocess.STDOUT)
    atexit.register(p.terminate)
    for _ in range(80):
        try: requests.get(f'http://127.0.0.1:{port}/', timeout=1); break
        except Exception: time.sleep(.2)
    return D, p
class C(requests.Session):
    def __init__(s, port): super().__init__(); s.B = f'http://127.0.0.1:{port}'
    def request(self, m, u, *a, **k):
        # Test-only: requests không gửi cookie Secure qua HTTP -> đồng bộ thủ công vào header
        v = self.cookies.get('gv_session')
        if v: self.headers['Cookie'] = 'gv_session=' + v
        elif 'Cookie' in self.headers: del self.headers['Cookie']
        return super().request(m, self.B + u, *a, **k)
F = dict(mon='toan', xa='Hoà Bình', truong='TH Kim Đồng', nam='2026 - 2027', lop='4', loai='Cuối học kì I', thoigian='40 phút', dapan='1')
from PIL import Image
buf = io.BytesIO(); Image.new('RGB', (600, 800), 'white').save(buf, 'PNG'); PNG = buf.getvalue()

# ---------- A. không có key ----------
D, p = server(8020, {})
c = C(8020)
assert c.get('/api/de/info').status_code == 401
assert c.post('/api/login', data=dict(teacher='Cô Lê Tâm', password='test-teacher-pw')).status_code == 200
info = c.get('/api/de/info').json(); assert info['ai'] is False and info['admin'] is False and info['ai_limit'] == 30
r = c.post('/api/de/extract', data=F, files=[('files', ('DeMau.docx', open(os.path.join(HERE, 'de_templates', 'toan.docx'), 'rb').read()))])
assert r.status_code == 200, r.text
j = r.json(); assert j['nq'] == 12 and j['mon'] == 'toan' and j['fields']['truong'] == 'TH Kim Đồng'
assert 'raw' not in j['exam'] and '_answers_orig' not in j['exam']
if j['pages']:
    assert len(j['pages']) == 4 and j['labels'][0] == 'Trắc nghiệm' and j['labels'][-1] == 'Đáp án'
    assert c.get(j['pages'][0]).headers['content-type'] == 'image/png'
else: assert '<div class="hp">' in j['html']
dl = c.get(j['download']); assert dl.status_code == 200 and dl.content[:2] == b'PK'
assert 'DeKiemTra_Toan_Lop4.docx' in dl.headers.get('content-disposition', '')
assert c.get('/api/de/info').json()['prefs']['truong'] == 'TH Kim Đồng', 'nhớ trường theo người dùng'
# sửa nội dung rồi dựng lại
ex = j['exam']; ex['sections'][0]['questions'][1]['text'] = 'Số liền trước của 100 là số nào?'
ex['sections'][0]['questions'][1]['blocks'][0]['items'][0] = '99'; ex['answers']['2']['text'] = 'A'
r = c.post(f"/api/de/{j['id']}/build", json=dict(exam=ex, fields=j['fields'], dapan=True)); assert r.status_code == 200, r.text
z = zipfile.ZipFile(io.BytesIO(c.get(j['download']).content)); x = z.read('word/document.xml').decode()
assert 'Số liền trước của 100 là số nào?' in x and 'Số 99 là số liền sau' not in x
# người khác không xem được đề này
c2 = C(8020); c2.post('/api/login', data=dict(teacher='Cô Linh', password='test-teacher-pw'))
assert c2.get(f"/api/de/{j['id']}").status_code == 404 and c2.get(j['download']).status_code == 404
assert c.get('/api/de/../../x').status_code in (404, 405)
# ảnh khi chưa bật AI
r = c.post('/api/de/extract', data=F, files=[('files', ('anh.png', PNG))]); assert r.status_code == 400 and r.json()['detail'] == 'Chưa bật đọc ảnh – liên hệ 0865291703', r.text
r = c.post('/api/de/extract', data=F, files=[('files', ('x.exe', b'MZ'))]); assert r.status_code == 400
r = c.post('/api/de/extract', data=dict(F, lop='7'), files=[('files', ('DeMau.docx', b'x'))]); assert r.status_code == 400
r = c.post('/api/de/extract', data=dict(F, truong=' '), files=[('files', ('DeMau.docx', b'x'))]); assert r.status_code == 400 and 'trường' in r.json()['detail']
# TV: 3 ô tải lên
tv = open(os.path.join(HERE, 'de_templates', 'tv.docx'), 'rb').read()
r = c.post('/api/de/extract', data=dict(F, mon='tv'), files=[('doc_hieu', ('tv.docx', tv))]); assert r.status_code == 200, r.text
jt = r.json(); assert jt['nq'] == 9 and jt['exam']['reading']['title'] == 'PHẦN THƯỞNG'
if jt['pages']: assert len(jt['pages']) == 7 and jt['labels'][-1] == 'Hướng dẫn chấm'
# thay đề mẫu: chỉ quản trị
assert c.post('/api/admin/de/template', data=dict(mon='toan'), files=dict(file=('m.docx', tv))).status_code == 403
a = C(8020); a.post('/api/login', data=dict(teacher='Anh Lê Đạt', password='test-teacher-pw'))
assert a.get('/api/de/info').json()['admin'] is True
r = a.post('/api/admin/de/template', data=dict(mon='toan'), files=dict(file=('m.docx', b'not a docx'))); assert r.status_code == 400
r = a.post('/api/admin/de/template', data=dict(mon='toan'), files=dict(file=('mau.docx', open(os.path.join(HERE, 'de_templates', 'toan.docx'), 'rb').read()))); assert r.status_code == 200 and r.json()['templates']['toan']['custom'] is True, r.text
assert os.path.exists(os.path.join(D, 'de_templates', 'toan.docx'))
r = a.post('/api/admin/de/template', data=dict(mon='toan', reset='1')); assert r.json()['templates']['toan']['custom'] is False
# tệp gốc không bị đồng bộ, nháp lưu trong thư mục người dùng
ud = os.path.join(D, 'users', 'gv-co-le-tam', 'de', jt['id']); assert os.path.exists(os.path.join(ud, 'exam.json')) and os.path.exists(os.path.join(ud, 'de.docx'))
assert os.listdir(os.path.dirname(ud)) == [jt['id']], 'chỉ giữ đề mới nhất'
p.terminate(); p.wait()
print('A. không key: Word/TV/sửa/tải/quyền OK')

# ---------- B. có key (giả lập) + giới hạn 1 lượt/ngày ----------
D, p = server(8021, dict(GEMINI_API_KEY=KEY, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}/v1beta', AI_DAILY_LIMIT='1'))
c = C(8021); c.post('/api/login', data=dict(teacher='other', name='Cô Hằng', password='test-guest-pw'))
info = c.get('/api/de/info').json(); assert info['ai'] is True and info['ai_used'] == 0 and info['ai_limit'] == 1
assert KEY not in json.dumps(info)
r = c.post('/api/de/extract', data=F, files=[('files', ('anh.png', PNG))]); assert r.status_code == 200, r.text
assert r.json()['nq'] == 1 and KEY not in r.text
assert c.get('/api/de/info').json()['ai_used'] == 1
r = c.post('/api/de/extract', data=F, files=[('files', ('anh2.png', PNG))]); assert r.status_code == 429 and 'hết 1 lượt' in r.json()['detail'], r.text
r = c.post('/api/de/extract', data=F, files=[('files', ('DeMau.docx', open(os.path.join(HERE, 'de_templates', 'toan.docx'), 'rb').read()))]); assert r.status_code == 200, 'Word không tính lượt AI'
p.terminate(); p.wait()
logs = open(os.path.join(D, 'server.log')).read() + ''.join(open(os.path.join(dp, f), errors='ignore').read() for dp, _, fs in os.walk(D) for f in fs if f.endswith(('.json', '.log')))
assert KEY not in logs, 'KHÔNG được ghi key ra log / dữ liệu'
print('B. có key: đọc ảnh + giới hạn lượt + không lộ key OK')

# ---------- C. đề B sau đề A: không lẫn nội dung, không còn tệp của A ----------
import docx as _docx
def small_docx():
    d = _docx.Document()
    d.add_paragraph('I. TRẮC NGHIỆM')
    d.add_paragraph('Câu 1. Bạn Bình có 7 quả cam, mẹ cho thêm 5 quả. Bình có tất cả mấy quả cam? (M1) (0,5đ)')
    d.add_paragraph('A. 12\tB. 13\tC. 11\tD. 2')
    d.add_paragraph('II. TỰ LUẬN')
    d.add_paragraph('Câu 2. Tính nhẩm: 30 + 40 = ? (M1) (1đ)')
    d.add_paragraph('.' * 200)
    b = io.BytesIO(); d.save(b); return b.getvalue()
D, p = server(8023, {})
c = C(8023); c.post('/api/login', data=dict(teacher='Cô Lê Thủy', password='test-teacher-pw'))
rA = c.post('/api/de/extract', data=F, files=[('files', ('DeMau.docx', open(os.path.join(HERE, 'de_templates', 'toan.docx'), 'rb').read()))]); assert rA.status_code == 200, rA.text
A = rA.json(); assert A['nq'] == 12
udir = os.path.join(D, 'users', 'gv-co-le-thuy', 'de')
assert os.listdir(udir) == [A['id']] and os.listdir(os.path.join(udir, A['id'], 'src'))
rB = c.post('/api/de/extract', data=F, files=[('files', ('de_B.docx', small_docx()))]); assert rB.status_code == 200, rB.text
Bj = rB.json(); assert Bj['nq'] == 2 and Bj['id'] != A['id'], Bj['nq']
assert os.listdir(udir) == [Bj['id']], ('đề cũ phải bị xoá hết', os.listdir(udir))
srcB = os.listdir(os.path.join(udir, Bj['id'], 'src')); assert srcB == ['0_all.docx'], srcB
assert c.get(f"/api/de/{A['id']}").status_code == 404 and c.get(A['download']).status_code == 404
if A['pages']: assert c.get(A['pages'][0]).status_code == 404
qs = [Q['text'] for s in Bj['exam']['sections'] for Q in s['questions']]
assert len(qs) == 2 and 'quả cam' in qs[0] and 'Tính nhẩm' in qs[1], qs
assert set(Bj['exam']['answers']) <= {'1', '2'} and not Bj['exam']['figures'], 'không mang đáp án/hình của đề A'
x = zipfile.ZipFile(io.BytesIO(c.get(Bj['download']).content)).read('word/document.xml').decode()
import re as _re
titles = _re.findall(r'Câu (\d+)\.', _re.sub(r'<[^>]+>', '', x))
assert 'quả cam' in x and 'Số 99 là số liền sau' not in x and 'Viết số thích hợp' not in x and 'tấm vải' not in x, 'đề B lẫn câu của đề A'
assert sorted(set(titles)) == ['1', '2'], titles
if Bj['pages']:
    assert all(Bj['id'] in u for u in Bj['pages']) and len(Bj['pages']) >= 1
    import pymupdf
    pdf = pymupdf.open(os.path.join(udir, Bj['id'], 'pages', 'de.pdf')); t = ''.join(pg.get_text() for pg in pdf)
    assert 'quả cam' in t and 'Số 99' not in t, 'xem trước phải là đề B'
# đề của người khác không bị xoá
c2 = C(8023); c2.post('/api/login', data=dict(teacher='Cô Linh', password='test-teacher-pw'))
r2 = c2.post('/api/de/extract', data=F, files=[('files', ('de_B.docx', small_docx()))]); assert r2.status_code == 200
assert os.listdir(udir) == [Bj['id']], 'tạo đề của người khác không đụng tới đề của mình'
assert c.post('/api/de/reset').json()['ok'] and os.listdir(udir) == [] and c2.get(f"/api/de/{r2.json()['id']}").status_code == 200
p.terminate(); p.wait()
print('C. đề B sau đề A: sạch, không lẫn OK')
gs.shutdown(); print('TẤT CẢ OK')
