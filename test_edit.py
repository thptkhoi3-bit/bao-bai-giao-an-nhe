"""Kiểm tra API sửa 1 dòng báo bài: chỉ đổi đúng dòng/tuần, không ảnh hưởng tuần khác và giáo viên khác."""
import os, tempfile, sys, docx
os.environ['DATA_DIR'] = tempfile.mkdtemp(); os.environ['SECRET_KEY'] = 'x' * 32
os.environ['TEACHER_PASSWORD_SHA256'] = 'feb9a081d082f4d8d31fa3c51cd960922c29d6fb9440d1ac1a00a8c59bde3cdb'
os.environ['GUEST_PASSWORD_SHA256'] = '4fcacb897dec6fd766c0538c222a2f58cfa5e10f13d682facac89f81d02e7011'
import subprocess, time, requests, atexit, core
PORT = 8011; B = f'http://127.0.0.1:{PORT}'
srv = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(PORT)], cwd=os.path.dirname(os.path.abspath(__file__)), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
atexit.register(srv.terminate)
for _ in range(50):
    try: requests.get(B + '/', timeout=1); break
    except Exception: time.sleep(.2)
class C(requests.Session):
    def __init__(s): super().__init__()
    def request(self, m, u, *a, **k):
        # Test-only: requests không gửi cookie Secure qua HTTP -> đồng bộ thủ công vào header
        v = self.cookies.get('gv_session')
        if v: self.headers['Cookie'] = 'gv_session=' + v
        elif 'Cookie' in self.headers: del self.headers['Cookie']
        return super().request(m, B + u, *a, **k)
def client(teacher, pw, name=''):
    c = C(); r = c.post('/api/login', data=dict(teacher=teacher, password=pw, name=name)); assert r.status_code == 200, r.text; return c
a = client('Anh Lê Đạt', 'test-teacher-pw'); g = client('other', 'test-guest-pw', 'Cô Thử')
master = os.path.join(os.environ['DATA_DIR'], 'users'); 
st = a.get('/api/status').json(); weeks = [w['week'] for w in st['weeks']]; w = weeks[-1]
def snap(c):
    return {x: c.get(f'/api/baobai/week/{x}').json() for x in weeks}
before = snap(a)
row = next(r for r in before[w] if r['mon'])
r = a.post('/api/baobai/edit', data=dict(week=w, idx=row['idx'], mon=row['mon'], ten='Bài kiểm thử sửa', tich='Tích hợp thử'))
assert r.status_code == 200, r.text
after = snap(a)
for x in weeks:
    if x != w: assert after[x] == before[x], f'tuần {x} bị đổi'
for i, (p, q) in enumerate(zip(before[w], after[w])):
    if i == row['idx']: assert q['ten'] == 'Bài kiểm thử sửa' and q['tich'] == 'Tích hợp thử' and not q['empty']
    else: assert p == q, f'dòng {i} bị đổi'
# định dạng TNR 12
path = [os.path.join(dp, f) for dp, _, fs in os.walk(master) for f in fs if f == 'baobai_goc.docx' and 'backup' not in dp]
M = docx.Document(next(p for p in path if 'cô' not in p.lower()))
cell = core.rows_of(core.week_tables(M)[w][0])[row['idx']]['cells'][3]
rp = cell.paragraphs[0].runs[0]._r.rPr.xml
assert 'Times New Roman' in rp and 'w:val="24"' in rp
# sai môn -> 409, sai dòng -> 404, tuần sai -> 404
assert a.post('/api/baobai/edit', data=dict(week=w, idx=row['idx'], mon='Môn khác', ten='x')).status_code == 409
assert a.post('/api/baobai/edit', data=dict(week=w, idx=9999, ten='x')).status_code == 404
assert a.post('/api/baobai/edit', data=dict(week=999, idx=0, ten='x')).status_code == 404
# giáo viên khác: không có báo bài gốc, không thấy/sửa được dữ liệu của Anh Lê Đạt
assert g.post('/api/baobai/edit', data=dict(week=w, idx=row['idx'], ten='hack')).status_code == 400
assert snap(a)[w] == after[w]
# chưa đăng nhập
assert C().post('/api/baobai/edit', data=dict(week=w, idx=0, ten='x')).status_code == 401
print('EDIT OK: chỉ dòng', row['idx'], 'tuần', w, 'thay đổi; các tuần khác', [x for x in weeks if x != w], 'giữ nguyên; tách biệt giáo viên OK')
