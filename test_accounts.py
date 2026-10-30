"""Đăng ký chờ duyệt, duyệt/từ chối/khóa/hết hạn/xóa, chỉ quản trị, tách dữ liệu, giới hạn số lần thử."""
import os, tempfile, sys, json, datetime
D = tempfile.mkdtemp(); os.environ.update(DATA_DIR=D, SECRET_KEY='x' * 32, RL_REGISTER='9', RL_LOGIN='12',
                  TEACHER_PASSWORD_SHA256='feb9a081d082f4d8d31fa3c51cd960922c29d6fb9440d1ac1a00a8c59bde3cdb',
                  GUEST_PASSWORD_SHA256='4fcacb897dec6fd766c0538c222a2f58cfa5e10f13d682facac89f81d02e7011')
os.environ.pop('DATA_REPO', None); os.environ.pop('DATA_TOKEN', None)
import subprocess, time, requests, atexit
PORT = 8015; B = f'http://127.0.0.1:{PORT}'
srv = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(PORT)], cwd=os.path.dirname(os.path.abspath(__file__)), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
atexit.register(srv.terminate)
for _ in range(50):
    try: requests.get(B + '/', timeout=1); break
    except Exception: time.sleep(.2)
class C(requests.Session):
    def __init__(s, ip): super().__init__(); s.headers['X-Forwarded-For'] = ip
    def request(self, m, u, *a, **k):
        # Test-only: requests không gửi cookie Secure qua HTTP -> đồng bộ thủ công vào header
        v = self.cookies.get('gv_session')
        if v: self.headers['Cookie'] = 'gv_session=' + v
        elif 'Cookie' in self.headers: del self.headers['Cookie']
        return super().request(m, B + u, *a, **k)
def reg(c, name, phone, pw='matkhau1', pw2=None, school=''):
    return c.post('/api/register', data=dict(name=name, phone=phone, school=school, password=pw, password2=pw2 or pw))
def lg(c, **d): return c.post('/api/login', data=d)
anon = C('10.0.0.1')
# --- kiểm tra dữ liệu vào ---
assert reg(anon, 'A', '0912345678').status_code == 400
assert reg(anon, 'Cô Hoa', '12345').status_code == 400
assert reg(anon, 'Cô Hoa', '0912345678', 'abc').status_code == 400
assert reg(anon, 'Cô Hoa', '0912345678', 'matkhau1', 'khac123').status_code == 400
r = reg(anon, 'Cô Hoa', '091 234 5678', school='TH Lê Lợi - lớp 3A'); assert r.status_code == 200, r.text
assert r.json()['message'] == 'Anh/chị đã đăng ký, vui lòng chờ duyệt. Liên hệ 0865291703.' and r.json()['status'] == 'pending'
assert reg(anon, 'Cô Hoa 2', '+84912345678').status_code == 409  # trùng SĐT (chuẩn hoá +84)
assert reg(anon, 'Thầy Nam', '0987654321').status_code == 200
assert reg(anon, 'Cô Mai', '0356789012').status_code == 200
acc = json.load(open(os.path.join(D, 'accounts.json'), encoding='utf-8'))
hoa = next(a for a in acc.values() if a['phone'] == '0912345678')
assert hoa['pw'].startswith('scrypt$') and 'matkhau1' not in json.dumps(acc), 'mật khẩu phải được băm bằng KDF'
assert len({a['pw'].split('$')[4] for a in acc.values()}) == 3, 'mỗi tài khoản muối riêng'
# --- chưa duyệt ---
c = C('10.0.0.2')
r = lg(c, teacher='registered', phone='0912345678', password='matkhau1'); assert r.status_code == 403 and 'chờ duyệt' in r.json()['detail'], r.text
assert lg(c, teacher='registered', phone='0912345678', password='sai').json()['detail'] == 'Sai số điện thoại hoặc mật khẩu'
# --- chỉ quản trị ---
admin = C('10.0.0.3'); assert lg(admin, teacher='Anh Lê Đạt', password='test-teacher-pw').status_code == 200
tam = C('10.0.0.4'); assert lg(tam, teacher='Cô Lê Tâm', password='test-teacher-pw').status_code == 200
fake = C('10.0.0.5'); assert lg(fake, teacher='other', name='Anh Lê Đạt', password='test-guest-pw').status_code == 200
for cl in (anon, tam, fake):
    assert cl.get('/api/admin/accounts').status_code in (401, 403)
    assert cl.post(f"/api/admin/accounts/{hoa['id']}", data={'action': 'approve'}).status_code in (401, 403)
    assert not cl.get('/api/me').json().get('admin') if cl is not anon else True
me = admin.get('/api/me').json(); assert me['admin'] and me['pending'] == 3, me
L = admin.get('/api/admin/accounts').json(); assert L['pending'] == 3 and all('pw' not in a for a in L['accounts'])
ids = {a['phone']: a['id'] for a in L['accounts']}
# --- duyệt: mặc định +1 năm ---
a = admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'approve'}).json()
assert a['status'] == 'approved' and a['expires'] == (datetime.date.today() + datetime.timedelta(days=365)).isoformat()
r = lg(c, teacher='registered', phone='0912 345 678', password='matkhau1'); assert r.status_code == 200 and r.json()['name'] == 'Cô Hoa', r.text
assert c.get('/api/me').json()['name'] == 'Cô Hoa' and c.get('/api/status').status_code == 200
assert os.path.isdir(os.path.join(D, 'users', ids['0912345678'])), 'thư mục dữ liệu riêng data/users/<id>'
# --- tách dữ liệu: Cô Hoa chưa có báo bài gốc, Anh Đạt có bản mẫu ---
assert c.get('/api/status').json()['master'] is False and admin.get('/api/status').json()['master'] is True
assert c.get('/api/baobai/download').status_code == 404
# --- từ chối ---
admin.post(f"/api/admin/accounts/{ids['0987654321']}", data={'action': 'reject'})
r = lg(C('10.0.0.6'), teacher='registered', phone='0987654321', password='matkhau1'); assert r.status_code == 403 and 'từ chối' in r.json()['detail']
# --- khóa: phiên đang mở bị chặn ngay ---
admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'lock'})
r = c.get('/api/status'); assert r.status_code == 401 and 'khóa' in r.json()['detail']
assert 'khóa' in lg(c, teacher='registered', phone='0912345678', password='matkhau1').json()['detail']
admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'unlock'})
assert lg(c, teacher='registered', phone='0912345678', password='matkhau1').status_code == 200
# --- hết hạn ---
y = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
assert admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'expiry', 'expires': y}).json()['expired']
assert c.get('/api/status').status_code == 401
r = lg(c, teacher='registered', phone='0912345678', password='matkhau1'); assert r.status_code == 403 and 'hết hạn' in r.json()['detail'] and '0865291703' in r.json()['detail']
fut = (datetime.date.today() + datetime.timedelta(days=30)).isoformat()
assert admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'expiry', 'expires': fut}).json()['expires'] == fut
assert lg(c, teacher='registered', phone='0912345678', password='matkhau1').status_code == 200
# duyệt với ngày tự chọn
assert admin.post(f"/api/admin/accounts/{ids['0356789012']}", data={'action': 'approve', 'expires': fut}).json()['expires'] == fut
assert admin.post(f"/api/admin/accounts/{ids['0356789012']}", data={'action': 'xx'}).status_code == 400
assert admin.post(f"/api/admin/accounts/{ids['0356789012']}", data={'action': 'expiry', 'expires': 'abc'}).status_code == 400
# --- xóa ---
assert admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'delete'}).json()['ok']
assert not os.path.exists(os.path.join(D, 'users', ids['0912345678']))
assert c.get('/api/status').status_code == 401
assert admin.post(f"/api/admin/accounts/{ids['0912345678']}", data={'action': 'approve'}).status_code == 404
# --- 4 giáo viên cố định + khách vẫn như cũ ---
for t in ('Anh Lê Đạt', 'Cô Lê Tâm', 'Cô Lê Thủy', 'Cô Linh'):
    assert lg(C('10.0.1.1'), teacher=t, password='test-teacher-pw').status_code == 200
assert lg(C('10.0.1.1'), teacher='other', name='Cô Thử', password='test-guest-pw').status_code == 200
# --- giới hạn theo IP ---
bad = C('10.9.9.9')
codes = [lg(bad, teacher='Cô Linh', password='sai').status_code for _ in range(13)]
assert codes[:12] == [403] * 12 and codes[12] == 429, codes
assert lg(bad, teacher='Cô Linh', password='test-teacher-pw').status_code == 429  # bị chặn tạm thời
assert lg(C('10.9.9.8'), teacher='Cô Linh', password='test-teacher-pw').status_code == 200  # IP khác không ảnh hưởng
sp = C('10.8.8.8'); codes = [reg(sp, f'Người {i}', f'09{i:08d}').status_code for i in range(10)]
assert codes[:9] == [200] * 9 and codes[9] == 429, codes
print('ACCOUNTS OK')
