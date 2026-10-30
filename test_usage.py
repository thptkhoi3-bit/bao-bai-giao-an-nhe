"""Nhật ký sử dụng theo ngày (giờ VN) + báo cáo cho 'Quản Gia' (/workspace/tools/usage_report.py, gh giả)."""
import os, sys, json, time, shutil, tempfile, threading, subprocess, atexit, datetime, requests
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import usage
VN = usage.VN; T = tempfile.mkdtemp()

# 1) đơn vị: gộp an toàn khi nhiều luồng ghi + flush cùng lúc, đúng ngày VN, không mất số liệu
class FakeSync:
    ready = True; n = 0
    def schedule(s): s.n += 1
fs = FakeSync(); U = usage.Usage(T, fs, interval=0.01)
clock = {'t': datetime.datetime(2026, 10, 17, 23, 59, 50, tzinfo=VN)}; U.clock = lambda: clock['t']; U.start()
def worker(i):
    for j in range(200): U.action('gv-a', 'Anh Lê Đạt', 'chat'); (U.login('gv-a', 'Anh Lê Đạt') if j % 50 == 0 else None)
th = [threading.Thread(target=worker, args=(i,)) for i in range(8)]; [t.start() for t in th]
for _ in range(20): U.flush(); time.sleep(0.005)
[t.join() for t in th]; U.flush()
clock['t'] = datetime.datetime(2026, 10, 18, 0, 0, 5, tzinfo=VN); U.login('gv-a', 'Anh Lê Đạt'); U.action('gv-a', 'Anh Lê Đạt', 'xoa_het'); U.flush()
j = json.load(open(os.path.join(T, 'usage', '2026-10-17.json'))); a = j['users']['gv-a']
assert a['actions'] == {'chat': 1600} and a['logins'] == 32 and a['first_login'] == '23:59:50', a
assert len(a['login_times']) == usage.MAX_TIMES
j2 = json.load(open(os.path.join(T, 'usage', '2026-10-18.json'))); assert j2['users']['gv-a']['logins'] == 1 and j2['users']['gv-a']['actions'] == {}
assert fs.n >= 1 and not [f for f in os.listdir(os.path.join(T, 'usage')) if f.endswith('.tmp')]
# ghi lỗi -> giữ lại sự kiện, lần sau ghi tiếp
U2 = usage.Usage(os.path.join(T, 'ro'), None); open(os.path.join(T, 'ro'), 'w').close()   # 'ro' là tệp -> không tạo được thư mục
U2.login('x', 'X'); assert U2.flush() == 0 and len(U2._buf) == 1
print('Gộp an toàn đa luồng, đúng ngày giờ VN, lỗi ghi không mất số liệu: OK')

# 2) máy chủ thật: đăng nhập, chức năng chính, đăng ký mới -> usage/<ngày VN>.json; không có mật khẩu/SĐT/IP; khởi động lại vẫn cộng dồn
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
G0 = A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx'; BBP = A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx'
TPL = '/workspace/GiaoAn_Goc.docx'
DD = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=DD, SECRET_KEY='t' * 32, USAGE_FLUSH='0.3')
for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_API_KEY'): env.pop(k, None)
B = 'http://127.0.0.1:8087'
def up():
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8087'], cwd=HERE, env=env, stdout=open(os.path.join(DD, 'l'), 'a'), stderr=subprocess.STDOUT); atexit.register(p.terminate)
    for _ in range(100):
        try: requests.get(B, timeout=1); return p
        except Exception: time.sleep(.2)
def login(t, pw='1993', name=''):
    s = requests.Session(); r = s.post(B + '/api/login', data=dict(teacher=t, password=pw, name=name)); return s, r
pr = up()
s, r = login('Anh Lê Đạt'); assert r.ok
assert s.post(B + '/api/baobai/master', files=[('file', ('m.docx', open(G0, 'rb')))]).ok and s.post(B + '/api/giaoan/template', files=[('file', ('t.docx', open(TPL, 'rb')))]).ok
assert s.post(B + '/api/baobai/fill', data=dict(week=7), files=[('file', ('p.docx', open(BBP, 'rb')))]).ok
assert s.post(B + '/api/giaoan/build', data=dict(week=7)).ok and s.post(B + '/api/giaoan/build', data=dict(week=7)).ok
assert s.post(B + '/api/giaoan/build', data=dict(week='')).status_code == 400          # lỗi -> không đếm
for m in ('Làm báo bài thế nào?', 'Cảm ơn'): assert s.post(B + '/api/chat', json={'message': m}).ok
login('Anh Lê Đạt'); assert login('Cô Linh', 'sai')[1].status_code == 403             # sai mật khẩu -> không ghi
s2, r = login('other', '0865291703', 'Cô Hoa Mai'); assert r.ok
assert requests.post(B + '/api/register', data=dict(name='Cô Thử Đăng Ký', phone='0912345678', school='TH Đông Hồ', password='abcdef', password2='abcdef')).ok
time.sleep(1.2)
day = datetime.datetime.now(VN).date().isoformat(); P = os.path.join(DD, 'usage', f'{day}.json')
raw = open(P, encoding='utf-8').read(); j = json.loads(raw)
a = j['users']['gv-anh-le-dat']; assert a['name'] == 'Anh Lê Đạt' and a['logins'] == 2, a
assert a['actions'] == {'dien_bao_bai': 1, 'tao_giao_an': 2, 'chat': 2}, a['actions']
assert j['users']['khach-co-hoa-mai']['logins'] == 1 and 'gv-co-linh' not in j['users']
assert len(j['registrations']) == 1 and j['registrations'][0]['name'] == 'Cô Thử Đăng Ký' and j['registrations'][0]['id'].startswith('dk-')
for bad in ('1993', '0865291703', '0912345678', 'abcdef', '127.0.0.1', 'scrypt', 'Làm báo bài'): assert bad not in raw, bad
pr.terminate(); pr.wait(); pr = up(); login('Anh Lê Đạt'); time.sleep(1.2)
assert json.load(open(P))['users']['gv-anh-le-dat']['logins'] == 3 and json.load(open(P))['users']['gv-anh-le-dat']['actions']['tao_giao_an'] == 2
print('Máy chủ: đăng nhập/điền báo bài/tạo giáo án/chat/đăng ký được ghi theo ngày VN; không lưu mật khẩu/SĐT/IP/nội dung; khởi động lại vẫn cộng dồn: OK')

# 3) không làm chậm request: ghi nhận chỉ là thêm vào bộ đệm
U3 = usage.Usage(tempfile.mkdtemp(), None, interval=999); t0 = time.perf_counter()
for i in range(20000): U3.action('k', 'n', 'chat')
assert (time.perf_counter() - t0) / 20000 < 5e-5 and len(U3._buf) == 20000
print('Ghi nhận ~%.1f µs/lần (không chạm ổ đĩa): OK' % ((time.perf_counter() - t0) / 20000 * 1e6))

# 4) báo cáo cho Quản Gia (gh giả, chỉ GET)
REP = '/workspace/tools/usage_report.py'
if os.path.exists(REP):
    G = tempfile.mkdtemp(); fake = os.path.join(G, 'gh'); calls = os.path.join(G, 'calls')
    open(os.path.join(G, 'usage.json'), 'w').write(open(P, encoding='utf-8').read())
    acc = {'dk-1': dict(id='dk-1', name='Cô Chờ Duyệt', phone='0900000001', pw='scrypt$x', school='TH A', status='pending', created='2026-10-16T08:00:00'),
           j['registrations'][0]['id']: dict(id=j['registrations'][0]['id'], name='Cô Thử Đăng Ký', phone='0912345678', pw='scrypt$y', school='TH Đông Hồ', status='pending', created=f'{day}T10:00:00'),
           'dk-3': dict(id='dk-3', name='Thầy Đã Duyệt', phone='0900000003', pw='z', status='approved', created='2026-10-01T08:00:00')}
    json.dump(acc, open(os.path.join(G, 'acc.json'), 'w'))
    open(fake, 'w').write(f"""#!/bin/sh
echo "$@" >> {calls}
case "$*" in *"-X GET"*"contents/data/usage/{day}.json"*) cat {G}/usage.json;; *"-X GET"*"contents/data/accounts.json"*) cat {G}/acc.json;; *) echo '{{"message":"Not Found"}}'; exit 1;; esac
"""); os.chmod(fake, 0o755)
    e2 = dict(os.environ, GH_BIN=fake)
    out = subprocess.run([sys.executable, REP], env=e2, capture_output=True, text=True).stdout
    print(out)
    assert '2 người dùng' in out and 'Anh Lê Đạt: 3 lần đăng nhập' in out and '1 điền báo bài, 2 tạo giáo án' in out and '2 tin nhắn trợ lý' in out
    assert 'Cô Hoa Mai: 1 lần đăng nhập' in out and 'Đang chờ anh duyệt: 2' in out and 'Cô Thử Đăng Ký (TH Đông Hồ)' in out and 'mới hôm nay' in out
    assert 'Thầy Đã Duyệt' not in out and '0912345678' not in out and 'scrypt' not in out
    assert all(' -X GET ' in ' ' + l + ' ' for l in open(calls).read().splitlines())
    out = subprocess.run([sys.executable, REP, '--date', '2026-01-01'], env=e2, capture_output=True, text=True).stdout
    assert out == 'KHONG_CO_NGUOI_DUNG\n', repr(out)
    open(os.path.join(G, 'e.json'), 'w').write('{"date":"x","users":{},"registrations":[]}')
    assert subprocess.run([sys.executable, REP, '--file', os.path.join(G, 'e.json')], capture_output=True, text=True).stdout == 'KHONG_CO_NGUOI_DUNG\n'
    assert subprocess.run([sys.executable, REP, '--file', os.path.join(G, 'none.json')], capture_output=True, text=True).stdout == 'KHONG_CO_NGUOI_DUNG\n'
    print('Báo cáo Quản Gia (gh giả, chỉ GET; không có dữ liệu -> KHONG_CO_NGUOI_DUNG): OK')
print('TẤT CẢ OK')
