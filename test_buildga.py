"""v26 – 'Tạo giáo án' không được lỗi: RAM không phình (Render 512 MB), dữ liệu cũ (trước bản nháp) được chuyển sang bản nháp,
tuần gửi kiểu '7'/'Tuần 7', thiếu tuần/thiếu mẫu/tệp hỏng -> báo tiếng Việt rõ ràng, giao diện hiện lỗi thay vì treo 'Đang ghép…'."""
import os, sys, time, json, shutil, tempfile, subprocess, atexit, requests
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
G0 = A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx'; BBP = A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx'
TPL = '/workspace/GiaoAn_Goc.docx'; GAP = A + '5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'
DD = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=DD, SECRET_KEY='t' * 32)
for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_API_KEY'): env.pop(k, None)
B = 'http://127.0.0.1:8085'; LOG = os.path.join(DD, 'server.log')
pr = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8085'], cwd=HERE, env=env, stdout=open(LOG, 'w'), stderr=subprocess.STDOUT); atexit.register(pr.terminate)
for _ in range(100):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
rss = lambda: int(open(f'/proc/{pr.pid}/status').read().split('VmRSS:')[1].split()[0]) // 1024
def sess(t, pw='1993', name=''):
    s = requests.Session(); s.post(B + '/api/login', data=dict(teacher=t, password=pw, name=name)).raise_for_status(); s.get(B + '/api/status'); return s
fl = lambda n, p: (n, (os.path.basename(p), open(p, 'rb')))
def ins(r): return sum(len(t['inserted']) for t in r['tiet'])

# 1) luồng đầy đủ: báo bài gốc + điền tuần 7 (xoá+điền) + giáo án gốc + bản nháp -> tạo
s = sess('Anh Lê Đạt')
assert s.post(B + '/api/baobai/master', files=[fl('file', G0)]).ok and s.post(B + '/api/giaoan/template', files=[fl('file', TPL)]).ok
assert s.post(B + '/api/baobai/fill', data=dict(week=7), files=[fl('file', BBP)]).ok
r = s.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.ok, r.text; assert ins(r.json()) == 0   # chưa có bản nháp: vẫn tạo được (bước 1, 2)
t = time.time(); r = s.post(B + '/api/giaoan/build', data=dict(week=7), files=[fl('files', GAP)]); dt = time.time() - t
assert r.ok, r.text; base = ins(r.json()); assert base >= 15, base
assert dt < 8, f'ghép quá chậm: {dt:.1f}s'
for wk in ('7', ' 7 ', 'Tuần 7'):
    r = s.post(B + '/api/giaoan/build', data=dict(week=wk)); assert r.ok and ins(r.json()) == base, (wk, r.text[:200])
for wk, msg in (('', 'chọn tuần'), ('30', 'không có tuần 30')):
    r = s.post(B + '/api/giaoan/build', data=dict(week=wk)); assert r.status_code == 400 and msg in r.json()['detail'], r.text
r = s.post(B + '/api/giaoan/build'); assert r.status_code == 400 and 'chọn tuần' in r.json()['detail'], r.text
print('Luồng đầy đủ + tuần dạng chữ/thiếu/không có: OK (%.1fs/lần ghép)' % dt)

# 2) RAM không phình theo số lần ghép (trước đây ~+40 MB mỗi lần -> Render 512 MB tắt máy giữa chừng)
m0 = rss()
for i in range(8):
    assert s.post(B + '/api/giaoan/build', data=dict(week=7)).ok and s.get(B + '/api/draft', params={'week': 7}).ok
m1 = rss(); assert m1 - m0 < 80, f'RAM tăng {m1 - m0} MB sau 8 lần ghép ({m0}->{m1})'
print(f'RAM ổn định: {m0} -> {m1} MB sau 8 lần ghép + đọc bản nháp')

# 3) dữ liệu cũ (trước khi có bản nháp): out/state.json + out/src_<id>/000.docx, không có thư mục draft
s2 = sess('Cô Lê Tâm'); ud = os.path.join(DD, 'users', 'gv-co-le-tam'); assert os.path.isdir(ud), os.listdir(os.path.join(DD, 'users'))
open(os.path.join(ud, 'baobai_goc.docx'), 'wb').write(s.get(B + '/api/baobai/download').content); shutil.copy(TPL, os.path.join(ud, 'giaoan_goc.docx'))
os.makedirs(os.path.join(ud, 'out', 'src_abc123'), exist_ok=True); shutil.copy(GAP, os.path.join(ud, 'out', 'src_abc123', '000.docx'))
json.dump(dict(src='abc123', week=7, plain='zzz', th=None, current='plain'), open(os.path.join(ud, 'out', 'state.json'), 'w'))
shutil.rmtree(os.path.join(ud, 'draft'), ignore_errors=True)
r = s2.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.ok, r.text; assert ins(r.json()) == base, ins(r.json())
d = s2.get(B + '/api/draft', params={'week': 7}).json(); assert len(d['files']) == 1 and 'trước đây' in d['files'][0]['name'], d['files']
r = s2.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.ok and len(s2.get(B + '/api/draft', params={'week': 7}).json()['files']) == 1   # không nhân đôi
s2.delete(B + '/api/draft', params={'week': 7}); r = s2.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.ok and ins(r.json()) == 0   # đã xoá thì không tự chép lại
print('Dữ liệu cũ (chưa có bản nháp) -> tự chuyển vào bản nháp 1 lần: OK')

# 4) bản nháp hỏng / thiếu thông tin (đồng bộ dở) -> vẫn ghép, không lỗi
dd = os.path.join(ud, 'draft', 'w8'); os.makedirs(dd, exist_ok=True)
open(os.path.join(dd, 'bad1.docx'), 'wb').write(b'not a docx'); shutil.copy(GAP, os.path.join(dd, 'ok1.docx'))
json.dump(dict(files=[dict(id='bad1', name='hỏng.docx', lessons=[]), dict(id='ok1'), dict(id='mất'), 'rác']), open(os.path.join(dd, 'draft.json'), 'w'))
assert s2.get(B + '/api/draft', params={'week': 8}).ok
r = s2.post(B + '/api/giaoan/build', data=dict(week=8)); assert r.ok, r.text
# 5) giáo án gốc hỏng -> 500 nhưng có thông báo tiếng Việt (JSON), có traceback trong log
open(os.path.join(ud, 'giaoan_goc.docx'), 'wb').write(b'PK\x03\x04 hong')
r = s2.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.status_code in (400, 500) and 'giáo án' in r.json()['detail'], r.text
os.remove(os.path.join(ud, 'giaoan_goc.docx')); r = s2.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.status_code == 400 and 'Chưa có giáo án gốc' in r.json()['detail']
s3 = sess('Cô Lê Thủy'); r = s3.post(B + '/api/giaoan/build', data=dict(week=7)); assert r.status_code == 400 and 'Chưa có báo bài gốc' in r.json()['detail']
assert 'Traceback' in open(LOG).read()
print('Bản nháp hỏng, mẫu hỏng/thiếu, chưa có báo bài: báo rõ ràng, không treo: OK')

# 6) giao diện thật: chọn tuần, thêm giáo án phụ, bấm Tạo giáo án -> có nút tải; lỗi máy chủ -> hiện lỗi tiếng Việt
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    br = p.chromium.launch()
    for vw, vh in ((1366, 768), (390, 844)):
        ctx = br.new_context(viewport={'width': vw, 'height': vh}); pg = ctx.new_page(); errs = []
        pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.add_init_script("try{sessionStorage.setItem('botRun','1')}catch(e){}")
        ctx.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in s.cookies.get_dict().items()]); pg.goto(B)
        pg.wait_for_selector('#app:not(.hide)'); pg.click('#t2'); pg.wait_for_function("document.getElementById('gweek').options.length>2")
        pg.select_option('#gweek', '7'); pg.set_input_files('#gfiles', GAP); pg.wait_for_selector('#dflist span')
        pg.click('#ggo'); pg.wait_for_selector('#gdl', timeout=60000)
        assert pg.locator('#gres .pill.ok').count() >= 12
        for status, body, ct, want in ((502, '<html>Bad Gateway</html>', 'text/html', 'Máy chủ đang khởi động lại'),
                                       (500, '{"detail":"Có lỗi khi tạo giáo án. Thầy cô kiểm tra lại"}', 'application/json', 'Có lỗi khi tạo giáo án'),
                                       (422, '{"detail":[{"loc":["body","week"]}]}', 'application/json', 'Thông tin gửi lên chưa đúng')):
            pg.route('**/api/giaoan/build', (lambda a: lambda rt: rt.fulfill(status=a[0], body=a[1], content_type=a[2]))((status, body, ct)))
            pg.click('#ggo'); pg.wait_for_selector('#gerr'); txt = pg.inner_text('#gerr')
            assert want in txt and 'Đang ghép' not in pg.inner_text('#gres'), txt
            pg.unroute('**/api/giaoan/build')
        pg.route('**/api/giaoan/build', lambda rt: rt.abort()); pg.click('#ggo'); pg.wait_for_selector('#gerr'); assert 'Mất kết nối' in pg.inner_text('#gerr')
        assert not errs, errs; ctx.close()
    br.close()
print('Giao diện: tạo giáo án OK; lỗi 502/500/422/mất mạng -> hiện thông báo tiếng Việt, không treo: OK')
print('TẤT CẢ OK')
