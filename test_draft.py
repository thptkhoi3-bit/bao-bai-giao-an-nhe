"""v22 – bản nháp giáo án phụ theo tuần: tích luỹ nhiều tệp, bài trùng -> bản mới thắng, bỏ 1 tệp, xoá nháp, bền qua khởi động lại, tách tuần."""
import os, sys, time, json, tempfile, subprocess, atexit, copy, requests, io
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import core, docx, storage
from docx.oxml.ns import qn
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
SHOT = os.environ.get('SHOT_DIR', '/workspace')
TPL = '/workspace/GiaoAn_Goc.docx'; PHU = A + '5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'
TMP = tempfile.mkdtemp(); BB = os.path.join(TMP, 'bb.docx')
core.fill_baobai(A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx', A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx', BB, 7)
tx = lambda d: [''.join(x.text or '' for x in e.iter(qn('w:t'))) for e in docx.Document(d).element.body]

def split(subjects, out, mark=None):
    S = docx.Document(PHU); b, ls = core.split_lessons(S); keep = set()
    prev = -1
    for l in ls:  # giữ cả phần đầu (tên môn) trước mỗi bài
        if l['subject'] in subjects: keep |= set(range(prev + 1, l['end'] + 1))
        prev = l['end']
    for i, e in enumerate(b):
        if i not in keep and e.tag != qn('w:sectPr'): e.getparent().remove(e)
    if mark: S.add_paragraph(mark)
    S.save(out); return out
TV = split({'tv'}, os.path.join(TMP, 'TV.docx')); TOAN = split({'toan'}, os.path.join(TMP, 'Toan.docx'))
TV2 = split({'tv'}, os.path.join(TMP, 'TV_sua.docx'), mark='ghi chú bản sửa')
NTV = sum(1 for l in core.split_lessons(docx.Document(PHU))[1] if l['subject'] == 'tv')
NTO = sum(1 for l in core.split_lessons(docx.Document(PHU))[1] if l['subject'] == 'toan')

DD = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=DD, SECRET_KEY='t' * 32)
for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_API_KEY'): env.pop(k, None)
B = 'http://127.0.0.1:8050'
def up():
    pr = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8050'], cwd=HERE, env=env, stdout=open(os.path.join(DD, 'server.log'), 'a'), stderr=subprocess.STDOUT); atexit.register(pr.terminate)
    for _ in range(100):
        try: requests.get(B, timeout=1); return pr
        except Exception: time.sleep(.2)
pr = up(); c = requests.Session()
def login():
    c.post(B + '/api/login', data=dict(teacher='Cô Lê Tâm', password='1993')).raise_for_status()
login()
c.post(B + '/api/baobai/master', files={'file': ('bb.docx', open(BB, 'rb'))}).raise_for_status()
c.post(B + '/api/giaoan/template', files={'file': ('ga.docx', open(TPL, 'rb'))}).raise_for_status()
add = lambda w, *fs: c.post(B + '/api/draft/add', data=dict(week=w), files=[('files', (os.path.basename(f), open(f, 'rb'))) for f in fs]).json()
build = lambda w=7: c.post(B + '/api/giaoan/build', data=dict(week=w)).json()
import re
def filled(r): return {core.subj(re.match(r'TIẾT\s+\d+(?:\s*[+,]\s*\d+)*\s+(.*)$', re.sub(r'\s+', ' ', t['tiet'])).group(1)) for t in r['tiet'] if t['inserted']}
def get(w=7): return c.get(B + f'/api/draft?week={w}').json()

# 1. chỉ Tiếng Việt -> chỉ TIẾT TV được điền
r = add(7, TV); assert r['reports'][0]['added'] and not r['reports'][0]['replaced'] and len(r['reports'][0]['added']) == NTV
d = r['draft']; assert len(d['files']) == 1 and d['subjects'] == ['Tiếng Việt'] and any(t['need'] and not t['filled'] for t in d['tiet'])
b1 = build(); assert filled(b1) == {'tv'}, filled(b1)
assert len(get()['files']) == 1, 'Tạo giáo án không được xoá bản nháp'
# 2. thêm Toán -> TV vẫn còn, Toán thêm vào
r = add(7, TOAN); assert len(r['reports'][0]['added']) == NTO and not r['reports'][0]['replaced']
d = r['draft']; assert len(d['files']) == 2 and set(d['subjects']) == {'Tiếng Việt', 'Toán'}
b2 = build(); assert filled(b2) == {'tv', 'toan'}, filled(b2)
emp2 = sum(1 for t in get()['tiet'] if t['need'] and not t['filled'])
# 3. tải lại TV (bản sửa) -> chỉ bài TV bị thay, Toán giữ nguyên
r = add(7, TV2); rp = r['reports'][0]; assert len(rp['replaced']) == NTV and not rp['added'], rp
fs = {f['name']: f for f in r['draft']['files']}
assert fs['TV.docx']['active'] == 0 and fs['TV.docx']['replaced'] == NTV and fs['Toan.docx']['active'] == NTO and fs['TV_sua.docx']['active'] == NTV
b3 = build(); assert filled(b3) == {'tv', 'toan'}
doc = c.get(B + '/api/giaoan/current').content; assert 'ghi chú bản sửa' not in str(tx(io.BytesIO(doc)))  # nội dung ngoài bài không bị chèn
import zipfile; assert sum(len(t['inserted']) for t in b3['tiet']) == sum(len(t['inserted']) for t in b2['tiet']), 'bài TV không được nhân đôi'
# 4. tuần khác tách biệt
add(8, TOAN); assert len(get(8)['files']) == 1 and len(get(7)['files']) == 3
# 5. bền qua khởi động lại
before = get(); pr.terminate(); pr.wait(); pr = up(); c.cookies.clear(); login()
after = get(); assert [f['id'] for f in after['files']] == [f['id'] for f in before['files']] and after['lessons'] == before['lessons']
uk = [x for x in os.listdir(os.path.join(DD, 'users'))]; dp = [os.path.relpath(os.path.join(r_, f), DD) for k in uk for r_, _, fl in os.walk(os.path.join(DD, 'users', k, 'draft')) for f in fl]
assert dp and all(storage.wanted(x.replace(os.sep, '/')) for x in dp), 'bản nháp phải được đồng bộ'
# 6. bỏ 1 tệp: bỏ bản TV sửa -> bản TV cũ dùng lại
fid = next(f['id'] for f in after['files'] if f['name'] == 'TV_sua.docx')
d = c.delete(B + f'/api/draft/file?week=7&id={fid}').json(); fs = {f['name']: f for f in d['files']}
assert 'TV_sua.docx' not in fs and fs['TV.docx']['active'] == NTV and sum(1 for t in d['tiet'] if t['need'] and not t['filled']) == emp2
assert c.delete(B + '/api/draft/file?week=7&id=deadbeef00').status_code == 404
# 7. tệp gộp tất cả môn: TV/Toán thay bằng bản gộp, kết quả = giáo án chuẩn
r = add(7, PHU); rp = r['reports'][0]; assert len(rp['replaced']) == NTV + NTO and rp['added']
b7 = build(); assert tx(io.BytesIO(c.get(B + '/api/giaoan/current').content)) == tx('/workspace/GiaoAn_Tuan7_HoanChinh_v2.docx'), 'gộp nhiều tệp phải ra đúng giáo án chuẩn'
# 8. xoá nháp tuần 7 -> tuần 8 còn nguyên; tạo giáo án vẫn chạy (không có bài)
d = c.delete(B + '/api/draft?week=7').json(); assert not d['files'] and len(get(8)['files']) == 1
assert not build()['tiet'] or not any(t['inserted'] for t in build()['tiet'])
# 9. API cũ (gửi tệp kèm build) vẫn chạy và cộng vào nháp
rb = c.post(B + '/api/giaoan/build', data=dict(week=7), files=[('files', ('p.docx', open(PHU, 'rb')))]).json()
assert rb['draft_added'] and len(get()['files']) == 1
c.delete(B + '/api/draft?week=7')
print('1-9. API: TV -> chỉ TV; + Toán giữ TV; tải lại thay đúng bài; tách tuần; bền qua khởi động lại + đồng bộ; bỏ 1 tệp; tệp gộp; xoá nháp OK')

from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    br = pw.chromium.launch()
    for vw, vh in ((1366, 768), (1280, 800)):
        pg = br.new_page(viewport={'width': vw, 'height': vh}); errs = []; pg.on('pageerror', lambda e: errs.append(str(e))); pg.on('console', lambda m: m.type == 'error' and print('CONSOLE', m.text))
        pg.add_init_script("try{sessionStorage.setItem('botRun','1')}catch(e){}")
        pg.goto(B); pg.context.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in c.cookies.get_dict().items()]); pg.goto(B)
        pg.wait_for_selector('#app:not(.hide)'); pg.click('.tabs button:nth-child(2)'); pg.select_option('#gweek', '7'); pg.wait_for_timeout(300)
        pg.wait_for_selector('#dstat:has-text("chưa có tệp")')
        pg.set_input_files('#gfiles', TV); pg.wait_for_selector('#dflist span'); pg.wait_for_timeout(200)
        pg.set_input_files('#gfiles', TOAN); pg.wait_for_function("document.querySelectorAll('#dflist span').length==2")
        assert 'Tiếng Việt' in pg.inner_text('#dsubs') and 'Toán' in pg.inner_text('#dsubs') and int(pg.inner_text('#demp')) > 0
        assert 'Đã thêm' in pg.inner_text('#gres') and pg.locator('#dclear').is_visible()
        assert pg.evaluate('document.documentElement.scrollHeight<=innerHeight+1'), 'trang bị cuộn'
        assert pg.evaluate("[...document.querySelectorAll('*')].every(e=>{const s=getComputedStyle(e);return !/blur/.test(s.filter+s.backdropFilter)})"), 'không được làm mờ'
        bot = pg.evaluate("(()=>{const b=document.getElementById('bot');if(!b)return null;const r=b.getBoundingClientRect();return [r.left,r.top,r.right,r.bottom]})()")
        for sel in ('#ggo', '#dview', '#dclear', '.drm'):
            r_ = pg.evaluate(f"(()=>{{const r=document.querySelector('{sel}').getBoundingClientRect();return [r.left,r.top,r.right,r.bottom]}})()")
            assert not bot or bot[2] <= r_[0] or bot[0] >= r_[2] or bot[3] <= r_[1] or bot[1] >= r_[3], f'mascot che {sel}'
        if vw == 1366: pg.screenshot(path=SHOT + '/v22_draft.png')
        pg.click('#ggo'); pg.wait_for_selector('#gres .dlbox', timeout=90000)
        assert '2 tệp' in pg.inner_text('#dstat'), 'tạo giáo án không được xoá nháp'
        pg.reload(); pg.wait_for_selector('#app:not(.hide)'); pg.click('.tabs button:nth-child(2)'); pg.select_option('#gweek', '7')
        pg.wait_for_selector('#dstat:has-text("2 tệp")')
        pg.select_option('#gweek', '8'); pg.wait_for_selector('#dstat:has-text("tuần 8")'); assert '1 tệp' in pg.inner_text('#dstat')
        pg.select_option('#gweek', '7'); pg.wait_for_selector('#dstat:has-text("tuần 7")'); pg.click('#dview')
        pg.click('#dflist .drm'); pg.wait_for_function("document.querySelectorAll('#dflist span').length==1")
        pg.once('dialog', lambda dl: dl.dismiss()); pg.click('#dclear'); pg.wait_for_timeout(300); assert pg.locator('#dflist span').count() == 1
        pg.once('dialog', lambda dl: dl.accept()); pg.click('#dclear'); pg.wait_for_selector('#dstat:has-text("chưa có tệp")')
        assert '1 tệp' in (pg.select_option('#gweek', '8') and pg.wait_for_selector('#dstat:has-text("tuần 8")') and pg.inner_text('#dstat'))
        assert not errs, errs; pg.close()
    br.close()
print('10. UI: danh sách nháp sau TV rồi Toán, không cuộn, không mờ, không bị che, giữ sau Tạo/tải lại, tách tuần, × bỏ tệp, xác nhận xoá OK')
print('TẤT CẢ OK')
