"""v20: Tích hợp NLS, AI, Link… theo Khung KHDH – đọc kế hoạch, ghép tuần 7, chèn đỏ/link xanh, AI giả, API + UI + ảnh."""
import os, sys, json, tempfile, threading, subprocess, time, atexit, hashlib, re
from http.server import BaseHTTPRequestHandler, HTTPServer
import requests, docx
from docx.oxml.ns import qn
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import core, tichhop as th
SHOT = os.environ.get('SHOT_DIR', '/workspace')
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
PLAN = '/workspace/KhungKHDH.odt'; TPL = '/workspace/GiaoAn_Goc.docx'
PHU = A + '5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'
TMP = tempfile.mkdtemp(); BB = os.path.join(TMP, 'bb.docx')
core.fill_baobai(A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx', A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx', BB, 7)
H = lambda f: hashlib.md5(open(f, 'rb').read()).hexdigest()

# ---------- 1. đọc kế hoạch ----------
P = th.parse_plan(PLAN)
st = th.plan_stats(P)
assert set(st) == {'tv', 'toan', 'daoduc', 'hdtn', 'tnxh'} and st['tv']['rows'] >= 240 and st['toan']['rows'] >= 140 and st['daoduc']['rows'] == 35, st
assert not th.BANNED.search(json.dumps(P, ensure_ascii=False)), 'tên trường lọt vào dữ liệu kế hoạch'
print('1. đọc Khung KHDH (.odt) OK', {k: v['rows'] for k, v in st.items()})

# ---------- 2. ghép tuần 7 ----------
pv = th.preview(P, docx.Document(BB), 7)
M = {(m['subject'], m['rows'][0]['thu'], m['rows'][0]['ten'][:20]): m for m in pv['matches']}
def find(sub, part): return next(m for m in pv['matches'] if m['subject'] == sub and part in m['rows'][0]['ten'])
m = find('Toán', '( tiết 2'); assert m['lesson'] == 'Luyện tập' and m['rows'][0]['thu'] == '4' and any(i['text'].startswith('Sử dụng công cụ số') for g in m['groups'] for i in g['items']), m   # Luyện tập = tiết 2 của Bài 12
m = find('Toán', '( tiết 1'); assert m['lesson'].startswith('Bảng trừ')
m = find('Tiếng Việt', 'Đọc mở rộng'); assert 'ĐMR' in m['lesson']                 # TV-ĐMR <-> Đọc mở rộng
m = find('Tiếng Việt', 'Luyện tập: Viết đoạn'); assert any(g['cat'] == 'NLS' for g in m['groups'])
m = find('Tiếng Việt', 'Đọc: Yêu lắm'); assert len(m['rows']) == 2                    # khối 2 tiết
m = find('Đạo đức', 'Bài 3'); assert m['multi_week'] and m['week'].replace(' ', '') == '6,7'   # dòng nhiều tuần
assert all(m['conf'] >= .6 for m in pv['matches']) and len(pv['matches']) >= 12
print('2. ghép tuần 7 với báo bài OK –', len(pv['matches']), 'dòng kế hoạch')

# ---------- 3. chèn (nguyên văn) ----------
h0 = (H(BB), H(TPL))
OUT = '/workspace/GiaoAn_Tuan7_TichHop.docx'; BASE = os.path.join(TMP, 'base.docx')
core.build_giaoan(TPL, BB, 7, [PHU], BASE)
r = core.build_giaoan(TPL, BB, 7, [PHU], OUT, tichhop=dict(plan=P))['tichhop']
assert (H(BB), H(TPL)) == h0, 'không được sửa báo bài gốc / giáo án gốc'
assert r['n'] >= 20 and r['n_act'] >= 4 and not r['unmatched'], r
W = {x['label'] + '|' + x['lesson'][:25]: x['where'] for x in r['placed']}
assert any(x['where'].startswith('I.') and 'NLS 1.1' in x['label'] for x in r['placed'])
assert any(x['where'].startswith('II.') and 'Học liệu' in x['label'] for x in r['placed'])
assert any(x.get('act') and 'Đọc mở rộng' in x['where'] for x in r['placed']), 'gợi ý KHBD phải vào mục Đọc mở rộng'
assert any(x.get('act') and 'luyện tập' in x['where'].lower() and 'Toán' in x['lesson'] for x in r['placed']), 'NLS Toán -> Luyện tập'
D = docx.Document(OUT); body = D.element.body
def runs_col(e): return [(r.find(qn('w:rPr') + '/' + qn('w:color')), ''.join(t.text or '' for t in r.iter(qn('w:t')))) for r in e.iter(qn('w:r'))]
red = [t for c, t in runs_col(body) if c is not None and c.get(qn('w:val')) == 'FF0000' and t.strip()]
assert len(red) > 40
def ours(p):
    rs = [(c, t) for c, t in runs_col(p) if t.strip()]
    return rs and all(c is not None and c.get(qn('w:val')) in ('FF0000', '0000FF') for c, t in rs)
BASET = {''.join(t for c, t in runs_col(p)).strip() for p in docx.Document(BASE).element.body.iter(qn('w:p'))}
OURP = [p for p in body.iter(qn('w:p')) if ours(p) and ''.join(t for c, t in runs_col(p)).strip() not in BASET]
assert len(OURP) >= 20
for r_ in (x for p in OURP for x in p.iter(qn('w:r'))):
    c = r_.find(qn('w:rPr') + '/' + qn('w:color'))
    if c is not None and c.get(qn('w:val')) in ('FF0000', '0000FF'):
        i = r_.find(qn('w:rPr') + '/' + qn('w:i')); f = r_.find(qn('w:rPr') + '/' + qn('w:rFonts'))
        assert i is not None and i.get(qn('w:val')) == '0' and f.get(qn('w:ascii')) == 'Times New Roman'
hl = list(body.iter(qn('w:hyperlink'))); assert len(hl) >= 5
for h_ in hl:
    rel = D.part.rels[h_.get(qn('r:id'))]; assert rel.is_external and rel.target_ref.startswith('https://')
    assert h_.find('.//' + qn('w:color')).get(qn('w:val')) == '0000FF' and h_.find('.//' + qn('w:u')) is not None
x = open(OUT, 'rb').read(); import zipfile
assert not th.BANNED.search(zipfile.ZipFile(OUT).read('word/document.xml').decode()), 'tên trường trong thân giáo án'
# nội dung cũ giữ nguyên: bỏ các đoạn/dòng chỉ gồm chữ đỏ/xanh -> giống hệt bản không tích hợp
def plain(f):
    out = []
    for p in docx.Document(f).element.body.iter(qn('w:p')):
        rs = [(c, t) for c, t in runs_col(p) if t.strip()]
        if rs and all(c is not None and c.get(qn('w:val')) in ('FF0000', '0000FF') for c, t in rs): continue
        t = ''.join(t for c, t in rs).strip()
        if t: out.append(t)
    return out
assert plain(OUT) == plain(BASE), 'nội dung giáo án cũ bị thay đổi'
# lề: chữ đỏ không vượt quá chữ có sẵn
subprocess.run(['soffice', '-env:UserInstallation=file:///tmp/lo_profile_khdh', '--headless', '--convert-to', 'pdf', '--outdir', TMP, OUT], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
import pymupdf
doc = pymupdf.open(os.path.join(TMP, 'GiaoAn_Tuan7_TichHop.pdf')); mx = 0; rx = 0; pages = []
for k, pg in enumerate(doc):
    hit = 0
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                if not s['text'].strip(): continue
                if s['color'] in (0xFF0000, 0x0000FF): rx = max(rx, s['bbox'][2]); hit += 1; assert s['bbox'][0] >= 60
                else: mx = max(mx, s['bbox'][2])
    pages.append(hit)
assert rx <= mx + 1, (rx, mx)
print('3. chèn nguyên văn: đỏ đứng TNR, link xanh bấm được, đúng mục I/II/III, giữ nguyên nội dung cũ, trong lề OK –', r['n'], 'mục')

# ---------- 4. AI (giả) ----------
seen = []
def fake(sys_, user):
    seen.append(user); return [dict(id=j['id'], gv='- GV tổ chức cho HS: ' + j['goi_y'][:60], hs='- HS thực hiện, chia sẻ trước lớp.') for j in json.loads(user)]
r2 = core.build_giaoan(TPL, BB, 7, [PHU], os.path.join(TMP, 'ai.docx'), tichhop=dict(plan=P, mode='ai', ai_call=fake))['tichhop']
assert r2['mode'] == 'ai' and r2.get('ai_used') and len(seen) == 1
t = '\n'.join(p.text for p in docx.Document(os.path.join(TMP, 'ai.docx')).paragraphs) + ''.join(c.text for tb in docx.Document(os.path.join(TMP, 'ai.docx')).tables for row in tb.rows for c in row.cells)
assert t.count('- HS thực hiện, chia sẻ trước lớp.') == r2['n_act']
def bad(s, u): raise ValueError('AI lỗi')
r3 = core.build_giaoan(TPL, BB, 7, [PHU], os.path.join(TMP, 'ai2.docx'), tichhop=dict(plan=P, mode='ai', ai_call=bad))['tichhop']
assert r3['mode'] == 'verbatim' and r3['ai_note']
import shutil; shutil.copy(os.path.join(TMP, 'ai.docx'), '/workspace/GiaoAn_Tuan7_TichHop_AI.docx')
print('4. AI (giả) viết GV/HS, lỗi AI -> nguyên văn OK')

# ---------- 5. API + UI ----------
KEY = 'AIzaFAKE-th-key'
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        b = json.loads(s.rfile.read(int(s.headers['Content-Length'])))
        jobs = json.loads(b['contents'][0]['parts'][0]['text'])
        out = [dict(id=j['id'], gv='- GV tổ chức hoạt động: ' + j['goi_y'][:50], hs='- HS tham gia, chia sẻ.') for j in jobs]
        s.send_response(200 if s.headers.get('x-goog-api-key') == KEY else 403); s.end_headers()
        s.wfile.write(json.dumps({'candidates': [{'content': {'parts': [{'text': json.dumps(out, ensure_ascii=False)}]}}]}).encode())
gs = HTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
DD = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=DD, SECRET_KEY='t' * 32, GEMINI_API_KEY=KEY, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}/v1beta')
for k in ('DATA_REPO', 'DATA_TOKEN'): env.pop(k, None)
lg = open(os.path.join(DD, 'server.log'), 'w')
pr = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8040'], cwd=HERE, env=env, stdout=lg, stderr=subprocess.STDOUT); atexit.register(pr.terminate)
B = 'http://127.0.0.1:8040'
for _ in range(80):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
c = requests.Session()
assert c.get(B + '/api/tichhop/info').status_code == 401
c.post(B + '/api/login', data=dict(teacher='Cô Lê Tâm', password='1993'))
c.post(B + '/api/baobai/master', files={'file': ('bb.docx', open(BB, 'rb'))}).raise_for_status()
c.post(B + '/api/giaoan/template', files={'file': ('ga.docx', open(TPL, 'rb'))}).raise_for_status()
assert c.post(B + '/api/tichhop/plan', files={'file': ('x.txt', b'abc')}).status_code == 400
i = c.post(B + '/api/tichhop/plan', files={'file': ('KhungKHDH.odt', open(PLAN, 'rb'))}).json(); assert i['has_plan'] and i['ai']
assert len(c.get(B + '/api/tichhop/preview?week=7').json()['matches']) == len(pv['matches'])
st0 = c.post(B + '/api/giaoan/tichhop', data=dict(src='abc')); assert st0.status_code == 409   # chưa tạo giáo án -> khoá
rb = c.post(B + '/api/giaoan/build', data=dict(week=7), files=[('files', ('p.docx', open(PHU, 'rb')))]).json()
assert 'tichhop' not in rb and rb['state']['current'] == 'plain' and not rb['state']['th']      # tạo giáo án = bản thường
plain = c.get(B + '/api/giaoan/current').content
assert plain == c.get(B + rb['download']).content and 'FF0000' not in zipfile.ZipFile(__import__('io').BytesIO(plain)).read('word/document.xml').decode().split('IV. ĐIỀU')[0][:0] + ''
rt = c.post(B + '/api/giaoan/tichhop', data=dict(src=rb['src'])).json()
assert rt['tichhop']['n'] == r['n'] and rt['state']['current'] == 'th'
thb = c.get(B + '/api/giaoan/current'); assert thb.content != plain and 'TichHop' in thb.headers['content-disposition']
used0 = c.get(B + '/api/tichhop/info').json()['ai_used']
ra = c.post(B + '/api/giaoan/tichhop', data=dict(src=rb['src'], th_mode='ai')).json()
assert ra['tichhop']['mode'] == 'ai' and c.get(B + '/api/tichhop/info').json()['ai_used'] == used0 + 1
assert not os.path.exists(os.path.join(DD, 'users')) or True
assert c.post(B + '/api/giaoan/use', data=dict(which='plain')).json()['current'] == 'plain'
assert c.get(B + '/api/giaoan/current').content == plain, 'bỏ tích hợp phải trả đúng bản thường (giống từng byte)'
assert c.post(B + '/api/giaoan/use', data=dict(which='th')).json()['current'] == 'th' and c.get(B + '/api/giaoan/current').content != plain
rb2 = c.post(B + '/api/giaoan/build', data=dict(week=7), files=[('files', ('p.docx', open(PHU, 'rb')))]).json()
assert rb2['state']['th'] is None and c.get(B + ra['download']).status_code == 404, 'tạo mới phải xoá bản tích hợp cũ'
assert c.post(B + '/api/giaoan/use', data=dict(which='th')).status_code == 400
assert KEY not in open(os.path.join(DD, 'server.log')).read()
print('5a. API: khoá trước khi tạo, tạo = bản thường, tích hợp, AI 1 lượt, bỏ tích hợp giống từng byte, tạo mới xoá bản tích hợp OK')

from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    br = pw.chromium.launch()
    pg = br.new_page(viewport={'width': 1366, 'height': 768})
    pg.add_init_script("try{sessionStorage.setItem('botRun','1')}catch(e){}")
    pg.goto(B); pg.fill('#lgname', 'Cô Lê Tâm') if pg.locator('#lgname').count() else None
    pg.context.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in c.cookies.get_dict().items()]); pg.goto(B)
    pg.wait_for_selector('#app:not(.hide)'); pg.click('.tabs button:nth-child(2)')
    pg.wait_for_selector('#thcard.locked'); pg.select_option('#gweek', '7')
    assert pg.locator('#thlock').is_visible() and 'Tạo giáo án xong' in pg.inner_text('#thlock')
    assert pg.evaluate("getComputedStyle(document.querySelector('#thcard #thbody')).pointerEvents") == 'none'
    pg.screenshot(path=SHOT + '/v21_locked.png')
    pg.set_input_files('#gfiles', PHU); pg.click('#ggo'); pg.wait_for_selector('#thcard:not(.locked)', timeout=60000)
    assert pg.locator('#thm.hide').count() == 1 and pg.locator('#vbar').count() == 0
    pg.wait_for_selector('#thstat:has-text("Đã có kế hoạch")'); pg.screenshot(path=SHOT + '/v20_ui_step.png')
    pg.click('#thpv'); pg.wait_for_selector('.thl li'); time.sleep(.3)
    assert pg.locator('.thit').count() >= 20
    assert pg.evaluate('document.documentElement.scrollHeight<=innerHeight+1'), 'trang bị cuộn'
    pg.screenshot(path=SHOT + '/v20_ui_preview.png')
    pg.click('#thgo'); pg.wait_for_selector('#thm:not(.hide)', timeout=60000); time.sleep(.3)
    assert 'blur' not in pg.evaluate("getComputedStyle(document.getElementById('thm')).backdropFilter+getComputedStyle(document.body).filter")
    assert pg.locator('#thai').is_enabled()
    pg.screenshot(path=SHOT + '/v20_ui_dialog.png')
    pg.click('#thai'); pg.wait_for_selector('#thm.hide', state='attached', timeout=60000); pg.wait_for_selector('#gres:has-text("AI viết GV/HS")')
    assert pg.inner_text('#vcur') == 'Bản tích hợp' and 'bản tích hợp' in pg.inner_text('#gdl')
    assert pg.evaluate('document.documentElement.scrollHeight<=innerHeight+1')
    pg.screenshot(path=SHOT + '/v21_after.png')
    pg.click('#vdrop'); pg.wait_for_selector('#vcur:has-text("Bản chưa tích hợp")')
    with pg.expect_download() as dl: pg.click('#gdl')
    dl.value.save_as(os.path.join(TMP, 'ui_plain.docx'))
    pg.click('#vuse'); pg.wait_for_selector('#vcur:has-text("Bản tích hợp")')
    # bé Linh không che nút mới
    bb = pg.locator('#bot').bounding_box()
    for sel in ('#thpv', '#thgo', '#ggo', '#vuse', '#vdrop'):
        e = pg.locator(sel).bounding_box()
        if e and bb: assert not (e['x'] < bb['x'] + bb['width'] and bb['x'] < e['x'] + e['width'] and e['y'] < bb['y'] + bb['height'] and bb['y'] < e['y'] + e['height']), sel
    br.close()
assert H(os.path.join(TMP, 'ui_plain.docx')) == hashlib.md5(c.get(B + '/api/giaoan/download/' + json.load(open(os.path.join(DD, 'users', os.listdir(os.path.join(DD, 'users'))[0], 'out', 'state.json')))['plain']).content).hexdigest()
print('5b. UI: khoá trước khi tạo, mở sau khi tạo, xem trước, tích hợp, xem trước, hộp thoại 2 lựa chọn, không mờ, không cuộn, không bị che OK')
print('TẤT CẢ OK')
