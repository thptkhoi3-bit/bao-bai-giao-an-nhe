"""v26 – Nhận xét học sinh (TT27) + Kho lưu theo tuần: API, AI giả lập (gom lô, không gửi tên, giới hạn lượt), câu mẫu khác nhau,
xuất xlsx/docx, kho theo tuần (thay bản cũ, giữ 10 tuần, giới hạn dung lượng, cô lập người dùng), giao diện + ảnh chụp."""
import os, sys, io, json, time, shutil, tempfile, threading, subprocess, atexit, datetime, requests
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import kho, nhanxet, openpyxl, docx
SHOT = os.environ.get('SHOT_DIR', '/workspace')
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
G0 = A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx'; BBP = A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx'
TPL = '/workspace/GiaoAn_Goc.docx'; GAP = A + '5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'

# ---------- 1) đơn vị: kho ----------
R = tempfile.mkdtemp(); src = os.path.join(R, 's.docx'); open(src, 'wb').write(b'x' * 1000)
t0 = datetime.datetime(2026, 10, 1, 8, 0, tzinfo=kho.VN)
for w in range(1, 13): kho.put(R, 'gv-a', w, 'giaoan', src, f'GA{w}.docx', now=t0 + datetime.timedelta(hours=w))
L = kho.listing(R, 'gv-a'); assert [g['label'] for g in L['groups']] == [f'Tuần {w}' for w in range(12, 2, -1)], L['groups']
assert not os.path.exists(os.path.join(R, 'kho', 'gv-a', 'w01')) and not os.path.exists(os.path.join(R, 'kho', 'gv-a', 'w02'))
a = kho.put(R, 'gv-a', 12, 'giaoan', src, 'GA12b.docx', now=t0 + datetime.timedelta(hours=20))
g12 = [g for g in kho.listing(R, 'gv-a')['groups'] if g['week'] == 'w12'][0]; assert len(g12['items']) == 1 and g12['items'][0]['name'] == 'GA12b.docx'
assert len(os.listdir(os.path.join(R, 'kho', 'gv-a', 'w12'))) == 1
kho.put(R, 'gv-a', 'khac', 'de', src, 'De.docx', now=t0 + datetime.timedelta(hours=21))
assert len(kho.listing(R, 'gv-a')['groups']) == 10 and kho.listing(R, 'gv-a')['groups'][-1]['label'] == 'Khác'
os.environ['KHO_MAX_MB'] = str(3500 / 1024 / 1024)
kho.put(R, 'gv-a', 12, 'baobai', src, 'BB.docx', now=t0 + datetime.timedelta(hours=22))
L = kho.listing(R, 'gv-a'); n = sum(len(g['items']) for g in L['groups']); assert n == 3 and L['total'] <= 3500, L
assert any(x['name'] == 'BB.docx' for g in L['groups'] for x in g['items'])
del os.environ['KHO_MAX_MB']
assert kho.find(R, 'gv-a', '../../x') is None and kho.find(R, 'gv-b', a['id']) is None
for bad in ('../gv-a', 'GV A', ''):
    try: kho.udir(R, bad); raise AssertionError(bad)
    except ValueError: pass
print('Kho (đơn vị): thay bản cũ cùng loại/tuần, giữ 10 tuần gần nhất, giới hạn dung lượng bỏ tệp cũ nhất, mã tệp an toàn: OK')

# ---------- 2) đơn vị: câu mẫu ----------
HO = ['Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng']; DEM = ['Văn', 'Thị', 'Minh', 'Gia', 'Thu', 'Đức', 'Bảo', 'Khánh']; TEN = ['An', 'Bình', 'Châu', 'Dũng', 'Hà', 'Huy', 'Linh', 'Nam', 'Ngọc', 'Vy']
names = [f'{HO[i % 5]} {DEM[i % 8]} {TEN[i // 4 % 10]}' for i in range(40)]; assert len(set(names)) == 40
st = [dict(name=n, muc={'tv': 'TCH'[i % 3], 'toan': 'HTC'[i % 3]}, pn={'pc': 'TDC'[i % 3], 'nl': 'DTC'[i % 3]}, note='hay giúp bạn' if i == 1 else '') for i, n in enumerate(names)]
c = nhanxet.template_comments(st, ['tv', 'toan'], 'Cuối năm học')
for k in ('tv', 'toan', 'pc', 'nl'): assert len({x[k] for x in c}) == 40, k
assert all('cần' in c[i]['tv'] for i in range(40) if st[i]['muc']['tv'] == 'C') and not any('cần' in c[i]['tv'] for i in range(40) if st[i]['muc']['tv'] == 'T')
assert 'Em hay giúp bạn.' in c[1]['pc'] and c == nhanxet.template_comments(st, ['tv', 'toan'], 'Cuối năm học')
assert nhanxet.names_from_rows([['STT', 'Họ và tên đệm', 'Tên'], ['1', 'Nguyễn Văn', 'An'], ['2', 'Lê Thị', 'Hà']]) == ['Nguyễn Văn An', 'Lê Thị Hà']
print('Câu mẫu: 40 em không trùng câu, đúng mức (CHT có “cần…”, HTT không), có ghi chú GV, ổn định: OK')

# ---------- 3) máy chủ + Gemini giả ----------
CALLS = []
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        body = json.loads(s.rfile.read(int(s.headers['Content-Length']))); CALLS.append(body)
        user = body['contents'][0]['parts'][0]['text']; rows = json.loads(user.split('Danh sách:\n', 1)[1])
        keys = list(body['generationConfig']['responseSchema']['items']['properties'])
        out = [{'i': r['i'], **{k: f'AI: em số {r["i"]} – {k} – {list(r["muc"].values())[0]}' for k in keys if k != 'i'}} for r in rows]
        b = json.dumps({'candidates': [{'content': {'parts': [{'text': json.dumps(out, ensure_ascii=False)}]}}]}).encode()
        s.send_response(200); s.send_header('Content-Type', 'application/json'); s.send_header('Content-Length', str(len(b))); s.end_headers(); s.wfile.write(b)
gs = ThreadingHTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
KEY = 'test-key-' + 'x' * 20
def server(port, extra):
    dd = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=dd, SECRET_KEY='t' * 32, USAGE_FLUSH='0.3', **extra)
    for k in ('DATA_REPO', 'DATA_TOKEN'): env.pop(k, None)
    if 'GEMINI_API_KEY' not in extra: env.pop('GEMINI_API_KEY', None)
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(port)], cwd=HERE, env=env, stdout=open(os.path.join(dd, 'log'), 'w'), stderr=subprocess.STDOUT); atexit.register(p.terminate)
    for _ in range(100):
        try: requests.get(f'http://127.0.0.1:{port}', timeout=1); return dd, p
        except Exception: time.sleep(.2)
def login(B, t, pw='1993', name=''):
    s = requests.Session(); s.post(B + '/api/login', data=dict(teacher=t, password=pw, name=name)).raise_for_status(); return s
J = lambda s, B, path, d: s.post(B + path, data=json.dumps(d), headers={'Content-Type': 'application/json'})

DD, P1 = server(8091, {}); B = 'http://127.0.0.1:8091'
s = login(B, 'Anh Lê Đạt'); T = tempfile.mkdtemp()
wb = openpyxl.Workbook(); ws = wb.active; ws.append(['DANH SÁCH LỚP 2A']); ws.append(['STT', 'Họ và tên', 'Ngày sinh']); [ws.append([i + 1, n, '1/1/2019']) for i, n in enumerate(names[:25])]; wb.save(T + '/l.xlsx')
d = docx.Document(); t = d.add_table(rows=1, cols=3); [setattr(t.rows[0].cells[i], 'text', h) for i, h in enumerate(['STT', 'Họ và tên đệm', 'Tên'])]
for i, (a_, b_) in enumerate([('Nguyễn Văn', 'An'), ('Trần Thị', 'Bình')]):
    r = t.add_row().cells; r[0].text, r[1].text, r[2].text = str(i + 1), a_, b_
d.save(T + '/l.docx'); open(T + '/l.csv', 'w', encoding='utf-8-sig').write('STT;Họ tên học sinh;Giới tính\n1;Lê Minh Châu;Nữ\n2;Phạm Gia Huy;Nam\n')
for fn, want in (('l.xlsx', 25), ('l.docx', ['Nguyễn Văn An', 'Trần Thị Bình']), ('l.csv', ['Lê Minh Châu', 'Phạm Gia Huy'])):
    r = s.post(B + '/api/nhanxet/parse', files={'file': (fn, open(T + '/' + fn, 'rb'))}); assert r.ok, r.text
    assert (len(r.json()['names']) == want) if isinstance(want, int) else r.json()['names'] == want, (fn, r.json())
r = s.post(B + '/api/nhanxet/parse', data={'text': '1. Nguyễn Văn An\n2) Trần Thị Bình\n\nLê Minh Châu'}); assert r.json()['names'] == ['Nguyễn Văn An', 'Trần Thị Bình', 'Lê Minh Châu']
assert s.post(B + '/api/nhanxet/parse', data={'text': '123\n---'}).status_code == 400
assert s.post(B + '/api/nhanxet/parse', files={'file': ('x.pdf', b'%PDF')}).status_code == 400
info = s.get(B + '/api/nhanxet').json(); assert info['state'] is None and not info['info']['ai'] and 'tienganh' not in info['info']['default_subjects']
STATE = dict(ki='Cuối học kì I', lop='2A', subjects=['tv', 'toan', 'daoduc', 'tienganh'], students=[dict(name=n, muc={'tv': 'TCH'[i % 3]}, pn={'pc': 'TDC'[i % 3]}, note='') for i, n in enumerate(names[:25])])
r = J(s, B, '/api/nhanxet/generate', {'state': STATE}); assert r.ok, r.text; g = r.json()
assert g['mode'] == 'mau' and g['ai_calls'] == 0 and 'chưa bật AI' in g['message']
S = g['state']['students']; assert len({x['nx']['tv'] for x in S}) == 25 and all(set(x['nx']) == {'tv', 'toan', 'daoduc', 'tienganh', 'pc', 'nl'} for x in S)
S[0]['nx']['tv'] = 'Câu thầy cô tự sửa cho em đầu tiên.'
r = J(s, B, '/api/nhanxet/export', {'state': g['state'], 'fmt': 'xlsx'}); assert r.ok, r.text; ex = r.json()
x = s.get(B + ex['download']); assert x.ok and x.headers['content-type'].startswith('application/vnd.openxmlformats-officedocument.spreadsheetml')
wb = openpyxl.load_workbook(io.BytesIO(x.content)); ws = wb.active
hdr = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]; assert hdr[:3] == ['STT', 'Họ và tên', 'Tiếng Việt'] and 'Phẩm chất' in hdr and 'Năng lực' in hdr and 'Tiếng Anh' in hdr, hdr
assert ws.max_row == 3 + 25 and ws.cell(4, 2).value == names[0] and ws.cell(4, 4).value == 'Câu thầy cô tự sửa cho em đầu tiên.' and ws.cell(4, 3).value == 'Hoàn thành tốt'
assert ws.cell(5, 3).value == 'Chưa hoàn thành' and 'CUỐI HỌC KÌ I' in ws.cell(1, 1).value
r = J(s, B, '/api/nhanxet/export', {'state': g['state'], 'fmt': 'docx'}); assert r.ok; y = s.get(B + r.json()['download']); D = docx.Document(io.BytesIO(y.content))
txt = '\n'.join(c.text for row in D.tables[0].rows for c in row.cells); assert names[24] in txt and 'Câu thầy cô tự sửa' in txt and 'Tiếng Anh (' in txt
assert J(s, B, '/api/nhanxet/export', {'state': g['state'], 'fmt': 'pdf'}).status_code == 400
assert s.get(B + '/api/nhanxet').json()['state']['students'][0]['nx']['tv'] == 'Câu thầy cô tự sửa cho em đầu tiên.'
print('Nhận xét (không AI): đọc xlsx/docx/csv/dán, câu mẫu khác nhau, sửa tay, xuất xlsx/docx đúng cột & nội dung, lưu lại: OK')

assert s.post(B + '/api/baobai/master', files=[('file', ('m.docx', open(G0, 'rb')))]).ok and s.post(B + '/api/giaoan/template', files=[('file', ('t.docx', open(TPL, 'rb')))]).ok
assert s.post(B + '/api/baobai/fill', data=dict(week=7), files=[('file', ('p.docx', open(BBP, 'rb')))]).ok
assert s.post(B + '/api/giaoan/build', data=dict(week=7), files=[('files', ('g.docx', open(GAP, 'rb')))]).ok
id1 = [x for g_ in s.get(B + '/api/kho').json()['groups'] for x in g_['items'] if x['type'] == 'giaoan'][0]['id']
assert s.post(B + '/api/giaoan/build', data=dict(week=7)).ok
r = s.post(B + '/api/de/extract', data=dict(mon='toan', lop='2', loai='Giữa học kì I', nam='2026 - 2027', xa='Hoà Bình', truong='TH Kim Đồng'), files=[('files', ('toan.docx', open(os.path.join(HERE, 'de_templates', 'toan.docx'), 'rb')))]); assert r.ok, r.text[:300]
K = s.get(B + '/api/kho').json(); w7 = [g_ for g_ in K['groups'] if g_['week'] == 'w07'][0]
types = {x['type'] for g_ in K['groups'] for x in g_['items']}; assert types >= {'baobai', 'giaoan', 'de', 'nhanxet_xlsx', 'nhanxet_docx'}, types
assert [x['type'] for x in w7['items']].count('giaoan') == 1 and all(x['id'] != id1 for x in w7['items']), w7
assert {x['name'] for x in w7['items']} >= {'BaoBai_Tuan7.docx', 'GiaoAn_Tuan7.docx'}
assert any(x['name'] == 'NhanXet_Lop2A_CuoiKiI.xlsx' for g_ in K['groups'] for x in g_['items']), K
assert s.get(B + '/api/kho/file/' + id1).status_code == 404
bb = [x for x in w7['items'] if x['type'] == 'baobai'][0]; f = s.get(B + '/api/kho/file/' + bb['id']); assert f.ok and docx.Document(io.BytesIO(f.content)).tables
s2 = login(B, 'Cô Linh'); assert s2.get(B + '/api/kho').json()['groups'] == []
assert s2.get(B + '/api/kho/file/' + bb['id']).status_code == 404 and s2.delete(B + '/api/kho/file/' + bb['id']).status_code == 404
assert requests.get(B + '/api/kho/file/' + bb['id']).status_code == 401 and requests.get(B + '/api/kho').status_code == 401
assert s.get(B + '/api/kho/file/..%2F..%2Faccounts.json').status_code == 404
assert s.delete(B + '/api/kho/file/' + bb['id']).ok and s.get(B + '/api/kho/file/' + bb['id']).status_code == 404
assert not [fn for fn in os.listdir(os.path.join(DD, 'kho', 'gv-anh-le-dat', 'w07')) if fn.startswith('baobai')]
time.sleep(1); day = datetime.datetime.now(kho.VN).date().isoformat()
U = json.load(open(os.path.join(DD, 'usage', day + '.json')))['users']['gv-anh-le-dat']['actions']; assert U.get('nhan_xet') == 1 and U.get('tao_giao_an') == 2, U
rep = subprocess.run([sys.executable, '/workspace/tools/usage_report.py', '--file', os.path.join(DD, 'usage', day + '.json'), '--date', day], capture_output=True, text=True).stdout
assert '1 nhận xét học sinh' in rep, rep
print('Kho (máy chủ): báo bài/giáo án/đề/nhận xét tự lưu theo tuần, tạo lại giữ bản mới, chỉ chủ tài khoản xem/tải/xoá, nhật ký “nhận xét học sinh”: OK')

DD2, P2 = server(8092, dict(GEMINI_API_KEY=KEY, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}/v1beta', AI_DAILY_LIMIT='2'))
B2 = 'http://127.0.0.1:8092'; s3 = login(B2, 'Cô Lê Tâm')
st40 = dict(STATE, students=[dict(name=n, muc={}, pn={}, note='chữ đẹp' if i == 3 else '') for i, n in enumerate(names)])
r = J(s3, B2, '/api/nhanxet/generate', {'state': dict(st40, students=st40['students'][:20])}); g = r.json()
assert r.ok and g['mode'] == 'ai' and g['ai_calls'] == 2 and len(CALLS) == 2, (g.get('mode'), g.get('ai_calls'), len(CALLS))
assert g['state']['students'][16]['nx']['tv'].startswith('AI: em số 16')
sent = json.dumps(CALLS, ensure_ascii=False); assert not any(n in sent for n in names) and 'chữ đẹp' in sent and KEY not in open(os.path.join(DD2, 'log')).read()
r = J(s3, B2, '/api/nhanxet/generate', {'state': st40}); g = r.json(); assert g['mode'] == 'mau' and g['ai_calls'] == 0 and 'hết lượt AI' in g['message'] and len(CALLS) == 2
assert len({x['nx']['toan'] for x in g['state']['students']}) == 40
print('Nhận xét (AI giả): 15 em/lượt, không gửi tên học sinh, tôn trọng AI_DAILY_LIMIT, hết lượt tự dùng câu mẫu: OK')

# ---------- 4) giao diện ----------
from playwright.sync_api import sync_playwright
M_JS = "()=>({sh:document.documentElement.scrollHeight,ih:innerHeight,sw:document.documentElement.scrollWidth,iw:innerWidth})"
COVER = """()=>{const b=document.querySelector('#bot .bfig').getBoundingClientRect();const hit=[];
 const vis=e=>{let r=e.getBoundingClientRect(),x0=r.left,y0=r.top,x1=r.right,y1=r.bottom;for(let p=e.parentElement;p;p=p.parentElement){const cs=getComputedStyle(p);if(/(auto|scroll|hidden)/.test(cs.overflow+cs.overflowX+cs.overflowY)){const q=p.getBoundingClientRect();x0=Math.max(x0,q.left);y0=Math.max(y0,q.top);x1=Math.min(x1,q.right);y1=Math.min(y1,q.bottom)}}return [x0,y0,x1,y1]};
 document.querySelectorAll('#p4 button,#p4 input,#p4 select,#p4 textarea,#p4 a,#p5 button,#p5 a,#t4,#t5').forEach(e=>{if(!e.offsetParent)return;const [x0,y0,x1,y1]=vis(e);if(x1-x0<1||y1-y0<1)return;
  if(x0<b.right-1&&x1>b.left+1&&y0<b.bottom-1&&y1>b.top+1)hit.push(e.id||e.className||e.tagName)});return hit}"""
BLUR = "()=>[...document.querySelectorAll('*')].filter(e=>{const c=getComputedStyle(e);return /blur/.test(c.filter+c.backdropFilter)}).length"
with sync_playwright() as pw:
    br = pw.chromium.launch(); ctx = br.new_context(viewport={'width': 1920, 'height': 1080}, accept_downloads=True); pg = ctx.new_page(); errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    sT = login(B, 'Cô Lê Thủy'); ctx.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in sT.cookies.get_dict().items()])
    pg.goto(B); pg.wait_for_selector('#app:not(.hide)'); pg.click('#t4'); pg.wait_for_selector('#nxki option', state='attached')
    pg.fill('#nxpaste', '\n'.join(f'{i + 1}. {n}' for i, n in enumerate(['Nguyễn Văn An', 'Trần Thị Bình', 'Lê Minh Châu', 'Phạm Gia Huy', 'Hoàng Thu Hà', 'Vũ Đức Minh', 'Đỗ Khánh Linh', 'Bùi Quang Nam', 'Ngô Bảo Ngọc', 'Đặng Tuấn Phong', 'Lý Thảo Vy', 'Phan Nhật Quang', 'Trịnh Mai Anh', 'Đinh Hải Đăng'])))
    pg.click('#nxload'); pg.wait_for_selector('.nxt tbody tr'); assert pg.locator('.nxt tbody tr').count() == 14
    sels = pg.locator('.nxt tbody tr').nth(0).locator('select'); sels.nth(0).select_option('T'); sels.nth(1).select_option('T')
    pg.locator('.nxt tbody tr').nth(2).locator('select').nth(0).select_option('C'); pg.locator('.nxt tbody tr').nth(1).locator('input').nth(1).fill('hay giúp bạn'); pg.keyboard.press('Tab')
    pg.select_option('#nxki', 'Cuối học kì I'); pg.wait_for_timeout(900)
    pg.click('#nxgo'); pg.wait_for_selector('.nxt textarea'); pg.wait_for_timeout(300)
    tv = [pg.locator(f'.nxt textarea[data-i="{i}"][data-k="tv"]').input_value() for i in range(14)]; assert len(set(tv)) == 14 and 'cần' in tv[2], tv
    ta = pg.locator('.nxt textarea[data-i="0"][data-k="toan"]'); ta.fill('Em tính nhẩm rất nhanh. Cô khen em!'); pg.keyboard.press('Tab'); pg.wait_for_timeout(900)
    pg.evaluate("document.querySelector('#p4>div').scrollTop=0;document.getElementById('nxtw').scrollLeft=0"); pg.wait_for_timeout(150)
    pg.screenshot(path=f'{SHOT}/v26_nhanxet.png')
    with pg.expect_download() as dl: pg.click('#nxx')
    assert dl.value.suggested_filename.endswith('.xlsx'), dl.value.suggested_filename; xp = os.path.join(T, 'ui.xlsx'); dl.value.save_as(xp)
    ws = openpyxl.load_workbook(xp).active; assert ws.cell(4, 2).value == 'Nguyễn Văn An' and 'Em tính nhẩm rất nhanh' in str([ws.cell(4, c).value for c in range(1, ws.max_column + 1)])
    for W, H in [(1280, 800), (1366, 768), (1440, 900), (1920, 1080)]:
        pg.set_viewport_size({'width': W, 'height': H}); pg.wait_for_timeout(200)
        for tb in ('#t4', '#t5'):
            pg.click(tb); pg.wait_for_timeout(400); m = pg.evaluate(M_JS); assert m['sh'] <= m['ih'] and m['sw'] <= m['iw'], (W, H, tb, m)
            hb = pg.evaluate("document.querySelector('header.v9 .hbar').getBoundingClientRect().height"); assert hb < 70, (W, tb, 'thanh trên bị xuống 2 dòng', hb)
            for side in ('false', 'true'):
                pg.evaluate(f"document.getElementById('bot').classList.toggle('right',{side})"); assert not pg.evaluate(COVER), (W, tb, side, pg.evaluate(COVER))
            assert pg.evaluate(BLUR) == 0
    pg.set_viewport_size({'width': 1920, 'height': 1080})
    sT.post(B + '/api/baobai/master', files=[('file', ('m.docx', open(G0, 'rb')))]); sT.post(B + '/api/giaoan/template', files=[('file', ('t.docx', open(TPL, 'rb')))])
    for w in (6, 7): sT.post(B + '/api/baobai/fill', data=dict(week=w), files=[('file', ('p.docx', open(BBP, 'rb')))])
    sT.post(B + '/api/giaoan/build', data=dict(week=7), files=[('files', ('g.docx', open(GAP, 'rb')))])
    pg.click('#t1'); pg.click('#t5'); pg.wait_for_function("document.querySelectorAll('.khg').length>=2"); pg.wait_for_timeout(300)
    pg.screenshot(path=f'{SHOT}/v26_kho.png')
    n0 = pg.locator('.khi').count(); pg.once('dialog', lambda d_: d_.accept()); pg.locator('.khi button').first.click(); pg.wait_for_function(f"document.querySelectorAll('.khi').length=={n0 - 1}")
    with pg.expect_download() as dl: pg.locator('.khi a').first.click()
    assert os.path.getsize(dl.value.path()) > 1000
    pg.evaluate("document.documentElement.setAttribute('data-theme','dark')"); pg.wait_for_timeout(100); assert pg.evaluate(BLUR) == 0
    ctx2 = br.new_context(viewport={'width': 390, 'height': 844}, is_mobile=True); m2 = ctx2.new_page(); ctx2.add_cookies(ctx.cookies())
    m2.goto(B); m2.wait_for_selector('#app:not(.hide)')
    for tb in ('#t4', '#t5'):
        m2.click(tb); m2.wait_for_timeout(500); m = m2.evaluate(M_JS); assert m['sw'] <= m['iw'], (tb, m)
    assert not errs, errs; br.close()
print('Giao diện: nhập danh sách, chọn mức, ghi chú, viết nhận xét, sửa, xuất Excel; Kho theo tuần, tải/xoá; 1 màn hình 4 cỡ, không bị trợ lý che, không mờ, điện thoại không tràn ngang: OK')
print('TẤT CẢ OK')
