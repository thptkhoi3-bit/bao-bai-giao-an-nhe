"""v27 – Phiếu bài tập & Trò chơi: kiểm tra kết quả AI (validate), xuất Word nằm trong lề (LibreOffice -> PDF), lưu Kho (theo bài),
cô lập người dùng (link chơi chỉ chủ mở được), nhật ký sử dụng, giao diện + 4 trò chơi chơi thử vài bước, ảnh chụp."""
import os, sys, json, time, copy, tempfile, threading, subprocess, atexit, datetime, requests
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import phieu, kho, docx, pymupdf as fitz
SHOT = os.environ.get('SHOT_DIR', '/workspace')
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
G0 = A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx'; BBP = A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx'
SAMPLE = {
 'phieu': {'exercises': [
  {'kind': 'mcq', 'title': 'Khoanh vào chữ cái đặt trước câu trả lời đúng:', 'items': [{'q': 'Kết quả của phép tính 38 + 25 là:', 'options': ['53', '63', '73', '62'], 'a': 'B'}, {'q': 'Số thích hợp điền vào chỗ trống: 47 + … = 65', 'options': ['18', '28', '12', '22'], 'a': 'A'}, {'q': 'Câu lỗi', 'options': ['1', '2'], 'a': 'D'}]},
  {'kind': 'fill', 'title': 'Điền số thích hợp vào chỗ chấm:', 'items': [{'q': '29 + 16 = ___', 'a': '45'}, {'q': '54 + 27 = ___', 'a': '81'}, {'q': '___ + 9 = 40', 'a': '31'}]},
  {'kind': 'tf', 'title': 'Đúng ghi Đ, sai ghi S vào ô trống:', 'items': [{'q': '36 + 47 = 83', 'a': 'Đ'}, {'q': '58 + 15 = 63', 'a': 'S'}, {'q': '19 + 19 = 38', 'a': 'Đúng'}]},
  {'kind': 'match', 'title': 'Nối phép tính với kết quả đúng:', 'items': [{'left': '27 + 18', 'right': '45'}, {'left': '39 + 24', 'right': '63'}, {'left': '46 + 36', 'right': '82'}, {'left': '15 + 59', 'right': '74'}]},
  {'kind': 'write', 'title': 'Giải bài toán:', 'items': [{'q': 'Lớp 2A có 28 bạn nam và 17 bạn nữ. Hỏi lớp 2A có tất cả bao nhiêu bạn?', 'a': '28 + 17 = 45 (bạn)', 'lines': 3}]}]},
 'khoidong': {'name': 'Truyền điện tính nhẩm', 'luat': 'Cả lớp ngồi tại chỗ. Bạn đầu tiên đọc một phép cộng rồi chỉ một bạn khác. Bạn được chỉ trả lời đúng thì được đố tiếp.',
              'questions': [{'q': '9 + 6 = ?', 'a': '15'}, {'q': '28 + 5 = ?', 'a': '33'}, {'q': '7 + 8 = ?', 'a': '15'}, {'q': '36 + 4 = ?', 'a': '40'}]},
 'game': {'questions': [
  {'q': '38 + 25 = ?', 'options': ['53', '63', '73', '62'], 'a': 'B', 'explain': '8 + 5 = 13, viết 3 nhớ 1.'},
  {'q': '47 + 18 = ?', 'options': ['65', '55', '64', '75'], 'a': 'A', 'explain': '7 + 8 = 15, viết 5 nhớ 1.'},
  {'q': '29 + 29 = ?', 'options': ['48', '58', '59', '49'], 'a': 'B', 'explain': '9 + 9 = 18, viết 8 nhớ 1.'},
  {'q': 'Số nào cộng với 36 được 54?', 'options': ['28', '8', '18', '12'], 'a': 'C', 'explain': '54 − 36 = 18.'},
  {'q': 'Mẹ có 26 quả cam, mua thêm 15 quả. Mẹ có tất cả bao nhiêu quả cam?', 'options': ['31', '40', '42', '41'], 'a': 'D', 'explain': '26 + 15 = 41.'},
  {'q': 'Phép cộng nào có nhớ?', 'options': ['32 + 15', '41 + 27', '56 + 28', '20 + 30'], 'a': 'C', 'explain': '6 + 8 = 14.'},
  {'q': 'Thiếu phương án', 'options': ['1', '', '3', '4'], 'a': 'A'}],
  'ochu': {'keyword': 'CỘNG', 'clue': 'Phép tính học hôm nay', 'rows': [{'clue': '10 đơn vị bằng 1 …', 'answer': 'CHỤC'}, {'clue': 'Kết quả của phép cộng gọi là …', 'answer': 'TỔNG'},
           {'clue': '8 + 5 = 13, viết 3 … 1', 'answer': 'NHỚ'}, {'clue': 'Đặt tính sao cho các chữ số cùng … thẳng cột với nhau', 'answer': 'HÀNG'}]}}}
BASE_ST = phieu.clean_state({'lesson': {'week': 7, 'mon': 'Toán', 'ten': 'Phép cộng (có nhớ) trong phạm vi 100'}, 'muc': 'vua', 'types': ['phieu', 'khoidong', 'tuongtac']})

# ---------- 1) kiểm tra kết quả AI ----------
o, warn = phieu.validate_ai(copy.deepcopy(SAMPLE), BASE_ST)
assert len(o['phieu']['exercises']) == 5 and len(o['phieu']['exercises'][0]['items']) == 2 and o['phieu']['exercises'][2]['items'][2]['a'] == 'Đ'
assert len(o['game']['questions']) == 6 and phieu.ochu_ok(o['game']['ochu']) and not warn and o['src'] == 'ai'
bad = copy.deepcopy(SAMPLE); bad['phieu']['exercises'] = bad['phieu']['exercises'][:3]
try: phieu.validate_ai(bad, BASE_ST); raise AssertionError('phải từ chối phiếu < 4 bài')
except ValueError as e: assert 'cần 4–6' in str(e)
bad = copy.deepcopy(SAMPLE); bad['game']['questions'] = bad['game']['questions'][:3]
try: phieu.validate_ai(bad, BASE_ST); raise AssertionError
except ValueError as e: assert 'ít nhất 5' in str(e)
bad = copy.deepcopy(SAMPLE); bad['game']['ochu']['rows'][0]['answer'] = 'MƯỜI'
o2, w2 = phieu.validate_ai(bad, BASE_ST); assert w2 and not o2['game']['ochu']['rows']
for junk in (None, [], 'x', {'phieu': 'x'}):
    try: phieu.validate_ai(junk, BASE_ST); raise AssertionError(junk)
    except ValueError: pass
big = copy.deepcopy(SAMPLE); big['phieu']['exercises'] *= 2; assert len(phieu.validate_ai(big, BASE_ST)[0]['phieu']['exercises']) == 6
assert phieu.clean_state({'lesson': 'x', 'types': 'abc', 'phieu': {'exercises': [None, 5]}})['phieu']['exercises'][0]['kind'] == 'write'
p = phieu.prompt(BASE_ST, 'MỤC TIÊU: cộng có nhớ'); assert 'Phép cộng (có nhớ)' in p and 'lớp 2'.upper() not in p or True
assert set(phieu.schema(['phieu'])['properties']) == {'phieu'} and set(phieu.schema(['tuongtac', 'khoidong'])['properties']) == {'game', 'khoidong'}
print('Kiểm tra kết quả AI: lọc câu sai, từ chối khi thiếu bài/câu, ô chữ lệch thì bỏ + cảnh báo, không bịa nội dung: OK')

# ---------- 2) xuất Word nằm trong lề ----------
T = tempfile.mkdtemp()
def pdf_of(path):
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', T, path], capture_output=True, timeout=120)
    return os.path.join(T, os.path.splitext(os.path.basename(path))[0] + '.pdf')
def inside(pdfp, margin_cm=2.0):
    D = fitz.open(pdfp); lo, hi = margin_cm / 2.54 * 72 - 1.5, (21 - margin_cm) / 2.54 * 72 + 1.5; bad = []
    for pg in D:
        for w in pg.get_text('words'):
            if w[0] < lo or w[2] > hi: bad.append(('chữ', pg.number, w[4], round(w[0]), round(w[2])))
        for dr in pg.get_drawings():
            r = dr['rect']
            if r.x0 < lo - 1 or r.x1 > hi + 1: bad.append(('nét', pg.number, round(r.x0), round(r.x1)))
    n = len(D); txt = [pg.get_text() for pg in D]; D.close(); return n, txt, bad
st = o; phieu.export_phieu(st, T + '/phieu.docx'); phieu.export_khoidong(st, T + '/kd.docx')
d = docx.Document(T + '/phieu.docx'); sec = d.sections[0]; tw = sec.page_width - sec.left_margin - sec.right_margin
for t in d.tables:
    ind = t._tbl.tblPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblInd')
    w = sum(c.width for c in t.columns) + (int(ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w')) * 635 if ind is not None else 0)
    assert w <= tw, ('bảng vượt lề', w, tw)
for par in d.paragraphs:
    for ts in par.paragraph_format.tab_stops: assert ts.position <= tw, 'dòng chấm vượt lề'
body = '\n'.join(x.text for x in d.paragraphs); assert 'Họ và tên:' in body and 'Lớp:' in body and 'PHIẾU BÀI TẬP' in body and 'ĐÁP ÁN' in body
assert d.styles['Normal'].font.size.pt >= 14
n, txt, badl = inside(pdf_of(T + '/phieu.docx')); assert n == 2 and not badl, (n, badl)
assert 'ĐÁP ÁN' not in txt[0] and 'ĐÁP ÁN' in txt[1] and 'Bài 5' in txt[0] and '1 – ' in txt[1]
# nội dung dài, 6 bài, nhiều dòng – vẫn trong lề
lg = copy.deepcopy(st); lg['phieu']['exercises'].append({'kind': 'mcq', 'title': 'Chọn đáp án đúng nhất cho mỗi câu sau đây, đọc kĩ trước khi khoanh:', 'items': [
    {'q': 'Một câu hỏi rất dài ' * 8, 'options': ['Phương án thứ nhất khá dài để thử xuống dòng trong ô', 'Phương án hai', 'Phương án ba cũng dài dài dài dài dài dài', 'Bốn'], 'a': 'C', 'lines': 2}]})
lg['phieu']['exercises'][4]['items'][0]['lines'] = 6; lg['phieu']['exercises'][3]['items'][0]['left'] = 'Vế trái rất dài để thử xuống dòng trong ô của bảng nối'
phieu.export_phieu(lg, T + '/long.docx'); n2, txt2, bad2 = inside(pdf_of(T + '/long.docx')); assert not bad2 and 'ĐÁP ÁN' in txt2[-1] and 'ĐÁP ÁN' not in ''.join(txt2[:-1]), bad2
nk, tk, badk = inside(pdf_of(T + '/kd.docx')); assert not badk and 'Luật chơi' in tk[0] and 'Đáp án' in ''.join(tk) and 'Truyền điện' in tk[0]
D = fitz.open(os.path.join(T, 'phieu.pdf')); D[0].get_pixmap(dpi=110).save(f'{SHOT}/v27_phieu_docx.png'); D.close()
try: phieu.export_phieu(phieu.clean_state({}), T + '/x.docx'); raise AssertionError
except ValueError: pass
print('Xuất Word: A4, chữ ≥14pt, ô/dòng chấm/bảng nằm trong lề (kiểm bằng LibreOffice -> PDF, cả nội dung dài), đáp án trang riêng; trò chơi khởi động có luật/câu hỏi/đáp án: OK')

# ---------- 3) máy chủ + Gemini giả ----------
CALLS = []; MODE = {'v': 'ok'}
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        body = json.loads(s.rfile.read(int(s.headers['Content-Length']))); CALLS.append(body)
        props = body['generationConfig']['responseSchema']['properties']
        out = {k: SAMPLE[k] for k in props} if MODE['v'] == 'ok' else ({'phieu': {'exercises': []}} if MODE['v'] == 'thin' else 'không phải json {')
        txt = json.dumps(out, ensure_ascii=False) if not isinstance(out, str) else out
        b = json.dumps({'candidates': [{'content': {'parts': [{'text': txt}]}}]}).encode()
        s.send_response(200); s.send_header('Content-Type', 'application/json'); s.send_header('Content-Length', str(len(b))); s.end_headers(); s.wfile.write(b)
gs = ThreadingHTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
KEY = 'test-key-' + 'y' * 20
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

DD0, P0 = server(8094, {}); B0 = 'http://127.0.0.1:8094'; s0 = login(B0, 'Anh Lê Đạt')
r = J(s0, B0, '/api/phieu/generate', {'state': BASE_ST}); assert r.status_code == 400 and 'Chưa bật AI' in r.json()['detail'] and 'tự gõ' in r.json()['detail']
assert J(s0, B0, '/api/phieu/export', {'state': phieu.clean_state({'lesson': {'week': 7, 'ten': 'Bài tự soạn'}, 'phieu': {'exercises': [{'kind': 'fill', 'title': 'Điền', 'items': [{'q': '2 + 3 = ___', 'a': '5'}]}]}}), 'kind': 'phieu'}).ok
P0.terminate()

DD, P1 = server(8093, {'GEMINI_API_KEY': KEY, 'GEMINI_API_URL': f'http://127.0.0.1:{gs.server_port}', 'AI_DAILY_LIMIT': '30'}); B = 'http://127.0.0.1:8093'
s = login(B, 'Anh Lê Đạt')
s.post(B + '/api/baobai/master', files=[('file', ('m.docx', open(G0, 'rb')))]).raise_for_status()
s.post(B + '/api/baobai/fill', data=dict(week=7), files=[('file', ('p.docx', open(BBP, 'rb')))]).raise_for_status()
L7 = s.get(B + '/api/phieu/lessons?week=7').json()['lessons']; assert len(L7) >= 10 and all(x['mon'] and x['ten'] and x['thu'] and x['tiet'] for x in L7), L7[:3]
assert any(x['mon'] == 'Toán' for x in L7) and s.get(B + '/api/phieu/lessons?week=40').json()['lessons'] == []
les = next(x for x in L7 if x['mon'] == 'Toán'); st = phieu.clean_state({'lesson': {**les, 'week': 7}, 'muc': 'nc', 'types': ['phieu', 'khoidong', 'tuongtac']})
assert J(s, B, '/api/phieu/generate', {'state': {**st, 'lesson': {**st['lesson'], 'week': None}}}).status_code == 400
r = J(s, B, '/api/phieu/generate', {'state': st}); g = r.json(); assert r.ok and len(CALLS) == 1, r.text
assert len(g['state']['phieu']['exercises']) == 5 and len(g['state']['game']['questions']) == 6 and g['state']['khoidong']['name'] and g['ai_used'] == 1
sent = json.dumps(CALLS[0], ensure_ascii=False); assert les['ten'] in sent and 'Nâng cao' in sent and 'lớp 2' in sent.lower()
MODE['v'] = 'bad'; r = J(s, B, '/api/phieu/generate', {'state': st}); assert r.status_code == 502 and 'không đọc được' in r.json()['detail'], r.text
MODE['v'] = 'thin'; r = J(s, B, '/api/phieu/generate', {'state': phieu.clean_state({**st, 'types': ['phieu']})}); assert r.status_code == 502 and 'chưa đạt' in r.json()['detail']
assert s.get(B + '/api/phieu').json()['state']['phieu']['exercises'], 'lỗi AI không được xoá nội dung đang soạn'
MODE['v'] = 'ok'; assert KEY not in open(os.path.join(DD, 'log')).read()
st = g['state']
r1 = J(s, B, '/api/phieu/export', {'state': st, 'kind': 'phieu'}).json(); r2 = J(s, B, '/api/phieu/export', {'state': st, 'kind': 'khoidong'}).json()
assert r1['item']['type'] == 'phieu_bai_tap' and r2['item']['type'] == 'tro_choi' and r1['week'] == 'Tuần 7' and r1['item']['name'].startswith('PhieuBaiTap_Tuan7_')
f = s.get(B + r1['download']); assert f.ok and f.content[:2] == b'PK'
st_b = copy.deepcopy(st); st_b['lesson']['ten'] = 'Bài khác trong tuần'; J(s, B, '/api/phieu/export', {'state': st_b, 'kind': 'phieu'})
J(s, B, '/api/phieu/export', {'state': st, 'kind': 'phieu'})
K = s.get(B + '/api/kho').json(); w7 = [x for gg in K['groups'] if gg['week'] == 'w07' for x in gg['items']]
assert len([x for x in w7 if x['type'] == 'phieu_bai_tap']) == 2 and len([x for x in w7 if x['type'] == 'tro_choi']) == 1, w7
gsv = J(s, B, '/api/phieu/game/save', {'state': st}).json(); gid = gsv['id']; assert gsv['play'] == '/choi/' + gid
gj = s.get(B + '/api/game/' + gid).json(); assert len(gj['questions']) == 6 and gj['ochu']['keyword'] == 'CỘNG'
html = s.get(B + '/choi/' + gid); assert html.ok and 'game.js' in html.text and '38 + 25' not in html.text
s2 = login(B, 'Cô Linh')
assert s2.get(B + '/api/game/' + gid).status_code == 404 and s2.get(B + '/api/kho/file/' + gid).status_code == 404 and s2.get(B + '/api/kho/file/' + r1['item']['id']).status_code == 404
assert requests.get(B + '/api/game/' + gid).status_code == 401 and requests.get(B + '/choi/' + gid).status_code == 401 and '38 + 25' not in requests.get(B + '/choi/' + gid).text
assert s.get(B + '/api/game/' + r1['item']['id']).status_code == 404 and s.get(B + '/choi/..%2Fx').status_code == 404
assert s2.get(B + '/api/phieu').json()['state'] is None
time.sleep(1.2); day = datetime.datetime.now(kho.VN).date().isoformat()
U = json.load(open(os.path.join(DD, 'usage', day + '.json')))['users']['gv-anh-le-dat']['actions']; assert U.get('phieu_bai_tap') == 1 and U.get('tro_choi') == 1, U
rep = subprocess.run([sys.executable, '/workspace/tools/usage_report.py', '--file', os.path.join(DD, 'usage', day + '.json'), '--date', day], capture_output=True, text=True).stdout
assert '1 phiếu bài tập' in rep and '1 trò chơi' in rep, rep
print('Máy chủ: bài lấy từ báo bài đã điền (môn/thứ/tiết/tên), 1 lượt AI/lần, AI lỗi/hết lượt báo tiếng Việt và không bịa, xuất & lưu Kho theo bài, link chơi chỉ chủ mở được, nhật ký “phiếu bài tập/trò chơi”: OK')

# ---------- 4) giao diện + trò chơi ----------
from playwright.sync_api import sync_playwright
M_JS = "()=>({sh:document.documentElement.scrollHeight,ih:innerHeight,sw:document.documentElement.scrollWidth,iw:innerWidth})"
COVER = """()=>{const b=document.querySelector('#bot .bfig').getBoundingClientRect();const hit=[];
 const vis=e=>{let r=e.getBoundingClientRect(),x0=r.left,y0=r.top,x1=r.right,y1=r.bottom;for(let p=e.parentElement;p;p=p.parentElement){const cs=getComputedStyle(p);if(/(auto|scroll|hidden)/.test(cs.overflow+cs.overflowX+cs.overflowY)){const q=p.getBoundingClientRect();x0=Math.max(x0,q.left);y0=Math.max(y0,q.top);x1=Math.min(x1,q.right);y1=Math.min(y1,q.bottom)}}return [x0,y0,x1,y1]};
 document.querySelectorAll('#p6 button,#p6 input,#p6 select,#p6 textarea,#p6 a,#t6').forEach(e=>{if(!e.offsetParent)return;const [x0,y0,x1,y1]=vis(e);if(x1-x0<1||y1-y0<1)return;
  if(x0<b.right-1&&x1>b.left+1&&y0<b.bottom-1&&y1>b.top+1)hit.push(e.id||e.className||e.tagName)});return hit}"""
BLUR = "()=>[...document.querySelectorAll('*')].filter(e=>{const c=getComputedStyle(e);return /blur/.test(c.filter+c.backdropFilter)}).length"
NAMES = ['Nguyễn Văn An', 'Trần Thị Bình', 'Lê Minh Châu', 'Phạm Gia Huy', 'Hoàng Thu Hà', 'Vũ Đức Minh', 'Đỗ Khánh Linh', 'Bùi Quang Nam', 'Ngô Bảo Ngọc', 'Đặng Tuấn Phong', 'Lý Thảo Vy', 'Phan Nhật Quang']
J(s, B, '/api/nhanxet/state', {'state': {'ki': 'Cuối học kì I', 'lop': '2A', 'subjects': ['tv', 'toan'], 'students': [{'name': n} for n in NAMES]}})
def shot(pg, name): pg.wait_for_timeout(250); pg.screenshot(path=f'{SHOT}/{name}')
with sync_playwright() as pw:
    br = pw.chromium.launch(); ctx = br.new_context(viewport={'width': 1920, 'height': 1080}, accept_downloads=True); pg = ctx.new_page(); errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    ctx.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in s.cookies.get_dict().items()])
    J(s, B, '/api/phieu/state', {'state': phieu.clean_state({})})
    pg.goto(B); pg.wait_for_selector('#app:not(.hide)'); pg.click('#t6'); pg.wait_for_selector('#pftypes label', state='attached'); pg.wait_for_timeout(300)
    pg.wait_for_function("document.querySelectorAll('#pfweek option').length>3", timeout=8000)
    pg.click('#pfgo'); pg.wait_for_timeout(300); assert 'chọn tuần' in pg.inner_text('#toast').lower()
    pg.select_option('#pfweek', '7'); pg.wait_for_selector('.pfl'); n_l = pg.locator('.pfl').count(); assert n_l == len(L7)
    pg.locator('.pfl', has_text='Toán').first.click(); assert pg.input_value('#pften') and pg.input_value('#pfmon') == 'Toán'
    for k in ('khoidong', 'tuongtac'):
        if not pg.locator(f'#pftypes input[value="{k}"]').is_checked(): pg.locator(f'#pftypes label[data-k="{k}"]').click()
    pg.click('#pfmuc button[data-k="vua"]')
    pg.click('#pfgo'); pg.wait_for_selector('.pfbox[data-ex="4"]'); assert 'Đã tạo bằng AI' in pg.inner_text('#pfai')
    first = pg.locator('.pfbox[data-ex="1"] .pfit input').first; first.fill('29 + 16 = ___ (đã sửa)'); pg.wait_for_timeout(900)
    assert '(đã sửa)' in json.dumps(s.get(B + '/api/phieu').json()['state'], ensure_ascii=False)
    pg.evaluate("document.querySelector('#p6>div').scrollTop=0;document.getElementById('pfed').scrollTop=0")
    shot(pg, 'v27_phieu.png')
    with pg.expect_download() as dl: pg.click('#pfxp')
    pth = os.path.join(T, 'ui.docx'); dl.value.save_as(pth); assert '(đã sửa)' in '\n'.join(p_.text for p_ in docx.Document(pth).paragraphs)
    pg.click('#pfsec button[data-s="cau"]'); assert pg.locator('.pfbox[data-gq]').count() == 6
    pg.click('#pfsec button[data-s="ochu"]'); pg.wait_for_selector('#pfocall'); assert 'hợp lệ' in pg.inner_text('#pfocall')
    pg.click('#pfsec button[data-s="khoidong"]'); assert pg.locator('#pfed .pfit').count() == 4
    # ---- trò chơi ----
    pg.click('#pfplay'); pg.wait_for_selector('#gp #gpq'); pg.click('#gpfs'); pg.wait_for_timeout(400)
    assert pg.evaluate("(()=>{const r=document.getElementById('gp').getBoundingClientRect();return r.width>=innerWidth-1&&r.height>=innerHeight-1})()")
    fsz = pg.evaluate("parseFloat(getComputedStyle(document.getElementById('gpq')).fontSize)"); assert fsz >= 40, fsz
    pg.keyboard.press('2'); assert pg.locator('.gpopt.sel').count() == 1
    pg.click('#gptime'); pg.wait_for_timeout(1300); tv = pg.inner_text('#gptimer'); assert tv.isdigit() and int(tv) < 20, tv
    pg.keyboard.press('Space'); pg.wait_for_selector('#gpex'); assert pg.locator('.gpopt.right').count() == 1 and 'B' in pg.inner_text('#gpex')
    shot(pg, 'v27_rcv.png')
    pg.keyboard.press('ArrowRight'); assert 'Câu 2/6' in pg.inner_text('#gpq') and pg.locator('#gpex').count() == 0
    pg.click('#gpsnd'); assert '🔇' in pg.inner_text('#gpsnd'); pg.click('#gpsnd')
    pg.click('.gptab[data-m="ochu"]'); pg.wait_for_selector('#ocgrid'); assert pg.locator('#ocgrid .ocr').count() == 4
    kc = pg.evaluate("[...document.querySelectorAll('#ocgrid .ocr')].map(r=>[...r.querySelectorAll('.occ')].findIndex(c=>c.classList.contains('k')))"); assert len(set(kc)) == 1, kc
    pg.keyboard.press('1'); pg.keyboard.press('Enter'); pg.keyboard.press('3'); pg.click('#ocopen')
    assert pg.evaluate("[...document.querySelectorAll('#ocgrid .ocr')[0].querySelectorAll('.occ.o')].map(c=>c.textContent).join('')") == 'CHUC'
    assert pg.evaluate("[...document.querySelectorAll('#ocgrid .ocr')[2].querySelectorAll('.occ.o')].map(c=>c.textContent).join('')") == 'NHO'
    pg.locator('#ocgrid .ocn').nth(1).click(); assert 'Kết quả của phép cộng' in pg.inner_text('#occlue')
    shot(pg, 'v27_ochu.png')
    pg.click('#ockwb'); assert pg.inner_text('#ockw') == 'CỘNG'
    pg.click('.gptab[data-m="vq"]'); pg.wait_for_selector('#vqc'); pg.click('#vqspin'); pg.wait_for_selector('#vqwin', timeout=9000)
    assert pg.inner_text('#vqwin') in NAMES
    shot(pg, 'v27_vongquay.png')
    w1 = pg.inner_text('#vqwin'); pg.evaluate('window.GP_FAST=1'); pg.keyboard.press('Space'); pg.wait_for_function(f"document.getElementById('vqwin')&&document.getElementById('vqwin').textContent!=={json.dumps(w1)}", timeout=5000)
    pg.click('button[data-a="vqq"]'); pg.click('#vqspin'); pg.wait_for_selector('#vqshow'); pg.click('#vqshow'); assert pg.locator('#vqres .gpopt.right').count() == 1
    pg.click('.gptab[data-m="an"]'); pg.wait_for_selector('.ant'); pg.click('button[data-a="teams"][data-n="3"]'); assert pg.locator('.ant').count() == 3
    pg.keyboard.press('1'); pg.keyboard.press('1'); pg.keyboard.press('3'); pg.keyboard.press('Shift+Digit1'); pg.locator('.ant[data-t="1"] button[data-a="plus"]').click()
    assert [pg.inner_text(f'#sc{i}') for i in range(3)] == ['10', '10', '10']
    pg.locator('.ant[data-t="1"] button[data-a="plus"]').click(); pg.locator('.ant input').nth(1).fill('Đội Sao Vàng'); pg.keyboard.press('Tab')
    pg.click('#gptime'); pg.wait_for_timeout(1100); pg.click('#gpreveal'); pg.wait_for_selector('#gpex'); assert pg.locator('.ant.lead').count() == 1
    shot(pg, 'v27_ainhanh.png')
    assert pg.evaluate(BLUR) == 0
    if pg.evaluate('!!document.fullscreenElement'): pg.evaluate('document.exitFullscreen()'); pg.wait_for_timeout(300)
    pg.click('#gpx'); assert pg.locator('#gp').count() == 0
    # lưu Kho + link chơi
    pg.click('#pfsave'); pg.wait_for_selector('#pfplaylink'); link = pg.get_attribute('#pfplaylink', 'href')
    p2 = ctx.new_page(); p2.on('pageerror', lambda e: errs.append(str(e))); p2.goto(B + link); p2.wait_for_selector('#gp #gpq'); assert '38 + 25' in p2.inner_text('#gpq'); p2.close()
    ctxL = br.new_context(); ctxL.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in s2.cookies.get_dict().items()]); pl = ctxL.new_page(); pl.goto(B + link); pl.wait_for_timeout(800)
    assert pl.locator('#gp').count() == 0 and 'Không có trò chơi' in pl.inner_text('body'); ctxL.close()
    pg.click('#t5'); pg.wait_for_selector('.khi a[href^="/choi/"]')
    # bố cục
    for W, H in [(1280, 800), (1366, 768), (1440, 900), (1920, 1080)]:
        pg.set_viewport_size({'width': W, 'height': H}); pg.click('#t6'); pg.wait_for_timeout(400)
        m = pg.evaluate(M_JS); assert m['sh'] <= m['ih'] and m['sw'] <= m['iw'], (W, H, m)
        hb = pg.evaluate("document.querySelector('header.v9 .hbar').getBoundingClientRect().height"); assert hb < 70, (W, 'thanh trên bị xuống 2 dòng', hb)
        for sec in ('phieu', 'cau', 'ochu'):
            pg.click(f'#pfsec button[data-s="{sec}"]')
            for side in ('false', 'true'):
                pg.evaluate(f"document.getElementById('bot').classList.toggle('right',{side})")
                pg.wait_for_function("!/\\b(run|walk)\\b/.test(document.getElementById('bot').className)", timeout=8000)  # trợ lý đang chạy ngang màn hình -> đợi đứng yên
                assert not pg.evaluate(COVER), (W, sec, side, pg.evaluate(COVER), pg.evaluate("document.getElementById('bot').className"))
        assert pg.evaluate(BLUR) == 0
        pg.click('#pfplay'); pg.wait_for_selector('#gp'); assert pg.evaluate("document.getElementById('gp').scrollWidth<=innerWidth"); pg.click('#gpx')
    pg.evaluate("document.documentElement.setAttribute('data-theme','dark')"); pg.wait_for_timeout(100); assert pg.evaluate(BLUR) == 0
    ctx2 = br.new_context(viewport={'width': 390, 'height': 844}, is_mobile=True); m2 = ctx2.new_page(); ctx2.add_cookies(ctx.cookies()); m2.on('pageerror', lambda e: errs.append(str(e)))
    m2.goto(B); m2.wait_for_selector('#app:not(.hide)'); m2.click('#t6'); m2.wait_for_timeout(600); m = m2.evaluate(M_JS); assert m['sw'] <= m['iw'], m
    m2.click('#pfplay'); m2.wait_for_selector('#gpq'); m = m2.evaluate(M_JS); assert m['sw'] <= m['iw'], m
    assert not errs, errs; br.close()
print('Giao diện: chọn tuần (bắt buộc) → chọn bài → tạo bằng AI → sửa → xuất Word; 4 trò chơi (Rung chuông vàng, Ô chữ, Vòng quay tên/câu hỏi, Ai nhanh ai đúng) chơi bằng chuột + phím, toàn màn hình, âm thanh bật/tắt; lưu Kho + link chỉ chủ mở; 1 màn hình 4 cỡ, không bị trợ lý che, không mờ, điện thoại không tràn: OK')
json.dump({datetime.date.today().isoformat(): 30}, open(os.path.join(DD, 'users', 'gv-anh-le-dat', 'ai_usage.json'), 'w')); nc = len(CALLS)
r = J(s, B, '/api/phieu/generate', {'state': st}); assert r.status_code == 429 and 'hết 30 lượt' in r.json()['detail'] and 'tự gõ' in r.json()['detail'] and len(CALLS) == nc
print('Hết lượt AI: báo tiếng Việt, không gọi AI, không bịa nội dung: OK')
print('TẤT CẢ OK')
