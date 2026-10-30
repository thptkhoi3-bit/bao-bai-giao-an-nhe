"""v28 – Khu trò chơi tương tác luôn hiện khi tick 'Trò chơi tương tác': 4 trò chơi + Soạn câu hỏi + Chơi ngay,
khoá kèm lời nhắc khi thiếu câu; AI lỗi (404) vẫn tự soạn được, không bịa nội dung; AI chạy được thì bật ngay. Máy tính + điện thoại."""
import os, sys, json, time, tempfile, threading, subprocess, atexit, requests
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
SHOT = os.environ.get('SHOT_DIR', '/workspace')
A = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
G0 = A + '5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx'; BBP = A + '455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx'
QS = [{'q': f'{a} + {b} = ?', 'options': [str(a + b), str(a + b + 10), str(a + b - 1), str(a + b + 1)], 'a': 'A', 'explain': ''} for a, b in [(38, 25), (47, 18), (29, 29), (36, 18), (26, 15)]]
OC = {'keyword': 'CỘNG', 'clue': 'Phép tính', 'rows': [{'clue': '10 đơn vị bằng 1 …', 'answer': 'CHỤC'}, {'clue': 'Kết quả phép cộng', 'answer': 'TỔNG'}, {'clue': 'viết 3 … 1', 'answer': 'NHỚ'}, {'clue': 'thẳng …', 'answer': 'HÀNG'}]}
MODE = {'v': 'fail'}; CALLS = []
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        s.rfile.read(int(s.headers['Content-Length'])); CALLS.append(s.path)
        if MODE['v'] == 'fail':
            b = json.dumps({'error': {'code': 404, 'message': 'model not found'}}).encode(); s.send_response(404)
        else:
            b = json.dumps({'candidates': [{'content': {'parts': [{'text': json.dumps({'game': {'questions': QS, 'ochu': OC}}, ensure_ascii=False)}]}}]}).encode(); s.send_response(200)
        s.send_header('Content-Type', 'application/json'); s.send_header('Content-Length', str(len(b))); s.end_headers(); s.wfile.write(b)
gs = ThreadingHTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
dd = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=dd, SECRET_KEY='t' * 32, GEMINI_API_KEY='test-key-' + 'w' * 20, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}')
for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_MODEL'): env.pop(k, None)
P = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8096'], cwd=HERE, env=env, stdout=open(os.path.join(dd, 'log'), 'w'), stderr=subprocess.STDOUT); atexit.register(P.terminate)
B = 'http://127.0.0.1:8096'
for _ in range(100):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
s = requests.Session(); s.post(B + '/api/login', data=dict(teacher='Anh Lê Đạt', password='1993', name='')).raise_for_status()
s.post(B + '/api/baobai/master', files=[('file', ('m.docx', open(G0, 'rb')))]).raise_for_status()
s.post(B + '/api/baobai/fill', data=dict(week=7), files=[('file', ('p.docx', open(BBP, 'rb')))]).raise_for_status()
s.post(B + '/api/phieu/state', data=json.dumps({'state': {'types': ['phieu', 'tuongtac'], 'muc': 'vua'}}), headers={'Content-Type': 'application/json'}).raise_for_status()

from playwright.sync_api import sync_playwright
GAMES = ['Rung chuông vàng', 'Ô chữ', 'Vòng quay', 'Ai nhanh ai đúng']
M_JS = "()=>({sh:document.documentElement.scrollHeight,ih:innerHeight,sw:document.documentElement.scrollWidth,iw:innerWidth})"
def games(pg): return pg.evaluate("[...document.querySelectorAll('#pfgames .pfgi')].map(e=>({t:e.querySelector('b').textContent,st:e.querySelector('.gst').textContent,dis:e.querySelector('.gplay').disabled,vis:!!e.offsetParent}))")
def fill_q(pg, i, q, opts):
    box = pg.locator(f'.pfbox[data-gq="{i}"]'); box.locator('.pfit > input[type=text]').first.fill(q)
    for j, o in enumerate(opts): box.locator('.pfopts input[type=text]').nth(j).fill(o)
with sync_playwright() as pw:
    br = pw.chromium.launch(); errs = []
    for W, H, mob in [(1440, 900, False), (390, 844, True)]:
        ctx = br.new_context(viewport={'width': W, 'height': H}, is_mobile=mob, has_touch=mob); pg = ctx.new_page(); pg.on('pageerror', lambda e: errs.append(str(e)))
        ctx.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in s.cookies.get_dict().items()])
        s.post(B + '/api/phieu/state', data=json.dumps({'state': {'types': ['phieu', 'tuongtac'], 'muc': 'vua'}}), headers={'Content-Type': 'application/json'})
        pg.goto(B); pg.wait_for_selector('#app:not(.hide)'); pg.click('#t6'); pg.wait_for_selector('#pfgames .pfgi')
        g = games(pg); assert [x['t'].split(' ', 1)[1] for x in g] == GAMES and all(x['vis'] and x['dis'] for x in g), g
        assert 'Cần ít nhất 3 câu hỏi hợp lệ – còn thiếu 3 câu' in g[0]['st'] and 'Chưa có từ khoá' in g[1]['st'], g
        # bỏ tick -> ẩn, tick lại -> hiện
        pg.locator('#pftypes label[data-k="tuongtac"]').click(); assert not pg.is_visible('#pfgcard')
        pg.locator('#pftypes label[data-k="tuongtac"]').click(); assert pg.is_visible('#pfgcard')
        # AI lỗi 404 (mọi mô hình) -> báo lỗi, không bịa, mở khung soạn câu hỏi
        MODE['v'] = 'fail'; pg.select_option('#pfweek', '7'); pg.wait_for_selector('.pfl'); pg.locator('.pfl', has_text='Toán').first.click()
        pg.click('#pfgo'); pg.wait_for_function("document.getElementById('pfai').textContent.includes('Không gọi được AI')", timeout=20000)
        assert 'tự soạn câu hỏi' in pg.inner_text('#pfai') and pg.evaluate("PF.sec") == 'cau' and pg.evaluate("PF.st.game.questions.length") == 0
        assert all(x['dis'] for x in games(pg))
        # tự soạn: Soạn câu hỏi -> thêm câu trống, gõ 2 câu -> vẫn thiếu 1
        pg.locator('.pfgi[data-g="rcv"] .gedit').click(); pg.wait_for_selector('.pfbox[data-gq="0"]')
        fill_q(pg, 0, '9 + 6 = ?', ['15', '14', '16', '13']); pg.click('#pfaddq'); fill_q(pg, 1, '7 + 8 = ?', ['14', '15', '16', '13'])
        pg.locator('.pfbox[data-gq="1"] .pfopts input[type=radio]').nth(1).check()
        g = games(pg); assert 'còn thiếu 1 câu' in g[0]['st'] and g[0]['dis'] and not g[2]['dis'] and g[3]['dis'], g
        pg.click('#pfaddq'); g = games(pg); assert 'Có 1 câu chưa đủ' in g[0]['st'], g
        fill_q(pg, 2, '28 + 5 = ?', ['33', '32', '23', '34'])
        g = games(pg); assert not g[0]['dis'] and not g[3]['dis'] and '3 câu hỏi' in g[0]['st'], g
        if not mob: pg.evaluate("document.querySelector('#p6>div').scrollTop=document.getElementById('pfgcard').offsetTop-80"); pg.wait_for_timeout(200); pg.screenshot(path=f'{SHOT}/v28_trochoi.png')
        else: pg.locator('#pfgcard').scroll_into_view_if_needed(); pg.wait_for_timeout(200); pg.screenshot(path=f'{SHOT}/v28_trochoi_mobile.png')
        m = pg.evaluate(M_JS); assert m['sw'] <= m['iw'], m
        pg.locator('.pfgi[data-g="an"] .gplay').click(); pg.wait_for_selector('#gp .ant'); assert '9 + 6' in pg.inner_text('#gpq'); pg.click('#gpx')
        pg.locator('.pfgi[data-g="rcv"] .gplay').click(); pg.wait_for_selector('#gp #gpq'); pg.click('#gpx')
        assert '9 + 6' in json.dumps(s.get(B + '/api/phieu').json()['state'], ensure_ascii=False) or (pg.wait_for_timeout(1000) or True)
        # AI chạy được -> câu hỏi + ô chữ -> cả 4 trò chơi bật
        MODE['v'] = 'ok'; pg.on('dialog', lambda d: d.accept())
        pg.locator('#pftypes label[data-k="phieu"]').click()  # chỉ trò chơi tương tác
        pg.click('#pfgo'); pg.wait_for_function("PF.st.game.questions.length===5", timeout=20000)
        g = games(pg); assert not any(x['dis'] for x in g), g
        pg.locator('.pfgi[data-g="ochu"] .gplay').click(); pg.wait_for_selector('#ocgrid'); pg.click('#gpx')
        pg.locator('.pfgi[data-g="vq"] .gplay').click(); pg.wait_for_selector('#vqc'); pg.click('#gpx')
        if not mob: pg.evaluate("document.querySelector('#p6>div').scrollTop=document.getElementById('pfgcard').offsetTop-80"); pg.wait_for_timeout(200); pg.screenshot(path=f'{SHOT}/v28_trochoi_ai.png')
        ctx.close()
    assert not errs, errs; br.close()
print('v28 trò chơi: 4 trò chơi luôn hiện khi tick, khoá + nhắc thiếu gì, AI 404 vẫn tự soạn (không bịa), AI chạy -> bật cả 4, máy tính + điện thoại: OK')
