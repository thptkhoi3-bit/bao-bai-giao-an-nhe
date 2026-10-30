"""v18: trợ lý AI Lê Tú Linh – API (Gemini giả, cần đăng nhập, giới hạn lượt, không key -> FAQ, không lộ key) + UI (chạy vào, lời chào,
mở khung chat, gửi/nhận, XSS, không che nút, không cuộn trang, mobile, không làm mờ) + ảnh chụp."""
import os, sys, json, tempfile, threading, subprocess, time, atexit
from http.server import BaseHTTPRequestHandler, HTTPServer
import requests
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
SHOT = os.environ.get('SHOT_DIR', '/workspace')
KEY = 'AIzaFAKE-chat-key-55'
SEEN = []
REPLY = 'Dạ, em là **Lê Tú Linh** ạ. Em gợi ý thầy cô vài trò chơi khởi động:\n- **Truyền điện**: trả lời đúng thì truyền cho bạn khác\n- **Ai nhanh hơn**: hai đội thi viết đáp án\n- *Hộp quà bí mật*: chuyền hộp theo nhạc'
XSS = 'Dạ, em là **Lê Tú Linh** ạ:\n- Trò chơi *Truyền điện*\n- Trò chơi **Ai nhanh hơn**\n<img src=x onerror=alert(1)><script>alert(2)</script>'
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        b = json.loads(s.rfile.read(int(s.headers['Content-Length']))); SEEN.append((s.path, s.headers.get('x-goog-api-key'), b))
        ok = s.headers.get('x-goog-api-key') == KEY
        last = b['contents'][-1]['parts'][0]['text']
        out = json.dumps({'candidates': [{'content': {'parts': [{'text': XSS if '<b>x</b>' in last else REPLY}]}}]}).encode()
        s.send_response(200 if ok else 403); s.end_headers(); s.wfile.write(out)
gs = HTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()

def server(port, extra):
    D = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=D, SECRET_KEY='c' * 32, **extra)
    for k in ('DATA_REPO', 'DATA_TOKEN', 'CHAT_DAILY_LIMIT'): env.pop(k, None)
    env.update(extra)
    if 'GEMINI_API_KEY' not in extra: env.pop('GEMINI_API_KEY', None)
    lg = open(os.path.join(D, 'server.log'), 'w')
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(port)], cwd=HERE, env=env, stdout=lg, stderr=subprocess.STDOUT)
    atexit.register(p.terminate)
    for _ in range(80):
        try: requests.get(f'http://127.0.0.1:{port}/', timeout=1); break
        except Exception: time.sleep(.2)
    return D, p
class C(requests.Session):
    def __init__(s, port): super().__init__(); s.B = f'http://127.0.0.1:{port}'
    def request(self, m, u, *a, **k): return super().request(m, self.B + u, *a, **k)
GURL = f'http://127.0.0.1:{gs.server_port}/v1beta'

# ---------- A. không có key: FAQ + thông báo thân thiện ----------
D, p = server(8030, {})
c = C(8030)
assert c.get('/api/chat').status_code == 401 and c.post('/api/chat', json={'message': 'hi'}).status_code == 401, 'phải đăng nhập'
assert c.post('/api/login', data=dict(teacher='Cô Lê Tâm', password='1993')).status_code == 200
i = c.get('/api/chat').json(); assert i['ai'] is False and i['name'] == 'Lê Tú Linh' and len(i['faq']) == 4 and i['limit'] == 50
for chip in ('Cách làm báo bài', 'Cách tạo đề kiểm tra', 'Gợi ý trò chơi khởi động', 'Soạn câu hỏi ôn tập'):
    r = c.post('/api/chat', json={'message': chip}).json(); assert r['src'] == 'faq' and len(r['reply']) > 80 and 'chưa được bật' in r['reply'], r
r = c.post('/api/chat', json={'message': 'Thời tiết hôm nay?'}).json(); assert r['reply'].startswith('Trợ lý AI chưa được bật – ') and '0865291703' in r['reply']
assert c.post('/api/chat', json={'message': '  '}).status_code == 400
assert len(c.get('/api/chat').json()['messages']) == 10
assert c.post('/api/chat/clear').json()['ok'] and c.get('/api/chat').json()['messages'] == []
p.terminate(); p.wait(); assert not SEEN
print('A. không key: cần đăng nhập, FAQ 4 gợi ý, thông báo chưa bật OK')

# ---------- B. có key (Gemini giả): ngữ cảnh, lưu lịch sử (tối đa 50), giới hạn lượt, không lộ key ----------
D, p = server(8031, dict(GEMINI_API_KEY=KEY, GEMINI_API_URL=GURL, GEMINI_MODEL='gemini-test', CHAT_DAILY_LIMIT='30'))
c = C(8031); c.post('/api/login', data=dict(teacher='Cô Lê Tâm', password='1993'))
r = c.post('/api/chat', json={'message': 'Chào em'}).json(); assert r['src'] == 'ai' and r['reply'] == REPLY and r['used'] == 1
path, key, body = SEEN[-1]
assert path == '/v1beta/models/gemini-test:generateContent' and key == KEY
sysp = body['system_instruction']['parts'][0]['text']; assert 'Lê Tú Linh' in sysp and 'xưng "em"' in sysp and 'Lê Đắc Đạt' in sysp and '2345' in sysp and 'GDPT 2018' in sysp
for k in range(20): c.post('/api/chat', json={'message': f'Câu {k}'})
body = SEEN[-1][2]; assert len(body['contents']) == 13 and body['contents'][-1]['parts'][0]['text'] == 'Câu 19' and body['contents'][-2]['role'] == 'model'
h = c.get('/api/chat').json(); assert len(h['messages']) == 42 and h['used'] == 21
for k in range(9): c.post('/api/chat', json={'message': f'x{k}'})
r = c.post('/api/chat', json={'message': 'quá lượt'}); assert r.status_code == 429 and '30 lượt' in r.json()['detail'], r.text
assert len(c.get('/api/chat').json()['messages']) == 50, 'lưu tối đa 50 tin'
# lượt chat tách riêng lượt đọc ảnh của Tạo đề
assert c.get('/api/de/info').json()['ai_used'] == 0
# người khác có lịch sử riêng
c2 = C(8031); c2.post('/api/login', data=dict(teacher='Cô Linh', password='1993')); assert c2.get('/api/chat').json()['messages'] == []
logs = open(os.path.join(D, 'server.log')).read()
for root, _, fs in os.walk(D):
    for f in fs:
        try: assert KEY not in open(os.path.join(root, f), errors='ignore').read(), f
        except IsADirectoryError: pass
assert KEY not in logs
p.terminate(); p.wait()
print('B. có key: Gemini giả, 12 tin ngữ cảnh, lưu 50 tin, giới hạn 30 lượt riêng, không lộ key OK')

# ---------- C. UI ----------
from playwright.sync_api import sync_playwright
D, p = server(8032, dict(GEMINI_API_KEY=KEY, GEMINI_API_URL=GURL))
B = 'http://127.0.0.1:8032'
BUB = 'Xin chào thầy cô! Thầy cô có cần em giúp hay giải đáp vấn đề gì không ạ?'
BLUR_JS = """()=>{const bad=[];document.querySelectorAll('*').forEach(e=>{const s=getComputedStyle(e);if(/blur/.test(s.filter+s.backdropFilter+(s.webkitBackdropFilter||'')))bad.push(e.id||e.className||e.tagName)});return bad}"""
M_JS = "()=>({sh:document.documentElement.scrollHeight,ih:innerHeight,sw:document.documentElement.scrollWidth,iw:innerWidth})"
COVER_JS = """()=>{const b=document.querySelector('#bot .bfig').getBoundingClientRect();const hit=[];
 document.querySelectorAll('button,a,input,select,textarea,label,[onclick]').forEach(e=>{if(e.closest('#bot,#botp'))return;const r=e.getBoundingClientRect();
  if(!r.width||!r.height||getComputedStyle(e).visibility==='hidden')return;if(e.offsetParent===null&&getComputedStyle(e).position!=='fixed')return;
  if(r.left<b.right-1&&r.right>b.left+1&&r.top<b.bottom-1&&r.bottom>b.top+1)hit.push((e.id||e.className||e.tagName)+':'+(e.textContent||'').trim().slice(0,20))});return {hit,b:[b.left,b.top,b.width,b.height]}}"""
def login(pg):
    pg.goto(B); pg.request.post(B + '/api/login', form={'teacher': 'Cô Lê Tâm', 'password': '1993'}); pg.goto(B); pg.wait_for_selector('#bot', state='visible', timeout=8000)
GREET = ['Xin chào thầy cô!', 'Thầy cô có cần em trợ giúp không ạ?']
POS = "()=>{const r=document.querySelector('#bot .bfig').getBoundingClientRect();return {l:r.left,r:innerWidth-r.right,h:r.height,cls:document.getElementById('bot').className}}"
def covers(pg, tag):
    for W, H in [(1280, 800), (1366, 768), (1440, 900), (1920, 1080)]:
        pg.set_viewport_size({'width': W, 'height': H}); pg.wait_for_timeout(150)
        for tab in ('#t1', '#t2', '#t3'):
            if pg.locator(tab).count(): pg.click(tab); pg.wait_for_timeout(200)
            cv = pg.evaluate(COVER_JS); assert not cv['hit'], (tag, W, H, tab, cv)
            m = pg.evaluate(M_JS); assert m['sh'] <= m['ih'] and m['sw'] <= m['iw'], (tag, W, H, tab, m)
    pg.set_viewport_size({'width': 1280, 'height': 800})
with sync_playwright() as pw:
    br = pw.chromium.launch()
    # trang đăng nhập: không có trợ lý
    pg = br.new_page(viewport={'width': 1280, 'height': 800}); pg.goto(B); pg.wait_for_timeout(1500)
    assert pg.is_hidden('#bot') and pg.is_hidden('#botp'), 'không hiện trợ lý ở trang đăng nhập'
    login(pg)
    assert 'run' in pg.get_attribute('#bot', 'class'), 'lần đầu trong phiên: chạy vào'
    frames = []; shot = lambda: frames.append(pg.screenshot(clip=dict(x=0, y=560, width=1280, height=240)))
    for k in range(7): pg.wait_for_timeout(250); shot()
    # lời chào 1 (≈5s), rồi lời chào 2 (≈5s), rồi tự ẩn
    pg.wait_for_selector('#bbub:not(.hide)', timeout=5000); pg.wait_for_timeout(450)
    assert pg.inner_text('#bbt') == GREET[0], pg.inner_text('#bbt')
    p0 = pg.evaluate(POS); assert p0['l'] < 20 and 80 <= p0['h'] <= 90, p0
    pg.screenshot(path=f'{SHOT}/v19_greet1.png')
    cv = pg.evaluate(COVER_JS); assert not cv['hit'], ('lúc chào ở góc trái', cv)
    for k in range(9): pg.wait_for_timeout(500); shot()
    pg.wait_for_function(f"document.getElementById('bbt').textContent==={json.dumps(GREET[1])}", timeout=3000); pg.wait_for_timeout(450)
    assert pg.is_visible('#bbub') and pg.evaluate(POS)['l'] < 20
    pg.screenshot(path=f'{SHOT}/v19_greet2.png')
    for k in range(9): pg.wait_for_timeout(500); shot()
    # đi bộ sang góc phải (4–6s), bong bóng ẩn, không chặn bấm vào nội dung phía dưới (bong bóng không nhận chuột)
    pg.wait_for_selector('#bot.walk', timeout=3000); assert pg.is_hidden('#bbub')
    t0 = time.time(); xs = []
    while 'right' not in (pg.get_attribute('#bot', 'class') or ''):
        shot(); xs.append(pg.evaluate(POS)['l']); pg.wait_for_timeout(350)
        assert time.time() - t0 < 9, 'phải đi sang phải trong 4–6s'
    dt = time.time() - t0; assert 3.0 <= dt <= 7.5, dt
    assert xs == sorted(xs) and xs[-1] > 600, xs
    for k in range(4): pg.wait_for_timeout(300); shot()
    p1 = pg.evaluate(POS); assert p1['r'] < 20 and 'idle' in p1['cls'] and 80 <= p1['h'] <= 90, p1
    assert pg.evaluate("getComputedStyle(document.querySelector('#bbub')).pointerEvents") == 'none'
    pg.screenshot(path=f'{SHOT}/v19_right.png')
    from PIL import Image; import io
    ims = [Image.open(io.BytesIO(f)).convert('RGB').resize((640, 120)).convert('P', palette=Image.ADAPTIVE) for f in frames]
    ims[0].save(f'{SHOT}/v19_walk.gif', save_all=True, append_images=ims[1:], duration=300, loop=0)
    # đứng góc phải: không che nút / không cuộn ở 4 cỡ, 3 tab; không làm mờ
    covers(pg, 'góc phải'); assert pg.evaluate(BLUR_JS) == []
    assert pg.evaluate(POS)['r'] < 20
    # rê chuột: hiện câu 2 thoáng qua
    pg.hover('#bot .bfig'); pg.wait_for_timeout(300); assert pg.is_visible('#bbub') and pg.inner_text('#bbt') == GREET[1]
    bb = pg.locator('#bbub').bounding_box(); assert bb['x'] + bb['width'] <= 1280
    pg.click('#bbx'); assert pg.is_hidden('#bbub')
    # mở khung chat ở góc phải, cạnh bé
    pg.click('#bot .bfig'); pg.wait_for_selector('#botp.open')
    assert pg.inner_text('#botp .bt b') == 'Lê Tú Linh' and 'AI · hỗ trợ thầy cô' in pg.inner_text('#botp .bt')
    pg.wait_for_function("document.querySelectorAll('#bchips button').length==4", timeout=5000)
    assert pg.locator('#bchips button').count() == 4 and pg.is_hidden('#bbub')
    r = pg.locator('#botp').bounding_box(); assert 340 <= r['width'] <= 380 and 480 <= r['height'] <= 530 and r['x'] > 1280 / 2 and 1280 - (r['x'] + r['width']) < 70, r
    # XSS: thoát HTML
    pg.fill('#bta', 'Em tên là gì? <b>x</b>'); pg.keyboard.press('Enter')
    pg.wait_for_selector('#bmsgs .bm.b >> nth=1', timeout=8000); pg.wait_for_selector('#btyp', state='detached', timeout=8000)
    last = pg.locator('#bmsgs .bm.b').last
    assert 'Lê Tú Linh' in last.inner_text() and last.locator('b').count() >= 2 and last.locator('li').count() == 2 and last.locator('i').count() == 1
    assert last.locator('img').count() == 0 and last.locator('script').count() == 0 and '<img src=x' in last.inner_text(), 'phải thoát HTML (XSS)'
    assert pg.locator('#bmsgs .bm.u').last.inner_text() == 'Em tên là gì? <b>x</b>'
    # hội thoại mẫu sạch cho ảnh chụp
    pg.click('#bclr'); pg.wait_for_timeout(300)
    pg.fill('#bta', 'Chào em, em tên là gì?'); pg.keyboard.press('Enter'); pg.wait_for_selector('#btyp', state='detached', timeout=8000); pg.wait_for_timeout(200)
    pg.locator('#bchips button').nth(2).click(); pg.wait_for_function("document.querySelectorAll('#bmsgs .bm.u').length>=2&&!document.getElementById('btyp')", timeout=8000)
    assert pg.evaluate(BLUR_JS) == [] and 'Truyền điện' in pg.locator('#bmsgs .bm.b').last.inner_text()
    pg.screenshot(path=f'{SHOT}/v19_chat.png')
    m = pg.evaluate(M_JS); assert m['sh'] <= m['ih'] and m['sw'] <= m['iw'], m
    # lần sau trong phiên: đứng sẵn góc phải, không chào / không đi
    pg.reload(); pg.wait_for_selector('#bot', state='visible'); pg.wait_for_timeout(700)
    p2 = pg.evaluate(POS); assert 'right' in p2['cls'] and 'run' not in p2['cls'] and 'walk' not in p2['cls'] and p2['r'] < 20 and pg.is_hidden('#bbub'), p2
    pg.click('#bot .bfig'); pg.wait_for_selector('#botp.open'); pg.wait_for_timeout(500); assert pg.locator('#bmsgs .bm.u').count() == 2
    pg.click('#bclr'); pg.wait_for_timeout(400); assert pg.locator('#bmsgs .bm.u').count() == 0
    pg.click('#bcls'); assert pg.is_hidden('#botp')
    pg.close()
    # bấm vào bé khi đang đi -> mở khung chat
    ctx2 = br.new_context(viewport={'width': 1280, 'height': 800}); pg = ctx2.new_page(); login(pg)
    pg.wait_for_selector('#bbub:not(.hide)', timeout=5000); pg.click('#bbx')          # đóng lời chào -> đi luôn
    pg.wait_for_selector('#bot.walk', timeout=3000); pg.wait_for_timeout(1200)
    bx = pg.locator('#bot .bfig').bounding_box(); pg.mouse.click(bx['x'] + bx['width'] / 2, bx['y'] + bx['height'] / 2)
    pg.wait_for_selector('#botp.open', timeout=2000)
    pg.wait_for_selector('#bot.right', timeout=8000)
    ctx2.close()
    # mobile (đã qua lượt chào trong phiên -> đứng sẵn góc phải)
    ctx = br.new_context(viewport={'width': 390, 'height': 844}, is_mobile=True, has_touch=True); pg = ctx.new_page()
    pg.add_init_script("sessionStorage.setItem('botRun','1')"); login(pg); pg.wait_for_timeout(500)
    p3 = pg.evaluate(POS); assert 58 <= p3['h'] <= 66 and p3['r'] < 16, p3
    m = pg.evaluate(M_JS); assert m['sw'] <= m['iw'], m
    pg.screenshot(path=f'{SHOT}/v19_mobile.png')
    pg.tap('#bot .bfig'); pg.wait_for_selector('#botp.open'); r = pg.locator('#botp').bounding_box(); assert r['width'] >= 389 and r['height'] >= 800, r
    pg.fill('#bta', 'Chào em'); pg.tap('#bsend'); pg.wait_for_selector('#btyp', state='detached', timeout=8000); pg.wait_for_timeout(200)
    assert pg.locator('#bmsgs .bm.b').count() >= 2
    m = pg.evaluate(M_JS); assert m['sw'] <= m['iw'], m
    # giảm chuyển động: không chạy/đi, đứng ngay góc phải
    ctx3 = br.new_context(viewport={'width': 1280, 'height': 800}, reduced_motion='reduce'); pg = ctx3.new_page(); login(pg); pg.wait_for_timeout(400)
    p4 = pg.evaluate(POS); assert 'run' not in p4['cls'] and 'walk' not in p4['cls'] and 'right' in p4['cls'] and p4['r'] < 20, p4
    br.close()
p.terminate()
print('C. UI: chạy vào, 2 lời chào ×5s, đi bộ sang góc phải, không che nút / không cuộn ở 4 cỡ, khung chat góc phải, markdown an toàn, lịch sử, bấm khi đang đi, mobile, giảm chuyển động OK')
print('TẤT CẢ OK')
