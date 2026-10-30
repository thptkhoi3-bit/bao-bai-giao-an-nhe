"""v16 UI: tab 'Tạo đề kiểm tra' – không cuộn trang ở 4 cỡ màn hình, mobile không tràn ngang, không làm mờ, tạo đề Toán/TV, xem lại + sửa, ảnh chụp."""
import os, sys, json, tempfile, threading, subprocess, time, atexit, io
from http.server import BaseHTTPRequestHandler, HTTPServer
import requests
from playwright.sync_api import sync_playwright
from PIL import Image, ImageDraw, ImageFont
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import exam
SHOT = os.environ.get('SHOT_DIR', '/workspace')
KEY = 'AIzaFAKE-ui-key'
SVG = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 70 42"><polygon points="6,36 34,6 64,36" fill="none" stroke="black" stroke-width="0.5"/>'
       '<line x1="34" y1="6" x2="34" y2="36" stroke="black" stroke-width="0.5"/><g font-family="Times New Roman" font-size="5"><text x="32" y="5">A</text><text x="1" y="40">B</text><text x="31" y="41">H</text><text x="64" y="40">C</text></g></svg>')
RESP = {'sections': [{'title': 'I. TRẮC NGHIỆM', 'instruction': 'Khoanh vào chữ cái đặt trước câu trả lời đúng', 'questions': [
    {'so': 1, 'text': 'Trong hình bên có bao nhiêu hình tam giác?', 'muc': 'M2', 'diem': '0,5', 'blocks': [{'t': 'figure', 'id': 'f1'}, {'t': 'options', 'items': ['2', '3', '4', '5'], 'layout': 'row'}]},
    {'so': 2, 'text': 'Số gồm 4 trăm, 5 chục và 4 đơn vị viết là:', 'muc': 'M1', 'diem': '0,5', 'blocks': [{'t': 'options', 'items': ['445', '544', '454', '545'], 'layout': 'row'}]}]},
    {'title': 'II. TỰ LUẬN', 'questions': [{'so': 3, 'text': 'Đặt tính rồi tính: 427 + 109', 'muc': 'M1', 'diem': '1', 'blocks': [{'t': 'lines', 'n': 3}]}]}],
    'answers_list': [{'so': 1, 'text': 'B', 'points': '0,5'}, {'so': 2, 'text': 'C', 'points': '0,5'}],
    'figures': [{'id': 'f1', 'kind': 'geometry', 'page': 0, 'box': [180, 300, 470, 700], 'svg': SVG, 'w_cm': 5.5}]}
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        s.rfile.read(int(s.headers['Content-Length'])); out = json.dumps({'candidates': [{'content': {'parts': [{'text': json.dumps(RESP)}]}}]}).encode()
        s.send_response(200); s.end_headers(); s.wfile.write(out)
gs = HTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
D = tempfile.mkdtemp(); PORT = 8022; B = f'http://127.0.0.1:{PORT}'
env = dict(os.environ, DATA_DIR=D, SECRET_KEY='z' * 32, GEMINI_API_KEY=KEY, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}/v1beta')
for k in ('DATA_REPO', 'DATA_TOKEN'): env.pop(k, None)
srv = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(PORT)], cwd=HERE, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL); atexit.register(srv.terminate)
for _ in range(80):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
# ảnh chụp đề giả (có hình tam giác) để thử đường đọc ảnh
IMG = os.path.join(D, 'anh_de.jpg'); im = Image.new('RGB', (1000, 1400), 'white'); dr = ImageDraw.Draw(im)
try: fnt = ImageFont.truetype('/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf', 30)
except Exception: fnt = None
dr.text((60, 120), 'Câu 1. Trong hình bên có bao nhiêu hình tam giác? (M2) (0,5đ)', fill='black', font=fnt)
dr.polygon([(340, 620), (500, 300), (680, 620)], outline='black', width=3); dr.line([(500, 300), (500, 620)], fill='black', width=3)
for t, xy in (('A', (490, 262)), ('B', (310, 625)), ('H', (490, 628)), ('C', (685, 625))): dr.text(xy, t, fill='black', font=fnt)
dr.text((60, 700), 'A. 2        B. 3        C. 4        D. 5', fill='black', font=fnt); im.save(IMG, quality=92)

SIZES = [(1280, 800), (1366, 768), (1440, 900), (1920, 1080)]
BLUR_JS = """()=>{const bad=[];document.querySelectorAll('*').forEach(e=>{const s=getComputedStyle(e);if(/blur/.test(s.filter+s.backdropFilter+(s.webkitBackdropFilter||'')))bad.push(e.id||e.className||e.tagName)});return bad}"""
M_JS = "()=>({sh:document.documentElement.scrollHeight,ih:innerHeight,sw:document.documentElement.scrollWidth,iw:innerWidth})"
def login(pg, who='Anh Lê Đạt'):
    pg.goto(B); pg.request.post(B + '/api/login', form={'teacher': who, 'password': '1993'}); pg.goto(B); pg.wait_for_selector('#t3'); pg.wait_for_timeout(600)
fails = []
with sync_playwright() as p:
    br = p.chromium.launch()
    # 1) bố cục desktop, chưa có đề
    for W, H in SIZES:
        pg = br.new_page(viewport={'width': W, 'height': H}, locale='vi-VN'); login(pg); pg.click('#t3'); pg.wait_for_timeout(250)
        m = pg.evaluate(M_JS)
        if m['sh'] > m['ih'] or m['sw'] > m['iw']: fails.append(f'{W}x{H} cuộn trang {m}')
        assert pg.is_visible('#p3') and not pg.is_visible('#p1')
        for mon in ('toan', 'tv'):
            pg.select_option('#dmon', mon)
            if mon == 'toan': pg.set_input_files('#dfiles', os.path.join(HERE, 'de_templates', 'toan.docx'))
            sc = pg.evaluate("()=>{const c=document.querySelector('#p3>div:first-child');return c.scrollHeight-c.clientHeight}")
            if sc > 2: fails.append(f'{W}x{H} {mon}: cột trái phải cuộn {sc}px')
        hb = pg.evaluate("()=>document.querySelector('header.v9 .hbar').getBoundingClientRect().height")
        if hb > 60: fails.append(f'{W}x{H}: thanh tab bị xuống 2 dòng ({hb}px)')
        pg.close()
    # 2) Toán từ Word (desktop 1920x1080) -> ảnh v16_de_toan.png
    pg = br.new_page(viewport={'width': 1920, 'height': 1080}, locale='vi-VN'); login(pg); pg.click('#t3')
    assert 'còn' in pg.inner_text('#dainote')
    pg.fill('#dxa', 'Hoà Bình'); pg.fill('#dtruong', 'TH Kim Đồng'); pg.select_option('#dlop', '3'); pg.select_option('#dloai', 'Cuối học kì I')
    pg.set_input_files('#dfiles', os.path.join(HERE, 'de_templates', 'toan.docx'))
    assert pg.locator('#dthumbs figure').count() == 1
    pg.click('#dgo'); pg.wait_for_selector('#dbadge:not(.hide)', timeout=120000); pg.wait_for_timeout(900)
    assert '12 câu' in pg.inner_text('#dbadge'), pg.inner_text('#dbadge')
    assert 'dis' not in pg.get_attribute('#ddl', 'class')
    tabs = pg.locator('#dptabs button').all_inner_texts(); assert any('Trắc nghiệm' in t for t in tabs) and any('Đáp án' in t for t in tabs), tabs
    m = pg.evaluate(M_JS); assert m['sh'] <= m['ih'] and m['sw'] <= m['iw'], m
    assert pg.evaluate(BLUR_JS) == [], 'có hiệu ứng làm mờ'
    pg.click('#dpage')   # xem cả trang
    pg.wait_for_selector('#toast', state='hidden', timeout=8000); pg.screenshot(path=f'{SHOT}/v16_de_toan.png')
    with pg.expect_download() as d: pg.click('#ddl')
    dl = d.value; outdoc = os.path.join(D, 'de_toan.docx'); dl.save_as(outdoc); assert dl.suggested_filename == 'DeKiemTra_Toan_Lop3.docx'
    # chế độ sửa: đổi câu 2 rồi cập nhật
    pg.click('#dmEd'); pg.wait_for_timeout(200); assert pg.is_visible('#ded') and pg.locator('#ded .dq').count() == 12
    ta = pg.locator('#ded .dq').nth(1).locator('.dqh textarea'); ta.fill('Số liền trước của 100 là số nào?')
    src0 = pg.get_attribute('#dpage', 'src')
    pg.click('#dupd'); pg.wait_for_function("s=>document.getElementById('dpage').getAttribute('src')!==s", arg=src0, timeout=120000)
    # đề B sau đề A: tệp/ô xem trước/đề cũ phải được xoá
    import docx as _docx
    d = _docx.Document(); d.add_paragraph('Câu 1. Bạn Bình có 7 quả cam, mẹ cho thêm 5 quả. Hỏi có mấy quả? (M1) (0,5đ)'); d.add_paragraph('A. 12\tB. 13\tC. 11\tD. 2')
    d.add_paragraph('Câu 2. Tính nhẩm: 30 + 40 = ? (M1) (1đ)'); d.add_paragraph('.' * 200); fB = os.path.join(D, 'de_B.docx'); d.save(fB)
    idA = pg.evaluate('DE.id')
    pg.set_input_files('#dfiles', fB)
    assert pg.locator('#dthumbs figure').count() == 1 and pg.inner_text('#dthumbs figcaption') == 'de_B.docx', 'tệp của đề A phải bị bỏ'
    assert pg.is_hidden('#dbadge') and 'dis' in pg.get_attribute('#ddl', 'class') and pg.inner_text('#ded') == '' and not pg.get_attribute('#dpage', 'src'), 'ô xem trước/đề cũ phải trống'
    pg.click('#dgo'); pg.wait_for_selector('#dbadge:not(.hide)', timeout=120000); pg.wait_for_timeout(500)
    assert '2 câu' in pg.inner_text('#dbadge'), pg.inner_text('#dbadge')
    idB = pg.evaluate('DE.id'); assert idB != idA and idB in pg.get_attribute('#dpage', 'src') and '&t=' in pg.get_attribute('#dpage', 'src')
    pg.click('#dmEd'); pg.wait_for_timeout(200); vals = pg.evaluate("()=>[...document.querySelectorAll('#ded textarea,#ded input')].map(e=>e.value).join('|')")
    assert pg.locator('#ded .dq').count() == 2 and 'quả cam' in vals and 'Số 99' not in vals and 'Số liền trước' not in vals, vals
    pg.wait_for_selector('#toast', state='hidden', timeout=8000); pg.screenshot(path=f'{SHOT}/v16_de_newB.png')
    # nút Tạo đề mới: xoá hết
    pg.click('#dnew'); pg.wait_for_timeout(500)
    assert pg.locator('#dthumbs figure').count() == 0 and pg.is_hidden('#dbadge') and pg.is_visible('#dempty') and 'dis' in pg.get_attribute('#ddl', 'class')
    assert pg.evaluate("async()=>(await fetch('/api/de/'+%r)).status" % idB) == 404
    # đổi môn sau khi có tệp: xoá danh sách tệp
    pg.set_input_files('#dfiles', fB); pg.select_option('#dmon', 'tv'); pg.select_option('#dmon', 'toan'); assert pg.locator('#dthumbs figure').count() == 0
    pg.close()
    # 3) Tiếng Việt: 3 ô -> v16_de_tv.png
    pg = br.new_page(viewport={'width': 1440, 'height': 900}, locale='vi-VN'); login(pg); pg.click('#t3')
    pg.select_option('#dmon', 'tv'); assert pg.is_visible('#dtv') and not pg.is_visible('#dtoan')
    assert pg.inner_text('#ddapanl') == 'Kèm hướng dẫn chấm'
    pg.set_input_files('#ddoc', os.path.join(HERE, 'de_templates', 'tv.docx'))
    assert 'has' in pg.get_attribute('#ddoc >> xpath=..', 'class')
    pg.click('#dgo'); pg.wait_for_selector('#dbadge:not(.hide)', timeout=120000); pg.wait_for_timeout(900)
    assert '9 câu' in pg.inner_text('#dbadge')
    tabs = pg.locator('#dptabs button').all_inner_texts(); assert any('Bài 1' in t for t in tabs) and any('Nghe - viết' in t for t in tabs) and any('Hướng dẫn chấm' in t for t in tabs), tabs
    m = pg.evaluate(M_JS); assert m['sh'] <= m['ih'] and m['sw'] <= m['iw'], m
    pg.wait_for_selector('#toast', state='hidden', timeout=8000); pg.screenshot(path=f'{SHOT}/v16_de_tv.png'); pg.close()
    # 4) Ảnh chụp đề (Gemini giả lập) -> xem lại: hình vẽ lại cạnh ảnh gốc -> v16_de_review.png
    pg = br.new_page(viewport={'width': 1440, 'height': 900}, locale='vi-VN'); login(pg); pg.click('#t3'); pg.select_option('#dmon', 'toan')
    pg.set_input_files('#dfiles', IMG); pg.click('#dgo'); pg.wait_for_selector('#dbadge:not(.hide)', timeout=120000); pg.wait_for_timeout(600)
    pg.click('#dmEd'); pg.wait_for_timeout(500)
    figs = pg.locator('#ded .dfig figure'); assert figs.count() == 2, figs.count()
    assert pg.evaluate("()=>[...document.querySelectorAll('#ded .dfig img')].every(i=>i.complete&&i.naturalWidth>50)")
    assert pg.evaluate(BLUR_JS) == []
    pg.wait_for_selector('#toast', state='hidden', timeout=8000); pg.screenshot(path=f'{SHOT}/v16_de_review.png'); pg.close()
    # 5) mobile
    for W, H in ((390, 844), (360, 740)):
        pg = br.new_page(viewport={'width': W, 'height': H}, locale='vi-VN', is_mobile=True, has_touch=True); login(pg, 'Cô Linh'); pg.click('#t3'); pg.wait_for_timeout(300); pg.fill('#dxa', 'An Bình'); pg.fill('#dtruong', 'TH An Bình')
        m = pg.evaluate(M_JS)
        if m['sw'] > m['iw']: fails.append(f'mobile {W}: tràn ngang {m}')
        pg.set_input_files('#dfiles', os.path.join(HERE, 'de_templates', 'toan.docx')); pg.click('#dgo'); pg.wait_for_selector('#dbadge:not(.hide)', timeout=120000); pg.wait_for_timeout(500)
        m = pg.evaluate(M_JS)
        if m['sw'] > m['iw']: fails.append(f'mobile {W} sau khi tạo: tràn ngang {m}')
        if W == 390: pg.screenshot(path=os.path.join(D, 'mobile.png'), full_page=True); pg.click('#dmEd'); pg.wait_for_timeout(200); m = pg.evaluate(M_JS); assert m['sw'] <= m['iw'], m
        pg.close()
    br.close()
# 6) ảnh trang 1 của file Word tạo ra
pgs = exam.render_pages(outdoc, os.path.join(D, 'out_pages'), dpi=150)
if pgs: import shutil; shutil.copy(pgs[0], f'{SHOT}/v16_out_p1.png'); assert len(pgs) == 4
gs.shutdown()
if fails: print('LỖI:', *fails, sep='\n  '); sys.exit(1)
print('tests_v16 OK')
