"""v25 – ô 'Tuần' ở tab Giáo án luôn hiện và có tuần của báo bài gốc (mẫu anh, mẫu trường khác, tài khoản mới, chưa có báo bài, mạng lỗi), sáng/tối, máy tính/điện thoại."""
import os, sys, time, tempfile, subprocess, atexit, requests
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import synth
SHOT = os.environ.get('SHOT_DIR', '/workspace'); T = tempfile.mkdtemp()
M = '/cursor/stores/store-0e61161c-f9bc-45cb-8d83-359add90b33e/baobai/BaoBai_Goc_moinhat.docx'
PHU = '/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'
DD = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=DD, SECRET_KEY='t' * 32)
for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_API_KEY'): env.pop(k, None)
B = 'http://127.0.0.1:8075'
pr = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8075'], cwd=HERE, env=env, stdout=open(os.path.join(T, 'l'), 'w'), stderr=subprocess.STDOUT); atexit.register(pr.terminate)
for _ in range(100):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
def sess(teacher, pw='1993', name=''):
    c = requests.Session(); c.post(B + '/api/login', data=dict(teacher=teacher, password=pw, name=name)).raise_for_status(); return c
cases = []
c = sess('Cô Lê Tâm'); c.post(B + '/api/baobai/master', files={'file': ('m.docx', open(M, 'rb'))}).raise_for_status()
c.post(B + '/api/draft/add', data=dict(week=7), files=[('files', ('p.docx', open(PHU, 'rb')))]).raise_for_status()
cases.append(('mẫu anh, đã điền + có nháp', c, 12, True))
c = sess('Cô Linh'); c.post(B + '/api/baobai/master', files={'file': ('a.docx', open(synth.baobai_A(os.path.join(T, 'A.docx')), 'rb'))}).raise_for_status(); cases.append(('mẫu trường khác A, chưa có nháp', c, 2, False))
c = sess('other', '0865291703', 'Cô Hoa'); c.post(B + '/api/baobai/master', files={'file': ('b.docx', open(synth.baobai_B(os.path.join(T, 'B.docx')), 'rb'))}).raise_for_status(); cases.append(('khách, 1 bảng cả năm B', c, 2, False))
c = sess('Cô Lê Thủy'); cases.append(('chưa có báo bài gốc', c, 0, False))
st = cases[0][1].get(B + '/api/status').json(); assert st['draft_weeks'] == [7] and len(st['weeks']) == 12
assert cases[3][1].get(B + '/api/status').status_code == 200

from playwright.sync_api import sync_playwright
INFO = """()=>{const s=document.getElementById('gweek'),r=s.getBoundingClientRect(),cs=getComputedStyle(s);const e=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);
 return {n:s.options.length,weeks:[...s.options].filter(o=>o.value).length,first:s.options[0]&&s.options[0].text,txt:[...s.options].map(o=>o.text).join('|'),
  shown:cs.display!=='none'&&cs.visibility!=='hidden'&&r.width>120&&r.height>=28,inview:r.left>=0&&r.right<=innerWidth+1&&r.top>=0&&r.bottom<=innerHeight,top:e===s||s.contains(e),hint:document.getElementById('ghint').textContent}}"""
with sync_playwright() as pw:
    br = pw.chromium.launch()
    for name, c, nweeks, has_draft in cases:
        for vw, vh, mob in ((1366, 768, False), (390, 844, True)):
            for theme in ('light', 'dark'):
                ctx = br.new_context(viewport={'width': vw, 'height': vh}, is_mobile=mob, device_scale_factor=2 if mob else 1); pg = ctx.new_page(); errs = []
                pg.on('pageerror', lambda e: errs.append(str(e)))
                pg.add_init_script(f"try{{localStorage.setItem('theme_v2','{theme}')}}catch(e){{}}")   # bé Linh chạy vào như thật
                ctx.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in c.cookies.get_dict().items()]); pg.goto(B)
                pg.wait_for_selector('#app:not(.hide)'); pg.click('#t2'); pg.wait_for_timeout(9000 if (vw, theme) == (1366, 'light') else 600)
                pg.evaluate("document.getElementById('gweek').scrollIntoView({block:'center'})"); pg.wait_for_timeout(200)
                i = pg.evaluate(INFO); tag = f'{name} {vw} {theme}: {i}'
                assert i['shown'] and i['inview'] and i['top'], 'ô Tuần bị ẩn/che: ' + tag
                assert i['weeks'] == nweeks, tag
                if nweeks: assert i['first'] == 'Chọn tuần…' and 'Tự nhận diện' not in i['txt'], tag
                else: assert 'Chưa có báo bài gốc' in i['first'] and 'Tab Báo bài'.lower() in i['hint'].lower(), tag
                if has_draft: assert 'Tuần 7' in i['txt'] and 'có bản nháp' in i['txt'], tag
                assert not errs, errs
                if nweeks:
                    pg.select_option('#gweek', '7' if nweeks == 12 else pg.evaluate("[...document.getElementById('gweek').options].find(o=>o.value).value"))
                    assert pg.is_enabled('#ggo'), tag
                if name.startswith('mẫu anh') and vw == 1366 and theme == 'light':
                    pg.select_option('#gweek', ''); pg.evaluate("(()=>{const s=document.getElementById('gweek');s.size=9;s.style.height='auto';s.style.position='relative';s.style.zIndex=50;s.focus()})()")
                    pg.wait_for_timeout(200); pg.screenshot(path=SHOT + '/v25_giaoan_tuan.png')
                ctx.close()
    # mạng lỗi lần đầu -> vẫn có ô Tuần, tự thử lại rồi hiện tuần
    ctx = br.new_context(viewport={'width': 1366, 'height': 768}); pg = ctx.new_page(); n = {'k': 0}
    def flaky(route):
        n['k'] += 1
        if n['k'] <= 2: return route.fulfill(status=500, body='{"detail":"lỗi máy chủ"}', content_type='application/json')
        route.continue_()
    pg.route('**/api/status', flaky); pg.add_init_script("try{sessionStorage.setItem('botRun','1')}catch(e){}")
    ctx.add_cookies([{'name': k, 'value': v, 'url': B} for k, v in cases[0][1].cookies.get_dict().items()]); pg.goto(B)
    pg.wait_for_selector('#app:not(.hide)'); pg.click('#t2'); pg.wait_for_timeout(300)
    assert pg.is_visible('#gweek') and pg.evaluate("document.getElementById('gweek').options.length") >= 1
    pg.wait_for_function("[...document.getElementById('gweek').options].filter(o=>o.value).length==12", timeout=15000)
    br.close()
print('Ô Tuần tab Giáo án: hiện, không bị che/cắt, đủ tuần (mẫu anh / trường khác / khách / chưa có báo bài – có hướng dẫn), đánh dấu có bản nháp, sáng/tối, 1366 & 390, tự thử lại khi mạng lỗi OK')
print('TẤT CẢ OK')
