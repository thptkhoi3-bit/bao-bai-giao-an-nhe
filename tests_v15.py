"""Giao diện đăng ký / đăng nhập tài khoản đã đăng ký / quản lý tài khoản (máy chủ tạm, dữ liệu tạm)."""
import os, sys, tempfile, subprocess, time, atexit, requests, datetime
D = tempfile.mkdtemp(); here = os.path.dirname(os.path.abspath(__file__))
env = dict(os.environ, DATA_DIR=D, SECRET_KEY='z' * 32); env.pop('DATA_REPO', None); env.pop('DATA_TOKEN', None)
srv = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8018'], cwd=here, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL); atexit.register(srv.terminate)
B = 'http://127.0.0.1:8018'
for _ in range(50):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
for i, (n, p, sc) in enumerate((('Thầy Trần Minh', '0987111222', 'TH Kim Đồng – lớp 4B'), ('Cô Phạm Thu Hà', '0356222333', 'TH Nguyễn Du – lớp 2A'), ('Cô Đỗ Lan', '0702333444', ''))):
    requests.post(B + '/api/register', data=dict(name=n, phone=p, school=sc, password='abcdef', password2='abcdef'), headers={'X-Forwarded-For': f'1.1.1.{i}'})
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(); E = []
    def page(w, h, mob=False):
        pg = b.new_page(viewport={'width': w, 'height': h}, device_scale_factor=2 if mob else 1, is_mobile=mob, locale='vi-VN'); pg.on('pageerror', lambda x: E.append(str(x)))
        pg.on('dialog', lambda d: d.accept()); pg.goto(B); pg.wait_for_selector('#login:not(.hide)'); pg.wait_for_timeout(300); return pg
    def noscroll(pg, mob=False):
        r = pg.evaluate('[document.documentElement.scrollHeight,innerHeight,document.documentElement.scrollWidth,innerWidth]'); assert r[2] <= r[3] and (mob or r[0] <= r[1]), r
        c = pg.evaluate("(()=>{const r=document.querySelector('.c14').getBoundingClientRect();return [r.top,r.bottom,r.right,innerHeight,innerWidth]})()")
        assert mob or (c[0] >= 0 and c[1] <= c[3] and c[2] <= c[4]), c
    # --- đăng ký ---
    for w, h, mob in ((1280, 800, 0), (1366, 768, 0), (1920, 1080, 0), (390, 844, 1)):
        pg = page(w, h, bool(mob)); pg.click('#toreg'); assert pg.is_visible('#rname') and not pg.is_visible('#lt') and pg.inner_text('#lch2') == 'ĐĂNG KÝ'
        pg.wait_for_timeout(100); noscroll(pg, mob)
        if w == 1280:
            pg.click('#rbtn'); assert 'họ và tên' in pg.inner_text('#rerr')
            pg.fill('#rname', 'Cô Nguyễn Thị Hoa'); pg.fill('#rphone', '0912 345 678'); pg.fill('#rschool', 'TH Lê Lợi – lớp 3A'); pg.fill('#rpw', 'abc'); pg.click('#rbtn'); assert '6 ký tự' in pg.inner_text('#rerr')
            pg.fill('#rpw', 'matkhau1'); pg.fill('#rpw2', 'matkhau2'); pg.click('#rbtn'); assert 'không khớp' in pg.inner_text('#rerr')
            pg.fill('#rpw2', 'matkhau1'); pg.evaluate("$('rerr').textContent=''"); pg.screenshot(path='/workspace/v15_register.png')
            pg.click('#rbtn'); pg.wait_for_selector('#rok:not(.hide)')
            assert pg.inner_text('#rokt') == 'Anh/chị đã đăng ký, vui lòng chờ duyệt. Liên hệ 0865291703.'; noscroll(pg)
            pg.screenshot(path='/tmp/v15_reg_done.png')
            pg.click('#rview .lsw button'); assert pg.input_value('#lt') == 'registered' and pg.is_visible('#lph') and pg.input_value('#lph') == '0912345678'
            pg.fill('#lp', 'matkhau1'); pg.click('#lbtn'); pg.wait_for_timeout(500); assert 'chờ duyệt' in pg.inner_text('#lerr'), pg.inner_text('#lerr')
        if mob: pg.screenshot(path='/tmp/v15_reg_mobile.png', full_page=True)
        pg.close()
    # --- quản trị ---
    pg = page(1280, 800); assert not pg.is_visible('#admbtn')
    pg.select_option('#lt', 'Cô Linh'); pg.fill('#lp', '1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.wait_for_timeout(400)
    assert not pg.is_visible('#admbtn'), 'giáo viên thường không thấy nút quản lý'; pg.click('#userbox .btn-ghost'); pg.wait_for_selector('#login:not(.hide)')
    pg.select_option('#lt', 'Anh Lê Đạt'); pg.fill('#lp', '1993'); pg.click('#lbtn'); pg.wait_for_selector('#admbtn:not(.hide)'); pg.wait_for_timeout(300)
    assert pg.inner_text('#admbadge') == '4'
    pg.click('#admbtn'); pg.wait_for_selector('#admrows tr[data-id]')
    assert pg.locator('#admrows tr[data-id]').count() == 4
    row = lambda name: pg.locator('#admrows tr', has_text=name)
    row('Cô Nguyễn Thị Hoa').locator('button', has_text='Duyệt').click(); pg.wait_for_timeout(400)
    exp = (datetime.date.today() + datetime.timedelta(days=365)).isoformat()
    assert row('Cô Nguyễn Thị Hoa').locator('input[type=date]').input_value() == exp and 'Đã duyệt' in row('Cô Nguyễn Thị Hoa').inner_text()
    row('Thầy Trần Minh').locator('button', has_text='Từ chối').click(); pg.wait_for_timeout(300)
    row('Cô Phạm Thu Hà').locator('button', has_text='Duyệt').click(); pg.wait_for_timeout(300)
    row('Cô Phạm Thu Hà').locator('button', has_text='Khóa').click(); pg.wait_for_timeout(300)
    assert 'Đã khóa' in row('Cô Phạm Thu Hà').inner_text() and 'Từ chối' in row('Thầy Trần Minh').inner_text() and pg.inner_text('#admbadge') == '1'
    pg.wait_for_timeout(300); pg.screenshot(path='/workspace/v15_admin.png')
    row('Cô Đỗ Lan').locator('button', has_text='Xóa').click(); pg.wait_for_timeout(400); assert row('Cô Đỗ Lan').count() == 0 and not pg.is_visible('#admbadge')
    pg.click('.admf .btn'); assert not pg.is_visible('#adm'); pg.close()
    # --- đăng nhập tài khoản đã đăng ký ---
    for w, h, f, mob in ((1280, 800, 'v15_login_registered', 0), (390, 844, None, 1)):
        pg = page(w, h, bool(mob)); pg.select_option('#lt', 'registered'); assert pg.is_visible('#lph') and not pg.is_visible('#ln')
        pg.fill('#lph', '0912345678'); pg.fill('#lp', 'saimk'); pg.click('#lbtn'); pg.wait_for_timeout(400)
        assert 'Sai số điện thoại' in pg.inner_text('#lerr'); noscroll(pg, mob)
        pg.fill('#lp', 'matkhau1')
        if f: pg.fill('#lp', ''); pg.evaluate("$('lerr').textContent=''"); pg.fill('#lp', 'matkhau1'); pg.screenshot(path=f'/workspace/{f}.png')
        pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); assert pg.inner_text('#uname') == 'Cô Nguyễn Thị Hoa' and not pg.is_visible('#admbtn')
        pg.wait_for_timeout(300); pg.click('#userbox .btn-ghost'); pg.wait_for_selector('#login:not(.hide)')
        pg.select_option('#lt', 'registered'); pg.fill('#lph', '0356222333'); pg.fill('#lp', 'abcdef'); pg.click('#lbtn'); pg.wait_for_timeout(400)
        assert 'khóa' in pg.inner_text('#lerr'); pg.close()
    pg = page(390, 844, True); pg.select_option('#lt', 'Anh Lê Đạt'); pg.fill('#lp', '1993'); pg.click('#lbtn'); pg.wait_for_selector('#admbtn:not(.hide)'); pg.click('#admbtn'); pg.wait_for_selector('#admrows tr[data-id]'); pg.wait_for_timeout(300)
    r = pg.evaluate("[document.documentElement.scrollWidth,innerWidth,document.querySelector('.admbox').getBoundingClientRect().right]"); assert r[0] <= r[1] and r[2] <= r[1], r; pg.screenshot(path='/tmp/v15_admin_mobile.png'); pg.close()
    assert not E, E
print('UI v15 OK')
