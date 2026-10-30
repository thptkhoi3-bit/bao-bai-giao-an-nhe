from playwright.sync_api import sync_playwright
def go(b,w,h,mob=False,theme=None):
    c=b.new_context(viewport={'width':w,'height':h},device_scale_factor=2 if mob else 1,is_mobile=mob)
    if theme: c.add_init_script(f"localStorage.setItem('theme_v2','{theme}')")
    pg=c.new_page(); pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)')
    if mob: pg.screenshot(path='/tmp/v5_login_dark_m.png')
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.select_option('#wsel','7'); pg.wait_for_timeout(700); return pg
sh=lambda pg:pg.evaluate('[document.documentElement.getAttribute("data-theme"),document.documentElement.scrollHeight,innerHeight]')
with sync_playwright() as p:
    b=p.chromium.launch()
    pg=go(b,1280,800); print('default',sh(pg)); pg.screenshot(path='/workspace/v5_light_1280.png')
    pg.click('#userbox .thb'); pg.wait_for_timeout(200); print('toggled',sh(pg)); pg.reload(); pg.wait_for_selector('#app:not(.hide)'); print('after reload',sh(pg))
    pg=go(b,1280,800,theme='dark'); pg.screenshot(path='/workspace/v5_dark_1280.png'); print('dark',sh(pg))
    m=go(b,390,844,True,'dark'); m.screenshot(path='/workspace/v5_dark_mobile.png'); print('mob',sh(m))
