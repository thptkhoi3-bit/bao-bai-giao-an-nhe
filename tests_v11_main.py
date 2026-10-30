from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1920,'height':1080}); pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)')
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.select_option('#wsel','8'); pg.dispatch_event('#wsel','change'); pg.wait_for_timeout(900)
    bl=pg.evaluate("[...document.querySelectorAll('*')].filter(e=>{const c=getComputedStyle(e);return /blur/.test(c.filter+c.backdropFilter)||c.maskImage!=='none'}).length")
    print('main 1920',pg.evaluate('[document.documentElement.scrollHeight,innerHeight]'),'blurred/masked elements:',bl); assert bl==0
    pg.screenshot(path='/workspace/v11_main_1920.png')
