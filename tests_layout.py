from playwright.sync_api import sync_playwright
A='/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
def go(b,w,h,mob=False):
    pg=b.new_page(viewport={'width':w,'height':h},device_scale_factor=2 if mob else 1,is_mobile=mob); pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)')
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.wait_for_timeout(700); return pg
sh=lambda pg:pg.evaluate('[document.documentElement.scrollHeight,innerHeight,document.documentElement.scrollWidth,innerWidth]')
with sync_playwright() as p:
    b=p.chromium.launch()
    for w,h,f in ((1280,800,'v4_baobai_1280'),(1920,1080,'v4_baobai_1920'),(1440,900,None)):
        pg=go(b,w,h); pg.select_option('#wsel','7'); pg.wait_for_timeout(400); print(w,'baobai',sh(pg))
        pg.set_input_files('#bbfile',A+'455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx'); pg.wait_for_timeout(200); print(w,'baobai+file',sh(pg))
        if f: pg.screenshot(path=f'/workspace/{f}.png')
        pg.click('#t2'); pg.select_option('#gweek','7'); pg.set_input_files('#gfiles',[A+'5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx']); pg.click('#ggo'); pg.wait_for_selector('#gres .dlbox',timeout=90000); pg.wait_for_timeout(500)
        print(w,'giaoan result',sh(pg))
        if w==1440: pg.screenshot(path='/workspace/v4_giaoan_1440.png')
        pg.close()
    m=go(b,390,844,True); m.screenshot(path='/workspace/v4_mobile.png'); print('mobile',sh(m)); b.close()
