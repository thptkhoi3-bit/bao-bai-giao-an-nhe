from playwright.sync_api import sync_playwright
import sys
A='/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
def go(b,w,h,mob=False):
    pg=b.new_page(viewport={'width':w,'height':h},device_scale_factor=2 if mob else 1,is_mobile=mob); errs=[];pg.on('pageerror',lambda e:errs.append(str(e)))
    pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)')
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.select_option('#wsel','8'); pg.dispatch_event('#wsel','change'); pg.wait_for_timeout(900); pg.errs=errs; return pg
M='[document.documentElement.scrollHeight,innerHeight,document.querySelector("#hdr").offsetHeight,document.querySelector("main").offsetHeight,(()=>{const l=document.querySelector("#p1 .grid>div, #p1>div");return [l.scrollHeight,l.clientHeight]})()]'
with sync_playwright() as p:
    b=p.chromium.launch()
    for w,h,f in ((1280,800,'v9_baobai_1280'),(1920,1080,'v9_baobai_1920'),(1440,900,None)):
        pg=go(b,w,h); print(w,h,pg.evaluate(M),pg.errs)
        if f: pg.screenshot(path=f'/workspace/{f}.png')
        pg.close()
    m=go(b,390,844,True); m.screenshot(path='/workspace/v9_mobile.png'); print('mob',m.evaluate('[document.documentElement.scrollWidth,innerWidth]'))
    pg=go(b,1280,800)
    pg.click('#pv td.ed >> nth=0'); pg.wait_for_selector('#edm:not(.hide)'); pg.fill('#edten','Thử nghiệm'); pg.screenshot(path='/workspace/v9_edit_modal.png'); pg.click('#edm .btn.light')
    pg.click('#t2'); pg.select_option('#gweek','7'); pg.set_input_files('#gfiles',[A+'5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx']); pg.click('#ggo'); pg.wait_for_selector('#gres .dlbox',timeout=90000); pg.wait_for_timeout(4500)
    print('giaoan',pg.evaluate('[document.documentElement.scrollHeight,innerHeight]')); pg.screenshot(path='/workspace/v9_giaoan_1280.png')
