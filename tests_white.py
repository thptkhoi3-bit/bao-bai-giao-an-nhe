from playwright.sync_api import sync_playwright
A='/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1280,'height':800}); pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)')
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.select_option('#wsel','7'); pg.wait_for_timeout(700)
    dark=pg.evaluate('''()=>[...document.querySelectorAll('main *')].filter(e=>e.offsetParent).map(e=>getComputedStyle(e).backgroundColor).filter(c=>{const m=c.match(/\\d+/g);return m&&(m[3]===undefined||+m[3]>0)&&(+m[0]+ +m[1]+ +m[2])<200}).length''')
    print('dark surfaces',dark, pg.evaluate('[document.documentElement.scrollHeight,innerHeight]'))
    pg.screenshot(path='/workspace/v8_ivory_baobai.png')
    pg.click('#t2'); pg.select_option('#gweek','7'); pg.set_input_files('#gfiles',[A+'5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx']); pg.click('#ggo'); pg.wait_for_selector('#gres .dlbox',timeout=90000); pg.wait_for_timeout(3500)
    print(pg.evaluate('[document.documentElement.scrollHeight,innerHeight]')); pg.screenshot(path='/workspace/v8_ivory_giaoan.png')
