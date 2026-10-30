from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b=p.chromium.launch()
    for w,h,f,mob in ((1280,800,'v14_login_1280',0),(1920,1080,'v14_login_1920',0),(1440,900,None,0),(1366,768,"v14_login_1366",0),(390,844,'v14_login_mobile',1)):
        pg=b.new_page(viewport={'width':w,'height':h},device_scale_factor=2 if mob else 1,is_mobile=bool(mob)); e=[];pg.on('pageerror',lambda x:e.append(str(x)))
        pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)'); pg.wait_for_timeout(600)
        r=pg.evaluate('[document.documentElement.scrollHeight,innerHeight,document.documentElement.scrollWidth,innerWidth]')
        print(w,h,r,e); assert r[2]<=r[3] and (mob or r[0]<=r[1])
        if f: pg.screenshot(path=f'/workspace/{f}.png',full_page=bool(mob))
        pg.select_option('#lt','other'); assert pg.is_visible('#ln'); pg.select_option('#lt','Anh Lê Đạt'); assert not pg.is_visible('#ln')
        pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.close()
    print('LOGIN v14 OK')
