from playwright.sync_api import sync_playwright
from PIL import Image
import sys
def run(p,args,shot=None):
    b=p.chromium.launch(args=args)
    c=b.new_context(viewport={'width':1280,'height':800},color_scheme='dark')
    c.add_init_script("localStorage.setItem('theme','dark')")
    pg=c.new_page(); pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)')
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)'); pg.select_option('#wsel','7'); pg.wait_for_timeout(800)
    th=pg.evaluate('document.documentElement.getAttribute("data-theme")')
    f=shot or '/tmp/fd.png'; pg.screenshot(path=f); b.close()
    im=Image.open(f).convert('RGB')
    # sample card / table body pixels
    pts=[(150,215),(200,300),(640,450),(900,700)]
    px=[im.getpixel(q) for q in pts]; ok=all(sum(v)>660 for v in px)
    print(args,'theme',th,'pixels',px,'WHITE' if ok else 'DARK'); return ok
with sync_playwright() as p:
    a=run(p,[])
    b=run(p,['--enable-features=WebContentsForceDark','--force-dark-mode'],'/workspace/v7_forced_dark_check.png')
    c=run(p,['--blink-settings=forceDarkModeEnabled=true','--enable-features=WebContentsForceDark'])
sys.exit(0 if a and b and c else 1)
