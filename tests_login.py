from playwright.sync_api import sync_playwright
J='''()=>{const s=document.getElementById('lt'),r=s.getBoundingClientRect(),c=getComputedStyle(s),o=getComputedStyle(s.options[1]);
const e=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);
return {vis:r.width>0&&r.height>0&&r.top>=0&&r.bottom<=innerHeight,top:Math.round(r.top),h:Math.round(r.height),opts:[...s.options].map(x=>x.text).length,
 col:c.color,bg:c.backgroundColor,optc:o.color,optbg:o.backgroundColor,hit:e===s||s.contains(e)||e.closest('.field')!==null,label:document.querySelector('label[for=lt]').textContent}}'''
with sync_playwright() as p:
  for bn in ('chromium','firefox','webkit'):
    try: b=getattr(p,bn).launch()
    except Exception as ex: print(bn,'unavailable',str(ex)[:80]); continue
    for w,h in ((1280,720),(1366,657),(1536,730),(1920,1080),(390,844)):
      for th in ('light','dark'):
        c=b.new_context(viewport={'width':w,'height':h}); c.add_init_script(f"localStorage.setItem('theme','{th}')")
        pg=c.new_page(); pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)'); pg.wait_for_timeout(300)
        r=pg.evaluate(J); assert r['opts']==7 and r['hit'], r
        pg.select_option('#lt','Cô Linh'); assert pg.input_value('#lt')=='Cô Linh'
        print(bn,w,h,th,r)
        if bn=='chromium' and th=='light' and w in(1280,390): pg.select_option('#lt',''); pg.screenshot(path=f'/workspace/fix_login_{"1280" if w==1280 else "mobile"}.png')
        c.close()
    b.close()
