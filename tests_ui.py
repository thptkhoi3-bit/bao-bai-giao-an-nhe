from playwright.sync_api import sync_playwright
import docx
from docx.oxml.ns import qn
A='/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
tx=lambda d:[''.join(x.text or '' for x in e.iter(qn('w:t'))) for e in docx.Document(d).element.body]
def run(w,h,mobile,shots):
  with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':w,'height':h},device_scale_factor=2 if mobile else 1,is_mobile=mobile,accept_downloads=True); pg=ctx.new_page()
    errs=[]; pg.on('pageerror',lambda e:errs.append(str(e)))
    pg.goto('http://localhost:8000'); pg.wait_for_selector('#login:not(.hide)'); pg.wait_for_timeout(400)
    if 'login' in shots: pg.select_option('#lt','Anh Lê Đạt'); pg.screenshot(path=shots['login'])
    pg.select_option('#lt','Anh Lê Đạt'); pg.fill('#lp','9999'); pg.click('#lbtn'); pg.wait_for_function("document.getElementById('lerr').textContent.length>0")
    print(w,'wrong pw msg:',pg.inner_text('#lerr'))
    pg.fill('#lp','1993'); pg.click('#lbtn'); pg.wait_for_selector('#app:not(.hide)')
    pg.set_input_files('#mfile',A+'5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx'); pg.evaluate("document.getElementById('mmore').open=true")
    pg.click('#mmore .btn'); pg.wait_for_function("document.getElementById('toast').textContent.includes('thay')")
    pg.set_input_files('#bbfile',A+'455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx')
    assert pg.is_disabled('#bbgo') and pg.is_visible('#bbhint'), 'phải chọn tuần trước'
    pg.wait_for_selector('#bbweek option[value="7"]', state='attached'); pg.select_option('#bbweek','7'); assert pg.is_enabled('#bbgo')
    pg.click('#bbgo'); pg.wait_for_selector('#bbres .dlbox',timeout=60000)
    print(w,'stats:',pg.inner_text('#bbres .stats').replace('\n',' '))
    with pg.expect_download() as d: pg.click('#bbres .dlbox a')
    d.value.save_as('/tmp/ui_bb.docx')
    print(w,'baobai==manual',tx('/tmp/ui_bb.docx')==tx('/cursor/stores/store-0e61161c-f9bc-45cb-8d83-359add90b33e/baobai/BaoBai_Goc_moinhat.docx'))
    if 'baobai' in shots: pg.evaluate("scrollTo(0,0)"); pg.screenshot(path=shots['baobai'],full_page=True)
    pg.click('#t2'); pg.select_option('#gweek','7'); pg.set_input_files('#gfiles',[A+'5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'])
    pg.wait_for_timeout(900); print(pg.evaluate("""()=>{const g=document.getElementById('ggo');g.scrollIntoView({block:'center'});const r=g.getBoundingClientRect();const e=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);return [scrollY,r.y,e&&e.outerHTML.slice(0,90)]}""")); pg.screenshot(path='/tmp/dbg.png'); pg.click('#ggo',timeout=5000); pg.wait_for_selector('#gres .dlbox',timeout=90000)
    print(w,'ga stats:',pg.inner_text('#gres .stats').replace('\n',' '))
    with pg.expect_download() as d: pg.click('#gres .dlbox a')
    d.value.save_as('/tmp/ui_ga.docx'); print(w,'giaoan==manual',tx('/tmp/ui_ga.docx')==tx('/workspace/GiaoAn_Tuan7_HoanChinh_v2.docx'))
    if 'giaoan' in shots: pg.evaluate("scrollTo(0,0)"); pg.screenshot(path=shots['giaoan'],full_page=True)
    if 'desk' in shots: pg.click('#t1'); pg.wait_for_timeout(300); pg.screenshot(path=shots['desk'])
    pg.click('#userbox .btn-ghost'); pg.wait_for_selector('#login:not(.hide)'); print(w,'logout ok; js errors:',errs); b.close()
run(390,844,True,{'login':'/workspace/ui_login.png','baobai':'/workspace/ui_baobai.png','giaoan':'/workspace/ui_giaoan.png'})
run(1280,860,False,{'desk':'/workspace/ui_desktop.png'})
