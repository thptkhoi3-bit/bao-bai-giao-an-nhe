"""EXAM_RULES §6 – không nội dung nào chạm/tràn khung viền trang.
Khứ hồi DeMau.docx (Toán) + mẫu TV: 0 lỗi mọi trang; ca hình rất lớn tổng hợp: 0 lỗi; mẫu gốc (chưa sửa) bị bắt lỗi; vòng sửa hoạt động."""
import os, sys, io, re, json, shutil, tempfile, zipfile, copy
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import exam, pymupdf
from lxml import etree
from PIL import Image
TPL = {m: os.path.join(HERE, 'de_templates', m + '.docx') for m in ('toan', 'tv')}
F = dict(xa='Hoà Bình', truong='TH Kim Đồng', nam='2026 - 2027', lop=3, loai='CUỐI HỌC KÌ I', thoigian='40 phút')
TMP = tempfile.mkdtemp(prefix='tframe_')
HAS_LO = bool(exam.soffice_bin())
q = exam.q

def run(mon, src, name):
    w = os.path.join(TMP, name); os.makedirs(w, exist_ok=True)
    ex = exam.extract([('all', src)], mon, w)
    out = os.path.join(TMP, name + '.docx'); pd = os.path.join(TMP, name + '_pages')
    warns, _, pages, bad = exam.build_checked(TPL[mon], ex, F, out, [os.path.join(w, 'media'), os.path.join(w, 'figs')], True, pd if HAS_LO else None)
    return ex, out, pd, pages, bad

# ---- 1. khung trang theo sectPr: mép trong khung không lấn vùng chữ, lề giữ nguyên ----
for m in ('toan', 'tv'):
    x = etree.fromstring(zipfile.ZipFile(TPL[m]).read('word/document.xml')); s = x.find(q('body')).find(q('sectPr'))
    mar0 = {k: v for k, v in s.find(q('pgMar')).attrib.items()}
    ch = exam.fit_page_frame(s); assert ch, 'mẫu gốc có khung lấn vùng chữ -> phải sửa'
    for side in ('top', 'bottom', 'left', 'right'):
        b = s.find(q('pgBorders')).find(q(side)); sp, wd = exam.border_inner_pt(b)
        assert sp >= exam.FRAME_MIN_SPACE_PT and sp + wd + exam.FRAME_SAFE_PT <= int(s.find(q('pgMar')).get(q(side))) / 20 + 1e-6, (m, side)
    assert dict(s.find(q('pgMar')).attrib) == mar0, 'lề giữ nguyên (chỉ dời khung ra ngoài)'
print('1. khung trang theo sectPr OK')

if not HAS_LO:
    print('(không có LibreOffice – bỏ qua phần render)'); print('TẤT CẢ OK'); sys.exit(0)

# ---- 2. mẫu gốc (chưa sửa) bị bộ kiểm tra bắt lỗi: Câu 8 tràn khung đáy trang 1 + chữ đầu trang dính khung ----
pd0 = os.path.join(TMP, 'orig'); exam.render_pages(TPL['toan'], pd0)
bad0 = exam.check_frame(os.path.join(pd0, 'de.pdf'), exam.docx_frame_pt(TPL['toan']))
p2 = [v for v in bad0 if v['page'] == 2 and v['side'] == 'top' and v['kind'] == 'hình']
assert len(p2) >= 4, ('phải bắt được hình phương án Câu 8 của mẫu gốc thò lên khung đầu trang 2', bad0[:5])
assert any(v['page'] == 1 and v['side'] == 'top' and v['kind'] == 'chữ' for v in bad0), 'phải bắt được chữ tiêu đề dính khung (khung lấn lề)'
print(f'2. mẫu gốc: bắt {len(bad0)} lỗi (hình Câu 8 thò lên khung trang 2, chữ đầu trang dính khung) OK')

# ---- 3. khứ hồi DeMau.docx (Toán) + TV: 0 lỗi mọi trang ----
for mon, name in (('toan', 'rt_toan'), ('tv', 'rt_tv')):
    ex, out, pd, pages, bad = run(mon, TPL[mon], name)
    assert pages and bad == [], (mon, bad[:5])
    pdf = os.path.join(pd, 'de.pdf')
    again = exam.check_frame(pdf, exam.docx_frame_pt(out)); assert again == [], again[:3]
    if mon == 'tv': assert len(pages) == 7, len(pages)
    else:
        doc = pymupdf.open(pdf)
        # Câu 8: tiêu đề + hàng hình + phương án A–D trên cùng 1 trang, mọi hình của trang đó trong khung
        p8 = [i for i, p in enumerate(doc) if 'Câu 8.' in p.get_text() and 'ĐÁP ÁN' not in p.get_text()]; assert len(p8) == 1, p8
        pg = doc[p8[0]]; t8 = pg.search_for('Câu 8.')[0]; t9 = (pg.search_for('II. TỰ LUẬN') or pg.search_for('Câu 9.'))[0]
        fr, border = exam.frame_rect(pg, exam.docx_frame_pt(out))
        shapes = [d['rect'] for d in pg.get_drawings() if d not in border and t8.y1 < d['rect'].y0 < t9.y0]
        assert len(shapes) >= 8, len(shapes)
        assert all(fr[1] + 2 <= r.y0 and r.y1 <= fr[3] - 2 for r in shapes)
        lab = [pg.search_for(f'{c}.') for c in 'ABCD']; assert all(any(t8.y1 < r.y0 < t9.y0 for r in L) for L in lab), 'phương án A–D cùng trang với Câu 8'
        # câu có hình không bị tách: tiêu đề Câu 12 cùng trang với hình
        p12 = [i for i, p in enumerate(doc) if 'Câu 12.' in p.get_text() and 'ĐÁP ÁN' not in p.get_text()][0]
        t12 = doc[p12].search_for('Câu 12.')[0]
        assert any(d['rect'].y0 > t12.y1 and d['rect'].height > 50 for d in doc[p12].get_drawings()), 'hình Câu 12 phải cùng trang tiêu đề'
        assert len(pages) == 4, len(pages)
        if os.environ.get('V17_SHOTS'):
            for i, p in enumerate(pages[:3]): shutil.copy(p, f'/workspace/v17_toan_p{i + 1}.png')
print('3. khứ hồi Toán (4 trang) + TV (7 trang): 0 lỗi, Câu 8 trọn 1 trang, Câu 12 không tách OK')

# ---- 4. ca hình rất lớn tổng hợp ----
import docx
from docx.shared import Cm
big = os.path.join(TMP, 'big.png'); Image.new('RGB', (1500, 2000), (60, 120, 200)).save(big, dpi=(72, 72))
wide = os.path.join(TMP, 'wide.png'); Image.new('RGB', (1600, 600), (220, 80, 80)).save(wide)
d = docx.Document(); d.add_paragraph('I. TRẮC NGHIỆM')
d.add_paragraph('Câu 1. Quan sát hình rồi chọn đáp án đúng: (M1) (0,5đ)')
d.add_paragraph().add_run().add_picture(big, width=Cm(30))                      # ảnh cùng dòng 30 cm (vượt bề rộng & chiều cao)
d.add_paragraph('A. 1\tB. 2\tC. 3\tD. 4')
d.add_paragraph('Câu 2. Hình nào lớn nhất? (M1) (0,5đ)')
pa = d.add_paragraph(); pa.add_run().add_picture(wide, width=Cm(14))
inl = pa._p.find('.//' + exam.WPD + 'inline'); WPN = exam.WPD[1:-1]
anc = etree.SubElement(etree.Element('x'), exam.WPD + 'anchor', {k: '0' for k in ('distT', 'distB', 'distL', 'distR', 'simplePos', 'locked')} | {'relativeHeight': '1', 'behindDoc': '0', 'layoutInCell': '1', 'allowOverlap': '1'})
etree.SubElement(anc, exam.WPD + 'simplePos', x='0', y='0')
for tag, rel, off in (('positionH', 'page', 9 * 360000), ('positionV', 'page', 25 * 360000)):   # neo theo trang, lệch phải, sát đáy
    e = etree.SubElement(anc, exam.WPD + tag, relativeFrom=rel); etree.SubElement(e, exam.WPD + 'posOffset').text = str(off)
for c in inl:
    if c.tag == exam.WPD + 'docPr': etree.SubElement(anc, exam.WPD + 'wrapNone')
    if c.tag != exam.WPD + 'effectExtent': anc.append(copy.deepcopy(c))
inl.getparent().replace(inl, anc)
d.add_paragraph('A. Hình 1\tB. Hình 2\tC. Hình 3\tD. Hình 4')
d.add_paragraph('Câu 3. Điền số vào bảng: (M2) (1đ)')
t = d.add_table(rows=2, cols=3); t.style = 'Table Grid'
tp = t._tbl.tblPr; tw = tp.find(q('tblW'))
if tw is None: tw = etree.SubElement(tp, q('tblW'))
tw.set(q('w'), '18000'); tw.set(q('type'), 'dxa')
for g in t._tbl.find(q('tblGrid')).findall(q('gridCol')): g.set(q('w'), '6000')
for tc in t._tbl.iter(q('tc')):
    pr = tc.get_or_add_tcPr(); w_ = pr.find(q('tcW'))
    if w_ is None: w_ = etree.SubElement(pr, q('tcW'))
    w_.set(q('w'), '6000'); w_.set(q('type'), 'dxa')
for i, c in enumerate(t._cells): c.text = f'ô số {i + 1} – nội dung dài để bảng rộng'
t.cell(1, 0).paragraphs[0].add_run().add_picture(wide, width=Cm(3))   # bảng có hình -> được giữ nguyên khối (raw)
d.add_paragraph('.' * 200)
src = os.path.join(TMP, 'big_src.docx'); d.save(src)
ex, out, pd, pages, bad = run('toan', src, 'big')
assert exam.count_questions(ex) == 3, exam.count_questions(ex)
assert [b['t'] for b in ex['sections'][0]['questions'][2]['blocks']][0] == 'raw'
assert pages and bad == [], bad[:5]
z = etree.fromstring(zipfile.ZipFile(out).read('word/document.xml')); body = z.find(q('body')); sect = body.find(q('sectPr'))
Wemu = exam.text_width_tw(sect) * 635; Hemu = exam.text_height_tw(sect) * 635
exts = [(int(e.get('cx')), int(e.get('cy'))) for e in z.iter(exam.WPD + 'extent')]
assert exts and all(cx <= Wemu and cy <= Hemu * exam.FIG_MAX_H + 1 for cx, cy in exts), exts
for a in z.iter(exam.WPD + 'anchor'):
    assert a.find(exam.WPD + 'positionV').get('relativeFrom') == 'paragraph' and a.find(exam.WPD + 'positionH').get('relativeFrom') in ('column', 'character')
tbl = [t for t in body.findall(q('tbl')) if 'ô số 1' in ''.join(x.text or '' for x in t.iter(q('t')))][0]
assert sum(int(g.get(q('w'))) for g in tbl.find(q('tblGrid')).findall(q('gridCol'))) <= exam.text_width_tw(sect)
print(f'4. ca hình rất lớn (ảnh 30 cm, hình nổi theo trang, bảng 18000tw): {len(pages)} trang, 0 lỗi OK')

# ---- 5. vòng sửa: lỗi giả ở đáy trang -> ngắt trang trước câu; lỗi không sửa được -> trả về (không giao im lặng) ----
calls = []
real = exam.check_frame
def fake(pdf, sf=None, tol=2.0):
    calls.append(1); sp = exam.question_spans(pdf); s2 = [s for s in sp if s[2] == 2][0]
    return [dict(page=s2[0], kind='hình', side='bottom', bbox=[60, s2[1] + 5, 100, s2[1] + 30], gap=-3)]
exam.check_frame = fake
try:
    ex2 = json.load(open(os.path.join(TMP, 'big', 'exam.json'))) if os.path.exists(os.path.join(TMP, 'big', 'exam.json')) else ex
    _, _, pages5, bad5 = exam.build_checked(TPL['toan'], ex, F, os.path.join(TMP, 'loop.docx'), [os.path.join(TMP, 'big', 'media')], True, os.path.join(TMP, 'loop_pages'))
finally: exam.check_frame = real
assert bad5 and len(calls) >= 3, (len(calls), bad5)
x5 = zipfile.ZipFile(os.path.join(TMP, 'loop.docx')).read('word/document.xml').decode()
assert 'pageBreakBefore' in x5, 'vòng 1 phải chèn ngắt trang trước Câu 2'
print(f'5. vòng sửa: {len(calls)} vòng, ngắt trang + co hình, lỗi còn lại được trả về OK')

# ---- 6. bảng thường (không hình) trong câu phải được giữ, vừa bề rộng chữ, trong khung ----
d = docx.Document(); d.add_paragraph('I. TRẮC NGHIỆM')
d.add_paragraph('Câu 1. Bảng sau cho biết số quyển sách mỗi lớp quyên góp: (M2) (1đ)')
t = d.add_table(rows=3, cols=5); t.style = 'Table Grid'
for g in t._tbl.find(q('tblGrid')).findall(q('gridCol')): g.set(q('w'), '3000')     # 15000tw > bề rộng chữ
for tc in t._tbl.iter(q('tc')):
    pr = tc.get_or_add_tcPr(); w_ = pr.find(q('tcW'))
    if w_ is None: w_ = etree.SubElement(pr, q('tcW'))
    w_.set(q('w'), '3000'); w_.set(q('type'), 'dxa')
for j, v in enumerate(['Lớp', '3A', '3B', '3C', '3D']): t.cell(0, j).text = v
for j, v in enumerate(['Số sách', '45', '38', '52', '41']): t.cell(1, j).text = v
for j, v in enumerate(['Đúng/Sai', 'Đ', 'S', '', '']): t.cell(2, j).text = v
d.add_paragraph('Lớp nào quyên góp nhiều nhất?')
d.add_paragraph('A. 3A\tB. 3B\tC. 3C\tD. 3D')
d.add_paragraph('Câu 2. Tính: 12 + 7 = ? (M1) (0,5đ)')
d.add_paragraph('A. 19\tB. 18\tC. 17\tD. 20')
src = os.path.join(TMP, 'tbl_src.docx'); d.save(src)
ex, out, pd, pages, bad = run('toan', src, 'tbl')
b1 = [b['t'] for b in ex['sections'][0]['questions'][0]['blocks']]
assert b1[0] == 'raw' and 'options' in b1, b1
assert pages and bad == [], bad[:5]
z = etree.fromstring(zipfile.ZipFile(out).read('word/document.xml')); body = z.find(q('body')); sect = body.find(q('sectPr'))
tb = [t for t in body.findall(q('tbl')) if 'Số sách' in ''.join(x.text or '' for x in t.iter(q('t')))]
assert len(tb) == 1, 'bảng số liệu phải có trong đề'
txt = ''.join(x.text or '' for x in tb[0].iter(q('t')))
assert all(s in txt for s in ('3A', '3D', '45', '52', 'Đúng/Sai')), txt
assert sum(int(g.get(q('w'))) for g in tb[0].find(q('tblGrid')).findall(q('gridCol'))) <= exam.text_width_tw(sect)
pdft = ''.join(p.get_text() for p in pymupdf.open(os.path.join(pd, 'de.pdf')))
assert 'Số sách' in pdft and '52' in pdft
print('6. bảng số liệu không hình trong câu: giữ nguyên, vừa bề rộng chữ, 0 lỗi khung OK')
shutil.rmtree(TMP, ignore_errors=True) if not os.environ.get("KEEP") else None
print('TẤT CẢ OK')
