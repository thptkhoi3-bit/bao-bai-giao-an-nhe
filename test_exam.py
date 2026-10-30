"""v16 – Tạo đề kiểm tra: khứ hồi Toán/TV từ đề mẫu, PDF chữ, Gemini giả lập (vẽ lại SVG + cắt ảnh), API, giới hạn lượt AI."""
import os, sys, io, re, json, zipfile, tempfile, shutil, subprocess, threading, logging, copy, time
from http.server import BaseHTTPRequestHandler, HTTPServer
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
for k in ('GEMINI_API_KEY', 'GEMINI_API_URL', 'DATA_REPO', 'DATA_TOKEN'): os.environ.pop(k, None)
import exam, storage
from lxml import etree
q = exam.q
TPL = {m: os.path.join(HERE, 'de_templates', m + '.docx') for m in ('toan', 'tv')}
FIELDS = dict(xa='Hoà Bình', truong='TH Kim Đồng', nam='2026 - 2027', loai='Cuối học kì I', lop=4, thoigian='35 phút')
LOGS = io.StringIO(); h = logging.StreamHandler(LOGS); h.setLevel(logging.DEBUG); logging.getLogger().addHandler(h); logging.getLogger().setLevel(logging.DEBUG)

def doc_xml(p): return etree.fromstring(zipfile.ZipFile(p).read('word/document.xml'))
def all_text(p):
    z = zipfile.ZipFile(p); out = []
    for n in z.namelist():
        if n.startswith('word/') and n.endswith('.xml'):
            r = etree.fromstring(z.read(n)); out += [exam.ptext(x) for x in r.iter(q('p'))]
    return '\n'.join(out)
def body(p): return [e for e in doc_xml(p).find(q('body')) if e.tag in (q('p'), q('tbl'))]
def pages(p):
    if not exam.soffice_bin(): return None
    d = tempfile.mkdtemp(); exam.render_pages(p, d, dpi=30); return len([f for f in os.listdir(d) if f.endswith('.png')])
def hdr_tables(p): return [e for e in body(p) if e.tag == q('tbl') and 'UBND' in exam.cell_text(e.find(q('tr')).find(q('tc')))]
def strip_text(e):
    e = copy.deepcopy(e)
    for t in e.iter(q('t')): t.text = ''; t.attrib.pop(exam.XS, None)
    return etree.tostring(e)

# ---------- 1. Điền đầu đề: chỉ đổi chữ trong run có sẵn ----------
for m in ('toan', 'tv'):
    root = doc_xml(TPL[m]); before = strip_text(root); n_p = len(list(root.iter(q('p')))); n_r = len(list(root.iter(q('r'))))
    log = exam.fill_header(root, FIELDS)
    assert log, m
    assert strip_text(root) == before, f'{m}: điền đầu đề làm đổi cấu trúc'
    assert len(list(root.iter(q('p')))) == n_p and len(list(root.iter(q('r')))) == n_r
    txt = ''.join(t.text or '' for t in root.iter(q('t')))
    assert 'PHẬT TÍCH' not in txt and 'HOÀ BÌNH' in txt and 'KIM ĐỒNG' in txt and 'CUỐI HỌC KÌ I' in txt, m
print('1. điền đầu đề OK')

# ---------- 2. Khứ hồi Toán ----------
wd = tempfile.mkdtemp(); ex = exam.extract([('all', shutil.copy(TPL['toan'], os.path.join(wd, 'nguon.docx')))], 'toan', wd)
qs = [Q for s in ex['sections'] for Q in s['questions']]
assert [Q['so'] for Q in qs] == list(range(1, 13)), [Q['so'] for Q in qs]
assert [s['title'] for s in ex['sections']] == ['I. TRẮC NGHIỆM', 'II. TỰ LUẬN']
assert all(Q['muc'] and Q['diem'] for Q in qs), 'mức + điểm'
assert qs[0]['blocks'][1] == dict(t='options', items=['653', '356', '635', '365'], layout='row')
assert qs[1]['blocks'] == [dict(t='options', items=['100', '98', '97', '101'], layout='row')] and qs[1]['diem'] == '0,25'
assert any(b['t'] == 'raw' for b in qs[7]['blocks']), 'Câu 8: hình giữ nguyên từ Word'
assert any(b['t'] == 'lines' for b in qs[10]['blocks'])
assert ex['answers']['2']['text'] == 'B' and ex['answers']['12']['points'] == '1'
out = os.path.join(wd, 'de.docx'); warns, hl = exam.build(TPL['toan'], ex, FIELDS, out, [os.path.join(wd, 'media')])
assert not warns, warns
import docx; d = docx.Document(out)
ps = [p.text for p in d.paragraphs]
titles = [t for t in ps if re.match(r'^Câu \d+\.', t)]
assert len(titles) == 12, titles
src_titles = [p.text for p in docx.Document(TPL['toan']).paragraphs if re.match(r'^Câu \d+', p.text)]
norm = lambda s: re.sub(r'\s+', '', s.replace('M 1', 'M1').replace('M 2', 'M2').replace('M1)(', 'M1) ('))
for a, b in zip(titles, src_titles): assert norm(a).replace(')(', ') (') .replace(' ', '') == norm(b).replace(' ', '') or a.split(' ', 2)[2][:15] in b, (a, b)
t = all_text(out)
assert 'PHẬT TÍCH' not in t and 'PHẬT TÍCH' not in ' '.join(ps).upper()
for s in ('UBND XÃ HOÀ BÌNH', 'KIM ĐỒNG', 'CUỐI HỌC KÌ I', 'NĂM HỌC 2026 - 2027', 'Lớp 4', '(Bài kiểm tra có 02 trang)', 'I. TRẮC NGHIỆM', 'II. TỰ LUẬN', 'ĐÁP ÁN', '35 phút'):
    assert s in t, s
# bảng tiêu đề: giống hệt mẫu đã điền (không đổi định dạng)
ref = doc_xml(TPL['toan']); exam.fill_header(ref, FIELDS)
assert etree.tostring(hdr_tables(out)[0], method='c14n') == etree.tostring([e for e in ref.find(q('body')) if e.tag == q('tbl')][0], method='c14n'), 'bảng tiêu đề phải giữ nguyên'
# bảng đáp án: 12 câu + dòng II. Tự luận
ans = [e for e in body(out) if e.tag == q('tbl')][-1]
rows = [exam.cell_text(tr.findall(q('tc'))[0]) for tr in ans.findall(q('tr'))]
assert sum(1 for r in rows if re.search(r'Câu \d+\.', r)) == 12 and any(r.startswith('II.') for r in rows), rows
# ngắt trang trước ĐÁP ÁN, không trang trống
bl = body(out); ia = next(i for i, e in enumerate(bl) if exam.ptext(e).strip() == 'ĐÁP ÁN' if e.tag == q('p'))
assert exam.has_pagebreak(bl[ia - 1]) or exam.has_pagebreak(bl[ia])
# phần khác của gói giữ nguyên (styles, numbering, header/footer)
zt, zo = zipfile.ZipFile(TPL['toan']), zipfile.ZipFile(out)
c14 = lambda b: etree.tostring(etree.fromstring(b), method='c14n')
for n in zt.namelist():
    if n.startswith('word/') and n.endswith('.xml') and n not in ('word/document.xml',) and n in zo.namelist():
        assert c14(zt.read(n)) == c14(zo.read(n)), f'{n} bị thay đổi'
    elif not n.endswith('.xml') and not n.endswith('.rels') and n in zo.namelist():
        assert zt.read(n) == zo.read(n), n
np_ = pages(out)
if np_ is not None: assert pages(TPL['toan']) == 3 and np_ == 4, np_   # §6: Câu 8 không còn tràn khung -> Câu 12 sang trang 3
print('2. khứ hồi Toán OK (12 câu, trang =', np_, ')')

# không kèm đáp án
out2 = os.path.join(wd, 'de2.docx'); exam.build(TPL['toan'], ex, FIELDS, out2, [os.path.join(wd, 'media')], with_answers=False)
assert 'ĐÁP ÁN' not in all_text(out2)
if exam.soffice_bin(): assert pages(out2) == 3   # §6: Câu 12 sang trang 3

# ---------- 3. Khứ hồi Tiếng Việt ----------
wd = tempfile.mkdtemp(); ex = exam.extract([('all', shutil.copy(TPL['tv'], os.path.join(wd, 'nguon.docx')))], 'tv', wd)
assert exam.count_questions(ex) == 9
assert ex['reading']['title'] == 'PHẦN THƯỞNG' and len(ex['reading']['paragraphs']) == 7 and ex['reading']['source'] == ['Phỏng theo BLAI-TƠN', '(Lương Hùng dịch)']
assert [t['label'] for t in ex['writing']['topics']] == ['Đề 1', 'Đề 2'] and ex['dictation']['title'] == 'Cô giáo tí hon' and ex['dictation']['source'] == 'Theo Nguyễn Thi'
assert ex['answers']['1']['text'] == 'C' and ex['answers']['8']['text'] == 'B'
out = os.path.join(wd, 'de.docx'); warns, hl = exam.build(TPL['tv'], ex, FIELDS, out, [os.path.join(wd, 'media')])
assert not warns, warns
t = all_text(out); txt = ' '.join(p.text for p in docx.Document(out).paragraphs)
assert 'PHẬT TÍCH' not in t
assert txt.count('TRƯỜNG TH & THCS KIM ĐỒNG') >= 1 and t.count('KIM ĐỒNG') >= 3, 'mọi phần đầu đề đều điền'
for s in ('PHẦN THƯỞNG', 'Cô giáo tí hon', 'Đề 1:', 'Đề 2:', 'HƯỚNG DẪN CHẤM', '2026-2027', '(Bài kiểm tra có 02 trang)'): assert s in t, s
cells = ' '.join(exam.cell_text(tc) for tb in [e for e in body(out) if e.tag == q('tbl')] for tc in tb.iter(q('tc')))
for n, a in (('1', 'C'), ('2', 'A'), ('8', 'B')): assert re.search(rf'Câu {n}: {a}\b(?! {a}\b)', cells), (n, cells[-600:])
assert len([x for x in re.findall(r'^Câu \d+\.', '\n'.join(p.text for p in docx.Document(out).paragraphs), re.M)]) == 9
hd = hdr_tables(out); assert len(hd) == 2
prev = hd[1].getprevious(); assert exam.has_pagebreak(prev) and not exam.ptext(prev).strip(), 'Bài số 2 bắt đầu trang mới'
assert not exam.is_empty_p(prev.getprevious()), 'không còn đoạn trống tràn'
np_ = pages(out)
if np_ is not None: assert np_ == 7, np_
print('3. khứ hồi Tiếng Việt OK (9 câu, trang =', np_, ')')

# ---------- 4. PDF chữ (tạo từ docx bằng LibreOffice) ----------
if exam.soffice_bin():
    pd = tempfile.mkdtemp()
    subprocess.run(['soffice', f'-env:UserInstallation=file://{pd}/prof', '--headless', '--convert-to', 'pdf', '--outdir', pd, TPL['toan'], TPL['tv']], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    wd = tempfile.mkdtemp(); ex = exam.extract([('all', os.path.join(pd, 'toan.pdf'))], 'toan', wd)
    qs = [Q for s in ex['sections'] for Q in s['questions']]
    assert [Q['so'] for Q in qs] == list(range(1, 13)) and [s['title'] for s in ex['sections']] == ['I. TRẮC NGHIỆM', 'II. TỰ LUẬN']
    assert qs[1]['blocks'][0]['items'] == ['100', '98', '97', '101'] and qs[10]['text'].endswith('mảnh vải?')
    assert any('Câu 8' in w for w in ex['warnings']), 'phương án hình trong PDF phải được cảnh báo'
    out = os.path.join(wd, 'de.docx'); exam.build(TPL['toan'], ex, FIELDS, out, [os.path.join(wd, 'figs')]); assert len([p for p in docx.Document(out).paragraphs if re.match(r'^Câu \d+\.', p.text)]) == 12
    wd = tempfile.mkdtemp(); ex = exam.extract([('all', os.path.join(pd, 'tv.pdf'))], 'tv', wd)
    assert exam.count_questions(ex) == 9 and ex['reading']['title'] == 'PHẦN THƯỞNG' and len(ex['writing']['topics']) == 2 and ex['dictation']['title'] == 'Cô giáo tí hon'
    print('4. PDF chữ OK')
else: print('4. (bỏ qua – không có LibreOffice)')

# ---------- 5. Ảnh khi chưa có key ----------
from PIL import Image, ImageDraw
wd = tempfile.mkdtemp(); img = os.path.join(wd, 'anh.png'); im = Image.new('RGB', (1000, 1400), 'white'); dr = ImageDraw.Draw(im)
dr.polygon([(200, 900), (500, 600), (800, 900)], outline='black', width=4); dr.rectangle((600, 1000, 900, 1300), fill=(220, 60, 60)); im.save(img)
try: exam.extract([('all', img)], 'toan', wd); raise SystemExit('phải báo chưa bật đọc ảnh')
except exam.NeedAI as e: assert str(e) == 'Chưa bật đọc ảnh – liên hệ 0865291703'
print('5. thiếu key -> thông báo OK')

# ---------- 6. Gemini giả lập: vẽ lại hình học (SVG) + cắt tranh ----------
KEY = 'AIzaFAKE-test-key-123456'
SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 60 40"><polygon points="5,35 30,5 55,35" fill="none" stroke="black" stroke-width="0.4"/><text x="2" y="39" font-size="4" font-family="Times New Roman">A</text></svg>'
RESP = {'sections': [{'title': 'I. TRẮC NGHIỆM', 'instruction': 'Khoanh vào chữ cái đặt trước câu trả lời đúng', 'questions': [
            {'so': 1, 'text': 'Hình bên có mấy góc?', 'muc': 'M1', 'diem': '0,5', 'blocks': [{'t': 'figure', 'id': 'f1'}, {'t': 'options', 'items': ['1', '2', '3', '4'], 'layout': 'row'}]},
            {'so': 2, 'text': 'Hình nào được tô màu?', 'muc': 'M2', 'diem': '1', 'blocks': [{'t': 'figure', 'id': 'f2'}, {'t': 'ds', 'items': ['a) Hình vuông', 'b) Hình tròn']}]}]},
        {'title': 'II. TỰ LUẬN', 'questions': [{'so': 3, 'text': 'Tính 25 + 17', 'muc': 'M1', 'diem': '1', 'blocks': [{'t': 'lines', 'n': 2}]}]}],
        'answers_list': [{'so': 1, 'text': 'C', 'points': '0,5'}],
        'figures': [{'id': 'f1', 'kind': 'geometry', 'page': 0, 'box': [420, 180, 660, 820], 'svg': SVG, 'w_cm': 5},
                    {'id': 'f2', 'kind': 'picture', 'page': 0, 'box': [700, 590, 940, 910]}]}
SEEN = []; MODE = {'code': 200}
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        body_ = json.loads(s.rfile.read(int(s.headers['Content-Length'])))
        SEEN.append(dict(path=s.path, key=s.headers.get('x-goog-api-key'), body=body_))
        if MODE['code'] != 200: s.send_response(MODE['code']); s.end_headers(); s.wfile.write(b'{}'); return
        out = json.dumps({'candidates': [{'content': {'parts': [{'text': json.dumps(RESP, ensure_ascii=False)}]}}]}).encode()
        s.send_response(200); s.send_header('Content-Type', 'application/json'); s.end_headers(); s.wfile.write(out)
gs = HTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
os.environ.update(GEMINI_API_KEY=KEY, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}/v1beta')
calls = []
ex = exam.extract([('all', img)], 'toan', wd, ai_hook=lambda: calls.append(1))
assert calls == [1] and len(SEEN) == 1
s0 = SEEN[0]; assert s0['path'] == '/v1beta/models/gemini-3.6-flash:generateContent' and s0['key'] == KEY
gc = s0['body']['generationConfig']; assert gc['responseMimeType'] == 'application/json' and 'responseSchema' in gc and gc['temperature'] == 0
assert 'Tiếng Việt' in s0['body']['contents'][0]['parts'][0]['text'] or 'tiểu học' in s0['body']['contents'][0]['parts'][0]['text']
assert any('inline_data' in p for p in s0['body']['contents'][0]['parts'])
assert exam.count_questions(ex) == 3 and ex['answers']['1']['text'] == 'C'
f1, f2 = ex['figures']['f1'], ex['figures']['f2']
assert f1['kind'] == 'geometry' and os.path.exists(f1['file']) and os.path.exists(f1['orig'])
w1 = Image.open(f1['file']).size; assert abs(w1[0] - round(5 / 2.54 * 300)) <= 3, w1   # 5 cm ở 300 dpi
px = Image.open(f1['file']).convert('L'); assert px.getextrema()[0] < 60, 'SVG phải có nét vẽ'
c2 = Image.open(f2['file']); assert abs(c2.size[0] - 0.34 * 1000) < 15 and abs(c2.size[1] - 0.26 * 1400) < 20, c2.size   # cắt đúng vùng, độ phân giải gốc
r_, g_, b_ = c2.convert('RGB').getpixel((c2.size[0] // 2, c2.size[1] // 2)); assert r_ > 200 and g_ < 90, 'vùng cắt phải chứa hình tô đỏ'
out = os.path.join(wd, 'de_ai.docx'); warns, _ = exam.build(TPL['toan'], ex, FIELDS, out, [os.path.join(wd, 'figs')])
z = zipfile.ZipFile(out); tpl_media = {n for n in zipfile.ZipFile(TPL['toan']).namelist() if n.startswith('word/media/')}
new_media = [n for n in z.namelist() if n.startswith('word/media/') and n not in tpl_media]
assert len(new_media) >= 2, new_media
dt = docx.Document(out); assert [p.text for p in dt.paragraphs if re.match(r'^Câu \d+\.', p.text)][0].startswith('Câu 1. Hình bên có mấy góc? (M1) (0,5đ)')
assert 'a) Hình vuông' in all_text(out); shutil.copy(out, os.path.join(tempfile.gettempdir(), 'v16_de_ai.docx'))
# lỗi dịch vụ
MODE['code'] = 429
try: exam.extract([('all', img)], 'toan', wd); raise SystemExit('phải lỗi 429')
except ValueError as e: assert 'quá tải' in str(e) or 'hết lượt' in str(e)
MODE['code'] = 200
assert KEY not in LOGS.getvalue(), 'KHÔNG được ghi key vào log'
print('6. Gemini giả lập OK (SVG 300dpi + cắt ảnh)')

# ---------- 7. Đồng bộ: không đẩy tệp gốc / ảnh xem trước ----------
assert not storage.wanted('users/gv-x/de/abcdef123456/src/0_all.png') and not storage.wanted('users/gv-x/de/abcdef123456/pages/p1.png')
assert storage.wanted('users/gv-x/de/abcdef123456/exam.json') and storage.wanted('users/gv-x/de/abcdef123456/de.docx')
print('7. đồng bộ OK')
gs.shutdown()
print('TẤT CẢ OK')
