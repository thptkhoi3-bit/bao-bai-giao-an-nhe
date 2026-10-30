"""Tạo đề kiểm tra: đọc nội dung đề (Word / PDF / ảnh) -> JSON có cấu trúc -> dựng .docx trên đề mẫu.

Quy tắc bắt buộc: xem EXAM_RULES.md (giữ nguyên cấu trúc mẫu, chỉ thay chữ trong run sẵn có cho các
trường giáo viên chọn; ngoại lệ đã duyệt: ngắt trang đầu phần + bỏ đoạn trống tràn đầu trang).
"""
import os, re, io, json, copy, math, base64, zipfile, shutil, subprocess, tempfile, threading, logging, uuid
from lxml import etree

log = logging.getLogger('exam')
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
q = lambda t: f'{{{W}}}{t}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
FIG_TAGS = ('drawing', 'pict', 'object')
MATH = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

# =====================================================================================
# 1. Tiện ích XML
# =====================================================================================
def ptext(p, tabs=True):
    out = []
    for e in p.iter(q('t'), q('tab'), q('br'), MATH + 't'):
        if e.tag in (q('t'), MATH + 't'): out.append(e.text or '')
        elif e.tag == q('tab'):
            if tabs and e.getparent().tag == q('r'): out.append('\t')
        elif e.tag == q('br') and e.get(q('type')) not in ('page', 'column'): out.append('\n')
    return ''.join(out)

def has_fig(e): return any(True for t in FIG_TAGS for _ in e.iter(q(t))) or any(True for _ in e.iter(MATH + 'oMath'))
def has_pagebreak(e): return any(b.get(q('type')) == 'page' for b in e.iter(q('br')))
def is_empty_p(e):
    return e.tag == q('p') and not ptext(e).strip() and not has_fig(e) and not has_pagebreak(e) and e.find('.//' + q('sectPr')) is None

def clean_ids(e):
    """Bỏ w14:paraId/textId ở phần tử nhân bản (tránh trùng id)."""
    for x in e.iter():
        for a in (f'{{{W14}}}paraId', f'{{{W14}}}textId'):
            if a in x.attrib: del x.attrib[a]
    return e

def runs_of(p): return [r for r in p.iter(q('r'))]
def rtext(r): return ''.join((x.text or '') if x.tag == q('t') else ('\t' if x.tag == q('tab') else '') for x in r)

def make_run(rpr, text):
    r = etree.Element(q('r'))
    if rpr is not None: r.append(copy.deepcopy(rpr))
    parts = re.split(r'(\t|\n)', text)
    for s in parts:
        if s == '\t': etree.SubElement(r, q('tab'))
        elif s == '\n': etree.SubElement(r, q('br'))
        elif s:
            t = etree.SubElement(r, q('t')); t.text = s
            if s != s.strip(): t.set(XS, 'preserve')
    return r

def new_par(proto, segments, keep_ppr=True):
    """Đoạn mới theo pPr của đoạn mẫu; segments = [(rPr, text), ...]."""
    p = etree.Element(q('p'))
    ppr = proto.find(q('pPr')) if proto is not None else None
    if keep_ppr and ppr is not None:
        ppr = copy.deepcopy(ppr)
        n = ppr.find(q('numPr'))
        if n is not None: ppr.remove(n)          # không dùng đánh số tự động: chữ "a)" ghi thẳng
        p.append(ppr)
    for rpr, text in segments:
        if text: p.append(make_run(rpr, text))
    return p

def run_rpr(r): return r.find(q('rPr')) if r is not None else None

def set_par_text(p, text, keep_lead_ws=True):
    """Đặt chữ cho đoạn sẵn có: giữ pPr, giữ run khoảng trắng đầu (nếu có), chữ vào run có chữ đầu tiên."""
    rs = [r for r in p.findall(q('r'))]
    lead = []
    for r in rs:
        if keep_lead_ws and rtext(r) and not rtext(r).strip(): lead.append(r); continue
        break
    body = [r for r in rs if r not in lead and rtext(r)]
    if not body:
        rpr = run_rpr(rs[-1]) if rs else None; p.append(make_run(rpr, text)); return p
    first = body[0]; m = re.match(r'^(\s*)', rtext(first)); ws = m.group(1) if keep_lead_ws else ''
    nr = make_run(run_rpr(first), ws + text); first.addprevious(nr)
    for r in body: r.getparent().remove(r)
    return p

# =====================================================================================
# 2. Điền phần đầu đề (chỉ thay chữ trong run sẵn có) – EXAM_RULES §1
# =====================================================================================
def _sub_runs(p, pat, rep, flags=0):
    changed = []
    for _ in range(10):
        ts = [t for r in p.iter(q('r')) for t in r.findall(q('t'))]
        full = ''.join(t.text or '' for t in ts)
        m = next((m for m in re.finditer(pat, full, flags) if m.expand(rep) != m.group(0)), None)
        if not m: break
        s, e = m.span(); new = m.expand(rep); pos = 0; first = True
        for t in ts:
            tx = t.text or ''; a, b = pos, pos + len(tx); pos = b
            if b <= s or a >= e or not tx: continue
            lo, hi = max(s, a) - a, min(e, b) - a
            t.text = tx[:lo] + (new if first else '') + tx[hi:]
            if t.text != t.text.strip() and t.get(XS) is None: t.set(XS, 'preserve')
            first = False
        changed.append((m.group(0), new))
    return changed

def clean_name(s, prefixes):
    s = ' '.join((s or '').split()).upper()
    for pre in prefixes:
        if s.startswith(pre + ' '): s = s[len(pre) + 1:]
    return s.strip()

def header_rules(f):
    """f: dict(xa, truong, nam, loai, lop, thoigian)."""
    R_ = []
    xa = clean_name(f.get('xa'), ['UBND', 'XÃ', 'PHƯỜNG', 'UBND XÃ', 'UBND PHƯỜNG'])
    xa = clean_name(xa, ['XÃ', 'PHƯỜNG'])
    if xa: R_.append((r'(?<=UBND XÃ )(?!\s)[^\n()]+?(?=\s*$)', xa, None))
    tr = clean_name(f.get('truong'), ['TRƯỜNG'])
    tr = re.sub(r'^(TH & THCS|TH&THCS|TIỂU HỌC|TH|THCS)\s+', '', tr)
    if tr:
        R_.append((r'(?<=TRƯỜNG TH & THCS)(\s+)(?!\s)([^()\n]+?)(?=\s*(\(|$))', r'\g<1>' + tr, None))
        R_.append((r'(?<=TRƯỜNG TH)(\s+)(?![\s&])([^()\n]+?)(?=\s*(\(|$))', r'\g<1>' + tr, None))
    m = re.match(r'\s*(20\d\d)\s*-\s*(20\d\d)', f.get('nam') or '')
    if m: R_.append((r'(?<=NĂM HỌC )20\d\d(\s*-\s*)20\d\d', m.group(1) + r'\g<1>' + m.group(2), None))
    lb = (f.get('loai') or '').upper().replace('BÀI KIỂM TRA', '').split()
    if len(lb) >= 4 and lb[0] in ('GIỮA', 'CUỐI'):
        R_.append((r'(?<=BÀI KIỂM TRA )(GIỮA|CUỐI)(?= HỌC KÌ)', lb[0], 'BÀI KIỂM TRA'))
        R_.append((r'(?<=HỌC KÌ )II?\b', lb[3], 'BÀI KIỂM TRA'))
    # Cộng thêm (Tạo đề AI): 2 loại không phải kiểm tra — thay cả cụm tiêu đề mẫu, không đụng nhánh trên.
    # (Nhánh cũ chỉ chạy khi loai bắt đầu bằng GIỮA/CUỐI; 2 giá trị này không lọt vào đó.)
    lb_full = (f.get('loai') or '').strip().upper()
    if lb_full == 'PHIẾU BÀI TẬP THEO CHỦ ĐỀ':
        R_.append((r'BÀI KIỂM TRA\s+(GIỮA|CUỐI)\s+HỌC KÌ\s+I{1,2}', 'PHIẾU BÀI TẬP', 'BÀI KIỂM TRA'))
    elif lb_full == 'ĐỀ ÔN TẬP':
        R_.append((r'BÀI KIỂM TRA\s+(GIỮA|CUỐI)\s+HỌC KÌ\s+I{1,2}', 'ĐỀ ÔN TẬP', 'BÀI KIỂM TRA'))
    elif lb_full == 'ĐỀ KIỂM TRA CUỐI NĂM':
        R_.append((r'BÀI KIỂM TRA\s+(GIỮA|CUỐI)\s+HỌC KÌ\s+I{1,2}', 'BÀI KIỂM TRA CUỐI NĂM', 'BÀI KIỂM TRA'))
    # Ô đầu đề "(Bài kiểm tra có .. trang)": phiếu/ôn tập không phải bài kiểm tra → đổi chữ cho khớp.
    if lb_full == 'PHIẾU BÀI TẬP THEO CHỦ ĐỀ':
        R_.append((r'\(Bài kiểm tra có', '(Phiếu bài tập có', 'Bài kiểm tra'))
    elif lb_full == 'ĐỀ ÔN TẬP':
        R_.append((r'\(Bài kiểm tra có', '(Đề ôn tập có', 'Bài kiểm tra'))
    if f.get('lop'): R_.append((r'(?<=Lớp )\d', str(f['lop']), r'Môn|MÔN'))
    tg = re.search(r'\d+', f.get('thoigian') or '')
    if tg:
        R_.append((r'(?<=làm bài )\d+(?= phút)', tg.group(0), None)); R_.append((r'(?<=làm bài: )\d+(?= phút)', tg.group(0), None))
    return R_

def fill_header(root, fields):
    log_ = []
    rules = header_rules(fields)
    for p in root.iter(q('p')):
        txt = ''.join(t.text or '' for t in p.iter(q('t')))
        for pat, rep, filt in rules:
            if filt and not re.search(filt, txt): continue
            log_ += _sub_runs(p, pat, rep)
    return log_

# =====================================================================================
# 3. Đọc .docx -> các "đơn vị" (đoạn / bảng) có chữ, cờ hình, phần tử gốc
# =====================================================================================
class Numbering:
    def __init__(s, zf):
        s.fmt = {}; s.cnt = {}
        try: n = etree.fromstring(zf.read('word/numbering.xml'))
        except KeyError: return
        absn = {}
        for a in n.findall(q('abstractNum')):
            lv = {}
            for l in a.findall(q('lvl')):
                nf = l.find(q('numFmt')); lt = l.find(q('lvlText')); st = l.find(q('start'))
                lv[l.get(q('ilvl'))] = (nf.get(q('val')) if nf is not None else 'decimal', lt.get(q('val')) if lt is not None else '%1.', int(st.get(q('val'))) if st is not None else 1)
            absn[a.get(q('abstractNumId'))] = lv
        for num in n.findall(q('num')):
            a = num.find(q('abstractNumId'))
            if a is not None: s.fmt[num.get(q('numId'))] = absn.get(a.get(q('val')), {})
    def label(s, p):
        pp = p.find(q('pPr')); n = pp.find(q('numPr')) if pp is not None else None
        if n is None: return ''
        nid = n.find(q('numId')); il = n.find(q('ilvl'))
        nid = nid.get(q('val')) if nid is not None else None; il = il.get(q('val')) if il is not None else '0'
        f = s.fmt.get(nid, {}).get(il)
        if not f or nid == '0': return ''
        k = (nid, il); s.cnt[k] = s.cnt.get(k, f[2] - 1) + 1; v = s.cnt[k]
        fmt = f[0]
        val = {'lowerLetter': chr(96 + v) if v <= 26 else str(v), 'upperLetter': chr(64 + v) if v <= 26 else str(v),
               'lowerRoman': ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'][v - 1] if v <= 10 else str(v),
               'upperRoman': ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'][v - 1] if v <= 10 else str(v),
               'bullet': '-'}.get(fmt, str(v))
        return f[1].replace('%1', val) + ' '

def table_cells(tbl):
    rows = []
    for tr in tbl.findall(q('tr')):
        rows.append([tc for tc in tr.findall(q('tc'))])
    return rows

def cell_text(tc): return '\n'.join(ptext(p) for p in tc.iter(q('p'))).strip()

def docx_units(path):
    zf = zipfile.ZipFile(path); root = etree.fromstring(zf.read('word/document.xml'))
    num = Numbering(zf); units = []
    for e in root.find(q('body')):
        if e.tag == q('p'):
            t = num.label(e) + ptext(e)
            units.append(dict(kind='p', text=t, fig=has_fig(e), el=e, empty=is_empty_p(e), pb=has_pagebreak(e)))
        elif e.tag == q('tbl'):
            rows = table_cells(e)
            units.append(dict(kind='tbl', rows=[[cell_text(c) for c in r] for r in rows], fig=has_fig(e), el=e, empty=False, pb=False,
                              text='\t'.join(cell_text(c) for c in rows[0]) if rows else ''))
    return units, zf, root

# =====================================================================================
# 4. Bộ phân tích theo luật (Word / PDF chữ) -> JSON đề
# =====================================================================================
TAG_RE = re.compile(r'\(\s*M\s*(\d)\s*(?:\)\s*\(|[-–]+)\s*([\d]+(?:[.,]\d+)?)\s*(?:đ|điểm)\s*\)', re.I)
QSTART = re.compile(r'^\s*Câu\s*(\d+)\s*[.:]\s*', re.I)
SEC_RE = re.compile(r'^\s*(I|II|III|IV|V|PHẦN\s+\w+)\s*[.:]\s*(.+)$')
OPT_RE = re.compile(r'(?:(?<=^)|(?<=[\s\t]))([A-D])\s*[.)]\s*')
DOTS = re.compile(r'^[\s.…·_]+$')

def is_dots(t): return bool(t.strip()) and bool(DOTS.match(t)) and len(t.strip()) >= 8

def split_opts(t):
    ms = list(OPT_RE.finditer(t))
    if not ms or ms[0].group(1) != 'A' or t[:ms[0].start()].strip(): return None
    letters = [m.group(1) for m in ms]
    if letters != list('ABCD')[:len(letters)]: return None
    out = []
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(t)
        out.append(' '.join(t[m.end():end].replace('\t', ' ').split()))
    return out

def parse_tag(t):
    m = TAG_RE.search(t)
    if not m: return t.strip(), None, None
    return (t[:m.start()] + t[m.end():]).strip(), 'M' + m.group(1), m.group(2).replace('.', ',')

def norm_space(t): return re.sub(r'[ \t]+', ' ', t).strip()

def new_exam(mon):
    e = dict(mon=mon, sections=[], answers={}, figures={}, raw={}, warnings=[])
    if mon == 'tv':
        e.update(reading=dict(intro='', title='', paragraphs=[], source=[], instruction=''), writing=dict(topics=[]), dictation=dict(title='', paragraphs=[], source=''))
    return e

def parse_units(units, mon, raw_store=None, slot=None):
    """units: list of dict(kind,text,fig,el,empty,rows). Trả về JSON đề."""
    ex = new_exam(mon); state = 'head'; sec = None; qn = None; pending_space = 0
    def cur_sec():
        nonlocal sec
        if sec is None: sec = dict(title='', instruction='', questions=[]); ex['sections'].append(sec)
        return sec
    def add_block(b):
        if qn is not None: qn['blocks'].append(b)
    def add_raw(u):
        rid = 'r%d' % len(ex['raw']); ex['raw'][rid] = raw_store(u['el']) if raw_store else None
        add_block(dict(t='raw', id=rid, text=u['text'].strip()[:200]))
    i = 0; part = 'doc'
    alltxt = ' '.join(u['text'] for u in units if u['kind'] == 'p')
    if mon == 'tv' and slot == 'nghe_viet' and not re.search(r'Đề\s*nghe', alltxt, re.I): part = 'dictation'
    if mon == 'tv' and slot == 'de_viet' and not re.search(r'Bài số 2', alltxt): part = 'viet'
    if mon == 'tv' and slot == 'doc_hieu' and 'Đọc thầm' not in alltxt: state = 'reading_title'
    loose_topic = []
    while i < len(units):
        u = units[i]; t = u['text'] if u['kind'] == 'p' else ''
        st = t.strip()
        # ---- bảng ----
        if u['kind'] == 'tbl':
            if state == 'answers': parse_answer_table(u, ex, raw_store); i += 1; continue
            flat = ' '.join(' '.join(r) for r in u['rows'])
            if re.search(r'UBND|BÀI KIỂM TRA|Nhận xét', flat) and state != 'answers':
                if mon == 'tv' and ex['sections'] and 'Bài số 2' in flat: part = 'viet'
                i += 1; continue
            if 'ĐỌC HIỂU' in flat.upper() and mon == 'tv': state = 'reading_intro'; i += 1; continue
            if len(u['rows']) > 8 and not any(any(c.strip() for c in r) for r in u['rows'][1:]): i += 1; continue  # lưới ô ly
            if len(u['rows']) == 1 and qn is not None:
                cells = [c.strip() for c in u['rows'][0]]
                opts = split_opts('\t'.join(cells)) if cells and re.match(r'^[A-D]\s*[.)]', cells[0]) else None
                if u['fig'] or (opts is None and any(re.match(r'^[A-D]\s*[.)]', c) for c in cells)):
                    add_raw(u); i += 1; continue
                if opts: add_block(dict(t='options', items=opts, layout='row')); i += 1; continue
            if qn is not None: add_raw(u); i += 1; continue   # mọi bảng trong câu (bảng số liệu, Đ/S, lưới điền…) giữ nguyên khối
            i += 1; continue
        # ---- đoạn trống ----
        if u['empty']:
            if qn is not None and qn['blocks'] and qn['blocks'][-1]['t'] in ('raw', 'figure', 'space'):
                b = qn['blocks'][-1]
                if b['t'] == 'space': b['n'] += 1
                else: add_block(dict(t='space', n=1))
            i += 1; continue
        if u['fig'] and not st and qn is not None:
            if u.get('figid'): add_block(dict(t='figure', id=u['figid']))
            else: add_raw(u)
            i += 1; continue
        if u['fig'] and not st: i += 1; continue
        if u.get('pb') and part == 'dictation': part = 'after'
        if u['el'] is None and re.match(r'^(UBND|TRƯỜNG|NĂM HỌC|Môn\s*:|MÔN\s*:|Thời gian|\(Bài kiểm tra có|Họ và tên|Điểm\b|Nhận xét|BÀI KIỂM TRA|Bài số \d|KTĐ|KTTV|đọc hiểu$)', st.strip()) and state != 'answers':
            if mon == 'tv' and re.match(r'^Bài số 2', st.strip()): part = 'viet'; qn = None
            i += 1; continue
        if not st and not u['fig']: i += 1; continue
        # ---- mốc phần ----
        U = st.upper()
        if re.match(r'^(ĐÁP ÁN|HƯỚNG DẪN CHẤM)', U): state = 'answers'; qn = None; i += 1; continue
        if state == 'answers':
            m = re.match(r'^(?:(?:I|II|III)\.\s*[^\t]*?\t?\s*)?Câu\s*(\d+)\s*[:.]?\s*(.*)$', st)
            if m:
                rest = m.group(2).split('\t'); cur_ans = ex['answers'].setdefault(m.group(1), dict(text='', points=''))
                if rest and rest[0].strip():
                    if len(rest) > 1 and re.fullmatch(r'\s*[\d]+(?:[.,]\d+)?\s*(đ|điểm)?\s*', rest[-1]): cur_ans['points'] = rest[-1].strip(); rest = rest[:-1]
                    cur_ans['text'] = ' '.join(x.strip() for x in rest if x.strip())
            elif ex['answers'] and u['el'] is None:
                cur_ans = ex['answers'][list(ex['answers'])[-1]]
                if re.fullmatch(r'[\d]+(?:[.,]\d+)?\s*(đ|điểm)?', st) and cur_ans['text'] and not cur_ans['points']: cur_ans['points'] = st
                elif not re.match(r'^(I|II|III)\.\s', st) and not re.match(r'^(Câu|Đáp án|Điểm)$', st): cur_ans['text'] = (cur_ans['text'] + '\n' + st.replace('\t', '   ')).strip()
            i += 1; continue
        if part == 'after' or (qn is None and re.match(r'^(UBND|TRƯỜNG|NĂM HỌC|MÔN\s*:)', st)): i += 1; continue
        if mon == 'tv':
            if re.match(r'^(I\.\s*Chính tả|II\.\s*Luyện viết|Bài số 2)', st) and part != 'dictation': part = 'viet'; qn = None; i += 1; continue
            if re.match(r'^(I|II)\.\s', st) and state in ('head', 'reading_intro', 'reading_title') and part == 'doc': i += 1; continue
            if re.match(r'^Đề\s*nghe\s*-?\s*viết', st, re.I): part = 'dictation'; qn = None; i += 1; continue
            if part == 'dictation':
                d = ex['dictation']
                if re.match(r'^(Theo|Phỏng theo)\b', st) or (st.startswith('(') and st.endswith(')')): d['source'] = (d['source'] + ' ' + st).strip()
                elif not d['title']: d['title'] = st
                else: d['paragraphs'].append(st)
                i += 1; continue
            if part == 'viet':
                m = re.match(r'^(Đề\s*\d+)\s*:\s*(.*)$', st)
                if m: ex['writing']['topics'].append(dict(label=m.group(1), text=m.group(2).strip(), hint='')); i += 1; continue
                if st.startswith('(Gợi ý') and ex['writing']['topics']: ex['writing']['topics'][-1]['hint'] = st; i += 1; continue
                if slot == 'de_viet' and not re.match(r'^(II\.|I\.|Chọn|Bài làm|Viết vào)', st): loose_topic.append(st)
                i += 1; continue
            if re.match(r'^Đọc thầm', st): ex['reading']['intro'] = st; state = 'reading_title'; i += 1; continue
            if state == 'reading_title': ex['reading']['title'] = st; state = 'reading'; i += 1; continue
            if state == 'reading':
                if re.match(r'^(Theo|Phỏng theo)\b', st) or (st.startswith('(') and st.endswith(')') and len(st) < 60):
                    ex['reading']['source'].append(st); i += 1; continue
                if re.match(r'^(Khoanh|Dựa vào|Em hãy đọc|Trả lời)', st) or QSTART.match(st):
                    state = 'q'
                    if not QSTART.match(st): ex['reading']['instruction'] = st; i += 1; continue
                else:
                    if ex['reading']['source']: state = 'q'
                    else: ex['reading']['paragraphs'].append(st); i += 1; continue
        msec = SEC_RE.match(st)
        if msec and not QSTART.match(st) and len(st) < 60 and st.upper() == st.upper() and re.search(r'[A-ZĐÀ-Ỹ]{3,}', msec.group(2)):
            if msec.group(2).isupper() or mon == 'toan':
                sec = dict(title=st, instruction='', questions=[]); ex['sections'].append(sec); qn = None; state = 'q'; i += 1; continue
        # ---- câu hỏi (có thể nằm sau chuỗi chấm trong cùng đoạn) ----
        pieces = re.split(r'(?=\bCâu\s*\d+\s*[.:])', t) if 'Câu' in t else [t]
        if len(pieces) > 1 and pieces[0].strip() == '': pieces = pieces[1:]
        for k, piece in enumerate(pieces):
            ps = piece.strip()
            if not ps: continue
            m = QSTART.match(ps)
            if m and (k > 0 or ps == st.strip() or True) and not (state == 'head'):
                stem, muc, diem = parse_tag(ps[m.end():])
                qn = dict(so=int(m.group(1)), text=norm_space(stem), muc=muc, diem=diem, blocks=[])
                cur_sec()['questions'].append(qn); state = 'q'; continue
            if m and state == 'head':
                state = 'q'; stem, muc, diem = parse_tag(ps[m.end():])
                qn = dict(so=int(m.group(1)), text=norm_space(stem), muc=muc, diem=diem, blocks=[]); cur_sec()['questions'].append(qn); continue
            if qn is None:
                if state in ('q', 'head') and sec is not None and not sec['instruction'] and not sec['questions']: sec['instruction'] = ps
                elif state == 'q' and sec is None and mon == 'toan': cur_sec()['instruction'] = ps
                continue
            if u['fig']: add_raw(u); break
            if is_dots(ps):
                b = qn['blocks'][-1] if qn['blocks'] else None
                n = dots_lines(ps)
                if b and b['t'] == 'lines': b['n'] += n
                else: add_block(dict(t='lines', n=n))
                continue
            opts = split_opts(ps)
            if opts and len(opts) >= 2 and not any(opts) :
                ex['warnings'].append(f"Câu {qn['so']}: phương án là hình – tệp này không lấy được hình, hãy dùng file Word hoặc ảnh chụp."); continue
            if opts and len(opts) >= 2: add_block(dict(t='options', items=opts, layout='row')); continue
            if opts and len(opts) == 1:
                b = qn['blocks'][-1] if qn['blocks'] else None
                if b and b['t'] == 'options' and b.get('layout') == 'col' and len(b['items']) < 6: b['items'].append(opts[0])
                else: add_block(dict(t='options', items=opts, layout='col'))
                continue
            m2 = re.match(r'^([B-F])\s*[.)]\s*(.*)$', ps)
            b = qn['blocks'][-1] if qn['blocks'] else None
            if m2 and b and b['t'] == 'options' and b.get('layout') == 'col' and chr(65 + len(b['items'])) == m2.group(1):
                b['items'].append(norm_space(m2.group(2))); continue
            stem, muc, diem = parse_tag(piece)
            add_block(dict(t='text', text=(stem if muc else piece).strip('\n').rstrip(), muc=muc, diem=diem))
        i += 1
    if mon == 'tv' and not ex['writing']['topics'] and loose_topic:
        ex['writing']['topics'].append(dict(label='Đề 1', text=' '.join(loose_topic), hint=''))
    for s_ in ex['sections']:
        for qq in s_['questions']:
            while qq['blocks'] and qq['blocks'][-1]['t'] == 'space': qq['blocks'].pop()
    return ex

_font_cache = {}
def _font(size=14, bold=False):
    from PIL import ImageFont
    k = (size, bold)
    if k not in _font_cache:
        cands = ['/usr/share/fonts/truetype/liberation/LiberationSerif-%s.ttf' % ('Bold' if bold else 'Regular'),
                 '/usr/share/fonts/truetype/liberation2/LiberationSerif-%s.ttf' % ('Bold' if bold else 'Regular'),
                 'C:/Windows/Fonts/times.ttf']
        f = next((c for c in cands if os.path.exists(c)), None)
        _font_cache[k] = ImageFont.truetype(f, size * 20) if f else None
    return _font_cache[k]

LINE_TW = 10000   # bề rộng dòng (twips) mặc định khi chưa biết
def text_tw(s, size=14):
    f = _font(size)
    if f is None: return len(s) * (70 if s.strip('.') == '' else 140)
    return f.getlength(s)  # font dựng ở 20*size px => đơn vị twips

def dots_lines(s, width=None, size=14):
    return max(1, math.ceil(text_tw(s.strip(), size) / (width or LINE_TW) - 0.15))

def parse_answer_table(u, ex, raw_store):
    if u.get('el') is None: return
    found = False
    for tr in u['el'].findall(q('tr')):
        tcs = tr.findall(q('tc'))
        if len(tcs) < 2: continue
        c0 = cell_text(tcs[0]); m = re.search(r'Câu\s*(\d+)', c0)
        if not m: continue
        ans = dict(text=cell_text(tcs[1]), points=cell_text(tcs[-1]) if len(tcs) > 2 else '')
        if raw_store: ans['raw'] = [raw_store(tcs[1]), raw_store(tcs[-1]) if len(tcs) > 2 else None]
        ex['answers'][m.group(1)] = ans; found = True
    if not found:   # Hướng dẫn chấm dạng đoạn "Câu n: ..." trong ô bảng
        for p in u['el'].iter(q('p')):
            m = re.match(r'^\s*Câu\s*(\d+)\s*[:.]\s*(.*)$', ptext(p))
            if m: ex['answers'].setdefault(m.group(1), dict(text=m.group(2).strip(), points=''))

# =====================================================================================
# 5. Trích nội dung từ tệp nguồn
# =====================================================================================
class NeedAI(Exception): pass
AI_OFF_MSG = 'Chưa bật đọc ảnh – liên hệ 0865291703'
IMG_EXT = ('.jpg', '.jpeg', '.png', '.webp')

def raw_storer(zf, media_dir):
    try: rels = etree.fromstring(zf.read('word/_rels/document.xml.rels'))
    except KeyError: rels = None
    rmap = {r.get('Id'): (r.get('Type'), r.get('Target'), r.get('TargetMode')) for r in rels} if rels is not None else {}
    os.makedirs(media_dir, exist_ok=True)
    def store(el):
        el = copy.deepcopy(el)
        for o in list(el.iter('{urn:schemas-microsoft-com:office:office}OLEObject')): o.getparent().remove(o)  # chỉ giữ ảnh xem trước của công thức
        for h in list(el.iter(q('hyperlink'))):
            par = h.getparent(); idx = par.index(h)
            for c in list(h): par.insert(idx, c); idx += 1
            par.remove(h)
        images = {}
        for x in el.iter():
            for a, v in list(x.attrib.items()):
                if not a.startswith('{%s}' % R): continue
                typ, tgt, mode = rmap.get(v, (None, None, None))
                if typ and typ.endswith('/image') and mode != 'External':
                    try: blob = zf.read('word/' + tgt.lstrip('/').replace('word/', '', 1) if not tgt.startswith('/') else tgt.lstrip('/'))
                    except KeyError: continue
                    fn = uuid.uuid4().hex[:12] + os.path.splitext(tgt)[1].lower()
                    with open(os.path.join(media_dir, fn), 'wb') as f: f.write(blob)
                    images[v] = fn
                else: del x.attrib[a]
        return dict(xml=etree.tostring(el).decode(), images=images)
    return store

def extract_docx(path, mon, workdir, slot=None):
    units, zf, root = docx_units(path)
    ex = parse_units(units, mon, raw_storer(zf, os.path.join(workdir, 'media')), slot)
    ex['source_kind'] = 'docx'
    return ex

MARK_RE = re.compile(r'^(Câu\s*\d+|[A-D]\s*[.)]|[a-đ]\s*\)|(I|II|III|IV)\s*\.|ĐÁP ÁN|HƯỚNG DẪN|Đề\s*\d|Đề nghe|Theo\b|Phỏng theo|\(Gợi ý|Bài làm|Đọc thầm|Khoanh)')
def pdf_text_units(path, figdir=None):
    """PDF chữ -> đoạn: gom dòng cùng hàng (nối bằng tab), nối dòng xuống dòng trong cùng đoạn; ảnh nhúng -> hình (cắt vùng)."""
    import pymupdf
    doc = pymupdf.open(path); units = []; chars = 0; nfig = 0
    for pi, page in enumerate(doc):
        d = page.get_text('dict'); rows = []
        for blk in d['blocks']:
            if blk['type'] == 1:
                x0, y0, x1, y1 = blk['bbox']
                if figdir and (x1 - x0) > 25 and (y1 - y0) > 25:
                    os.makedirs(figdir, exist_ok=True); fid = f'p{nfig}'; nfig += 1
                    fn = os.path.join(figdir, fid + '.png'); page.get_pixmap(clip=pymupdf.Rect(blk['bbox']), dpi=200).save(fn)
                    rows.append(dict(y0=y0, y1=y1, x0=x0, x1=x1, text='', fig=fid))
                continue
            for ln in blk['lines']:
                t = ''.join(sp['text'] for sp in ln['spans'])
                if not t.strip(): continue
                x0, y0, x1, y1 = ln['bbox']; rows.append(dict(y0=y0, y1=y1, x0=x0, x1=x1, text=t.strip()))
        rows.sort(key=lambda r: (round(r['y0']), r['x0']))
        lines = []
        for r in rows:   # cùng hàng (lệch < 3pt) -> 1 dòng
            if lines and not r.get('fig') and not lines[-1].get('fig') and abs(r['y0'] - lines[-1]['y0']) < 3 and r['x0'] > lines[-1]['x1'] - 1:
                L = lines[-1]; L['text'] += '\t' + r['text']; L['x1'] = r['x1']; L['y1'] = max(L['y1'], r['y1'])
            else: lines.append(dict(r))
        if not lines: continue
        right = max(l['x1'] for l in lines if not l.get('fig')) if any(not l.get('fig') for l in lines) else 0
        left = min(l['x0'] for l in lines if not l.get('fig')) if right else 0
        cur = None
        for l in lines:
            t = l['text']
            if re.match(r'^\s*Trang\s*\d+\s*$', t) or re.match(r'^\d+$', t.strip()) and l['y0'] > page.rect.height - 60: continue
            if l.get('fig'):
                if cur: units.append(cur); cur = None
                units.append(dict(kind='p', text='', fig=True, figid=l['fig'], el=None, empty=False, pb=False)); continue
            chars += len(t)
            h = l['y1'] - l['y0']
            cont = (cur is not None and not MARK_RE.match(t) and not is_dots(t) and not is_dots(cur['text']) and '\t' not in cur['text'] and '\t' not in t
                    and l['y0'] - cur['_y1'] < h * 0.6 and l['x0'] <= cur['_x0'] + 4 and cur['_x1'] > right - 40
                    and not TAG_RE.search(cur['text'][-20:]))
            tagonly = cur is not None and re.match(r'^\(\s*(M\s*\d|[\d.,]+\s*đ)', t) and l['y0'] - cur['_y1'] < h * 0.8
            if cont or tagonly:
                cur['text'] += ' ' + t; cur['_y1'] = l['y1']; cur['_x1'] = l['x1']
            else:
                if cur: units.append(cur)
                cur = dict(kind='p', text=t, fig=False, el=None, empty=False, pb=False, _y1=l['y1'], _x0=l['x0'], _x1=l['x1'])
        if cur: units.append(cur)
    return units, chars / max(1, len(doc)), len(doc)

def extract_pdf_text(path, mon, slot=None, workdir=None):
    figdir = os.path.join(workdir, 'figs') if workdir else None
    units, per_page, n = pdf_text_units(path, figdir)
    if per_page < 80: return None
    ex = parse_units(units, mon, None, slot); ex['source_kind'] = 'pdf'
    for u in units:
        if u.get('figid') and figdir: ex['figures'][u['figid']] = dict(kind='picture', file=os.path.join(figdir, u['figid'] + '.png'))
    return ex

def count_questions(ex): return sum(len(s['questions']) for s in ex['sections'])

# ------------------------------- Gemini ------------------------------------------------
def gemini_url(): return os.environ.get('GEMINI_API_URL', 'https://generativelanguage.googleapis.com/v1beta').rstrip('/')
def gemini_enabled(): return bool(os.environ.get('GEMINI_API_KEY', '').strip())

SCHEMA_TEXT = """{
 "sections": [ { "title": "I. TRẮC NGHIỆM", "instruction": "câu dẫn chung (nếu có)",
   "questions": [ { "so": 1, "text": "đề câu hỏi (KHÔNG gồm 'Câu 1.' và KHÔNG gồm mức/điểm)", "muc": "M1", "diem": "0,5",
     "blocks": [
       {"t":"text","text":"dòng phụ, ví dụ 'a) Số lớn nhất ... là:'","muc":null,"diem":null},
       {"t":"options","items":["653","356","635","365"],"layout":"row|col"},
       {"t":"lines","n":3},
       {"t":"ds","items":["a) Số nào nhân với 1 cũng bằng 1","b) ..."]},
       {"t":"figure","id":"f1"} ] } ] } ],
 "reading": {"intro":"Đọc thầm bài văn sau ...","title":"TÊN BÀI","paragraphs":["đoạn 1","đoạn 2"],"source":["Theo ..."],"instruction":"Khoanh vào ..."},
 "writing": {"topics":[{"label":"Đề 1","text":"...","hint":"(Gợi ý: ...)"}]},
 "dictation": {"title":"...","paragraphs":["..."],"source":"Theo ..."},
 "answers": {"1": {"text":"C","points":"0,5"}},
 "figures": [ {"id":"f1","kind":"geometry|picture","page":0,"box":[ymin,xmin,ymax,xmax],"svg":"<svg ...>...</svg>","w_cm":6} ]
}"""

PROMPT = """Bạn là trợ lý số hoá đề kiểm tra tiểu học Việt Nam (môn {mon}). Đọc CHÍNH XÁC nội dung trong các tệp/ảnh đính kèm
và trả về DUY NHẤT một đối tượng JSON đúng lược đồ dưới đây (không thêm chữ nào khác).
Quy tắc:
- Giữ nguyên chữ tiếng Việt có dấu, số, dấu phép tính (×, :, –, =, ?), dấu chấm điền (....) như trong đề. Không tự sửa, không bịa thêm câu.
- Số câu hỏi đúng bằng số câu có trong đề. "so" là số thứ tự câu. Mức (M1/M2/M3) và điểm ("0,5", "1") lấy từ ghi chú như "(M1) (0,5đ)" hoặc "(M1 - 0,5đ)"; không có thì để null.
- Phương án A/B/C/D: "items" chỉ chứa nội dung sau chữ cái. layout "row" nếu các phương án nằm trên một hàng, "col" nếu mỗi phương án một dòng.
- Dòng kẻ chấm để học sinh viết: {{"t":"lines","n": số dòng}}. Câu Đúng ghi Đ / Sai ghi S: dùng "ds".
- Hình vẽ: thêm {{"t":"figure","id":"fN"}} vào đúng chỗ, và mô tả trong "figures":
  * Hình hình học (đoạn thẳng, tam giác, tứ giác, hình tròn, hình khối vẽ nét, lưới ô vuông, đồng hồ, số đo): kind "geometry",
    "svg" là mã SVG ĐẦY ĐỦ, CHÍNH XÁC (viewBox theo đơn vị mm, nét đen stroke-width 0.4, không tô màu trừ khi đề có màu,
    chữ/nhãn dùng <text> font-family "Times New Roman"), vẽ lại đúng tỉ lệ, đúng số đường, đúng vị trí nhãn. "w_cm" = bề rộng in (cm).
  * Tranh, ảnh minh hoạ, hình phức tạp: kind "picture", "page" = số thứ tự ảnh/trang (từ 0), "box" = [ymin,xmin,ymax,xmax] chuẩn hoá 0–1000 bao sát hình.
- Môn Tiếng Việt: điền "reading" (bài đọc, nguồn, câu dẫn), "writing" (các đề viết đoạn + gợi ý), "dictation" (bài nghe - viết) nếu có; câu hỏi đọc hiểu để trong sections[0].
- Có đáp án / hướng dẫn chấm thì điền "answers" theo số câu.
Lược đồ:
{schema}
{extra}"""

def _gem_schema():
    S, A, O = 'STRING', 'ARRAY', 'OBJECT'
    blk = {'type': O, 'properties': {'t': {'type': S, 'enum': ['text', 'options', 'lines', 'ds', 'figure']}, 'text': {'type': S, 'nullable': True},
           'muc': {'type': S, 'nullable': True}, 'diem': {'type': S, 'nullable': True}, 'items': {'type': A, 'items': {'type': S}},
           'layout': {'type': S, 'nullable': True}, 'n': {'type': 'INTEGER', 'nullable': True}, 'id': {'type': S, 'nullable': True}}, 'required': ['t']}
    qs = {'type': O, 'properties': {'so': {'type': 'INTEGER'}, 'text': {'type': S}, 'muc': {'type': S, 'nullable': True}, 'diem': {'type': S, 'nullable': True},
          'blocks': {'type': A, 'items': blk}}, 'required': ['so', 'text', 'blocks']}
    sec = {'type': O, 'properties': {'title': {'type': S, 'nullable': True}, 'instruction': {'type': S, 'nullable': True}, 'questions': {'type': A, 'items': qs}}, 'required': ['questions']}
    fig = {'type': O, 'properties': {'id': {'type': S}, 'kind': {'type': S, 'enum': ['geometry', 'picture']}, 'page': {'type': 'INTEGER', 'nullable': True},
           'box': {'type': A, 'items': {'type': 'NUMBER'}}, 'svg': {'type': S, 'nullable': True}, 'w_cm': {'type': 'NUMBER', 'nullable': True}}, 'required': ['id', 'kind']}
    ans = {'type': A, 'items': {'type': O, 'properties': {'so': {'type': 'INTEGER'}, 'text': {'type': S}, 'points': {'type': S, 'nullable': True}}, 'required': ['so', 'text']}}
    return {'type': O, 'properties': {
        'sections': {'type': A, 'items': sec},
        'reading': {'type': O, 'nullable': True, 'properties': {'intro': {'type': S, 'nullable': True}, 'title': {'type': S, 'nullable': True}, 'paragraphs': {'type': A, 'items': {'type': S}}, 'source': {'type': A, 'items': {'type': S}}, 'instruction': {'type': S, 'nullable': True}}},
        'writing': {'type': O, 'nullable': True, 'properties': {'topics': {'type': A, 'items': {'type': O, 'properties': {'label': {'type': S}, 'text': {'type': S}, 'hint': {'type': S, 'nullable': True}}}}}},
        'dictation': {'type': O, 'nullable': True, 'properties': {'title': {'type': S, 'nullable': True}, 'paragraphs': {'type': A, 'items': {'type': S}}, 'source': {'type': S, 'nullable': True}}},
        'answers_list': ans, 'figures': {'type': A, 'items': fig}}, 'required': ['sections']}

def gemini_call(parts_media, mon, extra=''):
    """parts_media: list of (mime, bytes) hoặc ('text', str). Trả về dict JSON. KHÔNG ghi key vào log."""
    import urllib.request, urllib.error
    key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not key: raise NeedAI(AI_OFF_MSG)
    import gemini
    schema = SCHEMA_TEXT.replace('"answers": {"1": {"text":"C","points":"0,5"}}', '"answers_list": [{"so":1,"text":"C","points":"0,5"}]')
    parts = [{'text': PROMPT.format(mon='Toán' if mon == 'toan' else 'Tiếng Việt', schema=schema, extra=extra)}]
    for mime, data in parts_media:
        if mime == 'text': parts.append({'text': 'NỘI DUNG CHỮ CỦA TỆP:\n' + data[:60000]})
        else: parts.append({'inline_data': {'mime_type': mime, 'data': base64.b64encode(data).decode()}})
    body = {'contents': [{'role': 'user', 'parts': parts}],
            'generationConfig': {'temperature': 0, 'responseMimeType': 'application/json', 'responseSchema': _gem_schema()}}
    try:
        res = gemini.post(body, 180)
    except urllib.error.HTTPError as e:
        code = e.code; log.warning('gemini: HTTP %s', code)
        if gemini.special_msg(code): raise ValueError(gemini.special_msg(code))
        if code == 429: raise ValueError('Dịch vụ đọc ảnh đang quá tải hoặc hết lượt miễn phí hôm nay, vui lòng thử lại sau.')
        raise ValueError(f'Không đọc được ảnh (lỗi dịch vụ {code}). Vui lòng thử lại hoặc dùng file Word.')
    except Exception as e:
        log.warning('gemini: lỗi kết nối %s', type(e).__name__); raise ValueError('Không kết nối được dịch vụ đọc ảnh, vui lòng thử lại.')
    try:
        txt = ''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts'])
        txt = re.sub(r'^```(?:json)?\s*|\s*```$', '', txt.strip())
        return json.loads(txt)
    except Exception:
        log.warning('gemini: phản hồi không phải JSON hợp lệ'); raise ValueError('Dịch vụ đọc ảnh trả về dữ liệu không hợp lệ, vui lòng thử lại.')

def _compose_schema():
    """Lược đồ rút gọn cho soạn đề mới: sections + answers_list (có 'giai' cho hướng dẫn chấm tự luận)."""
    S, A, O = 'STRING', 'ARRAY', 'OBJECT'
    blk = {'type': O, 'properties': {'t': {'type': S, 'enum': ['text', 'options', 'lines', 'ds']}, 'text': {'type': S, 'nullable': True},
           'items': {'type': A, 'items': {'type': S}}, 'layout': {'type': S, 'nullable': True}, 'n': {'type': 'INTEGER', 'nullable': True}}, 'required': ['t']}
    qs = {'type': O, 'properties': {'so': {'type': 'INTEGER'}, 'text': {'type': S}, 'muc': {'type': S, 'nullable': True}, 'diem': {'type': S, 'nullable': True},
          'blocks': {'type': A, 'items': blk}}, 'required': ['so', 'text', 'blocks']}
    sec = {'type': O, 'properties': {'title': {'type': S, 'nullable': True}, 'instruction': {'type': S, 'nullable': True}, 'questions': {'type': A, 'items': qs}}, 'required': ['questions']}
    ans = {'type': A, 'items': {'type': O, 'properties': {'so': {'type': 'INTEGER'}, 'text': {'type': S}, 'points': {'type': S, 'nullable': True},
           'giai': {'type': S, 'nullable': True}}, 'required': ['so', 'text']}}
    rd = {'type': O, 'properties': {'intro': {'type': S, 'nullable': True}, 'title': {'type': S, 'nullable': True},
          'paragraphs': {'type': A, 'items': {'type': S}}, 'source': {'type': A, 'items': {'type': S}},
          'instruction': {'type': S, 'nullable': True}}}
    return {'type': O, 'properties': {'sections': {'type': A, 'items': sec}, 'answers_list': ans, 'reading': rd}, 'required': ['sections']}

def gemini_compose(prompt_text):
    """Soạn đề mới bằng AI (chỉ chữ, không đọc ảnh). Trả về dict JSON. KHÔNG ghi key vào log."""
    import urllib.request, urllib.error
    key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not key: raise NeedAI(AI_OFF_MSG)
    import gemini
    body = {'contents': [{'role': 'user', 'parts': [{'text': prompt_text}]}],
            'generationConfig': {'temperature': 0, 'responseMimeType': 'application/json', 'responseSchema': _compose_schema()}}
    try:
        res = gemini.post(body, 180)
    except urllib.error.HTTPError as e:
        code = e.code; log.warning('gemini: HTTP %s', code)
        if gemini.special_msg(code): raise ValueError(gemini.special_msg(code))
        if code == 429: raise ValueError('Dịch vụ AI đang quá tải hoặc hết lượt miễn phí hôm nay, vui lòng thử lại sau.')
        if code in (500, 502, 503, 504): raise ValueError(f'Dịch vụ AI đang quá tải (lỗi dịch vụ {code}), vui lòng thử lại sau ít phút.')
        raise ValueError(f'Không soạn được đề (lỗi dịch vụ {code}). Vui lòng thử lại.')
    except Exception as e:
        log.warning('gemini: lỗi kết nối %s', type(e).__name__); raise ValueError('Không kết nối được dịch vụ AI, vui lòng thử lại.')
    try:
        txt = ''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts'])
        txt = re.sub(r'^```(?:json)?\s*|\s*```$', '', txt.strip())
        return json.loads(txt)
    except Exception:
        log.warning('gemini: phản hồi không phải JSON hợp lệ'); raise ValueError('Dịch vụ AI trả về dữ liệu không hợp lệ, vui lòng thử lại.')

def normalize_compose(d, mon):
    """Chuẩn hoá JSON từ gemini_compose thành exam dict dùng được cho exam.build (không hình, không đọc ảnh)."""
    ex = new_exam(mon); ex['source_kind'] = 'ai'
    for s in d.get('sections') or []:
        qs = []
        for qq in s.get('questions') or []:
            blocks = []
            for b in qq.get('blocks') or []:
                t = b.get('t')
                if t == 'text' and (b.get('text') or '').strip(): blocks.append(dict(t='text', text=b['text'], muc=None, diem=None))
                elif t == 'options' and b.get('items'): blocks.append(dict(t='options', items=[str(x) for x in b['items']][:8], layout='col' if b.get('layout') == 'col' else 'row'))
                elif t == 'lines': blocks.append(dict(t='lines', n=max(1, min(20, int(b.get('n') or 3)))))
                elif t == 'ds' and b.get('items'): blocks.append(dict(t='ds', items=[str(x) for x in b['items']]))
            stem, muc, diem = parse_tag(str(qq.get('text') or ''))
            qs.append(dict(so=int(qq.get('so') or len(qs) + 1), text=stem, muc=qq.get('muc') or muc, diem=(qq.get('diem') or diem or None), blocks=blocks))
        ex['sections'].append(dict(title=s.get('title') or '', instruction=s.get('instruction') or '', questions=qs))
    rd = d.get('reading')
    if isinstance(rd, dict) and rd.get('paragraphs'):
        ex['reading'] = dict(intro=rd.get('intro') or '', title=rd.get('title') or '',
                             paragraphs=[str(x) for x in rd['paragraphs'] if str(x).strip()],
                             source=[str(x) for x in (rd.get('source') or []) if str(x).strip()],
                             instruction=rd.get('instruction') or '')
    for a in d.get('answers_list') or []:
        if a.get('so'):
            ans = dict(text=a.get('text') or '', points=a.get('points') or '')
            if (a.get('giai') or '').strip(): ans['giai'] = a['giai']
            ex['answers'][str(a['so'])] = ans
    return ex

def normalize_ai(d, mon, images, workdir):
    """Kiểm tra + chuẩn hoá JSON từ Gemini; vẽ lại hình hình học (SVG->PNG 300dpi), cắt tranh từ ảnh gốc."""
    ex = new_exam(mon); ex['source_kind'] = 'ai'
    for s in d.get('sections') or []:
        qs = []
        for qq in s.get('questions') or []:
            blocks = []
            for b in qq.get('blocks') or []:
                t = b.get('t')
                if t == 'text' and (b.get('text') or '').strip(): blocks.append(dict(t='text', text=b['text'], muc=b.get('muc'), diem=b.get('diem')))
                elif t == 'options' and b.get('items'): blocks.append(dict(t='options', items=[str(x) for x in b['items']][:8], layout='col' if b.get('layout') == 'col' else 'row'))
                elif t == 'lines': blocks.append(dict(t='lines', n=max(1, min(20, int(b.get('n') or 1)))))
                elif t == 'ds' and b.get('items'): blocks.append(dict(t='ds', items=[str(x) for x in b['items']]))
                elif t == 'figure' and b.get('id'): blocks.append(dict(t='figure', id=str(b['id'])))
            stem, muc, diem = parse_tag(str(qq.get('text') or ''))
            qs.append(dict(so=int(qq.get('so') or len(qs) + 1), text=stem, muc=qq.get('muc') or muc, diem=(qq.get('diem') or diem or None), blocks=blocks))
        ex['sections'].append(dict(title=s.get('title') or '', instruction=s.get('instruction') or '', questions=qs))
    if mon == 'tv':
        for k in ('reading', 'writing', 'dictation'):
            if isinstance(d.get(k), dict):
                for kk, vv in d[k].items():
                    if kk in ex[k] and vv: ex[k][kk] = vv
    for a in d.get('answers_list') or []:
        if a.get('so'): ex['answers'][str(a['so'])] = dict(text=a.get('text') or '', points=a.get('points') or '')
    for k, v in (d.get('answers') or {}).items() if isinstance(d.get('answers'), dict) else []:
        ex['answers'][str(k)] = dict(text=(v or {}).get('text', ''), points=(v or {}).get('points', ''))
    figdir = os.path.join(workdir, 'figs'); os.makedirs(figdir, exist_ok=True)
    for f in d.get('figures') or []:
        fid = str(f.get('id') or ''); 
        if not fid: continue
        meta = dict(kind=f.get('kind'), w_cm=f.get('w_cm'))
        try:
            pg = int(f.get('page') or 0); box = f.get('box')
            if box and len(box) == 4 and 0 <= pg < len(images):
                meta['orig'] = crop_image(images[pg], box, os.path.join(figdir, fid + '_goc.png'))
            if f.get('kind') == 'geometry' and f.get('svg'):
                meta['file'] = svg_to_png(f['svg'], os.path.join(figdir, fid + '.png'), f.get('w_cm'))
                meta['svg'] = f['svg']
            elif meta.get('orig'):
                meta['file'] = meta['orig']
        except Exception as e:
            log.warning('figure %s: %s', fid, type(e).__name__)
        if meta.get('file'): ex['figures'][fid] = meta
        else: ex['warnings'].append(f'Không dựng được hình {fid} – hãy kiểm tra lại câu có hình.')
    return ex

def svg_to_png(svg, out, w_cm=None, dpi=300):
    import cairosvg
    svg = svg.strip()
    if not svg.startswith('<svg'): svg = svg[svg.find('<svg'):]
    root = etree.fromstring(svg.encode())
    if root.get('viewBox') and not root.get('width'):
        vb = [float(x) for x in re.split(r'[\s,]+', root.get('viewBox').strip())]
        wmm = (w_cm or 0) * 10 or vb[2]; root.set('width', f'{wmm}mm'); root.set('height', f'{wmm * vb[3] / vb[2]}mm')
    elif w_cm and root.get('width') and root.get('height'):
        try:
            w0, h0 = float(re.sub(r'[a-z]+', '', root.get('width'))), float(re.sub(r'[a-z]+', '', root.get('height')))
            if not root.get('viewBox'): root.set('viewBox', f'0 0 {w0} {h0}')
            root.set('width', f'{w_cm * 10}mm'); root.set('height', f'{w_cm * 10 * h0 / w0}mm')
        except ValueError: pass
    cairosvg.svg2png(bytestring=etree.tostring(root), write_to=out, dpi=dpi, background_color='white')
    return out

def crop_image(src, box, out, pad=0.01):
    from PIL import Image
    im = Image.open(src); w, h = im.size
    y0, x0, y1, x1 = [float(v) / 1000 for v in box]
    x0, y0 = max(0, x0 - pad), max(0, y0 - pad); x1, y1 = min(1, x1 + pad), min(1, y1 + pad)
    if x1 - x0 < 0.01 or y1 - y0 < 0.01: raise ValueError('box')
    im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h))).convert('RGB').save(out)
    return out

def pdf_to_images(path, outdir, dpi=200, max_pages=12):
    import pymupdf
    doc = pymupdf.open(path); out = []
    for i, page in enumerate(doc):
        if i >= max_pages: break
        fn = os.path.join(outdir, f'pdfpage_{len(os.listdir(outdir))}_{i}.png'); page.get_pixmap(dpi=dpi).save(fn); out.append(fn)
    return out

def doc_text(path):
    units, _, _ = docx_units(path)
    out = []
    for u in units:
        if u['kind'] == 'p': out.append(u['text'])
        else: out += ['\t'.join(r) for r in u['rows']]
    return '\n'.join(out)

def extract(files, mon, workdir, ai_hook=None):
    """files: list of (slot, path). Trả về JSON đề. ai_hook(): gọi trước mỗi lần dùng Gemini (giới hạn lượt)."""
    os.makedirs(workdir, exist_ok=True); imgdir = os.path.join(workdir, 'pages_src'); os.makedirs(imgdir, exist_ok=True)
    parsed, images, ai_parts, ai_slots = [], [], [], []
    for slot, p in files:
        n_before = len(ai_parts)
        ext = os.path.splitext(p)[1].lower()
        if ext == '.docx':
            ex = extract_docx(p, mon, workdir, slot)
            if count_questions(ex) == 0 and not ex.get('reading', {}).get('paragraphs') and gemini_enabled():
                ai_parts.append(('text', doc_text(p)))
            else: parsed.append((slot, ex))
        elif ext == '.pdf':
            ex = extract_pdf_text(p, mon, slot, workdir)
            if ex is not None and (count_questions(ex) or (mon == 'tv' and (ex['reading']['paragraphs'] or ex['dictation']['paragraphs'] or ex['writing']['topics']))):
                parsed.append((slot, ex))
            else:
                if not gemini_enabled(): raise NeedAI(AI_OFF_MSG + ' (PDF dạng ảnh quét).')
                for im in pdf_to_images(p, imgdir): images.append(im); ai_parts.append(('image/png', open(im, 'rb').read()))
        elif ext in IMG_EXT:
            if not gemini_enabled(): raise NeedAI(AI_OFF_MSG)
            from PIL import Image
            fn = os.path.join(imgdir, f'img_{len(images)}.png'); Image.open(p).convert('RGB').save(fn)
            images.append(fn); ai_parts.append(('image/png', open(fn, 'rb').read()))
        elif ext == '.doc': raise ValueError('Tệp .doc (Word cũ) chưa hỗ trợ – hãy lưu lại thành .docx.')
        else: raise ValueError(f'Không hỗ trợ tệp {os.path.basename(p)}')
        ai_slots += [slot] * (len(ai_parts) - n_before)
    if ai_parts:
        if ai_hook: ai_hook()
        extra = 'Các ảnh được đánh số theo thứ tự đính kèm (page = 0, 1, 2…).' if images else ''
        names = {'doc_hieu': 'bài đọc + câu hỏi đọc hiểu', 'de_viet': 'đề viết đoạn (writing)', 'nghe_viet': 'bài nghe - viết (dictation)'}
        hints = [f'Tệp/ảnh thứ {k}: {names[sl]}' for k, sl in enumerate(ai_slots) if sl in names]
        if hints: extra += '\n' + '; '.join(hints) + '.'
        parsed.append(('ai', normalize_ai(gemini_call(ai_parts, mon, extra), mon, images, workdir)))
    return merge_exams([e for _, e in parsed], mon)

def merge_exams(exs, mon):
    if not exs: raise ValueError('Chưa có nội dung đề – hãy tải lên ít nhất một tệp.')
    out = new_exam(mon)
    for e in exs:
        for s in e['sections']:
            if not s['questions'] and not s['title']: continue
            if out['sections'] and not s['title'] and out['sections'][-1]['questions'] is not None: out['sections'][-1]['questions'] += s['questions']
            else: out['sections'].append(copy.deepcopy(s))
        for k in ('answers', 'figures', 'raw'): out[k].update(e.get(k) or {})
        out['warnings'] += e.get('warnings') or []
        if mon == 'tv':
            for k in ('reading', 'writing', 'dictation'):
                for kk, vv in e[k].items():
                    if vv and not out[k][kk]: out[k][kk] = vv
    out['source_kind'] = '+'.join(sorted({e.get('source_kind', '?') for e in exs}))
    return out

# =====================================================================================
# 6. Dựng đề trên đề mẫu
# =====================================================================================
def body_list(body): return [e for e in body if e.tag in (q('p'), q('tbl'))]

def text_width_tw(sect):
    pg = sect.find(q('pgSz')); mar = sect.find(q('pgMar'))
    try: return int(pg.get(q('w'))) - int(mar.get(q('left'))) - int(mar.get(q('right')))
    except Exception: return 10000

def run_with(p, pred):
    for r in p.findall(q('r')):
        if pred(rtext(r)): return r
    return None

_TL = threading.local()   # cỡ chữ mặc định của đề mẫu đang dựng – riêng từng luồng/yêu cầu, không dùng chung
def doc_default_size(doc):
    try:
        st = doc.part.styles.element
        nm = next((x for x in st.findall(q('style')) if x.get(q('default')) == '1' and x.get(q('type')) == 'paragraph'), None)
        for e in ([nm] if nm is not None else []) + [st.find(q('docDefaults'))]:
            sz = e.find('.//' + q('sz')) if e is not None else None
            if sz is not None: return int(sz.get(q('val'))) / 2
    except Exception: pass
    return 11.0

def pr_size(rpr, default=None):
    default = default or getattr(_TL, 'default_sz', 14.0)
    if rpr is None: return default
    sz = rpr.find(q('sz'))
    return int(sz.get(q('val'))) / 2 if sz is not None else default

def is_qtitle(e): return e.tag == q('p') and bool(QSTART.match(ptext(e)))

class Protos:
    """Lấy đoạn/bảng mẫu (pPr/rPr) từ phần câu hỏi của đề mẫu."""
    def __init__(s, region, width):
        s.width = width
        ps = [e for e in region if e.tag == q('p')]
        if not any(is_qtitle(p) for p in ps):
            # Mẫu không có câu hỏi mẫu (vd: đề Viết) — prototype rỗng, chỉ dùng cho các phần khác
            s.title = s.text = s.sub = s.lines = s.opt_col = s.opt_row = None
            s.r_label = s.r_tag = s.r_stem = s.r_text = s.r_lines = None
            s.tag_sep = ') ('; s.tag_unit = 'đ'; s.dec = ','
            s.dot_char = '.'; s.lines_single = False; s.opt_tbl = None
            s._empty = True
            return
        s._empty = False
        s.title = next((p for p in ps if is_qtitle(p) and TAG_RE.search(ptext(p)) and len(runs_of(p)) >= 3), None)
        if s.title is None: s.title = next(p for p in ps if is_qtitle(p))
        rs = [r for r in s.title.findall(q('r')) if rtext(r)]
        lab = next(r for r in rs if 'Câu' in rtext(r)); s.r_label = run_rpr(lab)
        tagr = next((r for r in rs if re.search(r'\(\s*M|^\s*\(\s*$', rtext(r))), rs[-1]); s.r_tag = run_rpr(tagr)
        stem = next((r for r in rs[rs.index(lab) + 1:] if len(re.findall(r'\w', rtext(r))) >= 4 and r is not tagr), None)
        s.r_stem = run_rpr(stem) if stem is not None else run_rpr(lab)
        m = TAG_RE.search(ptext(s.title)); full = m.group(0) if m else '(M1) (0,5đ)'
        mm = re.match(r'\(\s*M\s*\d(.*?)([\d]+(?:[.,]\d+)?)\s*(đ|điểm)\s*\)', full)
        s.tag_sep = mm.group(1) if mm else ') ('; s.tag_unit = mm.group(3) if mm else 'đ'
        s.dec = '.' if mm and '.' in mm.group(2) else ','
        def plain(p):
            t = ptext(p)
            return (t.strip() and not is_qtitle(p) and not SEC_RE.match(t.strip()) and not (t.strip().isupper() and len(t.strip()) > 3) and not has_fig(p) and not is_dots(t) and not split_opts(t.strip()) and 'Câu' not in t
                    and (p.find(q('pPr')) is None or p.find(q('pPr')).find(q('numPr')) is None))
        cands = [p for p in ps if plain(p)]
        def jc(p):
            j = p.find(q('pPr')).find(q('jc')) if p.find(q('pPr')) is not None else None
            return j.get(q('val')) if j is not None else 'left'
        good = [p for p in cands if not TAG_RE.search(ptext(p)) and jc(p) not in ('center', 'right')]
        good.sort(key=lambda p: (ptext(p).startswith((' ', '\t')), any(r.find(q('rPr')) is not None and r.find(q('rPr')).find(q('b')) is not None for r in p.findall(q('r')))))
        s.text = good[0] if good else next((p for p in cands if not TAG_RE.search(ptext(p))), cands[0] if cands else s.title)
        tr = next((r for r in s.text.findall(q('r')) if rtext(r).strip()), None); s.r_text = run_rpr(tr)
        sub = next((p for p in cands if TAG_RE.search(ptext(p))), None)
        s.sub = sub if sub is not None else s.text
        s.lines = next((p for p in ps if is_dots(ptext(p))), None)
        if s.lines is not None:
            dr = run_with(s.lines, lambda t: is_dots(t)); s.r_lines = run_rpr(dr); dt = ptext(s.lines).strip()
            s.dot_char = '…' if dt.count('…') > dt.count('.') else '.'
            s.lines_single = text_tw(dt, pr_size(s.r_lines)) > width * 1.5   # 1 đoạn chấm dài tự xuống dòng
        else: s.r_lines, s.dot_char, s.lines_single = s.r_text, '.', False
        # phương án dạng bảng 1 hàng
        s.opt_tbl = None
        for t in (e for e in region if e.tag == q('tbl')):
            rows = t.findall(q('tr'))
            if len(rows) == 1 and not has_fig(t):
                cells = rows[0].findall(q('tc'))
                if len(cells) == 4 and all(re.match(r'^\s*[A-D]\s*\.', cell_text(c)) for c in cells): s.opt_tbl = t; break
        # phương án theo cột / theo hàng (đoạn)
        single = [p for p in ps if (lambda o: o is not None and len(o) == 1)(split_opts(ptext(p).strip()))]
        multi = [p for p in ps if (lambda o: o is not None and len(o) >= 2)(split_opts(ptext(p).strip()))]
        s.opt_col = single[0] if single else (multi[0] if multi else s.text)
        s.opt_row = multi[0] if multi else s.opt_col
        def opt_runs(p):
            rs = [r for r in p.findall(q('r'))]
            li = next((i for i, r in enumerate(rs) if re.match(r'^\s*A\s*\.', rtext(r))), None)
            if li is None: return run_rpr(rs[0]) if rs else None, s.r_text, s.r_text, ''
            lead = ''.join(rtext(r) for r in rs[:li]) + re.match(r'^(\s*)', rtext(rs[li])).group(1)
            after = next((r for r in rs[li + 1:] if rtext(r).strip()), None)
            lr = run_rpr(rs[li]); txr = run_rpr(after) if after is not None and not re.match(r'^\s*[A-D]\s*\.', rtext(after)) else lr
            if re.match(r'^\s*A\s*\.\s*\S', rtext(rs[li])): txr = lr
            return run_rpr(rs[li - 1]) if li else lr, lr, txr, lead
        s.col_lead_r, s.col_letter_r, s.col_text_r, s.col_lead = opt_runs(s.opt_col)
        s.row_lead_r, s.row_letter_r, s.row_text_r, s.row_lead = opt_runs(s.opt_row)
        s.empty = next((p for p in ps if is_empty_p(p)), None)
        s.subpar, s.sub_ind = None, None   # đoạn "a) …" của mẫu (danh sách chữ cái) -> lấy pPr + thụt lề của mức danh sách
        for p in ps:
            ppr = p.find(q('pPr')); n = ppr.find(q('numPr')) if ppr is not None else None
            if n is not None and not is_qtitle(p) and not split_opts(ptext(p).strip()) and ptext(p).strip():
                s.subpar = p; s.sub_num = n; break

    def set_numbering(s, numbering_el):
        if s.subpar is None or numbering_el is None: return
        nid = s.sub_num.find(q('numId')); il = s.sub_num.find(q('ilvl'))
        nid = nid.get(q('val')) if nid is not None else None; il = il.get(q('val')) if il is not None else '0'
        num = next((x for x in numbering_el.findall(q('num')) if x.get(q('numId')) == nid), None)
        if num is None: return
        aid = num.find(q('abstractNumId')).get(q('val'))
        ab = next((x for x in numbering_el.findall(q('abstractNum')) if x.get(q('abstractNumId')) == aid), None)
        lvl = next((x for x in ab.findall(q('lvl')) if x.get(q('ilvl')) == il), None) if ab is not None else None
        ind = lvl.find(q('pPr')).find(q('ind')) if lvl is not None and lvl.find(q('pPr')) is not None else None
        if ind is not None:
            left = int(ind.get(q('left')) or ind.get(q('start')) or 0); hang = int(ind.get(q('hanging')) or 0)
            s.sub_ind = max(0, left - hang)

    def tag(s, muc, diem):
        if not muc and not diem: return ''
        d = (diem or '').replace('.', s.dec).replace(',', s.dec)
        m = re.sub(r'\D', '', muc or '') or '?'
        if not d: return f'(M{m})'
        return f'(M{m}{s.tag_sep}{d}{s.tag_unit})'

def empty_par(P):
    if P.empty is not None:
        e = clean_ids(copy.deepcopy(P.empty))
        for r in e.findall(q('r')): e.remove(r)
        return e
    return new_par(P.text, [])

def pagebreak_par(P):
    p = empty_par(P); r = etree.SubElement(p, q('r')); etree.SubElement(r, q('br')).set(q('type'), 'page'); return p

def opt_label(i): return chr(65 + i) + '.'

class Builder:
    def __init__(s, doc, exam, P, media_dirs, fix=None, height=14000):
        s.doc, s.exam, s.P, s.media_dirs = doc, exam, P, media_dirs; s.warn = []
        s.fix = fix or {}; s.height = height; s.cur_shrink = 1.0; s.raw_tbls = []
        try: s.num_ids = {n.get(q('numId')) for n in doc.part.numbering_part.element.findall(q('num'))}
        except Exception: s.num_ids = set()

    # ---- câu hỏi -> danh sách phần tử ----
    def question(s, Q):
        P = s.P; out = []
        tag = P.tag(Q.get('muc'), Q.get('diem'))
        segs = [(P.r_label, f"Câu {Q['so']}."), (P.r_stem, ' ' + (Q.get('text') or '').strip())]
        if tag: segs.append((P.r_tag, ' ' + tag))
        out.append(clean_ids(new_par(P.title, segs)))
        s.cur_shrink = (s.fix.get('shrink') or {}).get(Q['so'], 1.0)
        kinds = []
        for b in Q.get('blocks') or []:
            els = s.block(b); out += els; kinds += [b.get('t')] * len(els)
        s.cur_shrink = 1.0
        s.keep_together(out, kinds, Q['so'] in (s.fix.get('breaks') or ()))
        return out

    def keep_together(s, out, kinds, brk):
        """EXAM_RULES §6: tiêu đề câu luôn đi cùng dòng sau; câu có hình: tiêu đề + hình + phương án giữ liền (keepNext/keepLines)
        -> hình không vừa trang thì cả cụm sang trang sau. brk: vòng kiểm tra yêu cầu ngắt trang trước câu này."""
        figs = [i for i, e in enumerate(out[1:], 1) if has_fig(e) or kinds[i - 1] == 'figure']
        last = max(figs) if figs else 0
        j = last + 1
        while figs and j < len(out) and kinds[j - 1] in ('space', 'options', 'ds'):
            if kinds[j - 1] in ('options', 'ds'): last = j
            j += 1
        for i, e in enumerate(out[:max(1, last + 1)]):
            ps = [e] if e.tag == q('p') else [p for p in e.iter(q('p')) if not any(x.tag == q('txbxContent') for x in p.iterancestors())]
            for p in ps:
                ppr_set(p, 'keepLines')
                if i < last or (i == 0 and len(out) > 1): ppr_set(p, 'keepNext')
        if brk: ppr_set(out[0], 'pageBreakBefore')

    def block(s, b):
        P = s.P; t = b.get('t')
        if t == 'text':
            tag = P.tag(b.get('muc'), b.get('diem')); res = []
            lines = (b.get('text') or '').split('\n')
            for k, ln in enumerate(lines):
                segs = [(P.r_text, ln)]
                if tag and k == len(lines) - 1: segs.append((P.r_tag, ' ' + tag))
                if P.subpar is not None and re.match(r'^[a-đ]\s*\)\s', ln) and not tag:
                    p = new_par(P.subpar, segs); ppr = p.find(q('pPr'))
                    if P.sub_ind is not None and ppr is not None:
                        for x in ppr.findall(q('ind')): ppr.remove(x)
                        ind = etree.Element(q('ind')); ind.set(q('left'), str(P.sub_ind)); ind.set(q('firstLine'), '0')
                        pos = [c.tag for c in ppr]; after = max([i for i, t in enumerate(pos) if t in (q('pStyle'), q('keepNext'), q('keepLines'), q('pageBreakBefore'), q('framePr'), q('widowControl'), q('numPr'), q('suppressLineNumbers'), q('pBdr'), q('shd'), q('tabs'), q('suppressAutoHyphens'), q('kinsoku'), q('wordWrap'), q('overflowPunct'), q('topLinePunct'), q('autoSpaceDE'), q('autoSpaceDN'), q('bidi'), q('adjustRightInd'), q('snapToGrid'), q('spacing'))] or [-1])
                        ppr.insert(after + 1, ind)
                    res.append(p)
                else: res.append(new_par(P.sub if tag else P.text, segs))
            return res
        if t == 'options': return s.options(b.get('items') or [], b.get('layout') or 'row')
        if t == 'lines': return s.lines(int(b.get('n') or 1))
        if t == 'ds': return s.ds(b.get('items') or [])
        if t == 'figure': return s.figure(b.get('id'))
        if t == 'raw': return s.raw(b.get('id'))
        if t == 'space': return [empty_par(P) for _ in range(min(int(b.get('n') or 1), 20))]
        return []

    def options(s, items, layout):
        P = s.P; items = [str(x).strip() for x in items]
        sz = pr_size(P.col_text_r)
        if layout == 'row' and P.opt_tbl is not None and 1 < len(items) <= 4:
            cells = P.opt_tbl.find(q('tr')).findall(q('tc'))
            cw = [int((c.find(q('tcPr')).find(q('tcW')).get(q('w')) if c.find(q('tcPr')) is not None and c.find(q('tcPr')).find(q('tcW')) is not None else P.width / 4)) for c in cells]
            if all(text_tw('A. ' + it, sz) < cw[i] - 250 for i, it in enumerate(items)):
                t = clean_ids(copy.deepcopy(P.opt_tbl))
                for i, c in enumerate(t.find(q('tr')).findall(q('tc'))):
                    ps = c.findall(q('p')); p0 = ps[0]
                    for x in ps[1:]: c.remove(x)
                    rs = p0.findall(q('r')); lr = next((r for r in rs if re.match(r'^\s*[A-D]', rtext(r))), rs[0] if rs else None)
                    after = rs[rs.index(lr) + 1:] if lr is not None else []
                    tr_ = next((r for r in reversed(after) if rtext(r).strip()), None)
                    lrpr, trpr = run_rpr(lr), run_rpr(tr_) if tr_ is not None else run_rpr(lr)
                    for r in rs: p0.remove(r)
                    if i < len(items):
                        p0.append(make_run(lrpr, chr(65 + i))); p0.append(make_run(trpr, '. ' + items[i]))
                return [t]
        if layout == 'row' and 1 < len(items) <= 4:
            colw = P.width / len(items); tabw = 720; lead = P.row_lead
            pos = text_tw(lead.replace('\t', ''), sz) + lead.count('\t') * tabw
            segs = [(P.row_lead_r, lead)] if lead else []; ok = True
            for i, it in enumerate(items):
                target = i * colw
                if i:
                    if pos > target - 100: target = pos + 300
                    n = max(1, math.ceil((target - pos) / tabw)); segs.append((P.row_text_r, '\t' * n)); pos = math.ceil(pos / tabw + 0.0001) * tabw + (n - 1) * tabw
                segs += [(P.row_letter_r, opt_label(i)), (P.row_text_r, ' ' + it)]
                pos += text_tw(opt_label(i) + ' ' + it, sz)
                if pos > P.width - 200: ok = False
            if ok: return [new_par(P.opt_row, segs)]
        return [new_par(P.opt_col, ([(P.col_lead_r, P.col_lead)] if P.col_lead else []) + [(P.col_letter_r, opt_label(i)), (P.col_text_r, ' ' + it)]) for i, it in enumerate(items)]

    def lines(s, n):
        P = s.P; proto = P.lines if P.lines is not None else P.text; sz = pr_size(P.r_lines)
        per = max(10, int((P.width - 150) / max(1.0, text_tw(P.dot_char * 50, sz) / 50)))
        if P.lines_single: return [new_par(proto, [(P.r_lines, P.dot_char * (per * n))])]
        return [new_par(proto, [(P.r_lines, P.dot_char * per)]) for _ in range(n)]

    def ds(s, items):
        P = s.P; res = []
        box_w = 560; text_w = P.width - box_w - 200
        tbl = etree.SubElement(etree.Element('x'), q('tbl'))
        tp = etree.SubElement(tbl, q('tblPr')); etree.SubElement(tp, q('tblW')).attrib.update({q('w'): str(text_w + box_w), q('type'): 'dxa'})
        lay = etree.SubElement(tp, q('tblLayout')); lay.set(q('type'), 'fixed')
        grid = etree.SubElement(tbl, q('tblGrid'))
        for w in (text_w, box_w): etree.SubElement(grid, q('gridCol')).set(q('w'), str(w))
        for it in items:
            tr = etree.SubElement(tbl, q('tr')); trp = etree.SubElement(tr, q('trPr')); h = etree.SubElement(trp, q('trHeight')); h.set(q('val'), '460')
            for k, w in enumerate((text_w, box_w)):
                tc = etree.SubElement(tr, q('tc')); tcp = etree.SubElement(tc, q('tcPr'))
                etree.SubElement(tcp, q('tcW')).attrib.update({q('w'): str(w), q('type'): 'dxa'})
                b = etree.SubElement(tcp, q('tcBorders'))
                for side in ('top', 'left', 'bottom', 'right'):
                    e = etree.SubElement(b, q(side)); e.attrib.update({q('val'): 'single' if k else 'nil', q('sz'): '8', q('color'): '000000'})
                etree.SubElement(tcp, q('vAlign')).set(q('val'), 'center')
                p = new_par(P.text, [(P.r_text, it if k == 0 else '')])
                ppr = p.find(q('pPr'))
                if ppr is not None:
                    for x in ppr.findall(q('ind')) + ppr.findall(q('spacing')) + ppr.findall(q('jc')): ppr.remove(x)
                tc.append(p)
            sp = etree.SubElement(tbl, q('tr')); sp.append(copy.deepcopy(tr.find(q('trPr')))); sp.find(q('trPr')).find(q('trHeight')).set(q('val'), '80')
            for k, w in enumerate((text_w, box_w)):
                tc = etree.SubElement(sp, q('tc')); tcp = etree.SubElement(tc, q('tcPr')); etree.SubElement(tcp, q('tcW')).attrib.update({q('w'): str(w), q('type'): 'dxa'})
                b = etree.SubElement(tcp, q('tcBorders'))
                for side in ('top', 'left', 'bottom', 'right'): etree.SubElement(b, q(side)).attrib.update({q('val'): 'nil'})
                p = etree.SubElement(tc, q('p')); pp = etree.SubElement(p, q('pPr')); spc = etree.SubElement(pp, q('spacing')); spc.attrib.update({q('before'): '0', q('after'): '0', q('line'): '80', q('lineRule'): 'exact'})
        res.append(tbl)
        return res

    def media_path(s, fn):
        for d in s.media_dirs:
            p = os.path.join(d, fn)
            if os.path.exists(p): return p
        return None

    def figure(s, fid):
        from docx.text.paragraph import Paragraph
        from docx.shared import Emu
        from PIL import Image
        meta = (s.exam.get('figures') or {}).get(fid)
        path = s.media_path(meta.get('file') and os.path.basename(meta['file'])) if meta else None
        if not path: s.warn.append(f'Thiếu hình {fid}'); return []
        p = new_par(s.P.text, []); ppr = p.find(q('pPr'))
        if ppr is None: ppr = etree.SubElement(p, q('pPr'))
        for x in ppr.findall(q('ind')) + ppr.findall(q('jc')): ppr.remove(x)
        sp = ppr.find(q('spacing'))
        if sp is not None: ppr.remove(sp)
        jc = etree.SubElement(ppr, q('jc')); jc.set(q('val'), 'center')
        im = Image.open(path); dpi = im.info.get('dpi', (300, 300))[0] or 300
        if meta.get('kind') != 'geometry' and not meta.get('w_cm'): dpi = 200
        w_emu = int(im.size[0] / float(dpi) * 914400)
        if meta.get('w_cm'): w_emu = int(float(meta['w_cm']) * 360000)
        w_emu = max(360000 * 2, min(w_emu, int(s.P.width * 635 * FIG_SAFE)))
        h_max = s.height * 635 * FIG_MAX_H; h_emu = w_emu * im.size[1] / max(1, im.size[0])
        if h_emu > h_max: w_emu = int(w_emu * h_max / h_emu)
        w_emu = int(w_emu * s.cur_shrink)
        inline = s.doc.part.new_pic_inline(path, Emu(w_emu), None)
        r = etree.SubElement(p, q('r')); etree.SubElement(r, q('drawing')).append(inline)
        return [p]

    def raw(s, rid): return s.raw_dict((s.exam.get('raw') or {}).get(rid))

    def raw_dict(s, d):
        if not d: return []
        el = etree.fromstring(d['xml'])
        imgs = d.get('images') or {}
        for x in el.iter():
            for a, v in list(x.attrib.items()):
                if a.startswith('{%s}' % R):
                    path = s.media_path(imgs.get(v, '')) if imgs.get(v) else None
                    if path: x.set(a, s.doc.part.get_or_add_image(path)[0])
                    else: del x.attrib[a]
        for n in list(el.iter(q('numPr'))):
            nid = n.find(q('numId'))
            if nid is None or nid.get(q('val')) not in s.num_ids: n.getparent().remove(n)
        fit_drawings(el, s.P.width, s.height, s.cur_shrink)
        if el.tag == q('tbl'): s.raw_tbls.append(el)
        return [clean_ids(el)]

def renumber_docpr(root):
    """Chỉ đổi id trùng (giữ nguyên id của phần mẫu, ví dụ bảng tiêu đề)."""
    WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'
    es = list(root.iter(WP)); ids = [int(e.get('id')) for e in es if (e.get('id') or '').isdigit()]
    seen = set(); nxt = max(ids or [0]) + 1
    for e in es:
        v = e.get('id')
        if v in seen or not (v or '').isdigit(): e.set('id', str(nxt)); nxt += 1
        seen.add(e.get('id'))

_docx_lock = threading.Lock()
def open_docx(path):
    """Mở .docx bằng python-docx nhưng GIỮ khoảng trắng (python-docx mặc định xoá nút chữ chỉ có khoảng trắng)."""
    import docx, docx.oxml.parser as P
    with _docx_lock:
        old = P.oxml_parser
        keep = etree.XMLParser(remove_blank_text=False, resolve_entities=False); keep.set_element_class_lookup(P.element_class_lookup)
        P.oxml_parser = keep
        try: return docx.Document(path)
        finally: P.oxml_parser = old

def strip_before(el, keep_one_break, P):
    """EXAM_RULES §4: bỏ đoạn trống tràn trước el; nếu keep_one_break: giữ 1 đoạn trống + ngắt trang (w:br)."""
    prev = el.getprevious(); removed = []
    while prev is not None and is_empty_p(prev):
        nx = prev.getprevious(); removed.append(prev); prev = nx
    if prev is not None and prev.tag == q('p') and has_pagebreak(prev) and not ptext(prev).strip():
        keep_one_break = False   # đã có ngắt trang
    keep = removed[-1] if (removed and keep_one_break) else None
    for p in removed:
        if p is not keep: p.getparent().remove(p)
    if keep_one_break:
        if keep is None: keep = empty_par(P); el.addprevious(keep)
        r = etree.SubElement(keep, q('r')); etree.SubElement(r, q('br')).set(q('type'), 'page')

def set_after_label(p, text):
    """Thay phần chữ sau dấu ':' đầu tiên (giữ nhãn 'Đề 1:' / 'Câu 3:')."""
    full = ''.join(t.text or '' for t in p.iter(q('t')))
    if ':' not in full: return set_par_text(p, text)
    if re.search(r'(?<=:)\s*\S.*$', full): _sub_runs(p, r'(?<=:)\s*\S.*$', ' ' + text if text else '')
    elif text: p.append(make_run(run_rpr(p.findall(q('r'))[-1]) if p.findall(q('r')) else None, ' ' + text))

def regen_list(paras, texts, setter):
    """Dùng lại các đoạn mẫu cho texts; thừa thì xoá, thiếu thì nhân bản đoạn cuối."""
    if not paras: return
    last = paras[-1]
    for i, t in enumerate(texts):
        if i < len(paras): setter(paras[i], t)
        else:
            c = clean_ids(copy.deepcopy(last)); setter(c, t); (paras[i - 1] if i - 1 < len(paras) else last).addnext(c); paras.append(c)
    for p in paras[len(texts):]:
        if p.getparent() is not None: p.getparent().remove(p)

def build(template, exam, fields, out_path, media_dirs=(), with_answers=True, fix=None):
    """Dựng đề. Trả về (warns, hlog). fix: gợi ý của vòng kiểm tra khung (build_checked)."""
    r = _build(template, exam, fields, out_path, media_dirs, with_answers, fix)
    return (r[0], r[1], r[2]) if fix is not None else (r[0], r[1])

def _build(template, exam, fields, out_path, media_dirs=(), with_answers=True, fix=None):
    doc = open_docx(template); body = doc.element.body; sect = body.find(q('sectPr'))
    fit_page_frame(sect)   # EXAM_RULES §6: khung viền trang không lấn vào vùng chữ
    width = text_width_tw(sect)
    els = body_list(body)
    kind = (exam.get('kind') or 'both') if exam.get('mon') == 'tv' else None
    qidx = [i for i, e in enumerate(els) if is_qtitle(e)]
    if not qidx and not (exam.get('mon') == 'tv' and kind == 'viet'):
        raise ValueError('Đề mẫu không có câu hỏi mẫu')
    hdr_tbls = [e for e in els if e.tag == q('tbl') and 'UBND' in ptext(e.find(q('tr')).find(q('tc')).find(q('p'))) + cell_text(e.find(q('tr')).find(q('tc')))]
    ans_i = next((i for i, e in enumerate(els) if e.tag == q('p') and re.match(r'^\s*(ĐÁP ÁN|HƯỚNG DẪN CHẤM)', ptext(e))), len(els))
    end_q = next((i for i, e in enumerate(els) if i > qidx[0] and (e in hdr_tbls or i == ans_i)), ans_i) if qidx else ans_i
    _TL.default_sz = doc_default_size(doc)
    P = Protos(els[qidx[0]:end_q] if qidx else [], width)
    try: P.set_numbering(doc.part.numbering_part.element)
    except Exception: pass
    B = Builder(doc, exam, P, list(media_dirs), fix, text_height_tw(sect))
    if exam['mon'] == 'tv':
        if kind == 'doc': build_tv_doc(B, els, qidx[0], end_q, hdr_tbls, ans_i, with_answers)
        elif kind == 'viet': build_tv_viet(B, els, end_q, hdr_tbls, ans_i, with_answers)
        else: build_tv(B, els, qidx[0], end_q, hdr_tbls, ans_i, with_answers)
    else: build_toan(B, els, qidx[0], end_q, ans_i, with_answers)
    # Ngoại lệ §4: mỗi bảng tiêu đề sau bài 1 bắt đầu trang mới
    for t in hdr_tbls[1:]:
        if t.getparent() is None: continue
        strip_before(t, True, P)
        first = [tc.findall(q('p')) for tc in hdr_tbls[0].find(q('tr')).findall(q('tc'))]
        for ci, tc in enumerate(t.find(q('tr')).findall(q('tc'))):
            ps = tc.findall(q('p'))
            if len(ps) > 1 and is_empty_p(ps[0]) and ci < len(first) and first[ci] and not is_empty_p(first[ci][0]): tc.remove(ps[0])
    hlog = fill_header(doc.element, fields)
    # chốt an toàn: tên trường mẫu không được lọt ra đề (chỉ thay chữ trong run, giữ định dạng)
    nm = clean_name(fields.get('truong'), ['TRƯỜNG']); nm = re.sub(r'^(TH & THCS|TH&THCS|TIỂU HỌC|TH|THCS)\s+', '', nm)
    for p in doc.element.iter(q('p')):
        if 'PH' in ''.join(t.text or '' for t in p.iter(q('t'))).upper():
            hlog += _sub_runs(p, r'(?i)PH[ẬÁA]T\s+T[ÍI]CH', nm or '')
    ol, orr = frame_overhang_tw(sect); compat = doc_compat(doc)
    ml = int(sect.find(q('pgMar')).get(q('left')) or 0)
    for t in B.raw_tbls: fit_table(t, width, ol, orr, compat, ml, strict=True)
    for t in body.findall(q('tbl')): fit_table(t, width, ol, orr, compat, ml, strict=False)   # §6 (3): mọi bảng cấp thân nằm trong khung
    renumber_docpr(doc.element)
    doc.save(out_path)
    return B.warn, hlog, sect_frame_pt(sect)

def _insert_after(anchor, items):
    for it in items: anchor.addnext(it); anchor = it
    return anchor

def build_toan(B, els, q0, end_q, ans_i, with_answers):
    P = B.P; body = els[0].getparent(); ex = B.exam
    sec_i = [i for i, e in enumerate(els[:ans_i]) if e.tag == q('p') and SEC_RE.match(ptext(e).strip()) and not is_qtitle(e) and i > 0]
    if not sec_i or sec_i[0] > q0: sec_i = [q0 - 1] + sec_i   # đề mẫu không có tiêu đề phần
    tsecs = []
    for k, si in enumerate(sec_i):
        end = sec_i[k + 1] if k + 1 < len(sec_i) else ans_i
        region = els[si + 1:end]
        first_q = next((j for j, e in enumerate(region) if is_qtitle(e)), len(region))
        tsecs.append(dict(title=els[si], instr=[e for e in region[:first_q] if not is_empty_p(e)], drop=region[first_q:] + [e for e in region[:first_q] if is_empty_p(e)]))
    esecs = [s_ for s_ in ex['sections'] if s_['questions'] or s_.get('title')] or [dict(title='', instruction='', questions=[])]
    for ts in tsecs:
        for e in ts['drop']: body.remove(e)
    last_anchor = None
    for k, es in enumerate(esecs):
        if k < len(tsecs):
            ts = tsecs[k]; title = ts['title']
            if es.get('title') and ptext(title).strip() and es['title'].strip() != ptext(title).strip() and SEC_RE.match(ptext(title).strip()):
                set_par_text(title, es['title'].strip())
            anchor = title
            for j, ins in enumerate(ts['instr']):
                if j == 0 and es.get('instruction') and es['instruction'].strip() != ptext(ins).strip(): set_par_text(ins, es['instruction'].strip())
                if j == 0 and not (es.get('instruction') or '').strip() and es.get('questions'): body.remove(ins); continue
                anchor = ins
            if not ts['instr'] and (es.get('instruction') or '').strip():
                anchor = _insert_after(anchor, [new_par(P.text, [(P.r_text, es['instruction'].strip())])])
        else:
            proto = tsecs[-1]['title']; t = clean_ids(copy.deepcopy(proto)); set_par_text(t, es.get('title') or f'{"IVX"[0]}.'); last_anchor.addnext(t); anchor = t
            if (es.get('instruction') or '').strip(): anchor = _insert_after(anchor, [new_par(P.text, [(P.r_text, es['instruction'].strip())])])
        items = []
        for Q in es['questions']: items += B.question(Q)
        last_anchor = _insert_after(anchor, items)
    for ts in tsecs[len(esecs):]:
        body.remove(ts['title'])
        for e in ts['instr']:
            if e.getparent() is not None: body.remove(e)
    # Đáp án
    if ans_i < len(els):
        title = els[ans_i]; tbl = next((e for e in els[ans_i + 1:] if e.tag == q('tbl')), None)
        if not with_answers:
            for e in els[ans_i:]:
                if e.getparent() is not None: body.remove(e)
            # bỏ đoạn trống thừa cuối đề
            last = body_list(body)[-1]
            while last is not None and is_empty_p(last) and len(body_list(body)) > 1:
                pv = last.getprevious(); body.remove(last); last = pv if pv is not None and pv.tag == q('p') else None
            return
        strip_before(title, True, P)
        if tbl is not None: toan_answer_table(B, tbl, esecs)

def toan_answer_table(B, tbl, esecs):
    rows = tbl.findall(q('tr'))
    def c0(r): return cell_text(r.findall(q('tc'))[0])
    simple = next((r for r in rows if re.match(r'^Câu\s*\d+\.?$', c0(r)) and not r.findall(q('tc'))[1].findall('.//' + q('tbl'))), None)
    secrow = next((r for r in rows if re.match(r'^(I|II|III|IV)\s*\.', c0(r)) and 'Câu' not in c0(r)), None)
    firstrow = next((r for r in rows if re.match(r'^(I|II|III)\s*\.', c0(r)) and 'Câu' in c0(r)), None)
    if simple is None: B.warn.append('Không nhận ra bảng đáp án mẫu – giữ nguyên.'); return
    label_p = firstrow.findall(q('tc'))[0].findall(q('p'))[0] if firstrow is not None else None
    hdr = rows[0] if 'Câu' in c0(rows[0]) and not re.search(r'\d', c0(rows[0])) else None
    protos = dict(simple=copy.deepcopy(simple), sec=copy.deepcopy(secrow) if secrow is not None else None, label=copy.deepcopy(label_p) if label_p is not None else None)
    for r in rows:
        if r is not hdr: tbl.remove(r)
    def fill_cell(tc, text=None, raw=None):
        tcpr = tc.find(q('tcPr')); ps = tc.findall(q('p')); p0 = ps[0] if ps else None
        if raw:
            new = B.raw_dict(raw)[0]
            for x in list(tc):
                if x is not tcpr: tc.remove(x)
            for x in list(new):
                if x.tag != q('tcPr'): tc.append(x)
            return
        for x in ps[1:]: tc.remove(x)
        lines = (text or '').split('\n')
        for k, ln in enumerate(lines):
            p = p0 if k == 0 else clean_ids(copy.deepcopy(p0)); 
            if k: tc.append(p)
            rs = p.findall(q('r')); rpr = run_rpr(next((r for r in rs if rtext(r).strip()), rs[0] if rs else None))
            for r in rs: p.remove(r)
            if ln: p.append(make_run(rpr, ln))
    for k, es in enumerate(esecs):
        if k > 0 and protos['sec'] is not None:
            r = clean_ids(copy.deepcopy(protos['sec']))
            if k >= 2 or not re.match(r'^II', cell_text(r.findall(q('tc'))[0])):
                set_par_text(r.findall(q('tc'))[0].findall(q('p'))[0], (es.get('title') or '').strip().capitalize() or ' ')
            tbl.append(r)
        for j, Q in enumerate(es['questions']):
            r = clean_ids(copy.deepcopy(protos['simple'])); tcs = r.findall(q('tc'))
            set_par_text(tcs[0].findall(q('p'))[0], f"Câu {Q['so']}.")
            if k == 0 and j == 0 and protos['label'] is not None: tcs[0].insert(len(tcs[0].findall(q('tcPr'))), clean_ids(copy.deepcopy(protos['label'])))
            a = (B.exam.get('answers') or {}).get(str(Q['so'])) or {}
            raw = (a.get('raw') or [None, None])
            if raw[0] and a.get('text') == (B.exam.get('_answers_orig') or {}).get(str(Q['so']), a.get('text')): fill_cell(tcs[1], raw=raw[0])
            else: fill_cell(tcs[1], a.get('text') or '')
            pts = a.get('points') or (Q.get('diem') or '')
            if len(tcs) > 2:
                if raw[1] and pts == a.get('points'): fill_cell(tcs[2], raw=raw[1])
                else: fill_cell(tcs[2], pts)
            tbl.append(r)

def build_tv(B, els, q0, end_q, hdr_tbls, ans_i, with_answers):
    P = B.P; body = els[0].getparent(); ex = B.exam; rd = ex.get('reading') or {}
    # --- bài đọc ---
    intro_i = next((i for i in range(q0) if els[i].tag == q('p') and ptext(els[i]).strip().startswith('Đọc thầm')), None)
    if intro_i is not None:
        j = intro_i + 1
        while j < q0 and is_empty_p(els[j]): j += 1
        title_i = j
        src_is = [i for i in range(title_i + 1, q0) if els[i].tag == q('p') and re.match(r'^\s*(Theo|Phỏng theo|\()', ptext(els[i]).strip())]
        instr_i = next((i for i in range(title_i + 1, q0) if els[i].tag == q('p') and re.match(r'^\s*(Khoanh|Dựa vào|Trả lời|Em hãy)', ptext(els[i]).strip())), None)
        pend = src_is[0] if src_is else (instr_i or q0)
        pas = [els[i] for i in range(title_i + 1, pend) if els[i].tag == q('p') and not is_empty_p(els[i])]
        if rd.get('intro'): set_par_text(els[intro_i], rd['intro'])
        if rd.get('title'): set_par_text(els[title_i], rd['title'])
        if rd.get('paragraphs'): regen_list(pas, [x for x in rd['paragraphs'] if x.strip()], set_par_text)
        if rd.get('source') is not None and src_is: regen_list([els[i] for i in src_is], [x for x in rd['source'] if x.strip()] or [''], set_par_text)
        if instr_i is not None:
            if (rd.get('instruction') or '').strip(): set_par_text(els[instr_i], rd['instruction'])
            elif rd.get('paragraphs'): body.remove(els[instr_i])
    # --- câu hỏi ---
    anchor = els[q0].getprevious()
    for e in els[q0:end_q]: body.remove(e)
    items = []
    for s_ in ex['sections']:
        if s_.get('title') and s_['title'].strip(): items.append(new_par(P.text, [(P.r_stem, s_['title'].strip())]))
        for Q in s_['questions']: items += B.question(Q)
    _insert_after(anchor, items)
    # --- viết đoạn ---
    tops = (ex.get('writing') or {}).get('topics') or []
    allp = [e for e in body_list(body)]
    de_ps = [e for e in allp if e.tag == q('p') and re.match(r'^\s*Đề\s*\d+\s*:', ptext(e))]
    if tops and de_ps:
        pairs = []
        for d in de_ps:
            nx = d.getnext(); hint = nx if nx is not None and nx.tag == q('p') and ptext(nx).strip().startswith('(Gợi ý') else None
            pairs.append((d, hint))
        for i, tp in enumerate(tops):
            if i < len(pairs): d, h = pairs[i]
            else:
                d0, h0 = pairs[-1]; d = clean_ids(copy.deepcopy(d0)); h = clean_ids(copy.deepcopy(h0)) if h0 is not None else None
                (pairs[-1][1] if pairs[-1][1] is not None else pairs[-1][0]).addnext(d)
                if h is not None: d.addnext(h)
                _sub_runs(d, r'(?<=Đề )\d+', str(i + 1)); pairs.append((d, h))
            set_after_label(d, tp.get('text') or '')
            if h is not None:
                if (tp.get('hint') or '').strip(): set_par_text(h, tp['hint'].strip())
                else: body.remove(h); pairs[i] = (d, None)
        for d, h in pairs[len(tops):]:
            body.remove(d)
            if h is not None: body.remove(h)
    # --- nghe - viết ---
    dc = ex.get('dictation') or {}
    nv = next((e for e in body_list(body) if e.tag == q('p') and re.match(r'^\s*Đề\s*nghe', ptext(e))), None)
    if nv is not None and (dc.get('title') or dc.get('paragraphs')):
        seq = []; e = nv.getnext()
        while e is not None and e.tag == q('p') and (ptext(e).strip() or is_empty_p(e)) and not has_pagebreak(e) and len(seq) < 12:
            if ptext(e).strip(): seq.append(e)
            e = e.getnext()
        if seq:
            src = seq[-1] if re.match(r'^\s*(Theo|Phỏng theo|\()', ptext(seq[-1]).strip()) else None
            ttl = seq[0]; texts = [x for x in seq[1:] if x is not src]
            if dc.get('title'): set_par_text(ttl, dc['title'])
            if dc.get('paragraphs'): regen_list(texts, [x for x in dc['paragraphs'] if x.strip()], set_par_text)
            if src is not None:
                if dc.get('source'): _sub_runs(src, r'\S.*$', dc['source'].strip())
    # --- EXAM_RULES §4 (quy tắc chung đã duyệt): bỏ đoạn trống tràn lên đầu trang trước "Bài làm đề" ---
    for e in body_list(body):
        if e.tag == q('p') and ptext(e).strip().startswith('Bài làm đề'): strip_before(e, False, P)
    # --- hướng dẫn chấm ---
    hdc = next((e for e in body_list(body) if e.tag == q('p') and re.match(r'^\s*(HƯỚNG DẪN CHẤM|ĐÁP ÁN)', ptext(e))), None)
    if hdc is None: return
    if not with_answers:
        # bỏ từ đoạn ngắt trang trước HDC tới hết
        start = hdc
        e = hdc.getprevious()
        while e is not None and e.tag == q('p') and not (e.find('.//' + q('tbl')) is not None) and (is_empty_p(e) or re.match(r'^\s*(UBND|TRƯỜNG)', ptext(e)) or has_pagebreak(e)):
            start = e
            if has_pagebreak(e): break
            e = e.getprevious()
        e = start
        while e is not None and e.tag != q('sectPr'):
            nx = e.getnext(); body.remove(e); e = nx
        return
    tbl = next((x for x in body_list(body)[body_list(body).index(hdc):] if x.tag == q('tbl')), None)
    if tbl is None: return
    nq = [Q for s_ in ex['sections'] for Q in s_['questions']]
    ans = ex.get('answers') or {}
    cps = [p for p in tbl.iter(q('p')) if re.match(r'^\s*Câu\s*\d+\s*:', ptext(p))]
    if not cps: return
    def setter(p, Q):
        _sub_runs(p, r'(?<=Câu )\s*\d+', str(Q['so']))
        set_after_label(p, (ans.get(str(Q['so'])) or {}).get('text', ''))
    if len(cps) != len(nq): B.warn.append(f'Hướng dẫn chấm mẫu có {len(cps)} câu, đề có {len(nq)} câu – cột điểm giữ như mẫu, hãy kiểm tra lại.')
    regen_list(cps, nq, setter)

# =====================================================================================
# 6b. Đề Tiếng Việt tách riêng: Đọc hiểu / Viết (theo mẫu của Thầy Đạt, 10/2026)
# =====================================================================================
READ_INTROS = ('Đọc thầm', 'Cho văn bản sau')

def _tv_reading(B, els, q0):
    """Thay bài đọc mẫu bằng bài đọc của đề. Trả về True nếu tìm thấy vùng bài đọc."""
    body = els[0].getparent(); ex = B.exam; rd = ex.get('reading') or {}
    intro_i = next((i for i in range(q0) if els[i].tag == q('p') and ptext(els[i]).strip().startswith(READ_INTROS)), None)
    if intro_i is None: return False
    j = intro_i + 1
    while j < q0 and is_empty_p(els[j]): j += 1
    title_i = j
    src_is = [i for i in range(title_i + 1, q0) if els[i].tag == q('p') and re.match(r'^\s*(Theo|Phỏng theo|\()', ptext(els[i]).strip())]
    instr_i = next((i for i in range(title_i + 1, q0) if els[i].tag == q('p') and re.match(r'^\s*(Khoanh|Dựa vào|Trả lời|Em hãy)', ptext(els[i]).strip())), None)
    pend = src_is[0] if src_is else (instr_i or q0)
    pas = [els[i] for i in range(title_i + 1, pend) if els[i].tag == q('p') and not is_empty_p(els[i])]
    if rd.get('intro'): set_par_text(els[intro_i], rd['intro'])
    if rd.get('title'): set_par_text(els[title_i], rd['title'])
    if rd.get('paragraphs'): regen_list(pas, [x for x in rd['paragraphs'] if x.strip()], set_par_text)
    if rd.get('source') is not None and src_is: regen_list([els[i] for i in src_is], [x for x in rd['source'] if x.strip()] or [''], set_par_text)
    if instr_i is not None:
        if (rd.get('instruction') or '').strip(): set_par_text(els[instr_i], rd['instruction'])
        elif rd.get('paragraphs'): body.remove(els[instr_i])
    return True

def _tv_questions(B, els, q0, end_q):
    """Thay vùng câu hỏi mẫu bằng câu hỏi của đề."""
    P = B.P; body = els[0].getparent(); ex = B.exam
    anchor = els[q0].getprevious()
    for e in els[q0:end_q]: body.remove(e)
    items = []
    for s_ in ex['sections']:
        if s_.get('title') and s_['title'].strip(): items.append(new_par(P.text, [(P.r_stem, s_['title'].strip())]))
        for Q in s_['questions']: items += B.question(Q)
    _insert_after(anchor, items)

def _tv_hdc(B, with_answers):
    """Điền đáp án vào bảng Hướng dẫn chấm (nếu mẫu có bảng đáp án)."""
    body = B.doc.element.body; ex = B.exam
    hdc = next((e for e in body_list(body) if e.tag == q('p') and re.match(r'^\s*(HƯỚNG DẪN CHẤM|ĐÁP ÁN)', ptext(e))), None)
    if hdc is None: return
    if not with_answers:
        start = hdc
        e = hdc.getprevious()
        while e is not None and e.tag == q('p') and not (e.find('.//' + q('tbl')) is not None) and (is_empty_p(e) or re.match(r'^\s*(UBND|TRƯỜNG)', ptext(e)) or has_pagebreak(e)):
            start = e
            if has_pagebreak(e): break
            e = e.getprevious()
        e = start
        while e is not None and e.tag != q('sectPr'):
            nx = e.getnext(); body.remove(e); e = nx
        return
    tbl = next((x for x in body_list(body)[body_list(body).index(hdc):] if x.tag == q('tbl')), None)
    if tbl is None: return
    # chỉ điền khi bảng là bảng đáp án (có dòng "Câu N:"), không đụng bảng ma trận
    cps = [p for p in tbl.iter(q('p')) if re.match(r'^\s*Câu\s*\d+\s*:', ptext(p))]
    if not cps: return
    nq = [Q for s_ in ex['sections'] for Q in s_['questions']]
    ans = ex.get('answers') or {}
    def setter(p, Q):
        _sub_runs(p, r'(?<=Câu )\s*\d+', str(Q['so']))
        set_after_label(p, (ans.get(str(Q['so'])) or {}).get('text', ''))
    if len(cps) != len(nq): B.warn.append(f'Hướng dẫn chấm mẫu có {len(cps)} câu, đề có {len(nq)} câu – hãy kiểm tra lại.')
    regen_list(cps, nq, setter)

def _tv_matrix(B):
    """Tự sinh bảng ma trận từ mức độ (M1/M2/M3) và loại câu (TN/TL) của đề."""
    body = B.doc.element.body; ex = B.exam
    mt_title = next((e for e in body_list(body) if e.tag == q('p') and 'MA TRẬN' in ptext(e).upper()), None)
    if mt_title is None: return
    tbl = mt_title.getnext()
    while tbl is not None and tbl.tag != q('tbl'): tbl = tbl.getnext()
    if tbl is None: return
    nq = [Q for s_ in ex['sections'] for Q in s_['questions']]
    def is_tn(Q): return any((b.get('t') == 'options') for b in (Q.get('blocks') or []))
    def pts(Q):
        try: return float(str(Q.get('diem') or '0').replace(',', '.'))
        except Exception: return 0.0
    def fmt(x):
        s = f'{x:.2f}'.rstrip('0').rstrip('.')
        return s.replace('.', ',') if s else '0'
    # thống kê theo mức x loại
    stat = {}
    for Q in nq:
        m = (Q.get('muc') or 'M1').upper().strip() or 'M1'
        k = (m, 'TN' if is_tn(Q) else 'TL')
        s = stat.setdefault(k, dict(n=0, nums=[], p=0.0))
        s['n'] += 1; s['nums'].append(str(Q['so'])); s['p'] += pts(Q)
    rows = tbl.findall(q('tr'))
    # giữ hàng tiêu đề, viết lại các hàng dữ liệu (copy độ rộng cột từ hàng tiêu đề)
    hdr_tcw = []
    for tc in rows[0].findall(q('tc')):
        w = tc.find(q('tcPr'))
        w = w.find(q('tcW')) if w is not None else None
        hdr_tcw.append((w.get(q('w')) if w is not None else None, w.get(q('type')) if w is not None else None))
    for r in rows[1:]: tbl.remove(r)
    def add_row(cells, bold=False, size=11):
        tr = etree.Element(q('tr'))
        for ci, v in enumerate(cells):
            tc = etree.Element(q('tc'))
            tcPr = etree.Element(q('tcPr'))
            if hdr_tcw[ci][0]:
                tcW = etree.Element(q('tcW')); tcW.set(q('w'), hdr_tcw[ci][0])
                tcW.set(q('type'), hdr_tcw[ci][1] or 'dxa'); tcPr.append(tcW)
            tc.append(tcPr)
            p = etree.Element(q('p'))
            pPr = etree.Element(q('pPr'))
            jc = etree.Element(q('jc')); jc.set(q('val'), 'center'); pPr.append(jc)
            p.append(pPr)
            r = etree.Element(q('r'))
            rPr = etree.Element(q('rPr'))
            rf = etree.Element(q('rFonts'))
            for a in ('ascii', 'hAnsi', 'cs'): rf.set(q(a), 'Times New Roman')
            rPr.append(rf)
            sz = etree.Element(q('sz')); sz.set(q('val'), str(size * 2)); rPr.append(sz)
            if bold: rPr.append(etree.Element(q('b')))
            r.append(rPr)
            t = etree.Element(q('t')); t.text = v; r.append(t)
            p.append(r); tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    levels = ['M1', 'M2', 'M3']
    tot = dict(n=0, p=0.0, tn_n=0, tn_p=0.0, tl_n=0, tl_p=0.0)
    for lv in levels:
        tnd = stat.get((lv, 'TN'), dict(n=0, nums=[], p=0.0))
        tld = stat.get((lv, 'TL'), dict(n=0, nums=[], p=0.0))
        row_p = tnd['p'] + tld['p']
        tot['n'] += tnd['n'] + tld['n']; tot['p'] += row_p
        tot['tn_n'] += tnd['n']; tot['tn_p'] += tnd['p']
        tot['tl_n'] += tld['n']; tot['tl_p'] += tld['p']
        def cell(v): return '' if v == 0 or v == '0' else str(v)
        add_row([f'Mức {lv[1]}',
                 cell(tnd['n']), ', '.join(tnd['nums']), cell(fmt(tnd['p'])),
                 cell(tld['n']), ', '.join(tld['nums']), cell(fmt(tld['p'])),
                 cell(fmt(row_p))])
    def cell(v): return '' if v == 0 or v == '0' else str(v)
    add_row(['Tổng', cell(tot['tn_n']), '', cell(fmt(tot['tn_p'])),
             cell(tot['tl_n']), '', cell(fmt(tot['tl_p'])), cell(fmt(tot['p']))], bold=True)

def build_tv_doc(B, els, q0, end_q, hdr_tbls, ans_i, with_answers):
    """Đề Tiếng Việt – Đọc hiểu: bài đọc + câu hỏi + HDC + ma trận tự động."""
    _tv_reading(B, els, q0)
    _tv_questions(B, els, q0, end_q)
    _tv_hdc(B, with_answers)
    if with_answers: _tv_matrix(B)

def build_tv_viet(B, els, end_q, hdr_tbls, ans_i, with_answers):
    """Đề Tiếng Việt – Viết: nghe-viết + viết đoạn. HDC là thang điểm cố định."""
    body = els[0].getparent(); ex = B.exam
    # --- viết đoạn: thay các "Đề N:" mẫu ---
    tops = (ex.get('writing') or {}).get('topics') or []
    allp = [e for e in body_list(body)]
    de_ps = [e for e in allp if e.tag == q('p') and re.match(r'^\s*Đề\s*\d+\s*:', ptext(e))]
    if tops and de_ps:
        pairs = []
        for d in de_ps:
            nx = d.getnext(); hint = nx if nx is not None and nx.tag == q('p') and ptext(nx).strip().startswith('(Gợi ý') else None
            pairs.append((d, hint))
        for i, tp in enumerate(tops):
            if i < len(pairs): d, h = pairs[i]
            else:
                d0, h0 = pairs[-1]; d = clean_ids(copy.deepcopy(d0)); h = clean_ids(copy.deepcopy(h0)) if h0 is not None else None
                (pairs[-1][1] if pairs[-1][1] is not None else pairs[-1][0]).addnext(d)
                if h is not None: d.addnext(h)
                _sub_runs(d, r'(?<=Đề )\d+', str(i + 1)); pairs.append((d, h))
            set_after_label(d, tp.get('text') or '')
            if h is not None:
                if (tp.get('hint') or '').strip(): set_par_text(h, tp['hint'].strip())
                else: body.remove(h); pairs[i] = (d, None)
        for d, h in pairs[len(tops):]:
            body.remove(d)
            if h is not None: body.remove(h)
    elif not tops:
        # không có đề viết đoạn: bỏ cả cụm "Chọn 1 trong 2 đề sau" + các Đề N:
        for d, h in [(d, d.getnext() if d.getnext() is not None and d.getnext().tag == q('p') and ptext(d.getnext()).strip().startswith('(Gợi ý') else None) for d in de_ps]:
            body.remove(d)
            if h is not None: body.remove(h)
    # --- nghe - viết ---
    dc = ex.get('dictation') or {}
    nv = next((e for e in body_list(body) if e.tag == q('p') and re.match(r'^\s*Đề\s*nghe', ptext(e))), None)
    if nv is not None and (dc.get('title') or dc.get('paragraphs')):
        seq = []; e = nv.getnext()
        while e is not None and e.tag == q('p') and (ptext(e).strip() or is_empty_p(e)) and not has_pagebreak(e) and len(seq) < 12:
            if ptext(e).strip(): seq.append(e)
            e = e.getnext()
        if seq:
            src = seq[-1] if re.match(r'^\s*(Theo|Phỏng theo|\()', ptext(seq[-1]).strip()) else None
            ttl = seq[0]; texts = [x for x in seq[1:] if x is not src]
            if dc.get('title'): set_par_text(ttl, 'Đề nghe - viết: ' + dc['title'].strip())
            if dc.get('paragraphs'): regen_list(texts, [x for x in dc['paragraphs'] if x.strip()], set_par_text)
            if src is not None:
                if dc.get('source'): _sub_runs(src, r'\S.*$', dc['source'].strip())
    # --- HDC: thang điểm cố định, chỉ cắt bỏ khi không kèm đáp án ---
    _tv_hdc(B, with_answers)

# =====================================================================================
# 7. Xem trước trang A4 (LibreOffice -> PDF -> PNG)
# =====================================================================================
_render_lock = threading.Lock()
def soffice_bin():
    for b in ('soffice', 'libreoffice'):
        p = shutil.which(b)
        if p: return p
    return None

if not soffice_bin():
    # Log ngay khi import để người vận hành biết khi khởi động (không crash, chỉ tắt xem trước A4)
    log.warning('exam: KHÔNG tìm thấy LibreOffice (soffice) trong image — chức năng xem trước trang A4 của đề '
                'sẽ bị tắt, người dùng vẫn tạo và tải file Word bình thường. Muốn bật: build image với '
                '--build-arg WITH_LIBREOFFICE=1 (mặc định đã bật).')

def render_pages(docx_path, outdir, dpi=110, timeout=120, with_png=True):
    """Trả về list đường dẫn PNG các trang, hoặc None nếu không có LibreOffice."""
    sb = soffice_bin()
    if not sb: return None
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):
        if f.endswith('.png') or f.endswith('.pdf'): os.remove(os.path.join(outdir, f))
    with _render_lock, tempfile.TemporaryDirectory() as td:
        prof = os.path.join(tempfile.gettempdir(), 'lo_profile_exam')
        try:
            subprocess.run([sb, f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'pdf', '--outdir', td, docx_path],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=timeout, check=False)
        except subprocess.TimeoutExpired:
            log.warning('render: soffice quá thời gian'); return None
        pdfs = [f for f in os.listdir(td) if f.endswith('.pdf')]
        if not pdfs: return None
        import pymupdf
        pdf = os.path.join(td, pdfs[0]); shutil.copy(pdf, os.path.join(outdir, 'de.pdf'))
        doc = pymupdf.open(pdf)
        if not with_png:
            # Đường tải Word chỉ cần PDF để kiểm tra khung trang — bỏ render ảnh PNG cho nhanh.
            n = doc.page_count; doc.close(); return ['p%d' % (i + 1) for i in range(n)]
        out = []
        for i, page in enumerate(doc):
            fn = os.path.join(outdir, f'p{i + 1}.png'); page.get_pixmap(dpi=dpi).save(fn); out.append(fn)
        return out

def page_labels(pdf_path, mon):
    """Nhãn từng trang theo nội dung (Bài 1 / Bài 2 / Nghe - viết / Đáp án)."""
    try:
        import pymupdf
        doc = pymupdf.open(pdf_path)
    except Exception: return []
    labels, cur = [], 'Đề'
    marks = ([('HƯỚNG DẪN CHẤM', 'Hướng dẫn chấm'), ('Đề nghe', 'Nghe - viết'), ('Bài số 2', 'Bài 2: Viết'), ('Bài số 1', 'Bài 1: Đọc')] if mon == 'tv'
             else [('ĐÁP ÁN', 'Đáp án'), ('TỰ LUẬN', 'Tự luận'), ('TRẮC NGHIỆM', 'Trắc nghiệm')])
    order = [m[1] for m in reversed(marks)]
    for page in doc:
        t = page.get_text('text'); blank = not t.strip() and not page.get_images()
        found = sorted((t.find(k), v) for k, v in marks if k in t)
        if found:
            # trang mang tên phần bắt đầu sớm nhất trên trang (nếu phần trước còn tràn sang thì vẫn lấy phần mới bắt đầu ở đầu trang)
            first = found[0][1]
            if order.index(first) >= (order.index(cur) if cur in order else -1): cur = first
        labels.append('Trang trống!' if blank else cur)
    return labels

# =====================================================================================
# 8. Xem trước HTML (khi không có LibreOffice)
# =====================================================================================
def html_preview(ex, fields):
    import html as H
    e = H.escape; out = ['<div class="hp">']
    out.append(f"<div class='hph'><b>{e((fields.get('truong') or '').upper())}</b><br>BÀI KIỂM TRA {e((fields.get('loai') or '').upper())}<br>NĂM HỌC {e(fields.get('nam') or '')}<br>Môn: {'TOÁN' if ex['mon']=='toan' else 'TIẾNG VIỆT'} - Lớp {e(str(fields.get('lop') or ''))}</div>")
    rd = ex.get('reading') or {}
    if rd.get('title'):
        out.append(f"<p class='c'><b>{e(rd['title'])}</b></p>" + ''.join(f'<p>{e(x)}</p>' for x in rd.get('paragraphs') or []) + ''.join(f"<p class='r'><i>{e(x)}</i></p>" for x in rd.get('source') or []))
    for s_ in ex['sections']:
        if s_.get('title'): out.append(f"<p><b>{e(s_['title'])}</b></p>")
        if s_.get('instruction'): out.append(f"<p><b>{e(s_['instruction'])}</b></p>")
        for Q in s_['questions']:
            tag = f" <span class='tg'>({e(Q.get('muc') or '')}{' – ' + e(Q['diem']) + 'đ' if Q.get('diem') else ''})</span>" if Q.get('muc') or Q.get('diem') else ''
            out.append(f"<p><b>Câu {Q['so']}.</b> {e(Q.get('text') or '')}{tag}</p>")
            for b in Q.get('blocks') or []:
                t = b['t']
                if t == 'text': out.append(f"<p>{e(b.get('text') or '')}</p>")
                elif t == 'options': out.append("<p class='op'>" + ''.join(f"<span><b>{chr(65+i)}.</b> {e(x)}</span>" for i, x in enumerate(b['items'])) + '</p>')
                elif t == 'lines': out.append("<p class='dl'></p>" * int(b.get('n') or 1))
                elif t == 'ds': out.append(''.join(f"<p>{e(x)} <span class='bx'></span></p>" for x in b['items']))
                elif t == 'figure': out.append(f"<p class='c'><img src='fig/{e(b['id'])}' style='max-width:60%'></p>")
                elif t == 'raw': out.append(f"<p class='c raw'>[hình/khối giữ nguyên từ file Word]</p>")
    out.append('</div>')
    return ''.join(out)

# =====================================================================================
# 9. KHUNG TRANG – quy tắc tuyệt đối (EXAM_RULES §6): không nội dung nào chạm/tràn khung viền trang
# =====================================================================================
WPD = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
DML = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
VML = '{urn:schemas-microsoft-com:vml}'
EMU_PT = 12700
FRAME_SAFE_PT = 3.0          # khoảng cách tối thiểu giữa nội dung và mép trong khung (khi dựng)
FRAME_TOL_PT = 2.0           # kiểm tra: nội dung cách mép trong khung < 2pt là lỗi
FRAME_MIN_SPACE_PT = 20.0    # khung không ra sát mép giấy hơn 20pt (~7 mm, an toàn khi in)
FIG_SAFE = 0.96              # hình rộng tối đa = 96% bề rộng chữ
FIG_MAX_H = 0.45             # hình cao tối đa = 45% chiều cao vùng chữ
COMPOSITE = ('double', 'triple', 'thinThick', 'thickThin', 'Gap', 'threeD', 'emboss', 'engrave', 'inset', 'outset')
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
             'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']

def ppr_of(p):
    ppr = p.find(q('pPr'))
    if ppr is None: ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    return ppr

def ppr_set(p, name, attrs=None):
    """Đặt/ghi đè phần tử con của pPr đúng thứ tự lược đồ."""
    ppr = ppr_of(p); el = ppr.find(q(name))
    if el is None:
        el = etree.Element(q(name)); k = PPR_ORDER.index(name)
        pos = next((i for i, c in enumerate(ppr) if etree.QName(c).localname in PPR_ORDER and PPR_ORDER.index(etree.QName(c).localname) > k), len(ppr))
        ppr.insert(pos, el)
    for a, v in (attrs or {}).items(): el.set(q(a), str(v))
    return el

def border_inner_pt(b, vertical=False):
    """(khoảng cách tới mép giấy, độ dày) của 1 cạnh khung (offsetFrom=page). Nét đôi/ghép (thinThick…) dày ~1,5× sz
    (đo trên LibreOffice: mép trong = space + 1,5×sz/8). Nếu mép trong vượt lề, LibreOffice còn đẩy cả khung lẫn chữ xuống ~4pt
    -> chữ dính khung (lỗi của mẫu gốc)."""
    sz = int(b.get(q('sz')) or 4) / 8.0; sp = float(b.get(q('space')) or 0)
    k = 1.5 if any(c in (b.get(q('val')) or '') for c in COMPOSITE) else 1.0
    return sp, sz * k

def fit_page_frame(sect):
    """Nếu khung viền trang lấn vào vùng chữ: dời khung ra ngoài (không sát mép giấy hơn FRAME_MIN_SPACE_PT), còn thiếu thì tăng lề.
    Trả về list thay đổi (để ghi log/kiểm tra)."""
    pb = sect.find(q('pgBorders')); mar = sect.find(q('pgMar')); ch = []
    if pb is None or mar is None or (pb.get(q('offsetFrom')) or 'text') != 'page': return ch
    for side in ('top', 'bottom', 'left', 'right'):
        b = pb.find(q(side))
        if b is None or (b.get(q('val')) or 'nil') in ('nil', 'none'): continue
        sp, w = border_inner_pt(b, side in ('top', 'bottom'))
        m = int(mar.get(q(side)) or 0) / 20.0
        if sp + w + FRAME_SAFE_PT <= m: continue
        nsp = max(FRAME_MIN_SPACE_PT, min(sp, m - w - FRAME_SAFE_PT))
        if nsp < sp: b.set(q('space'), str(int(math.floor(nsp)))); ch.append(f'{side}: khung {sp:g}pt→{int(math.floor(nsp))}pt'); sp = int(math.floor(nsp))
        if sp + w + FRAME_SAFE_PT > m:
            nm = int(math.ceil((sp + w + FRAME_SAFE_PT) * 20)); mar.set(q(side), str(nm)); ch.append(f'{side}: lề {m:g}pt→{nm / 20:g}pt')
    return ch

def text_height_tw(sect):
    pg = sect.find(q('pgSz')); mar = sect.find(q('pgMar'))
    try: return int(pg.get(q('h'))) - int(mar.get(q('top'))) - int(mar.get(q('bottom')))
    except Exception: return 14000

def _ext_scale(container, f):
    """Thu phóng khung ngoài của 1 hình (wp:extent + a:xfrm cấp ngoài cùng); toạ độ con trong nhóm giữ nguyên (chExt) nên cả nhóm co theo."""
    e = container.find(WPD + 'extent')
    if e is not None:
        for a in ('cx', 'cy'): e.set(a, str(max(1, int(int(e.get(a)) * f))))
    g = container.find(DML + 'graphic')
    x = next(g.iter(DML + 'xfrm'), None) if g is not None else None
    ex = x.find(DML + 'ext') if x is not None else None
    if ex is not None:
        for a in ('cx', 'cy'): ex.set(a, str(max(1, int(int(ex.get(a)) * f))))

def _owner_p(el):
    """Đoạn chứa hình, bỏ qua đoạn bên trong hộp chữ (txbxContent)."""
    x = el.getparent()
    while x is not None:
        if x.tag == q('p') and not any(a.tag == q('txbxContent') for a in x.iterancestors()): return x
        x = x.getparent()
    return None

def _vis_text(p):
    """Chữ thấy được của đoạn, không tính chữ trong hộp chữ/hình."""
    return ''.join((t.text or '') for t in p.iter(q('t')) if not any(a.tag in (q('txbxContent'), q('drawing'), q('pict')) for a in t.iterancestors()))

def _p_size_pt(p):
    r = next((r for r in p.iter(q('r')) if r.find(q('rPr')) is not None and r.find(q('rPr')).find(q('sz')) is not None), None)
    return pr_size(run_rpr(r)) if r is not None else pr_size((p.find(q('pPr')).find(q('rPr')) if p.find(q('pPr')) is not None else None))

def _p_line_pt(p):
    sp = p.find(q('pPr')).find(q('spacing')) if p.find(q('pPr')) is not None else None
    sz = _p_size_pt(p); base = sz * 1.17
    if sp is None or sp.get(q('line')) is None: return base
    ln = int(sp.get(q('line'))); rule = sp.get(q('lineRule')) or 'auto'
    return base * ln / 240 if rule == 'auto' else (ln / 20 if rule == 'exact' else max(base, ln / 20))

def _ind_tw(p):
    ind = p.find(q('pPr')).find(q('ind')) if p.find(q('pPr')) is not None else None
    if ind is None: return 0
    return int(ind.get(q('left')) or ind.get(q('start')) or 0) + int(ind.get(q('right')) or ind.get(q('end')) or 0)

def fit_drawings(el, width_tw, height_tw, shrink=1.0, log_=None):
    """EXAM_RULES §6 (1)(2): hình chép từ đề nguồn
    - hình nổi (wp:anchor): neo theo đoạn (positionV paragraph, positionH column), đưa về trong bề rộng chữ, phần trên không âm;
      dòng của đoạn neo được nới cao đủ chứa hình (giữ chỗ như bao trên-dưới, không thêm dòng trống) và giữ liền với câu.
    - hình cùng dòng (wp:inline) và hình nổi: co theo bề rộng chữ (trừ lề an toàn) và chiều cao tối đa; shrink<1 co thêm (vòng kiểm tra)."""
    W_e = width_tw * 635; H_e = height_tw * 635 * FIG_MAX_H; n = 0
    for inl in list(el.iter(WPD + 'inline')):
        p = _owner_p(inl); avail = (width_tw - (_ind_tw(p) if p is not None else 0)) * 635 * FIG_SAFE
        e = inl.find(WPD + 'extent'); cx, cy = int(e.get('cx')), int(e.get('cy'))
        f = min(1.0, avail / max(cx, 1), H_e / max(cy, 1)) * shrink
        if f < 0.999: _ext_scale(inl, f); n += 1
    groups = {}
    for a in el.iter(WPD + 'anchor'):
        if any(x.tag == q('txbxContent') for x in a.iterancestors()): continue
        p = _owner_p(a)
        if p is not None: groups.setdefault(id(p), (p, []))[1].append(a)
    for p, anchors in groups.values():
        boxes = []
        for a in anchors:
            ph, pv, e = a.find(WPD + 'positionH'), a.find(WPD + 'positionV'), a.find(WPD + 'extent')
            cx, cy = int(e.get('cx')), int(e.get('cy'))
            def off(pos, size, span):
                if pos is None: return 0
                o = pos.find(WPD + 'posOffset'); al = pos.find(WPD + 'align')
                if o is not None: return int(o.text or 0)
                v = (al.text or 'left') if al is not None else 'left'
                return {'center': (span - size) // 2, 'right': span - size, 'bottom': 0, 'inside': 0, 'outside': span - size}.get(v, 0)
            x = off(ph, cx, W_e); y = off(pv, cy, 0)
            rh = ph.get('relativeFrom') if ph is not None else 'column'; rv = pv.get('relativeFrom') if pv is not None else 'paragraph'
            if rh not in ('column', 'margin', 'character'): x = 0     # theo trang/lề ngoài -> về đầu cột (sau đó kẹp trong bề rộng chữ)
            if rh == 'character': x = None
            if rv not in ('paragraph', 'line'): y = 0                # theo trang/lề -> neo đầu đoạn
            boxes.append([a, x, y, cx, cy, rh])
        chars = [b for b in boxes if b[1] is None]   # neo theo ký tự: giữ vị trí ngang, chỉ co nếu rộng hơn bề rộng chữ
        for b in chars: b[1] = 0
        x0 = min(b[1] for b in boxes); x1 = max(b[1] + b[3] for b in boxes); y0 = min(b[2] for b in boxes); y1 = max(b[2] + b[4] for b in boxes)
        textual = bool(_vis_text(p).strip())
        # đoạn có chữ: hình gắn với vị trí chữ -> chỉ đụng tới khi thật sự vượt bề rộng chữ; đoạn chỉ có hình: co theo bề rộng trừ lề an toàn
        avail = W_e * (1.0 if textual else FIG_SAFE)
        f = min(1.0, avail / max(x1 - x0, 1), H_e / max(y1 - y0, 1)) * shrink
        dx = -x0 if x0 < 0 else 0; dy = -y0 if y0 < 0 else 0     # không để hình thò lên trên/ra trái đoạn neo
        for b in boxes:
            b[1] = int((b[1] + dx) * f); b[2] = int((b[2] + dy) * f)
            if f < 0.999: _ext_scale(b[0], f); b[3], b[4] = int(b[3] * f), int(b[4] * f)
        over = max(b[1] + b[3] for b in boxes) - int(avail)
        if over > 0:
            for b in boxes: b[1] = max(0, b[1] - over)
        for b in boxes:
            a = b[0]
            for tag, rel, val in (('positionH', 'column', b[1]), ('positionV', 'paragraph', b[2])):
                pos = a.find(WPD + tag)
                if pos is None or (tag == 'positionH' and b in chars): continue
                pos.set('relativeFrom', rel)
                for c in list(pos): pos.remove(c)
                etree.SubElement(pos, WPD + 'posOffset').text = str(int(val))
        # giữ chỗ: dòng của đoạn neo cao đủ chứa hình (tương đương bao trên-dưới nhưng không thêm dòng trống, và
        # cả cụm sang trang sau nguyên vẹn – LibreOffice bỏ qua "giữ với đoạn sau" của hình bao trên-dưới)
        need = max(b[2] + b[4] for b in boxes) / EMU_PT + 2
        if _p_line_pt(p) < need:
            sp = p.find(q('pPr')).find(q('spacing')) if p.find(q('pPr')) is not None else None
            if len(_vis_text(p)) < 120 and not p.findall('.//' + q('br')):
                ppr_set(p, 'spacing', {'line': int(math.ceil(need * 20)), 'lineRule': 'exact'})
            else: ppr_set(p, 'spacing', {'after': int(math.ceil((need - _p_line_pt(p)) * 20)) + (int(sp.get(q('after')) or 0) if sp is not None else 0)})
        n += 1
    # VML (w:pict không có mc:Choice): chỉ co bề rộng/cao trong style
    for sh in el.iter(VML + 'shape', VML + 'group', VML + 'rect', VML + 'oval', VML + 'line'):
        if any(x.tag in (VML + 'group',) for x in sh.iterancestors()) or any(x.tag.endswith('Fallback') for x in sh.iterancestors()): continue
        st = sh.get('style') or ''; m = dict(re.findall(r'([\w-]+)\s*:\s*([^;]+)', st))
        def pt(v):
            mm = re.match(r'(-?[\d.]+)\s*(pt|in|cm|mm|px)?', v or '')
            if not mm: return None
            return float(mm.group(1)) * {'pt': 1, 'in': 72, 'cm': 28.3465, 'mm': 2.83465, 'px': 0.75, None: 1}[mm.group(2)]
        w_, h_ = pt(m.get('width')), pt(m.get('height'))
        if not w_ or not h_: continue
        f = min(1.0, (width_tw / 20 * FIG_SAFE) / w_, (height_tw / 20 * FIG_MAX_H) / h_) * shrink
        if f < 0.999:
            st = re.sub(r'width\s*:\s*[^;]+', f'width:{w_ * f:.2f}pt', st); st = re.sub(r'height\s*:\s*[^;]+', f'height:{h_ * f:.2f}pt', st); sh.set('style', st); n += 1
    if log_ is not None and n: log_.append(n)
    return n

def table_cell_left_tw(tbl, compat):
    """Word/LO (chế độ tương thích < 15): mép trái bảng = lề + tblInd − lề trái ô (mặc định 108tw)."""
    if compat >= 15: return 0
    tp = tbl.find(q('tblPr')); cm = tp.find(q('tblCellMar')) if tp is not None else None
    l = (cm.find(q('left')) if cm.find(q('left')) is not None else cm.find(q('start'))) if cm is not None else None
    return int(l.get(q('w')) or 108) if l is not None else 108

def fit_table(tbl, width_tw, over_l=0, over_r=0, compat=14, margin_l=0, strict=True):
    """EXAM_RULES §6 (3): bảng rộng ≤ bề rộng chữ; mép thật (kể cả lề ô thò ra ở chế độ tương thích cũ) không quá over_l/over_r tw ngoài lề
    (over_* = phần lề còn cách khung ≥ an toàn).
    Ưu tiên dời bảng (giữ nguyên kích thước ô, VD lưới ô ly), chỉ co cột theo tỉ lệ khi bảng rộng hơn chỗ có."""
    tp = tbl.find(q('tblPr')); grid = tbl.find(q('tblGrid'))
    if tp is None: return False
    cols = [int(g.get(q('w')) or 0) for g in grid.findall(q('gridCol'))] if grid is not None else []
    tw = tp.find(q('tblW'))
    w = sum(cols) or (int(tw.get(q('w')) or 0) if tw is not None and tw.get(q('type')) == 'dxa' else 0)
    if not w: return False
    jc = tp.find(q('jc')); jv = jc.get(q('val')) if jc is not None else 'left'
    ind = tp.find(q('tblInd'))
    iv = int(ind.get(q('w')) or 0) if ind is not None and (ind.get(q('type')) or 'dxa') == 'dxa' else 0
    cml = table_cell_left_tw(tbl, compat)
    if strict: over_l, over_r = min(over_l, cml), min(over_r, cml)   # bảng chép từ đề nguồn: ≤ bề rộng chữ (chỉ thò phần lề ô chuẩn)
    room = width_tw if strict else width_tw + over_l + over_r          # bảng của mẫu: chỉ sửa khi thật sự chạm khung
    changed = False
    if w > room:
        f = room / w
        for g in (grid.findall(q('gridCol')) if grid is not None else []): g.set(q('w'), str(int(int(g.get(q('w')) or 0) * f)))
        if tw is not None and tw.get(q('type')) == 'dxa': tw.set(q('w'), str(int(int(tw.get(q('w')) or 0) * f)))
        for tc in tbl.iter(q('tcW')):
            if (tc.get(q('type')) or 'dxa') == 'dxa': tc.set(q('w'), str(int(int(tc.get(q('w')) or 0) * f)))
        w = int(w * f); changed = True
    if jv in ('center', 'right', 'end'): return changed      # căn giữa/phải: vị trí do bề rộng quyết định
    fl = tp.find(q('tblpPr'))
    if fl is not None:   # bảng nổi: vị trí ngang = tblpX theo horzAnchor (margin/text/page)
        if fl.get(q('tblpXSpec')): return changed
        base = margin_l if fl.get(q('horzAnchor')) == 'page' else 0
        px = int(fl.get(q('tblpX')) or 0); left = px - base - cml; right = left + w
        shift = (-over_l - left) if left < -over_l else (-(right - width_tw - over_r) if right > width_tw + over_r else 0)
        if shift: fl.set(q('tblpX'), str(px + int(shift))); changed = True
        return changed
    left = iv - cml; right = left + w
    shift = (-over_l - left) if left < -over_l else (-(right - width_tw - over_r) if right > width_tw + over_r else 0)
    if shift:
        if ind is None: ind = etree.Element(q('tblInd')); ind.set(q('type'), 'dxa'); tp.insert(next((i for i, c in enumerate(tp) if etree.QName(c).localname in ('tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription')), len(tp)), ind)
        ind.set(q('w'), str(iv + int(shift))); ind.set(q('type'), 'dxa'); changed = True
    return changed

def frame_overhang_tw(sect):
    """Phần lề trái/phải mà nội dung được phép thò ra (vẫn cách mép trong khung ≥ FRAME_SAFE_PT)."""
    pb = sect.find(q('pgBorders')); mar = sect.find(q('pgMar')); out = []
    for side in ('left', 'right'):
        m = int(mar.get(q(side)) or 0) / 20.0; b = pb.find(q(side)) if pb is not None else None
        if b is None or (b.get(q('val')) or 'nil') in ('nil', 'none'): out.append(int(max(0, m - 14) * 20)); continue
        sp, w = border_inner_pt(b)
        inner = sp + w if (pb.get(q('offsetFrom')) or 'text') == 'page' else m - sp
        out.append(int(max(0, m - inner - FRAME_SAFE_PT - 0.5) * 20))
    return out

def doc_compat(doc):
    try:
        st = doc.settings.element
        for cs in st.iter(q('compatSetting')):
            if cs.get(q('name')) == 'compatibilityMode': return int(cs.get(q('val')))
        return 14 if st.find(q('compat')) is not None else 15
    except Exception: return 14

# ---------- kiểm tra sau khi dựng: PDF (LibreOffice) -> hộp bao từng nội dung so với khung ----------
def frame_rect(page, sect_frame=None, drawings=None):
    """Mép trong khung viền trang lấy từ chính các nét khung vẽ trên trang.
    Có sect_frame (mép trong tính từ sectPr, pt): nét khung = nét dài nằm giữa mép giấy và mép trong dự kiến (+1pt) – không nhầm với
    viền bảng/lưới ô ly sát lề. Không có: nét dài gần hết trang, nối gần mép trên tới gần mép dưới (hoặc trái tới phải)."""
    W_, H_ = page.rect.width, page.rect.height
    dr = drawings if drawings is not None else page.get_drawings()
    def pick(expected):
        hz, vt, border = [], [], []
        for d in dr:
            r = d['rect']; w = (d.get('width') or 0) / 2
            if expected:
                L, T, Rr, B_ = expected
                if r.width > W_ * 0.85 and r.height < 8 and (r.y1 + w <= T + 1 or r.y0 - w >= B_ - 1): hz.append((r, w)); border.append(d)
                elif r.height > H_ * 0.85 and r.width < 8 and (r.x1 + w <= L + 1 or r.x0 - w >= Rr - 1): vt.append((r, w)); border.append(d)
            else:
                E = 36
                if r.width > W_ * 0.85 and r.height < 8 and r.x0 < E and r.x1 > W_ - E and (r.y1 < E or r.y0 > H_ - E): hz.append((r, w)); border.append(d)
                elif r.height > H_ * 0.85 and r.width < 8 and r.y0 < E and r.y1 > H_ - E and (r.x1 < E or r.x0 > W_ - E): vt.append((r, w)); border.append(d)
        top = [r.y1 + w for r, w in hz if r.y1 < H_ / 2]; bot = [r.y0 - w for r, w in hz if r.y0 > H_ / 2]
        lef = [r.x1 + w for r, w in vt if r.x1 < W_ / 2]; rig = [r.x0 - w for r, w in vt if r.x0 > W_ / 2]
        return [max(lef) if lef else None, max(top) if top else None, min(rig) if rig else None, min(bot) if bot else None], border
    got, border = pick(sect_frame) if sect_frame else ([None] * 4, [])
    if None in got:   # cạnh không khớp vị trí dự kiến (VD LibreOffice đẩy khung xuống khi khung lấn lề) -> nhận theo hình dạng nét khung
        g2, b2 = pick(None)
        for i in range(4):
            if got[i] is None and g2[i] is not None: got[i] = g2[i]
        ids = {id(d) for d in border}; border += [d for d in b2 if id(d) not in ids]
    if None not in got: return tuple(got), border
    return sect_frame, border

def docx_frame_pt(docx_path):
    """Khung (mép trong) theo sectPr cuối của 1 tệp .docx."""
    x = etree.fromstring(zipfile.ZipFile(docx_path).read('word/document.xml'))
    s = x.find(q('body')).find(q('sectPr'))
    return sect_frame_pt(s) if s is not None else None

def sect_frame_pt(sect):
    """Khung (mép trong) tính từ sectPr: pgBorders (offsetFrom page/text) + lề."""
    pg = sect.find(q('pgSz')); mar = sect.find(q('pgMar')); pb = sect.find(q('pgBorders'))
    Wp, Hp = int(pg.get(q('w'))) / 20, int(pg.get(q('h'))) / 20
    m = {s: int(mar.get(q(s)) or 0) / 20 for s in ('left', 'top', 'right', 'bottom')}
    if pb is None: return None
    v = {}
    for s in ('left', 'top', 'right', 'bottom'):
        b = pb.find(q(s))
        if b is None: v[s] = 0; continue
        sp, w = border_inner_pt(b, s in ('top', 'bottom'))
        v[s] = sp + w if (pb.get(q('offsetFrom')) or 'text') == 'page' else m[s] - sp
    return (v['left'], v['top'], Wp - v['right'], Hp - v['bottom'])

def content_boxes(page, border, drawings=None):
    """Hộp bao mọi thứ được vẽ: nét/hình (trừ nét khung), ảnh, từng ký tự (hộp glyph chính xác)."""
    import pymupdf
    out = []
    bid = {id(d) for d in border}
    for d in (drawings if drawings is not None else page.get_drawings()):
        if id(d) in bid: continue
        r = d['rect']; w = (d.get('width') or 0) / 2
        if d.get('type') == 'f' and (d.get('fill') in ((1, 1, 1), (1.0, 1.0, 1.0))) and r.width > page.rect.width * 0.9: continue   # nền trang trắng
        out.append(('hình', (r.x0 - w, r.y0 - w, r.x1 + w, r.y1 + w)))
    for im in page.get_image_info(): out.append(('ảnh', tuple(im['bbox'])))
    for b in page.get_text('rawdict', flags=pymupdf.TEXT_ACCURATE_BBOXES)['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                for c in s['chars']:
                    if c['c'].strip(): out.append(('chữ', tuple(c['bbox'])))
    return out

def check_frame(pdf_path, sect_frame=None, tol=FRAME_TOL_PT):
    """Trả về list lỗi: {page, kind, side, bbox, gap}. gap < 0: tràn ra ngoài khung; 0..tol: chạm khung."""
    import pymupdf
    doc = pymupdf.open(pdf_path); bad = []
    for i, page in enumerate(doc):
        dr = page.get_drawings(); fr, border = frame_rect(page, sect_frame, dr)
        if fr is None: continue
        L, T, Rr, B_ = fr
        for kind, (x0, y0, x1, y1) in content_boxes(page, border, dr):
            if x1 - x0 <= 0.01 and y1 - y0 <= 0.01: continue
            gaps = dict(left=x0 - L, top=y0 - T, right=Rr - x1, bottom=B_ - y1)
            side, g = min(gaps.items(), key=lambda kv: kv[1])
            if g < tol: bad.append(dict(page=i + 1, kind=kind, side=side, bbox=[round(v, 1) for v in (x0, y0, x1, y1)], gap=round(g, 2)))
    return bad

BOUNDARY = re.compile(r'^\s*(ĐÁP ÁN|HƯỚNG DẪN CHẤM|Bài số|Bài làm|Đề nghe|Đọc thầm|Đọc thành tiếng|UBND|PHẦN|[IV]+\s*\.)', re.I)
def question_spans(pdf_path):
    """Mốc theo thứ tự trên trang: tiêu đề 'Câu n.' của phần đề -> (page, y0, n); tiêu đề phần/bài/đáp án -> (page, y0, None)."""
    import pymupdf
    doc = pymupdf.open(pdf_path); out = []; done = False
    for i, page in enumerate(doc):
        for b in page.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans'])
                if re.match(r'^\s*(ĐÁP ÁN|HƯỚNG DẪN CHẤM)', t): done = True
                m = QSTART.match(t)
                if m and not done and l['bbox'][0] < page.rect.width * 0.25: out.append((i + 1, l['bbox'][1], int(m.group(1))))
                elif BOUNDARY.match(t) or (m and done): out.append((i + 1, l['bbox'][1], None))
    return sorted(out, key=lambda s: (s[0], s[1]))

def owner_question(spans, v):
    """Câu chứa nội dung lỗi: mốc gần nhất phía trên; là tiêu đề phần/đáp án (None) -> không thuộc câu nào (không tự sửa được)."""
    cand = [s for s in spans if (s[0], s[1]) <= (v['page'], v['bbox'][1] + 0.5)]
    return cand[-1] if cand and cand[-1][2] is not None else None

def build_checked(template, exam, fields, out_path, media_dirs=(), with_answers=True, pages_dir=None, max_rounds=5, pages_png=True):
    """Dựng -> render (LibreOffice) -> kiểm tra khung -> sửa (ngắt trang trước câu / co hình) -> dựng lại, tối đa max_rounds lần.
    Trả về (warns, hlog, pages|None, violations). violations rỗng = sạch; khác rỗng = KHÔNG giao im lặng (UI cảnh báo, tải cần xác nhận)."""
    fix = dict(breaks=set(), shrink={})
    pages, bad = None, []
    for rnd in range(max_rounds):
        warns, hlog, sframe = build(template, exam, fields, out_path, media_dirs, with_answers, fix=fix)
        if not pages_dir: return warns, hlog, None, []
        pages = render_pages(out_path, pages_dir, with_png=pages_png)
        if pages is None: return warns, hlog, None, []
        pdf = os.path.join(pages_dir, 'de.pdf')
        bad = check_frame(pdf, sframe)
        if not bad: return warns, hlog, pages, []
        spans = question_spans(pdf); changed = False; done = set()
        first_on_page = {}
        for s in spans:
            if s[2] is not None or s[0] not in first_on_page: first_on_page.setdefault(s[0], s[2])
        for v in bad:
            oq = owner_question(spans, v)
            if oq is None or oq[2] in done: continue
            so = oq[2]; done.add(so)
            if v['side'] == 'bottom' and so not in fix['breaks'] and first_on_page.get(oq[0]) != so:
                fix['breaks'].add(so); changed = True        # cả câu sang trang sau
            else:
                f = fix['shrink'].get(so, 1.0)
                if f > 0.45: fix['shrink'][so] = round(f * 0.8, 3); changed = True   # co hình của câu
        if os.environ.get('EXAM_DEBUG_FRAME'): print('vòng', rnd + 1, len(bad), fix, bad[:3])
        log.info('khung trang: vòng %d, %d lỗi, sửa=%s', rnd + 1, len(bad), {k: sorted(v) if isinstance(v, set) else v for k, v in fix.items()})
        if not changed: break
    return warns, hlog, pages, bad

def frame_summary(bad):
    pages = sorted({v['page'] for v in bad})
    return f"Còn {len(bad)} chỗ chạm/tràn khung viền trang (trang {', '.join(map(str, pages))}). Hãy sửa nội dung (bớt hình/thu nhỏ) rồi bấm “Cập nhật đề” trước khi tải."
