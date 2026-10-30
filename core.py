"""Lõi xử lý docx: báo bài & giáo án."""
import copy, io, re, unicodedata, datetime, logging, os
import docx
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from lxml import etree

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
EMB = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'

def nfc(s): return unicodedata.normalize('NFC', s or '')
def etext(e): return nfc(''.join(x.text or '' for x in e.iter(qn('w:t'))))
def ctext(c): return nfc(c.text)

# ---------------- môn học ----------------
SUBJ_ALIAS = {
    'tv': ['tiếng việt'], 'toan': ['toán'],
    'ontv': ['ôn tv', 'ôn tiếng việt', 'tc (ôn tv)', 'tự chọn', 'clb tiếng việt', 'clb đọc sách', 'luyện tiếng việt'],
    'ontoan': ['ôn toán', 'clb toán', 'luyện toán'],
    'hdtn': ['hđtn', 'hoạt động trải nghiệm', 'shl', 'sinh hoạt lớp', 'shdc', 'chào cờ', 'sinh hoạt dưới cờ'],
    'tnxh': ['tnxh', 'tự nhiên và xã hội', 'tự nhiên xã hội'],
    'daoduc': ['đạo đức'], 'thuvien': ['thư viện đọc', 'thư viện'],
    'chamsoc': ['chăm sóc hs', 'chăm sóc học sinh', 'hoạt động giáo dục sau giờ chính thức'],
    'gdtc': ['gdtc', 'giáo dục thể chất'], 'amnhac': ['âm nhạc'], 'mithuat': ['mĩ thuật', 'mỹ thuật'],
    'tinhoc': ['tin học'], 'tienganh': ['tiếng anh'], 'clb': ['clb sở thích'],
}
_A = {v: k for k, vs in SUBJ_ALIAS.items() for v in vs}
def subj(s):
    s = re.sub(r'\s+', ' ', nfc(s).strip().lower()).rstrip('.:')
    if s in _A: return _A[s]
    s2 = ABBR.get(s, s); s2 = BB_ALIAS.get(s2, s2)
    return _A.get(s2, s)
BB_ALIAS = {'tc (ôn tv)': 'ôn tv', 'tc (ôn toán)': 'ôn toán', 'chăm sóc học sinh': 'chăm sóc hs', 'mỹ thuật': 'mĩ thuật'}
def subj_bb(s):
    s = re.sub(r'\s+', ' ', nfc(s).strip().lower()); return BB_ALIAS.get(s, s)
def subj_in_text(s):
    s = nfc(s).lower()
    m = re.search(r'môn\s+(.+?)(?:\s+lớp|\s*-|$)', s)
    if m and subj(m.group(1)) in SUBJ_ALIAS: return subj(m.group(1))
    m = re.search(r'tiết\s*\d[\d\s+,&\-–]*\s*:\s*([^()\d:]+)', s)
    if m and subj(m.group(1)) in SUBJ_ALIAS: return subj(m.group(1))
    m = re.search(r'tiết\s*\d[\d\s+,&\-–]*\s+([^()\d:]+?)\s*:', s)
    if m and subj(m.group(1)) in SUBJ_ALIAS: return subj(m.group(1))
    r = subj(s)
    if r in SUBJ_ALIAS: return r
    # phân môn ('TV (Đọc)', 'TV (Viết)', 'Đọc TV'...) -> mã họ ('tv') để ghép theo môn không bị sót
    f = _fam_key(skey(s))
    if f == 'thư viện': return 'thuvien'
    for code in SUBJ_ALIAS:
        if _fam_of(code) == f: return code
    return None

# ---------------- run helpers ----------------
def tnr_rpr(sz=24):
    return parse_xml(f'<w:rPr {WNS}><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')

def set_cell_text(cell, text, rpr=None):
    p = cell.paragraphs[0]._p
    for r in p.findall(qn('w:r')): p.remove(r)
    r = etree.SubElement(p, qn('w:r'))
    if rpr is None: rpr = tnr_rpr()
    r.append(rpr)
    t = etree.SubElement(r, qn('w:t')); t.text = text; t.set(XML_SPACE, 'preserve')

def replace_span(p, start, end, new):
    """Thay ký tự [start,end) của đoạn văn bản, giữ định dạng run."""
    pos = 0; first = True
    for r in p.runs:
        s = r.text; a, b = pos, pos + len(s); pos = b
        if b <= start or a >= end: continue
        i, j = max(start, a) - a, min(end, b) - a
        r.text = s[:i] + (new if first else '') + s[j:]; first = False

# ---------------- nhận dạng cấu trúc bảng (theo mẫu của từng người dùng) ----------------
import contextvars
CMAP = contextvars.ContextVar('CMAP', default={})   # {chữ ký tiêu đề: {role: cột, 'hdr': số dòng tiêu đề}} – do người dùng tự ghép cột
ROLES = ('thu', 'buoi', 'tiet', 'mon', 'ten', 'tich', 'dd', 'tuan')
ROLE_VI = dict(thu='Thứ / Ngày', buoi='Buổi', tiet='Tiết', mon='Môn', ten='Tên bài', tich='Tích hợp / Điều chỉnh', dd='Đồ dùng', tuan='Tuần')
SYN = {
    'tich': ['điều chỉnh', 'tích hợp', 'ghi chú', 'lồng ghép', 'nd đc', 'đc-th', 'đc th'],
    'dd': ['đddh', 'đ.d.d.h', 'đồ dùng', 'đdht', 'thiết bị', 'đddh/tb', 'phương tiện'],
    'ten': ['tên bài', 'tên bài dạy', 'bài dạy', 'tên bài học', 'nội dung bài', 'nội dung dạy', 'nội dung', 'bài học', 'đầu bài', 'tên chủ đề'],
    'mon': ['môn', 'phân môn', 'hđgd', 'môn học', 'môn/hđ', 'hoạt động giáo dục'],
    'tiet': ['tiết', 'tt', 'tiết tkb', 'tiết ppct', 'tiết ct', 'stt', 'ppct'],
    'buoi': ['buổi', 'ca', 'sáng/chiều', 'thời gian'],
    'thu': ['thứ', 'ngày', 'thứ/ngày', 'ngày/thứ', 'thứ ngày'],
    'tuan': ['tuần'],
}
EXACT = {'tt', 'ca', 'stt', 'ppct', 'tuần'}
def _hn(s): return re.sub(r'\s+', ' ', re.sub(r'[\n\r\t:()\[\]]', ' ', nfc(s).lower())).strip()
def head_role(h):
    h = _hn(h)
    if not h: return None
    for role in ('tich', 'dd', 'ten', 'mon', 'tiet', 'buoi', 'thu', 'tuan'):
        for k in SYN[role]:
            if k in EXACT:
                if h == k or h.startswith(k + ' ') and role == 'tiet': return role
            elif re.search(r'(^|[\s/,.\-])' + re.escape(k) + r'($|[\s/,.\-])', h):
                if role == 'tich' and re.search(r'tên bài', h) and not re.search(r'điều chỉnh|tích hợp', h): continue
                if role == 'ten' and k == 'nội dung' and re.search(r'điều chỉnh|tích hợp|ghi chú', h): continue
                if role == 'thu' and k == 'ngày' and re.search(r'tiết|tên', h): continue
                return role
    return None

def _uniq_cells(row):
    """[(cột lưới, ô)] – bỏ ô lặp do gộp ngang."""
    out = []; seen = []
    for i, c in enumerate(row.cells):
        if any(c._tc is e for e in seen): continue
        seen.append(c._tc); out.append((i, c))
    return out
def _is_sep(row):
    u = {id(c._tc) for c in row.cells}
    return len(u) == 1 and len(row.cells) > 1
def _is_hdr_dup(r, hrow, cols):
    """Dòng dữ liệu lặp y hệt dòng tiêu đề (có mẫu in hàng tiêu đề 2 lần) -> không phải bài học."""
    if hrow is None or 'mon' not in cols or 'ten' not in cols: return False
    def g(row, k):
        j = cols[k]
        return _hn(ctext(row.cells[j])) if j < len(row.cells) else ''
    hm = g(hrow, 'mon')
    if not hm or g(r, 'mon') != hm or g(r, 'ten') != g(hrow, 'ten'): return False
    return 'tiet' not in cols or g(r, 'tiet') == g(hrow, 'tiet')

def table_title_week(t, nrows, sy):
    """Số tuần + khoảng ngày ghi trong các dòng đầu của chính bảng (tựa đề nằm trong bảng)."""
    wk = None; rng = None
    for r in t.rows[:max(0, nrows)]:
        for _, c in _uniq_cells(r):
            tx = ctext(c)
            if wk is None:
                m = WEEK_RE.search(tx)
                if m: wk = int(m.group(1))
            if rng is None: rng = parse_range(tx, sy)
        if wk is not None and rng is not None: break
    return wk, rng

def head_sig(t, n=1):
    return (' | '.join(' ¦ '.join(_hn(ctext(c)) for _, c in _uniq_cells(r)) for r in t.rows[:n])[:300] + f' #{len(t.rows[0].cells)}') if t.rows else ''

def detect(t):
    """-> dict(cols={role: cột}, hdr=số dòng tiêu đề, ok, sig) hoặc None nếu không giống bảng kế hoạch."""
    if not t.rows: return None
    sig = head_sig(t); ov = CMAP.get().get(sig)
    if ov:
        cols = {k: int(v) for k, v in ov.items() if k in ROLES and v not in (None, '', -1, '-1')}
        d = dict(cols=cols, hdr=int(ov.get('hdr', 1)), ok=True, sig=sig, manual=True)
        if ov.get('week'): d['week'] = int(ov['week'])
        if ov.get('range'): d['range'] = tuple(ov['range'])
        return d
    best = None
    for hdr in (1, 2, 3):
        if hdr > len(t.rows): break
        cols = {}
        ncol = len(t.rows[0].cells)
        for j in range(ncol):
            seen = []; parts = []
            for r in t.rows[:hdr]:
                if j < len(r.cells):
                    el = r.cells[j]._tc
                    if any(el is e for e in seen): continue
                    seen.append(el); parts.append(ctext(r.cells[j]))
            # dòng dưới (con) ưu tiên hơn dòng trên (nhóm) khi đặt tên
            role = None
            for p in reversed(parts):
                role = head_role(p)
                if role: break
            if role and role not in cols: cols[role] = j
        sc = len(cols) + (3 if 'mon' in cols and 'ten' in cols else 0)
        if not best or sc > best[0]: best = (sc, dict(cols=cols, hdr=hdr))
    d = best[1]; d['ok'] = 'mon' in d['cols'] and 'ten' in d['cols']; d['sig'] = sig; d['manual'] = False
    return d if d['cols'] else None

class Blk:
    """Một tuần trong bảng của người dùng: bảng + các dòng dữ liệu + sơ đồ cột."""
    def __init__(s, table, sch, rows):
        s.table, s.sch, s.rows = table, sch, rows
    @property
    def cols(s): return s.sch['cols']

DAYNUM = {'hai': '2', 'ba': '3', 'tư': '4', 'tu': '4', 'năm': '5', 'sáu': '6', 'bảy': '7', 'cn': '8', 'chủ nhật': '8'}
def norm_thu(s, year=None):
    t = nfc(s).strip().lower()
    if not t: return ''
    m = re.match(r'^(?:thứ|t)?\s*([2-7])\b', t)
    if m: return m.group(1)
    m = re.search(r'thứ\s*(hai|ba|tư|năm|sáu|bảy)', t) or re.match(r'^(hai|ba|tư|năm|sáu|bảy)\b', t)
    if m: return DAYNUM[m.group(1)]
    m = re.search(r'(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?', t)
    if m:
        y = int(m.group(3)) if m.group(3) else (year or datetime.date.today().year)
        if y < 100: y += 2000
        try: return str(datetime.date(y, int(m.group(2)), int(m.group(1))).weekday() + 2)
        except ValueError: pass
    return t

def _rows(blk_or_t):
    if isinstance(blk_or_t, Blk): return blk_or_t
    sch = detect(blk_or_t)
    if not sch: return Blk(blk_or_t, dict(cols={}, hdr=1, ok=False), [])
    return Blk(blk_or_t, sch, [r for r in blk_or_t.rows[sch['hdr']:] if not _is_sep(r)])

WEEK_RE = re.compile(r'\btuần\s*(?:thứ\s*)?[:.\-–]?\s*0?(\d{1,2})\b', re.I)
DATE_RE = r'(\d{1,2})\s*[/.\-]\s*(\d{1,2})(?:\s*[/.\-]\s*(\d{2,4}))?'
RANGE_RE = re.compile(DATE_RE + r'\s*(?:→|->|–|—|-|đến|tới|~)\s*(?:ngày\s*)?' + DATE_RE, re.I)
def school_year(d):
    txt = ' '.join(p.text for p in d.paragraphs[:200])
    m = re.search(r'năm\s*học\s*:?\s*(20\d\d)', txt, re.I)
    if m: return int(m.group(1))
    m = re.search(r'\d{1,2}/(\d{1,2})/(20\d\d)', txt)
    if m: return int(m.group(2)) - (1 if int(m.group(1)) < 8 else 0)
    t = datetime.date.today(); return t.year - (1 if t.month < 8 else 0)
def parse_range(t, sy):
    m = RANGE_RE.search(nfc(t))
    if not m: return None
    g = m.groups(); out = []
    for d, mo, y in (g[0:3], g[3:6]):
        if y: yy = int(y) + (2000 if int(y) < 100 else 0); y = y if len(y) == 4 else str(yy)
        else: y = str(sy if int(mo) >= 8 else sy + 1)
        if not (1 <= int(mo) <= 12 and 1 <= int(d) <= 31): return None
        out.append(f'{int(d)}/{int(mo)}/{y}' if len(g) else '')
    return tuple(out)

def _hdr_week(d, sy):
    """Số tuần ghi ở đầu trang (header) / chân trang (footer) của Word – nhiều trường xuất file
    để 'TUẦN n' ở header thay vì trong thân văn bản; week_tables phải ngó cả ở đó."""
    wk = None; rng = None; texts = []
    for s in d.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header, s.even_page_footer):
            try:
                texts += [p.text for p in part.paragraphs]
                for t in part.tables:
                    for r in t.rows:
                        texts += [c.text for c in r.cells]
            except Exception: continue
    for t in texts:
        m = WEEK_RE.search(t or '')
        if m: wk = int(m.group(1))
        r2 = parse_range(t or '', sy)
        if r2: rng = r2
    return wk, rng

def week_tables(d):
    """-> {tuần: (Blk, (ngày bắt đầu, ngày kết thúc) or None)} – theo mẫu của người dùng:
    tiêu đề 'TUẦN n'/'Tuần n (từ … đến …)', một bảng mỗi tuần hoặc một bảng cả năm (dòng 'Tuần n' hoặc cột Tuần)."""
    out = _scan_weeks(d, None, None)
    if out: return out
    sy = school_year(d); wk, rng = _hdr_week(d, sy)   # thân không ghi tuần -> thử header/footer
    return _scan_weeks(d, wk, rng) if wk is not None else out

def _scan_weeks(d, wk, rng):
    out = {}; sy = school_year(d)
    tbls = {t._tbl: t for t in d.tables}
    for e in d.element.body:
        if e.tag == qn('w:p'):
            t = etext(e)
            m = WEEK_RE.search(t)
            if m: wk = int(m.group(1)); rng = None
            r2 = parse_range(t, sy)
            if r2: rng = r2
        elif e.tag == qn('w:tbl') and e in tbls:
            try: wk, rng = _week_tbl(tbls[e], out, wk, rng, sy)
            except Exception: continue      # bảng lạ (chữ ký, khung…) không được làm hỏng cả tệp
    return out

def _week_tbl(t, out, wk, rng, sy):
    """Đọc 1 bảng; trả về (tuần, khoảng ngày) hiện hành sau bảng."""
    lead = 0
    while lead < len(t.rows) - 1 and _is_sep(t.rows[lead]):
        tx = ctext(t.rows[lead].cells[0]); m = WEEK_RE.search(tx)
        if m: wk = int(m.group(1)); rng = None
        rng = parse_range(tx, sy) or rng; lead += 1
    sch = detect_from(t, lead)
    if not sch or not sch['ok']: return wk, rng
    tw, trng = table_title_week(t, lead + sch['hdr'], sy)
    hrow = t.rows[lead + sch['hdr'] - 1] if sch['hdr'] >= 1 else None
    body = [r for r in t.rows[lead + sch['hdr']:] if not _is_hdr_dup(r, hrow, sch['cols'])]
    seps = [i for i, r in enumerate(body) if _is_sep(r) and WEEK_RE.search(ctext(r.cells[0]))]
    if seps or 'tuan' in sch['cols']:     # một bảng cho nhiều tuần
        cw = None; cr = None; groups = {}
        for r in body:
            if _is_sep(r):
                tx = ctext(r.cells[0]); m = WEEK_RE.search(tx)
                if m: cw = int(m.group(1)); cr = parse_range(tx, sy)
                continue
            if 'tuan' in sch['cols']:
                tx = ctext(r.cells[sch['cols']['tuan']]); m = re.search(r'(\d{1,2})', tx)
                if m and int(m.group(1)) != cw: cw = int(m.group(1)); cr = parse_range(tx, sy)
            if cw is not None: groups.setdefault(cw, ([], cr))[0].append(r)
        for w, (rs, cr) in groups.items():
            if w not in out: out[w] = (Blk(t, sch, rs), cr)
    elif sch.get('week') or tw or wk:
        w2 = sch.get('week') or tw or wk
        if w2 not in out:
            out[w2] = (Blk(t, sch, [r for r in body if not _is_sep(r)]), sch.get('range') or trng or rng)
    return tw or wk, trng or rng

def detect_from(t, lead=0):
    if not lead: return detect(t)
    class _T:  # bảng nhìn từ dòng 'lead'
        rows = t.rows[lead:]
    return detect(_T)

def rows_of(t, year=None):
    b = _rows(t); C = b.cols; res = []; prev = {}
    for r in b.rows:
        c = r.cells; g = lambda k: (ctext(c[C[k]]) if k in C and C[k] < len(c) else '')
        cm = {k: (c[C[k]] if k in C and C[k] < len(c) else None) for k in ('thu', 'buoi', 'tiet', 'mon', 'ten', 'tich', 'dd')}
        thu = norm_thu(g('thu'), year) or prev.get('thu', ''); buoi = g('buoi').strip()
        if not buoi and 'buoi' in C: buoi = prev.get('buoi', '')
        d = dict(thu=thu, buoi=buoi, tiet=g('tiet').strip(), mon=re.sub(r'\s+', ' ', g('mon')).strip(), ten=g('ten'), tich=g('tich'), dd=g('dd'), c=cm, cells=c, row=r)
        prev = dict(thu=thu, buoi=buoi); res.append(d)
    # ô gộp dọc: chỉ dòng đầu tiên được ghi
    for k in ('ten', 'tich', 'dd'):
        seen = set()
        for d in res:
            x = d['c'][k]
            if x is None: continue
            if id(x._tc) in seen: d['c'][k] = None
            else: seen.add(id(x._tc))
    return res

# ---- môn: chuẩn hoá tên/viết tắt (không gộp SHDC/SHL vào HĐTN) ----
ABBR = {'tv': 'tiếng việt', 't.việt': 'tiếng việt', 'tviet': 'tiếng việt', 'toán học': 'toán', 'tnxh': 'tự nhiên và xã hội', 'tự nhiên xã hội': 'tự nhiên và xã hội',
        'tn&xh': 'tự nhiên và xã hội', 'tn và xh': 'tự nhiên và xã hội', 'hdtn': 'hđtn', 'hoạt động trải nghiệm': 'hđtn', 'hđ trải nghiệm': 'hđtn', 'đđ': 'đạo đức', 'dd': 'đạo đức',
        'gdtc': 'giáo dục thể chất', 'thể dục': 'giáo dục thể chất', 'td': 'giáo dục thể chất', 'an': 'âm nhạc', 'mt': 'mĩ thuật', 'mỹ thuật': 'mĩ thuật', 'ta': 'tiếng anh',
        'th': 'tin học', 'ôn tiếng việt': 'ôn tv', 'ôn t.việt': 'ôn tv', 'luyện tv': 'ôn tv', 'luyện tiếng việt': 'ôn tv', 'luyện toán': 'ôn toán', 'chào cờ': 'shdc',
        'sinh hoạt dưới cờ': 'shdc', 'sinh hoạt lớp': 'shl', 'tnxh.': 'tự nhiên và xã hội'}
def skey(s):
    s = subj_bb(s); s = re.sub(r'[.:]+$', '', s).strip()
    return ABBR.get(s, s)
def skey_fuzzy(a, keys):
    """-> (khoá khớp, chắc chắn?)"""
    import difflib
    k = skey(a)
    if k in keys: return k, True
    m = difflib.get_close_matches(k, list(keys), n=2, cutoff=0.75)
    if m: return m[0], False
    return None, False

def _fam_key(k):
    """Khoá HỌ môn – bỏ phần phân môn để hai trường gọi cùng một môn vẫn về một họ:
    'tv (đọc)' -> 'tiếng việt', 'tv( viết)' -> 'tiếng việt', 'đọc tv' -> 'tiếng việt'."""
    f = re.sub(r'\(.*?\)', '', k or '').strip()
    f = re.sub(r'\s+', ' ', f)
    if f == 'đọc tv' or 'thư viện' in f: return 'thư viện'   # 'Đọc TV' = 'Thư viện đọc': tiết đọc thư viện, KHÔNG phải Tiếng Việt
    m = re.match(r'^(đọc|viết|nghe|nói)\s+(.+)$', f)
    if m: f = m.group(2).strip()
    f = re.sub(r'^(ôn|luyện)\s+', '', f)
    f = ABBR.get(f, f)
    if any(p.strip() in ('hđtn', 'shl', 'shdc', 'sinh hoạt lớp', 'sinh hoạt dưới cờ') for p in re.split(r'[-–/+,]', f)):
        return 'hđtn'   # SHL (sinh hoạt lớp) / SHDC CHÍNH LÀ HĐTN — mỗi trường viết tắt một kiểu
    return {'shdc': 'hđtn', 'shl': 'hđtn'}.get(f, f)   # Sinh hoạt dưới cờ / Sinh hoạt lớp thuộc họ HĐTN

_CODE2NAME = {}
for _k, _v in _A.items(): _CODE2NAME.setdefault(_v, _k)

def _fam_of(x):
    """Họ môn cho MỌI kiểu gọi: tên đầy đủ ('Toán'), phân môn ('TV (Đọc)'), mã môn ('toan', 'ontv', 'daoduc')."""
    if not x: return ''
    xl = x.strip().lower()
    if xl == 'thuvien': return 'thư viện'
    f = _fam_key(skey(x))
    if f == skey(x) and xl in _CODE2NAME:
        f = _fam_key(_CODE2NAME[xl])
    return f

def keyed(rows):
    cnt = {}; out = []
    for r in rows:
        k = (r['thu'], '', skey(r['mon'])); cnt[k] = cnt.get(k, 0) + 1
        out.append((k + (cnt[k],), r))
    return out

RED = 'FF0000'
def cell_rpr(cell, row=None, col_cells=()):
    """Định dạng chữ của chính bảng người dùng: run trong ô -> rPr của đoạn -> ô cùng cột -> ô cùng dòng -> TNR 12."""
    def first_rpr(c):
        if c is None: return None
        for r in c._tc.iter(qn('w:r')):
            if etext(r).strip():
                rp = r.find(qn('w:rPr'))
                return copy.deepcopy(rp) if rp is not None else parse_xml(f'<w:rPr {WNS}/>')
        return None
    rp = first_rpr(cell)
    for c in col_cells:  # chữ đang hiện trong cùng cột của bảng người dùng
        if rp is not None: break
        if c is not cell: rp = first_rpr(c)
    if rp is None and cell is not None:
        pp = cell.paragraphs[0]._p.find(qn('w:pPr'))
        if pp is not None and pp.find(qn('w:rPr')) is not None:
            rp = copy.deepcopy(pp.find(qn('w:rPr'))); rp.tag = qn('w:rPr')
    for c in list(row or []):
        if rp is not None: break
        rp = first_rpr(c)
    if rp is None: rp = tnr_rpr()
    for x in rp.findall(qn('w:color')): rp.remove(x)
    for x in rp.findall(qn('w:highlight')): rp.remove(x)
    return rp
def red(rp):
    c = etree.SubElement(rp, qn('w:color')); c.set(qn('w:val'), RED); return rp
def is_red(cell):
    if cell is None: return False
    for r in cell._tc.iter(qn('w:r')):
        c = r.find(qn('w:rPr') + '/' + qn('w:color'))
        if c is not None and (c.get(qn('w:val')) or '').upper() == RED and etext(r).strip(): return True
    return False

def clear_cell(cell):
    """Xoá chữ trong ô, giữ viền/gộp ô/định dạng đoạn đầu."""
    ps = cell._tc.findall(qn('w:p'))
    for p in ps[1:]: cell._tc.remove(p)
    p = ps[0]
    for x in list(p):
        if x.tag != qn('w:pPr'): p.remove(x)

def week_is_empty(t):
    return all(not r['ten'].strip() for r in rows_of(t) if r['mon'])

def preview(t):
    return [dict(idx=i, thu=r['thu'], buoi=r['buoi'] or r['tiet'], mon=r['mon'], ten=r['ten'].strip(), tich=r['tich'].strip(),
                 empty=bool(r['mon']) and not r['ten'].strip(), red=is_red(r['c']['ten']) or is_red(r['c']['tich']))
            for i, r in enumerate(rows_of(t))]

def _col(rows, k): return [r['c'][k] for r in rows if r['c'].get(k) is not None]

def fill_baobai(master_path, phu_path, out_path, week=None, to_week=None):
    # week: tuần của BẢN PHỤ; to_week: tuần của BẢN GỐC (mặc định = week).
    # Việc gióng nguồn->đích do tầng trên quyết (main.resolve_targets); động cơ chỉ chép đúng cặp được giao.
    """Điền tuần trống của bản gốc theo bản phụ – mỗi tệp nhận dạng cấu trúc riêng.
    Ghép theo Thứ + Môn (+ Tiết nếu cả hai có), rồi theo trình tự môn trong tuần.
    Ô ghép không chắc chắn (khác thứ, khác tiết, môn gần đúng) -> chữ ĐỎ + danh sách."""
    M = docx.Document(master_path); S = docx.Document(phu_path)
    mw, sw = week_tables(M), week_tables(S)
    if not mw: raise ValueError('Không nhận ra bảng kế hoạch tuần trong báo bài gốc')
    if not sw: raise ValueError('Không nhận ra bảng kế hoạch tuần trong báo bài phụ')
    if not week: raise ValueError('Chưa rõ tuần của báo bài phụ')
    src = int(week); dst = int(to_week) if to_week else src
    if src not in sw: raise ValueError(f'Không tìm thấy tuần {src} trong bản phụ')
    if dst not in mw: raise ValueError(f'Không tìm thấy tuần {dst} trong bản gốc')
    weeks = [(src, dst)]   # chỉ đúng cặp tuần được giao; tuần khác giữ nguyên từng byte
    report = {'weeks': [], 'skipped_weeks': []}
    for src_w, w in weeks:
        G = rows_of(mw[w][0]); P = [r for r in rows_of(sw[src_w][0]) if r['ten'].strip()]
        # bản làm việc: xoá cột Tên bài + Tích hợp của ĐÚNG tuần này (giữ định dạng để ghi lại)
        keep_fmt = {}
        for g in G:
            for k in ('ten', 'tich'):
                cell = g['c'][k]
                if cell is None: continue
                if g[k].strip():   # giữ định dạng chữ cũ của chính ô này để ghi lại
                    keep_fmt[id(cell._tc)] = cell_rpr(cell, g['cells'], [x['c'][k] for x in G if x['c'][k] is not None]); clear_cell(cell)
                g[k] = ''
        gk = {skey(r['mon']) for r in G if r['mon']}
        for r in P:
            r['k'], r['sure'] = skey_fuzzy(r['mon'], gk)
        # HỌ MÔN: một bên tách PHÂN MÔN (TV (Đọc), TV (Viết), Đọc TV...), bên kia ghi GỘP (Tiếng Việt)
        # – đưa cả hai bên về khoá họ rồi mới ghép nhóm (vẫn neo Thứ/Buổi/Tiết ở lượt ghép bên dưới).
        _allk = gk | {skey(r['mon']) for r in P}
        _fams = {_fam_key(x) for x in _allk}
        def _canon(k):
            if not k: return k
            f = _fam_key(k)
            return f if f != k and (f in _allk or f in _fams) else k
        _gk2 = {_canon(x) for x in gk}
        for r in P:
            if r['k']: r['k'] = _canon(r['k'])
            else:
                f = _canon(skey(r['mon']))
                if f in _gk2: r['k'], r['sure'] = f, True
        for g in G: g['gk'] = _canon(skey(g['mon']))
        has_tich = 'tich' in mw[w][0].cols and 'tich' in sw[src_w][0].cols
        has_dd = 'dd' in mw[w][0].cols and 'dd' in sw[src_w][0].cols
        targets = [r for r in G if r['mon']]
        used = set(); asg = {}
        from collections import Counter
        cg = Counter(g['gk'] for g in targets); cp = Counter(p['k'] for p in P)
        # CÁCH ĐIỀN THEO HƯỚNG DẪN CỦA GV: dán TỪNG MÔN MỘT — lấy các tiết của môn đó trong báo bài
        # phụ theo đúng THỨ TỰ LẦN LƯỢT, dán lần lượt vào các ô của môn đó trong báo bài gốc cho đến
        # hết. KHÔNG neo theo Thứ/Tiết (neo ngày là nguồn dán nhầm khi thứ tự hai bên lệch nhau).
        for k in cg:
            gi = [i for i, g in enumerate(targets) if g['gk'] == k]; pj = [j for j, p in enumerate(P) if p['k'] == k]
            # Riêng họ HĐTN khi bản phụ THIẾU tiết: điền TỪ DƯỚI LÊN — tiết SHDC đầu tuần thường
            # được bỏ trống, neo từ cuối giúp tiết SHL (sinh hoạt lớp) ở cuối tuần khớp đúng bài,
            # tránh dồn lệch lên trên và báo thiếu tiết oan (theo hướng dẫn của GV).
            if _fam_key(k) == 'hđtn' and len(pj) < len(gi): pairs = zip(reversed(gi), reversed(pj))
            else: pairs = zip(gi, pj)
            for i, j in pairs:
                used.add(j); asg[i] = (P[j], True, [])
        filled, empty, reds = [], [], []
        colc = {k: _col(G, k) for k in ('ten', 'tich', 'dd')}
        for i, g in enumerate(targets):
            lab = f"T{g['thu']} {g['buoi'] or ('tiết ' + g['tiet'] if g['tiet'] else '')} {g['mon']}".replace('  ', ' ')
            if i not in asg:
                if not g['ten'].strip():
                    empty.append(lab)
                    # Môn CHUYÊN (Tiếng Anh, GDTC, Tin học, Mĩ thuật, Âm nhạc): nhiều trường để trống
                    # không điền tên bài -> ô trống của các môn này KHÔNG tính vào danh sách ô đỏ cần kiểm tra.
                    _sm = (g['mon'] or '').lower()
                    _is_shdc = ('dưới cờ' in _sm) or skey(g['mon']) in ('shdc', 'chào cờ')
                    if cp.get(g.get('gk')) and _fam_key(g.get('gk') or '') not in SPEC_FAMS and not _is_shdc:
                        reds.append(dict(row=lab, why='thiếu trong báo bài phụ (môn này có ở phụ nhưng không đủ tiết)'))
                continue
            p, sure, why = asg[i]
            if not p['sure']: why = [f"môn gần đúng (phụ: {p['mon']})"] + why  # khớp mờ tên môn (gõ sai/chưa chuẩn) -> luôn đỏ để thầy kiểm tra, kể cả khi đã về cùng họ môn
            sure = sure and not why
            did = []
            for k, f in (('ten', 'tên bài'), ('tich', 'tích hợp'), ('dd', 'đồ dùng')):
                if k != 'ten' and not (has_tich if k == 'tich' else has_dd): continue
                cell = g['c'][k]
                if cell is None or not p[k].strip(): continue
                if k == 'dd' and g[k].strip() and nfc(g[k]).strip() == nfc(p[k]).strip(): continue
                rp = copy.deepcopy(keep_fmt[id(cell._tc)]) if id(cell._tc) in keep_fmt else cell_rpr(cell, g['cells'], colc[k])
                set_cell_text(cell, nfc(p[k]), rp if sure else red(rp)); did.append(f)
            if not sure and did: reds.append(dict(row=lab, why=', '.join(why) or 'không chắc chắn', ten=p['ten'].strip()))
            filled.append(dict(thu=g['thu'], buoi=g['buoi'], mon=g['mon'], ten=p['ten'].strip(), tich=p['tich'].strip(),
                               tu=f"T{p['thu']} {p['buoi'] or p['tiet']} {p['mon']}", fields=did, sure=sure, why=why))
        unused = [f"T{p['thu']} {p['buoi'] or p['tiet']} {p['mon']}: {p['ten'].strip()}" for j, p in enumerate(P) if j not in used]
        report['weeks'].append(dict(week=w, from_week=src_w, filled=filled, empty=empty, unused=unused, red=reds, preview=preview(mw[w][0])))
    M.save(out_path)
    return report

def edit_row(master_path, w, idx, ten, tich, mon=None, out_path=None):
    """Sửa Tên bài / Tích hợp của đúng 1 dòng trong tuần w (giữ định dạng bảng của người dùng, bỏ màu đỏ). Không đụng tuần khác."""
    M = docx.Document(master_path); wt = week_tables(M)
    if w not in wt: raise KeyError(w)
    rows = rows_of(wt[w][0])
    if not (0 <= idx < len(rows)): raise IndexError(idx)
    r = rows[idx]
    if mon is not None and nfc(mon).strip() != r['mon']: raise ValueError('Dòng đã thay đổi, hãy tải lại trang')
    for k, v in (('ten', ten), ('tich', tich)):
        if v is not None and r['c'][k] is not None:
            set_cell_text(r['c'][k], nfc(v).strip(), cell_rpr(r['c'][k], r['cells'], _col(rows, k)) if r[k].strip() else tnr_rpr(24))
    M.save(out_path or master_path)
    return preview(week_tables(docx.Document(out_path or master_path))[w][0])

def common_weeks(master_path, phu_path):
    """Tuần có trong cả báo bài phụ và gốc (để giáo viên chọn)."""
    mw = week_tables(docx.Document(master_path)); sw = week_tables(docx.Document(phu_path))
    return [dict(week=w, range=mw[w][1]) for w in sorted(sw) if w in mw]

def master_weeks(master_path):
    M = docx.Document(master_path)
    return [dict(week=w, range=r, empty=week_is_empty(t)) for w, (t, r) in sorted(week_tables(M).items())]

def master_week_preview(master_path, w):
    M = docx.Document(master_path); return preview(week_tables(M)[w][0])

def inspect(path):
    """Cho giao diện ghép cột thủ công: các bảng + tiêu đề cột + đoán vai trò."""
    d = docx.Document(path); out = []
    for i, t in enumerate(d.tables[:40]):
        if len(t.rows) < 2 or len(t.rows[0].cells) < 3: continue
        sch = detect(t) or dict(cols={}, hdr=1, ok=False, sig=head_sig(t))
        heads = [' / '.join(dict.fromkeys(ctext(r.cells[j]).strip() for r in t.rows[:max(1, sch['hdr'])] if j < len(r.cells) and ctext(r.cells[j]).strip())) for j in range(len(t.rows[0].cells))]
        sample = [[ctext(c).strip()[:30] for c in r.cells] for r in t.rows[sch['hdr']:sch['hdr'] + 3]]
        out.append(dict(i=i, sig=sch['sig'], heads=heads, sample=sample, guess=sch['cols'], hdr=sch['hdr'], ok=sch['ok']))
        if len(out) >= 4: break
    return out

# ---------------- giáo án ----------------
def first_table(d):
    for t in d.tables:
        sch = detect(t)
        if sch and sch['ok'] and 'thu' in sch['cols']: return t
    for t in d.tables:
        if t.rows and 'thứ' in ctext(t.rows[0].cells[0]).lower(): return t

def _sz_of(cell):
    if cell is None: return None
    for r in cell._tc.iter(qn('w:r')):
        e = r.find(qn('w:rPr') + '/' + qn('w:sz'))
        if e is not None and etext(r).strip(): return e.get(qn('w:val'))
    pp = cell.paragraphs[0]._p.find(qn('w:pPr'))
    e = pp.find(qn('w:rPr') + '/' + qn('w:sz')) if pp is not None else None
    return e.get(qn('w:val')) if e is not None else None

def transplant_frame(G, M, w):
    """BƯỚC 1: bê NGUYÊN bảng báo bài của tuần w (trong báo bài gốc của người dùng) sang thay cho
    khung báo bài trong giáo án — không sửa một ký tự nào trong bảng (đúng yêu cầu: bê y nguyên
    phần báo bài đã tạo vào giáo án). Giáo án không có khung báo bài thì bỏ qua, trả False."""
    mw = week_tables(M)
    if w not in mw: return False
    src = mw[w][0].table._tbl
    wtG = week_tables(G)
    dst = None
    if w in wtG: dst = wtG[w][0].table._tbl
    else:
        ft = first_table(G)
        if ft is not None: dst = ft._tbl
    if dst is None: return False
    new = copy.deepcopy(src)
    dst.getparent().replace(dst, new)
    _frame_fit(new, G)
    return True

def _frame_fit(tbl, G):
    """Chuẩn hoá bảng báo bài SAU KHI bê nguyên vào giáo án (theo yêu cầu của GV):
    (1) toàn bộ chữ trong bảng về cỡ 12; (2) bề rộng bảng = đúng bề rộng chữ của trang giáo án
    (bằng vạch kẻ trên/dưới), cột chia lại theo đúng tỉ lệ cũ, layout cố định. Nội dung không đổi."""
    from docx.oxml import OxmlElement
    for r in tbl.iter(qn('w:r')):
        rPr = r.get_or_add_rPr()
        for tg in ('w:sz', 'w:szCs'):
            e = rPr.find(qn(tg))
            if e is None: e = OxmlElement(tg); rPr.append(e)
            e.set(qn('w:val'), '24')
    try:
        sec = G.sections[0]
        tw = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
    except Exception:
        return
    if tw <= 0: return
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None: return
    w = tblPr.find(qn('w:tblW'))
    if w is None: w = OxmlElement('w:tblW'); tblPr.append(w)
    w.set(qn('w:type'), 'dxa'); w.set(qn('w:w'), str(tw))
    lay = tblPr.find(qn('w:tblLayout'))
    if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
    lay.set(qn('w:type'), 'fixed')
    grid = tbl.find(qn('w:tblGrid'))
    widths = None
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if cols:
            try: old_ws = [int(c.get(qn('w:w')) or 0) for c in cols]
            except Exception: old_ws = [0] * len(cols)
            tot = sum(old_ws) or len(cols)
            widths = [max(1, round(tw * (x or 1) / tot)) for x in old_ws]
            widths[-1] += tw - sum(widths)
            for c, nw in zip(cols, widths): c.set(qn('w:w'), str(nw))
    if widths:
        for tr in tbl.findall(qn('w:tr')):
            ci = 0
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None: continue
                span = 1
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    try: span = int(gs.get(qn('w:val')) or 1)
                    except Exception: span = 1
                tcw = tcPr.find(qn('w:tcW'))
                if tcw is not None and ci < len(widths):
                    tcw.set(qn('w:type'), 'dxa'); tcw.set(qn('w:w'), str(sum(widths[ci:ci + span])))
                ci += span

def center_tiet_subjects(G):
    """Căn tên môn của mỗi khối TIẾT vào CHÍNH GIỮA trang (thẳng hàng hình quyển sách ở vạch
    phân cách phía trên): bỏ cách kê bằng khoảng trắng của file mẫu, đặt 1 tab căn giữa tại
    1/2 bề rộng chữ rồi chèn đúng 1 tab trước tên môn. Không đổi chữ và định dạng chữ."""
    from docx.oxml import OxmlElement
    try:
        sec = G.sections[0]
        center = int((sec.page_width - sec.left_margin - sec.right_margin) / 635) // 2
    except Exception:
        return 0
    if center <= 0: return 0
    def _tabbed(r_el, text):
        for ch in list(r_el):
            if ch.tag != qn('w:rPr'): r_el.remove(ch)
        r_el.append(OxmlElement('w:tab'))
        t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
        r_el.append(t)
    n = 0
    for p in G.paragraphs:
        if not p.text.startswith('TIẾT'): continue
        # Nhãn 'TIẾT n' có thể tách nhiều run ('TIẾT '+'2 + 3'), tên môn cũng vậy ('T'+'oán'):
        # tách đúng ranh giới nhãn/tên môn rồi mới căn giữa.
        label_runs, subj_runs = _tiet_split(p._p)
        if not subj_runs: continue
        subj = ''.join(''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in subj_runs).strip()
        if not subj: continue
        keep = {id(r) for r in label_runs} | {id(subj_runs[-1])}
        for r_el in p._p.findall(qn('w:r')):
            if id(r_el) in keep: continue
            for t in r_el.findall(qn('w:t')): t.text = ''
        _tabbed(subj_runs[-1], subj)
        pPr = p._p.get_or_add_pPr()
        for tb in pPr.findall(qn('w:tabs')): pPr.remove(tb)
        tabs = OxmlElement('w:tabs'); tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'center'); tab.set(qn('w:pos'), str(center)); tabs.append(tab)
        sp = pPr.find(qn('w:spacing'))
        if sp is not None: sp.addprevious(tabs)
        else: pPr.append(tabs)
        n += 1
    return n

def step1_copy_table(G, M, w):
    """BƯỚC 1 (TỰ ĐỘNG): đặt nguyên bảng báo bài đã tạo vào khung báo bài trong giáo án."""
    transplant_frame(G, M, w)
    return []

def _dparse(s):
    try: return datetime.datetime.strptime(str(s).strip(), '%d/%m/%Y').date()
    except Exception: return None

def week_range(M, w, G=None):
    """Khoảng ngày (Thứ Hai, Thứ Sáu) của tuần w: lấy từ báo bài; tuần đó thiếu thì suy từ tuần
    khác trong báo bài; vẫn thiếu thì suy từ tiêu đề / dòng Thứ Hai của giáo án mẫu (dịch đúng
    số tuần chênh). Trả (date, date) hoặc None."""
    mw = week_tables(M)
    def _norm(r):
        if not r: return None
        d1 = _dparse(r[0])
        if d1 is None: return None
        d1 -= datetime.timedelta(days=d1.weekday())
        return (d1, d1 + datetime.timedelta(days=4))
    if w in mw:
        got = _norm(mw[w][1])
        if got: return got
    for w2 in sorted(mw):
        got = _norm(mw[w2][1])
        if got:
            sh = datetime.timedelta(days=7 * (w - w2))
            return (got[0] + sh, got[1] + sh)
    if G is not None:
        tw = None; trng = None; tday = None
        for p in G.paragraphs:
            t = nfc(p.text)
            m = re.search(r'TUẦN\s+(\d+)', t, re.I)
            if m and tw is None:
                tw = int(m.group(1))
                m2 = re.search(r'(\d{1,2}/\d{1,2}/\d{4})\s*(?:→|->|–|-|ĐẾN(?:\s+NGÀY)?)\s*(\d{1,2}/\d{1,2}/\d{4})', t, re.I)
                if m2: trng = _norm((m2.group(1), m2.group(2)))
            m3 = re.search(r'Thứ\s+Hai\s*,?\s*ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)', t, re.I)
            if m3 and tday is None:
                try: tday = datetime.date(int(m3.group(3)), int(m3.group(2)), int(m3.group(1)))
                except ValueError: pass
            if tw is not None and (trng or tday): break
        base = trng[0] if trng else tday
        if base is not None and tw is not None:
            mon = base + datetime.timedelta(days=7 * (w - tw))
            return (mon, mon + datetime.timedelta(days=4))
    return None

DAYS = ['Hai', 'Ba', 'Tư', 'Năm', 'Sáu', 'Bảy']
DAY_RE = re.compile(r'^\s*Thứ\s+(Hai|Ba|Tư|Năm|Sáu|Bảy|[2-7])\b', re.I)
def day_idx(x):
    x = nfc(x).lower()
    return int(x) - 2 if x.isdigit() else [d.lower() for d in DAYS].index(x)
def step2_dates(G, w, rng):
    """Sửa ngày theo mẫu của người dùng: 'Thứ Hai ngày 19 tháng 10 năm 2026', 'Thứ 2, ngày 19/10/2026', tiêu đề 'TUẦN n … a → b'."""
    if not rng: return ['Không có khoảng ngày của tuần trong báo bài']
    d0 = datetime.datetime.strptime(rng[0], '%d/%m/%Y').date()
    d0 -= datetime.timedelta(days=d0.weekday())
    log = []
    for p in G.paragraphs:
        t = nfc(p.text)
        m = re.search(r'Thứ\s+(Hai|Ba|Tư|Năm|Sáu|Bảy|[2-7])\s*,?\s*ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)', t, re.I) or \
            re.search(r'Thứ\s+(Hai|Ba|Tư|Năm|Sáu|Bảy|[2-7])\s*[,:\-–]?\s*(?:ngày\s*)?(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})', t, re.I)
        if m:
            dd = d0 + datetime.timedelta(days=day_idx(m.group(1)))
            for g, v in ((4, str(dd.year)), (3, str(dd.month)), (2, str(dd.day))):
                if m.group(g) != v and int(m.group(g)) != int(v): replace_span(p, m.start(g), m.end(g), v if len(m.group(g)) == len(v) or len(m.group(g)) == 1 else v.zfill(len(m.group(g))))
            log.append(nfc(p.text).strip()); continue
        m = re.search(r'TUẦN\s+(\d+)\s*[,(:\-–]?\s*(?:TỪ\s+NGÀY|TỪ)?\s*(\d{1,2}/\d{1,2}/\d{4})\s*(?:→|->|–|-|ĐẾN(?:\s+NGÀY)?)\s*(\d{1,2}/\d{1,2}/\d{4})', t, re.I)
        if m:
            for g, v in ((3, rng[1]), (2, rng[0]), (1, str(w))):
                if m.group(g) != v: replace_span(p, m.start(g), m.end(g), v)
    return log

# --- tách bài học ---
SEP_RE = re.compile(r'^(-{3,}|[.…]{5,}|thứ\s+\S+\s+ngày|buổi\s+(sáng|chiều)|kế hoạch bài dạy|tuần\s+\d+$|gv chuyên)', re.I)
def split_lessons(S):
    b = list(S.element.body); T = [etext(e).strip() for e in b]
    isI = [e.tag == qn('w:p') and re.match(r'^I\.', T[i]) for i, e in enumerate(b)]
    isIV = [e.tag == qn('w:p') and re.match(r'^IV\.', T[i]) for i, e in enumerate(b)]
    lessons = []; last_end = 0; prev_i = -1
    for i in range(len(b)):
        if not isI[i]: continue
        nxt = next((k for k in range(i + 1, len(b)) if isI[k]), len(b))
        iv = next((k for k in range(i, nxt) if isIV[k]), None)
        end = iv if iv is not None else nxt
        _floor = max(prev_i + 1, i - 6)
        if 0 < last_end < i: _floor = max(_floor, last_end)
        gap = list(range(_floor, i))
        start = i; subject = None
        for k in reversed(gap):
            t = T[k]
            if not t: continue
            if re.match(r'^(II|III|IV)\.', t): break
            if SEP_RE.match(t) or subj_in_text(t):
                subject = subject or subj_in_text(t); 
                if subject is None:
                    for kk in reversed([x for x in gap if x < k]):
                        if T[kk] and SEP_RE.match(T[kk]): break
                        s2 = subj_in_text(T[kk])
                        if s2: subject = s2; break
                break
            start = k
        if subject is None:
            for k in reversed(gap):
                if T[k] and subj_in_text(T[k]): subject = subj_in_text(T[k]); break
                if T[k] and re.match(r'^-{3,}', T[k]): break
        title = ' '.join(T[k] for k in range(start, i) if T[k])
        # bỏ 'Thời gian thực hiện' khỏi tiêu đề so khớp
        title = re.sub(r'Thời gian thực hiện.*', '', title, flags=re.I).strip()
        # kiểu đầu bài 'Môn: Tiếng Việt  Bài: ...' — lấy môn từ đó, bỏ phần 'Môn: ...' khỏi tiêu đề
        mm = re.match(r'[Mm]ôn\s*(?:học)?\s*[:\-]\s*(.+?)\s+(?=Bài\s*\d*\s*[:.])', title)
        if mm:
            if subject is None: subject = subj_in_text(mm.group(1))
            title = title[mm.end():].strip()
        lessons.append(dict(idx=len(lessons), start=start, end=end, title=title, subject=subject))
        last_end = end + 1 if iv is not None else end
        prev_i = i
        while last_end < len(b) and T[last_end] and re.match(r'^[.…\s]+$', T[last_end]): last_end += 1
    return b, lessons

# --- so khớp ---
STOP = set('bài của các và ở về em đọc nói nghe viết lt luyện tập : ; . , ! ? - – ( ) cđ shtcđ tiết'.split())
def toks(s):
    s = nfc(s).lower()
    tags = set('#t' + x for x in re.findall(r'(?:tiết|\bt)\s*(\d+)', s))
    for m in re.finditer(r'tiết\s*(\d+)\s*\+\s*(\d+)', s): tags |= {'#t' + m.group(1), '#t' + m.group(2)}
    s = re.sub(r'\(\s*(?:tiết|t)\s*\d+(?:\s*\+\s*\d+)?\s*\)|tiết\s*\d+(?:\s*\+\s*\d+)?|\(\s*\d+\s*tiết\s*\)', ' ', s)
    w = [x for x in re.findall(r'[\wđ]+', s) if x not in STOP]
    return set(w), tags

def score(ten_list, lesson_title, extra=''):
    a, ta = toks(' '.join(ten_list) + ' ' + extra); b, tb = toks(lesson_title)
    if not a or not b: return 0
    inter = len(a & b)
    s = 2 * inter / (len(a) + len(b))
    if inter >= 3 and inter / min(len(a), len(b)) >= 0.85: s = max(s, 0.85)  # một bên chứa gần trọn bên kia -> cùng một bài
    if ta and tb and not (ta & tb) and len(ten_list) == 1: s *= 0.5
    return s

def bai_expansion(M, mon_key, ten):
    m = re.search(r'bài\s*(\d+)', nfc(ten).lower())
    if not m: return ''
    n = m.group(1)
    rs = getattr(M, '_bb_rows', None)
    if rs is None:  # đọc báo bài 1 lần cho mỗi tài liệu (trước đây đọc lại cả năm cho từng TIẾT -> rất chậm, tốn RAM trên Render)
        rs = [(subj(r['mon']), nfc(r['ten']).lower(), r['ten']) for w, (t, _) in week_tables(M).items() for r in rows_of(t)]
        M._bb_rows = rs
    return ' '.join(ten for k, low, ten in rs if k == mon_key and re.search(r'bài\s*' + n + r'\b', low))

def tiet_headings(G):
    day = None; res = []
    for e in G.element.body:
        if e.tag != qn('w:p'): continue
        t = re.sub(r'\s+', ' ', etext(e)).strip()
        m = DAY_RE.match(t)
        if m and re.search(r'ngày', t, re.I) and len(t) < 90: day = day_idx(m.group(1)) + 2; continue
        m = re.match(r'TIẾT\s*(\d+(?:\s*[+,&\-–]\s*\d+)*)\s*[:.\-–]?\s*(\S.*)$', t, re.I)
        if m and day and len(t) < 90:
            nums = re.findall(r'\d+', m.group(1))
            mon = re.sub(r'^(?:môn\s*(?:học)?\s*[:\-]?\s*)', '', m.group(2).strip(), flags=re.I).strip(' ()[]:-–')
            mon = re.sub(r'\s*[\(\[].*$', '', mon).strip() or mon
            res.append(dict(el=e, thu=str(day), label=t, periods=len(nums), mon=mon))
    return res

THRESH = 0.34
_SECS_CACHE = {}
def load_secs(phu_paths, exclude=None):
    """exclude: {path: set(chỉ số bài)} – bài đã bị bản mới hơn trong bản nháp thay thế.
    Tệp phụ là bất biến sau khi tải lên → cache kết quả tách bài theo (đường dẫn, mtime, cỡ)."""
    secs = []
    for pi, pth in enumerate(phu_paths):
        try: st = os.stat(pth); k = (pth, st.st_mtime_ns, st.st_size)
        except OSError: k = None
        hit = _SECS_CACHE.get(k) if k else None
        if hit is not None:
            S, b, ls0 = hit; ls = [copy.deepcopy(l) for l in ls0]
        else:
            try: S = docx.Document(pth); b, ls = split_lessons(S)
            except Exception as e:  # tệp hỏng / không phải .docx thật: bỏ qua, các tệp khác vẫn ghép
                logging.getLogger('app').warning('load_secs: không đọc được %s: %r', os.path.basename(pth), e); continue
            if k:
                if len(_SECS_CACHE) > 64: _SECS_CACHE.clear()
                _SECS_CACHE[k] = (S, b, [copy.deepcopy(l) for l in ls])
        ex = (exclude or {}).get(pth, ())
        for l in ls:
            if l['idx'] in ex: continue
            l['doc'] = pi; l['S'] = S; l['body'] = b; l['file'] = pth; secs.append(l)
    return secs

def _norm_title(t):
    """Chuẩn hoá mạnh tên bài để so khớp NGUYÊN VĂN: bỏ 'Bài N:', '(tiết n)', 'Tn' ở cuối, dấu câu."""
    t = nfc(t).lower()
    t = re.sub(r'^\s*bài\s*\d*\s*[:.]\s*', '', t)
    t = re.sub(r'\(\s*(?:tiết|t)\s*\d+(?:\s*\+\s*\d+)?\s*\)', ' ', t)
    t = re.sub(r'\bt\d+\s*$', '', t)
    t = re.sub(r'[^\wđ ]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def _bai_nums(t):
    """Các số bài ('Bài 13' -> {'13'}) trong một tên bài."""
    return set(re.findall(r'bài\s*(\d+)', nfc(t).lower()))

def title_evidence(h, title):
    """Bằng chứng một bài phụ khớp với cột TÊN BÀI của một TIẾT (khung sườn dựng từ báo bài):
    (điểm tựa tên bài như assign vẫn chấm, cờ trùng số bài). Luật đặt bài (anh chốt):
    chỉ đặt khi điểm >= THRESH hoặc trùng số bài — lệch hẳn thì để trống, không ép đặt."""
    s = max(score(h['ten'], title), score(h['ten'], title, h.get('extra') or '') * 0.9)
    if s < 0.95 and _norm_title(title) and _norm_title(title) in [_norm_title(t) for t in h['ten']]: s = 0.95
    nums = set()
    for t in h['ten']: nums |= _bai_nums(t)
    return round(s, 2), bool(nums & _bai_nums(title))

def assign(G, M, secs, manual=None):
    table_rows = rows_of(first_table(G))
    heads = tiet_headings(G)
    # gán dòng bảng tuần cho từng TIẾT (theo thứ + môn, theo thứ tự)
    used = set()
    for h in heads:
        k = subj(h['mon']); kfam = _fam_of(h['mon']); h['rows'] = []
        for i, r in enumerate(table_rows):
            if len(h['rows']) >= h['periods']: break
            if i in used or r['thu'] != h['thu'] or _fam_of(r['mon']) != kfam: continue
            used.add(i); h['rows'].append(r)
        h['ten'] = [r['ten'].strip() for r in h['rows'] if r['ten'].strip()]
        h['extra'] = ' '.join(bai_expansion(M, k, t) for t in h['ten'])
        h['lessons'] = []
    taken = set()
    # Lượt 1 – TỪNG MÔN, THEO THỨ TỰ TIẾT TRONG KHUNG SƯỜN: mỗi TIẾT nhận bài phụ cùng họ môn
    # có BẰNG CHỨNG tên bài cao nhất (title_evidence: điểm >= THRESH hoặc trùng số bài; hoà điểm
    # thì bài xuất hiện trước trong phụ thắng). Không còn ép đặt theo thứ tự xuất hiện bất kể tên
    # bài (anh chốt lại: chép giáo án phải nhìn vào cột tên bài của khung sườn) — TIẾT không có
    # bài khớp thì để trống, bài lệch hẳn nằm ở danh sách chưa dùng cho GV gán tay.
    fam_heads = {}
    for h_i, h in enumerate(heads):
        if h['ten']: fam_heads.setdefault(_fam_of(h['mon']), []).append(h_i)
    fam_less = {}
    for l_i, l in enumerate(secs):
        if l['subject'] and l['title'].strip(): fam_less.setdefault(_fam_of(l['subject']), []).append(l_i)
    for fam, hs in fam_heads.items():
        ls = fam_less.get(fam) or []
        for h_i in hs:
            h = heads[h_i]; cap = max(1, len(h['ten']))
            while len(h['lessons']) < cap:
                best = None
                for l_i in ls:
                    if l_i in taken: continue
                    s, num = title_evidence(h, secs[l_i]['title'])
                    if s < THRESH and not num: continue
                    if best is None or s > best[0]: best = (s, num, l_i)
                if best is None: break
                s, num, l_i = best; taken.add(l_i)
                h['lessons'].append(dict(l=secs[l_i], score=s, li=l_i, hi=h_i, seq=True, num=num))
    # Lượt 2 – khớp theo ĐIỂM TÊN BÀI cho phần còn lại (bài không rõ môn, TIẾT còn trống, bài còn thừa)
    cand = []
    for h_i, h in enumerate(heads):
        if not h['ten']: continue
        hf = _fam_of(h['mon'])
        for l_i, l in enumerate(secs):
            if l_i in taken: continue
            if l['subject'] and _fam_of(l['subject']) != hf: continue
            s, num = title_evidence(h, l['title'])
            if s >= THRESH or num: cand.append((max(s, 0.5) if num else s, -h_i, -l_i, h_i, l_i))
    cand.sort(reverse=True)
    for s, _, _, h_i, l_i in cand:
        h = heads[h_i]
        if l_i in taken or len(h['lessons']) >= max(1, len(h['ten'])): continue
        # bài gộp nhiều tiết chiếm hết chỗ
        if h['lessons'] and any(len(toks(x['title'])[1]) > 1 or re.search(r'\d\s*tiết', x['title'].lower()) for x in [y['l'] for y in h['lessons']]): continue
        s2, num2 = title_evidence(h, secs[l_i]['title'])
        taken.add(l_i); h['lessons'].append(dict(l=secs[l_i], score=s2, li=l_i, hi=h_i, num=num2))
    # ghép theo THỨ TỰ cho TIẾT còn trống: cùng họ môn, số bài phụ còn thừa == số TIẾT trống của họ
    # (hoặc 1-1) thì gán theo đúng thứ tự xuất hiện — đánh dấu 'ordered' để báo cáo nhắc kiểm tra lại.
    fam_empty = {}
    for h_i, h in enumerate(heads):
        if not h['lessons'] and h['ten']: fam_empty.setdefault(_fam_of(h['mon']), []).append(h_i)
    fam_free = {}
    for l_i, l in enumerate(secs):
        if l_i not in taken and l['title'].strip() and l['subject']:
            fam_free.setdefault(_fam_of(l['subject']), []).append(l_i)
    for fam, hs in fam_empty.items():
        ls = fam_free.get(fam) or []
        pairs_z = list(zip(hs, ls))
        ok_all = bool(pairs_z)
        for a, b in pairs_z:
            s_ab, num_ab = title_evidence(heads[a], secs[b]['title'])
            if s_ab < THRESH and not num_ab: ok_all = False; break
        if ls and (len(ls) == len(hs) or (len(ls) == 1 and len(hs) == 1)) and ok_all:
            for h_i, l_i in pairs_z:
                if len(heads[h_i]['lessons']) >= max(1, len(heads[h_i]['ten'])): continue
                s, num = title_evidence(heads[h_i], secs[l_i]['title'])
                taken.add(l_i); heads[h_i]['lessons'].append(dict(l=secs[l_i], score=s, li=l_i, hi=h_i, ordered=True, num=num))
    # gán tay: người dùng chỉ định bài phụ -> TIẾT (bỏ qua ngưỡng điểm, vẫn tôn trọng bài đã dùng)
    for pair in (manual or []):
        try: si, hi = int(pair[0]), int(pair[1])
        except (TypeError, ValueError, IndexError): continue
        if si < 0 or si in taken or not (0 <= hi < len(heads)) or si >= len(secs): continue
        taken.add(si); heads[hi]['lessons'].append(dict(l=secs[si], score=1.0, li=si, hi=hi, manual=True))
    for h in heads:
        for x in h['lessons']:
            x['alt'] = round(max([c[0] for c in cand if c[4] == x['li'] and c[3] != x['hi'] and c[3] not in
                                 [y['hi'] for hh in heads for y in hh['lessons'] if y['li'] == x['li']]] or [0]), 2)
    return heads, table_rows, taken

RED_LOW = 0.4
def uncertain(x):
    """Lý do ghép bài phụ -> TIẾT không chắc chắn (hoặc None)."""
    if x.get('ai'):
        if x.get('num'): return None
        if x.get('score', 1) < RED_LOW: return f"AI ghép nhưng tên bài khớp thấp ({x.get('score')}) — kiểm tra lại"
        return None   # AI ghép đã qua cổng bằng chứng tên bài: báo ở dòng riêng, không tô đỏ
    if x.get('ordered'): return 'ghép theo thứ tự (tên bài chưa khớp) — kiểm tra lại'
    if x.get('num'): return None   # trùng số bài với cột tên bài của TIẾT: bằng chứng đủ
    if x['score'] < RED_LOW: return f"độ khớp thấp ({x['score']})"
    if x['score'] < 0.95 and x.get('alt', 0) >= x['score'] - 0.02: return f"khớp ngang nhiều TIẾT ({x['score']} / {x['alt']})"
    return None
def paint_red(p):
    for r in p.iter(qn('w:r')):
        rp = r.find(qn('w:rPr'))
        if rp is None: rp = etree.SubElement(r, qn('w:rPr')); r.remove(rp); r.insert(0, rp)
        for c in rp.findall(qn('w:color')): rp.remove(c)
        red(rp)

AI_NOTE = {'v': None}   # mã chẩn đoán aiassign lần chạy gần nhất (build_giaoan đưa vào báo cáo)

def step3_insert(G, M, w, phu_paths, prefix='P2', built=None, exclude=None, manual=None):
    secs = load_secs(phu_paths, exclude)
    heads, table_rows, taken = assign(G, M, secs, manual=manual)
    AI_NOTE['v'] = 'not_run'
    try:
        import aiassign
        aiassign.arbitrate(heads, secs, taken)   # AI gỡ phần tồn đọng; lỗi/không khóa -> giữ nguyên quy tắc
        AI_NOTE['v'] = aiassign.LAST.get('note') or 'unknown'
    except Exception as e:
        AI_NOTE['v'] = 'call_error:' + type(e).__name__
    ctx = InsertCtx(G, prefix)
    report = []
    for h in heads:
        h['lessons'].sort(key=lambda x: (x['l']['doc'], x['l']['start']))
        anchor = h['el']
        for x in h['lessons']:
            anchor = ctx.insert(anchor, x['l'])
        why = [w_ for w_ in (uncertain(x) for x in h['lessons']) if w_]
        if why: paint_red(h['el'])
        report.append(dict(thu=h['thu'], tiet=h['label'], ten=h['ten'], red='; '.join(why),
                           ai=[x['l']['title'] for x in h['lessons'] if x.get('ai')],
                           inserted=[f"{x['l']['title']} ({'gán tay' if x.get('manual') else ('AI ghép' if x.get('ai') else ('theo thứ tự' if x.get('ordered') else 'độ khớp ' + str(x['score'])))})" for x in h['lessons']]))
    unused = [dict(i=i, label=f"[{l['subject'] or '?'}] {l['title']}") for i, l in enumerate(secs) if i not in taken]
    ctx.fit_all()
    if built is not None: built.update(heads=heads, rows=table_rows)
    return report, unused

class InsertCtx:
    def __init__(s, G, prefix):
        s.G = G; s.pf = prefix; s.styles_done = {}; s.num_done = {}; s.inserted = []
        sec = G.sections[-1]
        s.AV = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
        s.st = G.styles.element
        s.did = max([int(x.get('id')) for x in G.element.iter(qn('wp:docPr'))] + [0]) + 100000
        s.base_done = set()
    def base_styles(s, S, key):
        if key in s.base_done: return
        s.base_done.add(key)
        dd = S.styles.element.find(qn('w:docDefaults'))
        rp = dd.find('.//' + qn('w:rPrDefault') + '/' + qn('w:rPr')) if dd is not None else None
        rp = copy.deepcopy(rp) if rp is not None else parse_xml(f'<w:rPr {WNS}/>')
        nrm = next((x for x in S.styles.element.findall(qn('w:style')) if x.get(qn('w:type')) == 'paragraph' and x.get(qn('w:default')) == '1'), None)
        if nrm is not None and nrm.find(qn('w:rPr')) is not None:
            for c in nrm.find(qn('w:rPr')):
                o = rp.find(c.tag)
                if o is not None: rp.remove(o)
                rp.append(copy.deepcopy(c))
        rf = rp.find(qn('w:rFonts'))
        if rf is not None:
            for a in list(rf.attrib):
                if 'Theme' in a or 'theme' in a: del rf.attrib[a]
            asc = rf.get(qn('w:ascii'))
            if asc:
                for a in ('hAnsi', 'cs', 'eastAsia'):
                    if not rf.get(qn('w:' + a)): rf.set(qn('w:' + a), asc)
        if rp.find(qn('w:color')) is None: rp.append(parse_xml(f'<w:color {WNS} w:val="auto"/>'))
        if rp.find(qn('w:sz')) is None: rp.append(parse_xml(f'<w:sz {WNS} w:val="22"/>'))
        pp = dd.find('.//' + qn('w:pPrDefault') + '/' + qn('w:pPr')) if dd is not None else None
        sp = pp.find(qn('w:spacing')) if pp is not None else None
        spx = etree.tostring(sp).decode() if sp is not None else f'<w:spacing {WNS} w:after="0" w:line="240" w:lineRule="auto"/>'
        sp_tbl = f'<w:spacing {WNS} w:after="0" w:line="240" w:lineRule="auto"/>'
        normal_id = next((x.get(qn('w:styleId')) for x in s.st.findall(qn('w:style')) if x.get(qn('w:type')) == 'paragraph' and x.get(qn('w:default')) == '1'), None)
        for sid, spc in ((key + 'Normal', spx), (key + 'NormalTbl', sp_tbl)):
            e = parse_xml(f'<w:style {WNS} w:type="paragraph" w:customStyle="1" w:styleId="{sid}"><w:name w:val="{sid}"/>'
                          + (f'<w:basedOn w:val="{normal_id}"/>' if normal_id else '') + '<w:pPr/></w:style>')
            ppr = e.find(qn('w:pPr')); ppr.append(parse_xml(spc)); ppr.append(parse_xml(f'<w:ind {WNS} w:left="0" w:firstLine="0"/>')); ppr.append(parse_xml(f'<w:jc {WNS} w:val="left"/>'))
            e.append(copy.deepcopy(rp)); s.st.append(e)
    def key_for(s, l):
        return s.pf if l['doc'] == 0 else f"{s.pf}d{l['doc']}"
    def copy_style(s, S, key, sid, seen=None):
        full = f'{key}_{sid}'
        if (key, sid) in s.styles_done: return full
        sty = {x.get(qn('w:styleId')): x for x in S.styles.element.findall(qn('w:style'))}
        if sid not in sty: return None
        s.styles_done[(key, sid)] = full
        e = copy.deepcopy(sty[sid]); e.set(qn('w:styleId'), full)
        nm = e.find(qn('w:name'))
        if nm is not None: nm.set(qn('w:val'), f'{key} {sid}')
        if e.get(qn('w:default')): del e.attrib[qn('w:default')]
        for t in ('w:link', 'w:next', 'w:rsid'):
            x = e.find(qn(t))
            if x is not None: e.remove(x)
        bo = e.find(qn('w:basedOn'))
        if bo is not None:
            v = bo.get(qn('w:val')); bt = sty.get(v)
            if bt is not None and bt.get(qn('w:default')) == '1' and bt.get(qn('w:type')) == 'paragraph': bo.set(qn('w:val'), key + 'Normal')
            elif bt is not None and bt.get(qn('w:default')) == '1' and bt.get(qn('w:type')) == 'table':
                tn = next((x.get(qn('w:styleId')) for x in s.st.findall(qn('w:style')) if x.get(qn('w:type')) == 'table' and x.get(qn('w:default')) == '1'), None)
                if tn: bo.set(qn('w:val'), tn)
                else: e.remove(bo)
            else:
                r = s.copy_style(S, key, v)
                if r: bo.set(qn('w:val'), r)
                else: e.remove(bo)
        elif e.get(qn('w:type')) == 'paragraph':
            e.insert(1, parse_xml(f'<w:basedOn {WNS} w:val="{key}Normal"/>'))
        clampind(e.find(qn('w:pPr')))
        for x in e.iter(qn('w:numId')): x.set(qn('w:val'), s.map_num(S, key, x.get(qn('w:val'))))
        s.st.append(e); return full
    def map_num(s, S, key, old):
        if old == '0': return '0'
        if (key, old) in s.num_done: return s.num_done[(key, old)]
        tnp = s.G.part.numbering_part.element; snp = S.part.numbering_part.element
        nums = {n.get(qn('w:numId')): n for n in snp.findall(qn('w:num'))}
        anum = {a.get(qn('w:abstractNumId')): a for a in snp.findall(qn('w:abstractNum'))}
        if old not in nums: return '0'
        na = max([int(a.get(qn('w:abstractNumId'))) for a in tnp.findall(qn('w:abstractNum'))] + [0]) + 1
        nn = max([int(a.get(qn('w:numId'))) for a in tnp.findall(qn('w:num'))] + [0]) + 1
        n = nums[old]; a = copy.deepcopy(anum[n.find(qn('w:abstractNumId')).get(qn('w:val'))])
        a.set(qn('w:abstractNumId'), str(na))
        for x in a.findall(qn('w:numStyleLink')) + a.findall(qn('w:styleLink')): a.remove(x)
        for lvl in a.findall(qn('w:lvl')):
            ps = lvl.find(qn('w:pStyle'))
            if ps is not None: lvl.remove(ps)
            clampind(lvl.find(qn('w:pPr')))
        fn = tnp.findall(qn('w:num'))
        fn[0].addprevious(a) if fn else tnp.append(a)
        m = copy.deepcopy(n); m.set(qn('w:numId'), str(nn)); m.find(qn('w:abstractNumId')).set(qn('w:val'), str(na)); tnp.append(m)
        s.num_done[(key, old)] = str(nn); return str(nn)
    def insert(s, anchor, l):
        S = l['S']; key = s.key_for(l); s.base_styles(S, key); n0 = len(s.inserted)
        for e in l['body'][l['start']:l['end']]:
            if e.tag == qn('w:sectPr'): continue
            c = copy.deepcopy(e)
            for x in list(c.iter(qn('w:sectPr'))): x.getparent().remove(x)
            for br in list(c.iter(qn('w:br'))):
                if br.get(qn('w:type')) == 'page': br.getparent().remove(br)
            for x in list(c.iter(qn('w:lastRenderedPageBreak'))) + list(c.iter(qn('w:bookmarkStart'))) + list(c.iter(qn('w:bookmarkEnd'))):
                x.getparent().remove(x)
            for x in c.iter(qn('w:pageBreakBefore')): x.set(qn('w:val'), '0')
            for x in list(c.iter(qn('w:footnoteReference'))) + list(c.iter(qn('w:endnoteReference'))): x.getparent().remove(x)
            for p in ([c] if c.tag == qn('w:p') else []) + list(c.iter(qn('w:p'))):
                inT = any(a.tag == qn('w:tbl') for a in p.iterancestors()) or c.tag == qn('w:tbl')
                pp = p.find(qn('w:pPr'))
                if pp is None: pp = parse_xml(f'<w:pPr {WNS}/>'); p.insert(0, pp)
                if pp.find(qn('w:pStyle')) is None:
                    pp.insert(0, parse_xml(f'<w:pStyle {WNS} w:val="{key}NormalTbl"/>' if inT else f'<w:pStyle {WNS} w:val="{key}Normal"/>'))
                    pp.find(qn('w:pStyle')).set(qn('w:val'), key + ('NormalTbl' if inT else 'Normal'))
                clampind(pp)
            for t in ('w:pStyle', 'w:rStyle', 'w:tblStyle'):
                for x in c.iter(qn(t)):
                    v = x.get(qn('w:val'))
                    if v in (key + 'Normal', key + 'NormalTbl'): continue
                    r = s.copy_style(S, key, v)
                    if r: x.set(qn('w:val'), r)
                    else: x.getparent().remove(x)
            for x in c.iter(qn('w:numId')): x.set(qn('w:val'), s.map_num(S, key, x.get(qn('w:val'))))
            for bl in c.iter(qn('a:blip')):
                rid = bl.get(EMB)
                if rid and rid in S.part.related_parts:
                    nr, _ = s.G.part.get_or_add_image(io.BytesIO(S.part.related_parts[rid].blob)); bl.set(EMB, nr)
            for x in c.iter():
                for a in list(x.attrib):
                    if a.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}') and x.tag != qn('a:blip'):
                        rid = x.get(a)
                        rel = S.part.rels.get(rid)
                        if rel is not None and rel.is_external: x.set(a, s.G.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
                        elif rel is not None and 'image' in rel.reltype: x.set(a, s.G.part.get_or_add_image(io.BytesIO(rel.target_part.blob))[0])
            for x in c.iter(qn('wp:docPr')): x.set('id', str(s.did)); s.did += 1
            anchor.addnext(c); anchor = c; s.inserted.append(c)
        l['els'] = s.inserted[n0:]
        return anchor
    def fit_all(s):
        for c in s.inserted:
            tbls = ([c] if c.tag == qn('w:tbl') else []) + list(c.iter(qn('w:tbl')))
            for t in tbls:
                if not any(a.tag == qn('w:tbl') for a in t.iterancestors()): fit_table(t, s.AV)
            if c.tag == qn('w:p'): scale_images(c, s.AV - 20)

def clampind(pp):
    if pp is None: return
    i = pp.find(qn('w:ind'))
    if i is None: return
    for a in ('leftChars', 'rightChars', 'hangingChars', 'firstLineChars', 'startChars', 'endChars'):
        if i.get(qn('w:' + a)) is not None: del i.attrib[qn('w:' + a)]
    for a in ('left', 'right', 'start', 'end'):
        if i.get(qn('w:' + a)) and int(float(i.get(qn('w:' + a)))) < 0: i.set(qn('w:' + a), '0')
    L = int(float(i.get(qn('w:left')) or i.get(qn('w:start')) or 0))
    if i.get(qn('w:hanging')) and int(float(i.get(qn('w:hanging')))) > L: i.set(qn('w:hanging'), str(L))
    if i.get(qn('w:firstLine')) and int(float(i.get(qn('w:firstLine')))) < -L: i.set(qn('w:firstLine'), str(-L))

def scale_images(el, max_tw):
    mx = max_tw * 635
    for ex in el.iter(qn('wp:extent')):
        cx, cy = int(ex.get('cx')), int(ex.get('cy'))
        if cx > mx:
            f = mx / cx; ncx, ncy = int(mx), int(cy * f)
            ex.set('cx', str(ncx)); ex.set('cy', str(ncy))
            for a in ex.getparent().iter(qn('a:ext')):
                if a.getparent().tag == qn('a:xfrm') and a.getparent().getparent().tag == qn('pic:spPr'):
                    a.set('cx', str(ncx)); a.set('cy', str(ncy))

def fit_table(t, AV):
    pr = t.find(qn('w:tblPr'))
    if pr is None: pr = parse_xml(f'<w:tblPr {WNS}/>'); t.insert(0, pr)
    for tag in ('w:tblW', 'w:tblInd', 'w:tblpPr'):
        x = pr.find(qn(tag))
        if x is not None: pr.remove(x)
    pos = 1 if pr.find(qn('w:tblStyle')) is not None else 0
    pr.insert(pos, parse_xml(f'<w:tblW {WNS} w:w="{AV}" w:type="dxa"/>'))
    pr.insert(pos + 1, parse_xml(f'<w:tblInd {WNS} w:w="108" w:type="dxa"/>'))
    jc = pr.find(qn('w:jc'))
    if jc is not None: pr.remove(jc)
    g = t.find(qn('w:tblGrid'))
    old = [int(float(x.get(qn('w:w')) or 0)) for x in g] if g is not None else []
    tot = sum(old)
    if not tot: return
    new = [round(w * AV / tot) for w in old]; new[-1] = AV - sum(new[:-1])
    for x, w in zip(g, new): x.set(qn('w:w'), str(w))
    for tr in t.findall(qn('w:tr')):
        ci = 0
        trp = tr.find(qn('w:trPr'))
        if trp is not None:
            gb = trp.find(qn('w:gridBefore'))
            if gb is not None: ci = int(gb.get(qn('w:val')))
            for x in trp.findall(qn('w:wBefore')) + trp.findall(qn('w:wAfter')): trp.remove(x)
        for tc in tr.findall(qn('w:tc')):
            tp = tc.find(qn('w:tcPr'))
            if tp is None: tp = parse_xml(f'<w:tcPr {WNS}/>'); tc.insert(0, tp)
            sp = tp.find(qn('w:gridSpan')); n = int(sp.get(qn('w:val'))) if sp is not None else 1
            cw = sum(new[ci:ci + n]); ci += n
            w = tp.find(qn('w:tcW'))
            if w is None: w = parse_xml(f'<w:tcW {WNS}/>'); tp.insert(0, w)
            w.set(qn('w:w'), str(cw)); w.set(qn('w:type'), 'dxa')
            mar = tp.find(qn('w:tcMar')); ml = 216
            if mar is not None:
                ml = sum(int(float(x.get(qn('w:w')) or 0)) for x in mar if x.tag in (qn('w:left'), qn('w:right'), qn('w:start'), qn('w:end')))
            scale_images(tc, cw - ml - 20)
            for p in tc.iter(qn('w:p')): clampind(p.find(qn('w:pPr')))

def _strip_highlight(el):
    """Bỏ tô sáng (w:highlight) khỏi mọi run của 1 đoạn — đó là dấu đánh dấu vị trí
    trong file mẫu của thầy cô, không phải nội dung, nên không được lọt vào bản hoàn chỉnh."""
    n = 0
    for r in el.findall(qn('w:r')):
        rp = r.find(qn('w:rPr'))
        if rp is not None:
            hl = rp.find(qn('w:highlight'))
            if hl is not None: rp.remove(hl); n += 1
    return n

_TIET_LAB_RE = re.compile(r'^(TIẾT\s*\d+(?:\s*[+,&\-–]\s*\d+)*)\s*[:.\-–]?\s*', re.I)

def _tiet_split(el):
    """Tách các run chữ của 1 tiêu đề TIẾT thành (label_runs, subj_runs).
    Nhãn 'TIẾT n' có thể tách thành nhiều run trong mẫu của thầy cô (vd 'TIẾT '+'2 + 3');
    tên môn là toàn bộ phần chữ còn lại sau nhãn. Run nào vắt qua đúng ranh giới nhãn/tên môn
    sẽ được tách làm đôi tại ranh giới. (Không có nhãn khớp -> coi như không có nhãn.)"""
    rs = _runs_text(el)
    texts = [''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in rs]
    m = _TIET_LAB_RE.match(''.join(texts))
    lab_len = len(m.group(0)) if m else 0
    label_runs, subj_runs = [], []
    acc = 0
    for r, rt in zip(rs, texts):
        r_end = acc + len(rt)
        if r_end <= lab_len:
            label_runs.append(r)
        elif acc >= lab_len:
            subj_runs.append(r)
        else:
            cut = lab_len - acc  # 0 < cut < len(rt): run vắt qua ranh giới
            import copy as _cp
            nr = _cp.deepcopy(r)
            _run_set(r, rt[:cut].rstrip())
            _run_set(nr, rt[cut:].lstrip())
            r.addnext(nr)
            label_runs.append(r)
            subj_runs.append(nr)
        acc = r_end
    return label_runs, subj_runs

def apply_heading_names(heads):
    """Thay phần tên môn ở cuối mỗi tiêu đề TIẾT bằng TÊN BÀI của báo bài cho TIẾT đó
    (đúng vị trí thầy cô đánh dấu tô sáng xanh trong file mẫu) + bỏ tô sáng đánh dấu.
    Tên bài lấy từ chính các dòng báo bài đã gióng vào TIẾT (h['ten']) — không đoán.
    TIẾT chưa có tên bài trong báo bài -> giữ nguyên tên môn. Tiết Sinh hoạt dưới cờ
    giữ nguyên: đã có dòng riêng ghi tên buổi sinh hoạt ngay dưới tiêu đề.
    Trả về số tiêu đề đã thay tên bài."""
    n = 0
    for h in heads or []:
        el = h.get('el')
        if el is None: continue
        _strip_highlight(el)
        names = []
        for t in (h.get('ten') or []):
            t = (t or '').strip()
            if t and t not in names: names.append(t)
        if not names: continue
        if skey(h.get('mon') or '') == 'shdc' or \
           any(nfc(x).lower().startswith('sinh hoạt dưới cờ') for x in names): continue
        _, subj_runs = _tiet_split(el)
        if not subj_runs: continue
        _run_set(subj_runs[0], ' + '.join(names))
        for r in subj_runs[1:]: _run_set(r, '')
        n += 1
    return n

def build_giaoan(template_path, master_path, w, phu_paths, out_path, tichhop=None, exclude=None, manual=None):
    """tichhop (tuỳ chọn): dict(plan, cats, sel, picks, mode, ai_call) – chèn tích hợp theo Khung KHDH vào các bài đã ghép.
    manual (tuỳ chọn): [(sec_idx, head_idx), ...] – gán tay bài phụ chưa dùng vào TIẾT do người dùng chọn."""
    G = docx.Document(template_path); M = docx.Document(master_path)
    mw = week_tables(M)
    if w not in mw: raise ValueError(f'Bản báo bài không có tuần {w}')
    n_ten = sum(1 for r in rows_of(mw[w][0]) if r['ten'].strip())
    if n_ten == 0:
        raise ValueError(f'Báo bài tuần {w} chưa điền tên bài nào. Thầy cô vào mục Báo bài điền tên bài cho tuần {w} trước, rồi mới tạo giáo án nhé!')
    log1 = step1_copy_table(G, M, w)
    _rng = week_range(M, w, G)
    log2 = step2_dates(G, w, (_rng[0].strftime('%d/%m/%Y'), _rng[1].strftime('%d/%m/%Y')) if _rng else mw[w][1])
    report, unused = ([], [])
    built = {}
    if phu_paths: report, unused = step3_insert(G, M, w, phu_paths, built=built, exclude=exclude, manual=manual)
    _heads = built.get('heads')
    if _heads is None:
        try: _heads = tiet_headings(G)
        except Exception: _heads = []
    n_head_names = apply_heading_names(_heads)
    th = None
    if tichhop and tichhop.get('plan'):
        import tichhop as TH
        th = TH.apply(G, M, w, tichhop['plan'], built, cats=tichhop.get('cats'), sel=tichhop.get('sel'), picks=tichhop.get('picks'),
                      mode=tichhop.get('mode') or 'verbatim', ai_call=tichhop.get('ai_call'))
    center_tiet_subjects(G)
    G.save(out_path)
    empty = [r['tiet'] for r in report if not r['inserted']]
    reds = [dict(tiet=r['tiet'], thu=r['thu'], why=r['red']) for r in report if r.get('red')]
    ai_list = [dict(tiet=r['tiet'], thu=r['thu'], ten=t) for r in report for t in (r.get('ai') or [])]
    out = dict(step1=log1, dates=log2, tiet=report, unused=unused, empty_tiet=empty, red=reds, ai=ai_list, ai_note=AI_NOTE['v'], heading_names=n_head_names)
    if th is not None: out['tichhop'] = th
    return out


# ================= BƯỚC 2: TỰ TẠO GIÁO ÁN MẪU TRỐNG TỪ BÁO BÀI =================
SPEC_FAMS = {_fam_key(x) for x in ('tiếng anh', 'giáo dục thể chất', 'tin học', 'mĩ thuật', 'âm nhạc')}  # môn chuyên: trống không tính ô đỏ
SPEC_SUBJ = {'mĩ thuật', 'mỹ thuật', 'gdtc', 'thể dục', 'tin học', 'tiếng anh', 'âm nhạc'}
DAYNAME = {'2': 'Hai', '3': 'Ba', '4': 'Tư', '5': 'Năm', '6': 'Sáu', '7': 'Bảy'}

def _runs_text(el):
    """Các run có chữ thật (không chỉ khoảng trắng) trong 1 đoạn (phần tử w:p)."""
    out = []
    for r in el.findall(qn('w:r')):
        ts = r.findall(qn('w:t'))
        if ts and any((t.text or '').strip() for t in ts): out.append(r)
    return out

def _run_set(r, text):
    ts = r.findall(qn('w:t'))
    if not ts: return
    ts[0].text = text
    ts[0].set(qn('xml:space'), 'preserve')
    for t in ts[1:]: t.text = ''

def _set_para(el, text):
    rs = _runs_text(el)
    if not rs: return
    _run_set(rs[0], text)
    for r in rs[1:]: _run_set(r, '')

def _set_head(el, label, subj):
    rs = _runs_text(el)
    if not rs: return
    _run_set(rs[0], label)
    if len(rs) > 1:
        _run_set(rs[-1], subj)
        for r in rs[1:-1]: _run_set(r, '')

def gen_blank_giaoan(master_path, week, sample_path, out_path):
    """BƯỚC 2 – Tạo giáo án MẪU TRỐNG cho 1 tuần từ báo bài, theo cấu trúc file mẫu tham chiếu:
    tiêu đề tuần (ngày tháng ĐÚNG theo báo bài) + khung báo bài (điền sẵn từ báo bài) +
    từng ngày: dòng Thứ/ngày (đã sửa đúng) + các khối TIẾT theo đúng môn & thứ tự trong báo bài
    (gộp TIẾT khi 1 bài học 2 tiết liền nhau), nội dung để trống (IV. ĐIỀU CHỈNH... / (Giáo viên chuyên)),
    cuối ngày có mục CHĂM SÓC HỌC SINH, cuối tệp giữ bảng ký tên của mẫu."""
    M = docx.Document(master_path); mw = week_tables(M)
    if week not in mw: raise ValueError(f'Báo bài không có tuần {week}')
    blk, rng = mw[week]
    rows = [r for r in rows_of(blk) if r['mon'].strip()]
    if not rows: raise ValueError(f'Báo bài tuần {week} chưa có dòng môn học nào')
    G = docx.Document(sample_path)
    wtG = week_tables(G)
    frame = None
    if wtG: frame = next(iter(wtG.values()))[0].table
    elif first_table(G) is not None: frame = first_table(G)
    if frame is None: raise ValueError('Tệp mẫu tham chiếu không có khung báo bài')
    det = detect(frame); cols = det['cols']
    paras = list(G.paragraphs)
    def find(pred):
        for p in paras:
            if pred(p.text.strip()): return p
        return None
    title_p = find(lambda t: re.search(r'TUẦN\s*\d+', t))
    day_p = find(lambda t: bool(DAY_RE.match(t)))
    head_p = find(lambda t: t.startswith('TIẾT'))
    iv1_p = find(lambda t: t.startswith('IV.'))
    iv2_p = None
    if iv1_p is not None:
        i0 = paras.index(iv1_p)
        for p in paras[i0 + 1:i0 + 3]:
            tx = p.text.strip()
            if tx and set(tx) <= {'.', '…'}: iv2_p = p; break
    spec_p = find(lambda t: t == '(Giáo viên chuyên)')
    tpt_p = find(lambda t: 'tổng phụ trách' in t)
    shdc_p = find(lambda t: t == 'Sinh hoạt dưới cờ')
    sep_p = find(lambda t: t.startswith('----------------') and len(t) < 60)
    cham_p = find(lambda t: t.startswith('CHĂM SÓC'))
    longsep_p = find(lambda t: t.startswith('--------------------------------------------'))
    TPL = {}
    for k, p in dict(day=day_p, head=head_p, iv1=iv1_p, iv2=iv2_p, spec=spec_p, tpt=tpt_p,
                     shdc=shdc_p, sep=sep_p, cham=cham_p, longsep=longsep_p).items():
        if p is not None: TPL[k] = copy.deepcopy(p._p)
    others = [t for t in G.tables if t._tbl is not frame._tbl]
    sig = others[-1] if others else None
    # --- khoảng ngày của tuần: báo bài -> tuần khác -> mẫu (dịch tuần); tiêu đề sửa TẠI CHỖ
    _rng = week_range(M, week, G)
    d1 = _rng[0] if _rng else None
    d2 = _rng[1] if _rng else None
    if title_p is not None:
        t = nfc(title_p.text)
        m = re.search(r'TUẦN\s+(\d+)\s*[,(:\-–]?\s*(?:TỪ\s+NGÀY|TỪ)?\s*(\d{1,2}/\d{1,2}/\d{4})\s*(?:→|->|–|-|ĐẾN(?:\s+NGÀY)?)\s*(\d{1,2}/\d{1,2}/\d{4})', t, re.I)
        if m and d1 and d2:
            for g, v in ((3, d2.strftime('%d/%m/%Y')), (2, d1.strftime('%d/%m/%Y')), (1, str(week))):
                if m.group(g) != v: replace_span(title_p, m.start(g), m.end(g), v)
        elif m:
            if m.group(1) != str(week): replace_span(title_p, m.start(1), m.end(1), str(week))
        else:
            _set_para(title_p._p, f'TUẦN {week}' + (f'           TỪ NGÀY {d1:%d/%m/%Y} → {d2:%d/%m/%Y}' if d1 and d2 else ''))
    # --- khung báo bài: BÊ NGUYÊN bảng báo bài của người dùng sang (không dựng lại, không sửa)
    if not transplant_frame(G, M, week):
        raise ValueError('Không đặt được bảng báo bài vào giáo án mẫu')
    tbl = first_table(G)._tbl
    # --- xoá phần thân cũ (sau khung, trước bảng ký tên)
    body = G.element.body
    children = list(body.iterchildren())
    i1 = children.index(tbl); i2 = children.index(sig._tbl) if sig is not None else len(children)
    for el in children[i1 + 1:i2]:
        if el.tag != qn('w:sectPr'): body.remove(el)
    anchor = sig._tbl if sig is not None else body.find(qn('w:sectPr'))
    def emit(el):
        if anchor is not None: anchor.addprevious(el)
        else: body.append(el)
    by_day = {}
    for r in rows: by_day.setdefault(r['thu'], []).append(r)
    nblocks = 0
    for thu in sorted(by_day, key=lambda x: int(x) if x.isdigit() else 99):
        drows = by_day[thu]
        if 'day' in TPL:
            el = copy.deepcopy(TPL['day'])
            dtext = 'Thứ ' + DAYNAME.get(thu, thu)
            if d1 and thu.isdigit():
                dt = d1 + datetime.timedelta(days=int(thu) - 2)
                dtext += f' ngày {dt.day} tháng {dt.month} năm {dt.year}'
            _set_para(el, dtext); emit(el)
        groups = []
        for r in drows:
            if (groups and groups[-1][0]['mon'].strip().lower() == r['mon'].strip().lower()
                    and r['ten'].strip() and groups[-1][0]['ten'].strip() == r['ten'].strip()):
                groups[-1].append(r)
            else: groups.append([r])
        c = 1
        for g in groups:
            r0 = g[0]; k = len(g)
            label = f'TIẾT {c}' if k == 1 else 'TIẾT ' + ' + '.join(str(c + i) for i in range(k))
            c += k; nblocks += 1
            sk = skey(r0['mon'])
            is_shdc = (sk == 'shdc' or r0['ten'].strip().lower().startswith('sinh hoạt dưới cờ'))
            disp = 'Hoạt động trải nghiệm' if is_shdc else ('Sinh hoạt lớp' if sk == 'shl' else r0['mon'])
            if 'head' in TPL:
                el = copy.deepcopy(TPL['head']); _set_head(el, label, disp); emit(el)
            if is_shdc:
                if 'shdc' in TPL:
                    el = copy.deepcopy(TPL['shdc']); _set_para(el, r0['ten'].strip() or 'Sinh hoạt dưới cờ'); emit(el)
                if 'tpt' in TPL: emit(copy.deepcopy(TPL['tpt']))
            elif nfc(disp).lower() in SPEC_SUBJ or nfc(r0['mon']).lower() in SPEC_SUBJ:
                if 'spec' in TPL: emit(copy.deepcopy(TPL['spec']))
            else:
                if 'iv1' in TPL: emit(copy.deepcopy(TPL['iv1']))
                if 'iv2' in TPL: emit(copy.deepcopy(TPL['iv2']))
            if 'sep' in TPL: emit(copy.deepcopy(TPL['sep']))
        if 'longsep' in TPL: emit(copy.deepcopy(TPL['longsep']))
        if 'cham' in TPL: emit(copy.deepcopy(TPL['cham']))
        if 'longsep' in TPL: emit(copy.deepcopy(TPL['longsep']))
    for p in G.paragraphs:
        if p.text.startswith('TIẾT'): _strip_highlight(p._p)
    center_tiet_subjects(G)
    G.save(out_path)
    return dict(week=int(week), days=len(by_day), blocks=nblocks, rows=len(rows))


# ================= THÔNG TIN GIÁO ÁN -> ĐẦU TRANG =================
def _hdr_paras(doc):
    """Mọi đoạn văn nằm trong đầu trang / chân trang của tài liệu (kể cả trong bảng ở đầu trang)."""
    seen = set(); out = []
    for sec in doc.sections:
        for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                     sec.even_page_header, sec.even_page_footer):
            for p in part.paragraphs:
                if id(p._p) not in seen: seen.add(id(p._p)); out.append(p)
            for t in part.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if id(p._p) not in seen: seen.add(id(p._p)); out.append(p)
    return out

def apply_ga_info(doc, info):
    """Ghi thông tin (Trường / Họ tên GV / Lớp / Năm học) vào ĐẦU TRANG giáo án theo cách PHẪU THUẬT:
    chỉ thay đúng GIÁ TRỊ tại chỗ (tên GV, lớp, năm học, tên trường) — giữ nguyên mọi khoảng cách,
    vị trí, định dạng gốc của file mẫu. Mục nào người dùng chưa điền thì TUYỆT ĐỐI không đụng vào.
    Trả về số giá trị đã thay."""
    info = info or {}
    gv = (info.get('gv') or '').strip(); lop = (info.get('lop') or '').strip()
    truong = (info.get('truong') or '').strip(); namhoc = (info.get('namhoc') or '').strip()
    n = 0
    for p in _hdr_paras(doc):
        t = p.text
        if ('GV:' in t) or ('GV :' in t) or (('Năm học' in t) and ('Lớp' in t)):
            if gv:
                m = re.search(r'(GV\s*:\s*)(.+?)(?=\s*[–—-]\s*Lớp|\s*Năm học|$)', t)
                if m and m.group(2).strip() != gv:
                    replace_span(p, m.start(2), m.end(2), gv); n += 1; t = p.text
            if lop:
                m = re.search(r'(Lớp\s*)([^\s,;]+)', t)
                if m and m.group(2) != lop:
                    replace_span(p, m.start(2), m.end(2), lop); n += 1; t = p.text
            if namhoc:
                m = re.search(r'(Năm học\s*)([\d][\d\s\-–]*)', t)
                if m and m.group(2).strip() != namhoc:
                    replace_span(p, m.start(2), m.start(2) + len(m.group(2).rstrip()), namhoc); n += 1
        elif truong and t.strip().lower().startswith('trường'):
            name = truong[6:].strip() if truong.lower().startswith('trường') else truong
            m = re.search(r'(Trường\s+)(.+?)\s*$', t)
            if m and m.group(2).strip() != name:
                replace_span(p, m.start(2), m.end(2), name); n += 1
    return n

def apply_ga_info_file(path, info):
    doc = docx.Document(path); n = apply_ga_info(doc, info)
    if n: doc.save(path)
    return n
