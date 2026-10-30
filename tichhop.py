"""Tích hợp NLS, AI, Link… theo Khung Kế hoạch dạy học (KHDH).

parse_plan(docx)      -> dict kế hoạch (chỉ cột 'Nội dung điều chỉnh, bổ sung' + tên bài/tuần/tiết; không lấy phần đầu văn bản)
match_week(plan,M,w)  -> ghép từng dòng kế hoạch của tuần w với dòng báo bài (tuần + môn + tên bài, nhãn tích hợp)
apply(G, ...)         -> chèn vào giáo án đã ghép (chữ đỏ, đứng; link xanh, gạch chân, bấm được)
"""
import re, os, json, copy, shutil, subprocess, tempfile, unicodedata, logging
import docx
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree
import core

log = logging.getLogger('tichhop')
WNS = core.WNS
RED, BLUE = 'FF0000', '0000FF'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

def nfc(s): return unicodedata.normalize('NFC', s or '')
def cl(s): return re.sub(r'[ \t\xa0\u200b]+', ' ', nfc(s)).strip()
BANNED = re.compile(r'ph[aáạậắâ]t\s*t[iíì]ch', re.I)   # tên trường trong tệp kế hoạch – không bao giờ chép ra
def clean_line(s):
    s = cl(s)
    return '' if BANNED.search(s) else s

# ---------------- danh mục ----------------
CATS = [  # key, nhãn hiển thị, regex nhận diện ĐẦU dòng (sau '- ', '* ')
    ('NLS', 'Năng lực số (NLS)', r'(tích hợp\s+)?(năng lực số|nls)\b'),
    ('AI', 'Trí tuệ nhân tạo (AI)', r'(tích hợp\s+)?ai\b'),
    ('CDS', 'Kĩ năng công dân số', r'(tích hợp\s+|lồng ghép\s+)?(gd\s+)?(kĩ|kỹ) năng công dân số|(lồng ghép\s+)?gd hành trang công dân số'),
    ('LINK', 'Học liệu số / Link', r'(bs\s+|bổ sung\s+)?học liệu|link\b|clip\b|video\b|https?://|www\.|youtu'),
    ('TB', 'Thiết bị, đồ dùng', r'thiết bị|đồ dùng dạy học|máy ảnh|chụp hình'),
    ('QPAN', 'Quốc phòng – an ninh', r'(giáo dục\s+|gd\s*|lồng ghép\s+)?(qpan|anqp|gdqpan|quốc phòng)'),
    ('GDDP', 'Giáo dục địa phương', r'(lồng ghép\s+)?(gdđp|giáo dục địa phương|gd địa phương)'),
    ('STEM', 'STEM', r'(bài học\s+)?stem\b'),
    ('QCN', 'Quyền con người / trẻ em', r'(giáo dục\s+|gd\s+)?quyền'),
    ('BVMT', 'Bảo vệ môi trường', r'(giáo dục\s+|gd\s+)?(bảo vệ môi trường|bvmt)'),
    ('KNS', 'Kĩ năng sống', r'(giáo dục\s+|gd\s+)?(kns|kĩ năng sống|kỹ năng sống)'),
    ('ATGT', 'An toàn giao thông', r'(giáo dục\s+|gd\s+)?(atgt|an toàn giao thông)'),
    ('DDLS', 'Đạo đức, lối sống', r'(giáo dục\s+|gd\s+)?(đạo đức|lý tưởng|lí tưởng)'),
    ('LM', 'Liên môn', r'(tích hợp\s+|gd\s+|giáo dục\s+)?liên môn'),
    ('KHAC', 'Khác (dinh dưỡng, tiết kiệm…)', r'(giáo dục\s+|gd\s+)?(dinh dưỡng|tiết kiệm|nội quy)|về tích hợp'),
]
CAT_KEYS = [c[0] for c in CATS]
CAT_NAMES = {c[0]: c[1] for c in CATS}
_CRX = [(k, re.compile(r'^' + rx, re.I)) for k, _, rx in CATS]
URL_RX = re.compile(r'(?:https?://|www\.|youtu\.be/|youtube\.com/)[^\s<>"\]\[)]+', re.I)
def strip_bullet(s): return re.sub(r'^[-–•*+]\s*', '', s).strip()
def head_cat(line):
    s = strip_bullet(line).lower()
    for k, rx in _CRX:
        if rx.search(s): return k
    return None
def is_url_only(line): return bool(URL_RX.fullmatch(strip_bullet(line).rstrip('.,;')))
TIET_RX = re.compile(r'^\(?\s*ti[eế]t\s*(\d+)\s*\)?\s*:?\s*$', re.I)

# ---------------- đọc tệp kế hoạch ----------------
SUBJECTS = [('tv', 'Tiếng Việt', r'tiếng việt'), ('toan', 'Toán', r'\btoán\b'), ('daoduc', 'Đạo đức', r'đạo đức'),
            ('hdtn', 'Hoạt động trải nghiệm', r'hoạt động trải nghiệm'), ('tnxh', 'Tự nhiên và Xã hội', r'tự nhiên và xã hội|tự nhiên xã hội')]
SUBJ_NAME = {k: n for k, n, _ in SUBJECTS}

def to_docx(path):
    """.odt/.doc -> .docx (LibreOffice). Trả về đường dẫn docx tạm (người gọi xoá)."""
    if path.lower().endswith('.docx'): return path
    import exam
    sb = exam.soffice_bin()
    if not sb: raise ValueError('Máy chủ chưa có LibreOffice để đọc tệp .odt – thầy cô lưu tệp thành .docx rồi tải lên nhé')
    td = tempfile.mkdtemp(prefix='khdh_')
    prof = os.path.join(tempfile.gettempdir(), 'lo_profile_khdh')
    subprocess.run([sb, f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'docx', '--outdir', td, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180, check=False)
    out = [os.path.join(td, f) for f in os.listdir(td) if f.endswith('.docx')]
    if not out: shutil.rmtree(td, ignore_errors=True); raise ValueError('Không đọc được tệp kế hoạch (.odt)')
    return out[0]

def _week(s):
    s = cl(s); m = re.findall(r'\d+', s)
    if not m: return None
    if '/' in s and len(m) >= 2 and not ',' in s: return dict(raw=s, weeks=[int(m[0])])
    return dict(raw=s, weeks=sorted({int(x) for x in m if 0 < int(x) <= 40}))

def _cols(rows):
    hdr = [' '.join(cl(r.cells[i].text).lower() for r in rows[:2]) for i in range(len(rows[0].cells))]
    def f(*keys, avoid=()):
        for i, h in enumerate(hdr):
            if any(k in h for k in keys) and not any(a in h for a in avoid): return i
        return None
    c = dict(week=f('tuần'), lesson=f('tên bài'), tiet=f('tiết', avoid=('tên bài',)), adj=f('nội dung điều chỉnh') if f('nội dung điều chỉnh') is not None else f('điều chỉnh', avoid=('yêu cầu',)),
             topic=f('chủ đề', avoid=('yêu cầu', 'yccđ')), note=f('ghi chú', avoid=('điều chỉnh',)))
    return c

def parse_plan(path):
    src = to_docx(path); tmp = src != path
    try: d = docx.Document(src)
    finally:
        if tmp: shutil.rmtree(os.path.dirname(src), ignore_errors=True)
    tbl_by_el = {t._tbl: t for t in d.tables}
    subj = None; out = {}
    for el in d.element.body:
        if el.tag == qn('w:p'):
            t = core.etext(el).strip().lower()
            if len(t) < 90 and re.match(r'^(\d+\.\s*)?(môn\s+)?', t):
                for k, _, rx in SUBJECTS:
                    if re.match(r'^\d+\.\s*(môn\s+)?' + rx, t): subj = k
                if re.match(r'^\d+\.\s*(môn\s+)?(âm nhạc|mĩ thuật|mỹ thuật|giáo dục thể chất|tin học|tiếng anh)', t): subj = None
            continue
        if el.tag != qn('w:tbl') or el not in tbl_by_el or not subj: continue
        T = tbl_by_el[el]; rows = list(T.rows)
        if len(rows) < 3: continue
        C = _cols(rows)
        if C['lesson'] is None or C['adj'] is None or C['week'] is None: continue
        if subj in out: continue
        res = []; wk = None; topic = None; bai = None
        for ri, r in enumerate(rows):
            cells = r.cells
            txt = [cl(c.text) for c in cells]
            if ri < 3 and ('tên bài' in txt[C['lesson']].lower() or 'chương trình' in txt[C['lesson']].lower()): continue
            vals = {x for x in txt if x}
            if len(vals) == 1 and C['topic'] is not None and txt[C['lesson']] == txt[C['topic']]:
                topic = clean_line(txt[C['topic']]); continue
            w = _week(txt[C['week']]) if txt[C['week']] else None
            if w: wk = w
            if C['topic'] is not None and txt[C['topic']]: bai = clean_line(txt[C['topic']])
            lesson = clean_line(txt[C['lesson']])
            adj_cell = cells[C['adj']]
            lines = [] if txt[C['adj']] == txt[C['lesson']] else [x for p in adj_cell.paragraphs for x in (clean_line(y) for y in re.split(r'[\n\r\v]+', p.text)) if x]
            tiet = cl(txt[C['tiet']]) if C['tiet'] is not None else ''
            if not lesson and not lines: continue
            if not lesson and res and wk == res[-1]['week']:   # dòng nối tiếp (ô gộp) -> cộng dồn vào dòng trên
                res[-1]['lines'] += lines; continue
            res.append(dict(row=ri, week=wk, topic=topic, bai=bai, lesson=lesson, tiet=tiet, lines=lines))
        out[subj] = res
    if not out: raise ValueError('Không thấy bảng kế hoạch (cột "Tên bài học" và "Nội dung điều chỉnh, bổ sung") trong tệp')
    plan = dict(version=1, subjects={})
    for k, rows in out.items():
        for i, r in enumerate(rows):
            r['pid'] = f'{k}{r["row"]}'; r['groups'] = groups(r['lines'], k)
        plan['subjects'][k] = rows
    return plan

def plan_stats(plan):
    st = {}
    for k, rows in plan['subjects'].items():
        ws = sorted({w for r in rows if r['week'] for w in r['week']['weeks']})
        st[k] = dict(name=SUBJ_NAME[k], rows=len(rows), with_items=sum(1 for r in rows if r['groups']),
                     weeks=[ws[0], ws[-1]] if ws else [])
    return st

# ---------------- tách nhóm / mục ----------------
def _label_split(line):
    s = strip_bullet(line); m = re.match(r'^(.{2,80}?)\s*:\s*(.*)$', s)
    if m and head_cat(m.group(1)) or (m and re.match(r'^(mục tiêu|gợi ý)', m.group(1), re.I)): return m.group(1).strip(), m.group(2).strip()
    return '', s

PLACE_RX = re.compile(r'(?:(?:đặt|tích hợp)\s+)?trong\s+(?:phần|hoạt động|hđ|tiết)\b[^;.]{0,80}', re.I)
def placement(text):
    """-> (chữ cái mục A–D hoặc None, cụm từ vị trí trong kế hoạch hoặc None)"""
    t = nfc(text)
    m = PLACE_RX.search(t[:200])
    ph = m.group(0).strip() if m else None
    if not ph:
        m2 = re.match(r'^\s*(?:trong|sau)\s+(?:phần|khi)\s+([^,;.]{3,40})', t, re.I)
        ph = m2.group(0).strip() if m2 and re.match(r'^\s*trong', t, re.I) else None
    if not ph: return None, None
    return letter(ph), ph

def letter(s):
    s = nfc(s).lower()
    if re.search(r'khởi động|mở đầu|ổn định', s): return 'A'
    if re.search(r'vận dụng|tiếp nối|củng cố|cam kết|dặn dò', s): return 'D'
    if re.search(r'luyện tập|thực hành|luyện đọc|sau khi đọc|luyện viết|viết vào vở', s): return 'C'
    if re.search(r'khám phá|hình thành|tìm hiểu|đọc hiểu|trả lời câu hỏi|trước khi đọc|trong khi đọc|nghe kể|quan sát|đọc văn bản|chia sẻ|nghe - viết|nghe viết', s): return 'B'
    if re.search(r'mở rộng|trải nghiệm|tổng kết', s): return 'D'
    return None

def default_letter(cat, subj):
    if cat in ('AI', 'NLS', 'CDS'): return 'C' if subj == 'toan' else 'D'
    return 'B'

def groups(lines, subj):
    """Tách ô 'Nội dung điều chỉnh' thành các nhóm: mỗi nhóm 1 loại tích hợp + mục tiêu + gợi ý + link."""
    G = []; cur = None; pend_tiet = None
    def new(cat, line):
        g = dict(cat=cat or 'KHAC', label='', items=[], tiet=pend_tiet); G.append(g); return g
    for idx, ln in enumerate(lines):
        mt = TIET_RX.match(ln) or re.match(r'^ti[eế]t\s*(\d+)\s*:\s*$', ln, re.I)
        if mt:
            rest_has_head = any(head_cat(x) and not is_url_only(x) for x in lines[idx + 1:])
            if not rest_has_head and G:
                for g in G:
                    if g['tiet'] is None: g['tiet'] = int(mt.group(1))
            else: pend_tiet = int(mt.group(1))
            continue
        low = strip_bullet(ln).lower()
        if is_url_only(ln):
            url = URL_RX.search(ln).group(0)
            if cur is not None and cur['cat'] in ('LINK', 'TB') and cur['items']:
                cur['items'][-1]['urls'].append(url)
            else:
                g = cur if cur is not None else new('LINK', ln)
                g['items'].append(dict(kind='link', label='Học liệu số', text='', urls=[url]))
                if cur is None: cur = g
            continue
        if re.match(r'^mục tiêu tích hợp', low):
            lab, body = _label_split(ln)
            if cur is None: cur = new('KHAC', ln)
            cur['items'].append(dict(kind='muctieu', label='Mục tiêu tích hợp', text=body, urls=URL_RX.findall(body)))
            continue
        if re.match(r'^gợi ý (trong khbd|hoạt động)', low):
            lab, body = _label_split(ln)
            body = re.sub(r'^hoạt động\s*:\s*', '', body, flags=re.I)
            if re.match(r'^mỗi hoạt động chỉ đặt|^chia hai nội dung', body, re.I):  # ghi chú cho GV, không phải hoạt động
                continue
            if cur is None or cur['cat'] in ('LINK', 'TB'):
                cur = new(head_cat(body) or 'KHAC', ln); cur['label'] = 'Tích hợp'
            parts = [body]
            if len(re.findall(r'đặt\s+trong', body, re.I)) >= 2: parts = [p.strip() for p in body.split(';') if p.strip()]
            for p in parts:
                L, ph = placement(p)
                cur['items'].append(dict(kind='goiy', label=lab or 'Gợi ý', text=p, urls=[], letter=L, hint=ph))
            continue
        c = head_cat(ln)
        if c and not (cur is not None and ln.lstrip().startswith(('-', '+')) and cur['items'] and c in ('LINK',) and cur['cat'] in ('LINK', 'TB')):
            if c == 'LINK' and cur is not None and cur['cat'] in ('TB', 'LINK') and cur['items'] and cur['items'][-1]['kind'] == 'link' and not cur['items'][-1]['urls']:
                lab, body = _label_split(ln); it = cur['items'][-1]
                if not it['text']: it['label'] = lab or it['label']; it['text'] = body if lab else strip_bullet(ln)
                else: it['text'] = (it['text'].rstrip(' ;:') + '; ' + strip_bullet(ln)).strip()
                it['urls'] += URL_RX.findall(ln); continue
            cur = new(c, ln); lab, body = _label_split(ln)
            cur['label'] = lab or strip_bullet(ln)[:60]
            kind = 'link' if c in ('LINK', 'TB') else 'tichhop'
            cur['items'].append(dict(kind=kind, label=lab, text=body if lab else strip_bullet(ln), urls=URL_RX.findall(ln)))
            continue
        # dòng tiếp theo của nhóm
        if cur is None:
            cur = new('KHAC', ln); cur['items'].append(dict(kind='tichhop', label='', text=strip_bullet(ln), urls=URL_RX.findall(ln))); continue
        it = cur['items'][-1] if cur['items'] else None
        if it is None or it['kind'] == 'goiy':
            cur['items'].append(dict(kind='link' if cur['cat'] in ('LINK', 'TB') else 'tichhop', label='', text=strip_bullet(ln), urls=URL_RX.findall(ln)))
        else:
            it['text'] = (it['text'].rstrip() + ('; ' if it['text'].strip() and not it['text'].rstrip().endswith(':') else ' ') + strip_bullet(ln)).strip()
            it['urls'] += URL_RX.findall(ln)
    for g in G:
        g['items'] = [it for it in g['items'] if it['text'].strip(' :;') or it['urls']]
        for it in g['items']:
            it['text'] = it['text'].strip()
            if it['kind'] == 'goiy' and not it.get('letter'): it['letter'] = default_letter(g['cat'], subj); it['auto'] = True
    return [g for g in G if g['items']]

# ---------------- ghép kế hoạch <-> báo bài ----------------
NORM = [(r'\btv\s*-\s*đmr\b|\bđmr\b', ' đọc mở rộng '), (r'nghe\s*-\s*viết', 'nghe viết'), (r'luyện viết đoạn', 'luyện tập viết đoạn'),
        (r'\bmrvt\b', 'mở rộng vốn từ'), (r'\bltvc\b', 'luyện từ và câu'), (r'\bhđmr\b', 'hoạt động mở rộng')]
STOPW = set('bài của các và ở về em : ; . , ! ? - – ( ) tiết lớp với cho trong một những là có'.split())
def ntoks(s):
    s = nfc(s).lower()
    for a, b in NORM: s = re.sub(a, b, s)
    s = re.sub(r'\(\s*ti[eế]t\s*\d+\s*\)|ti[eế]t\s*\d+', ' ', s)
    return {x for x in re.findall(r'[\wđ]+', s) if x not in STOPW and not x.isdigit()}
def sim(a, b):
    A, B = ntoks(a), ntoks(b)
    if not A or not B: return 0.0
    return 2 * len(A & B) / (len(A) + len(B))
def bai_no(s):
    m = re.search(r'\bbài\s*(\d+)', nfc(s).lower()); return int(m.group(1)) if m else None
def tiet_no(s):
    m = re.search(r'ti[eế]t\s*(\d+)', nfc(s).lower()); return int(m.group(1)) if m else None
def n_tiet(s):
    m = re.search(r'(\d+)\s*ti[eế]t', nfc(s).lower()); return int(m.group(1)) if m else 1
GENERIC = re.compile(r'^(luyện tập( chung)?|ôn tập|kiểm tra|thực hành)\.?$', re.I)

CODE_RX = re.compile(r'\b\d\.\d\.CB\d[a-z]\b|\b\d\.[A-D]\d\.\d\b')
LABEL_KW = [('NLS', r'nls|năng lực số'), ('AI', r'\bai\b'), ('QPAN', r'qpan|quốc phòng'), ('BVMT', r'môi trường|bvmt'),
            ('QCN', r'qcn|quyền'), ('LINK', r'học liệu|link'), ('DDLS', r'đạo đức|lối sống'), ('GDDP', r'gdđp|địa phương'),
            ('KNS', r'kns|kĩ năng sống'), ('ATGT', r'atgt|giao thông'), ('STEM', r'stem')]

def rowkey(k): return f'{k[0]}|{k[1]}|{k[2]}|{k[3]}'

def match_week(plan, M, w):
    """M: docx báo bài gốc (Document). Trả về dict(matches=[...], unmatched=[...], unlabeled=[...])."""
    wt = core.week_tables(M)
    if w not in wt: raise ValueError(f'Báo bài gốc không có tuần {w}')
    brows = [(rowkey(k), r) for k, r in core.keyed(core.rows_of(wt[w][0])) if r['mon'] and r['ten'].strip()]
    # thứ tự tiết của từng 'Bài N' Toán qua các tuần (để phân biệt nhiều tiết 'Luyện tập')
    toan_ord = {}
    cnt = {}
    for ww, (t, _) in sorted(wt.items()):
        if ww > w: break
        for k, r in core.keyed(core.rows_of(t)):
            if core.subj(r['mon']) != 'toan' or not r['ten'].strip(): continue
            n = bai_no(r['ten'])
            if n is None: continue
            cnt[n] = cnt.get(n, 0) + 1
            if ww == w: toan_ord[rowkey(k)] = (n, tiet_no(r['ten']) or cnt[n])
    matches = []; unmatched = []; used_b = {}
    for sk, prow in plan['subjects'].items():
        P = [p for p in prow if p['week'] and w in p['week']['weeks']]
        B = [(key, r) for key, r in brows if core.subj(r['mon']) == sk]
        # khoảng tiết của mỗi dòng kế hoạch trong 'Bài N' (Toán)
        rng = {}
        if sk == 'toan':
            acc = {}
            for p in prow:
                n = bai_no(p['bai'] or '') if p['bai'] else None
                if n is None: continue
                a = acc.get(n, 0); k = n_tiet(p['tiet']); rng[p['pid']] = (n, a + 1, a + k); acc[n] = a + k
        for p in P:
            if not p['groups']: continue
            ptxt = p['lesson'] + (' ' + (p['bai'] or '') if (GENERIC.match(p['lesson'].strip()) or sk in ('toan', 'daoduc', 'tnxh')) else '')
            cands = []
            for key, r in B:
                s = sim(r['ten'], ptxt)
                if sk == 'toan' and p['pid'] in rng and key in toan_ord:
                    n, a, b = rng[p['pid']]; bn, bt = toan_ord[key]
                    if n != bn: s = 0
                    else: s = 0.45 + 0.3 * s + (0.25 if a <= bt <= b else -0.3)
                if sk == 'hdtn' and re.search(r'hđ theo chủ đề|hoạt động theo chủ đề', ' '.join(p['lines']), re.I) and not re.search(r'shl|sinh hoạt lớp|shdc|dưới cờ', r['mon'], re.I):
                    s += 0.1
                lab = nfc(r['tich']).lower()
                if lab:
                    codes = set(CODE_RX.findall(r['tich'])); body = ' '.join(p['lines'])
                    if codes and codes & set(CODE_RX.findall(body)): s += 0.35
                    elif any(re.search(rx, lab) and any(g['cat'] == c for g in p['groups']) for c, rx in LABEL_KW): s += 0.12
                if s >= 0.3: cands.append((round(min(s, 1.0), 2), key, r))
            if not cands:
                unmatched.append(dict(pid=p['pid'], subject=SUBJ_NAME[sk], lesson=p['lesson'], week=p['week']['raw'], tiet=p['tiet'],
                                      reason='Không thấy bài này trong báo bài tuần ' + str(w), groups=p['groups']))
                continue
            order = {key: i for i, (key, _) in enumerate(brows)}
            cands.sort(key=lambda x: (-x[0], order[x[1]]))
            best = cands[0][0]
            near = [c for c in cands if c[0] >= best - 0.08]
            near.sort(key=lambda x: order[x[1]])
            matches.append(dict(pid=p['pid'], sk=sk, subject=SUBJ_NAME[sk], lesson=p['lesson'], bai=p['bai'], week=p['week']['raw'], tiet=p['tiet'],
                                multi_week=len(p['week']['weeks']) > 1, conf=best,
                                rows=[dict(key=k, thu=r['thu'], buoi=r['buoi'], mon=r['mon'], ten=r['ten'].strip(), tich=r['tich'].strip(), score=s) for s, k, r in near],
                                others=[dict(key=k, mon=r['mon'], ten=r['ten'].strip(), thu=r['thu'], score=s) for s, k, r in cands if s < best - 0.08][:3],
                                groups=p['groups']))
            for _, k, _r in near: used_b.setdefault(k, []).append(p['pid'])
    unlabeled = [dict(key=k, thu=r['thu'], buoi=r['buoi'], mon=r['mon'], ten=r['ten'].strip(), tich=r['tich'].strip())
                 for k, r in brows if r['tich'].strip() and k not in used_b and core.subj(r['mon']) in plan['subjects']]
    # đánh số mục
    for m in matches:
        n = 0
        for gi, g in enumerate(m['groups']):
            for ii, it in enumerate(g['items']):
                it['iid'] = f"{m['pid']}.{gi}.{ii}"
                it['dest'] = 'I4' if it['kind'] in ('tichhop', 'muctieu') else 'II' if it['kind'] == 'link' else it.get('letter') or 'B'
                if it['kind'] == 'link' and g['cat'] not in ('LINK', 'TB') and not it['text']: it['dest'] = 'II'
    return dict(week=w, matches=matches, unmatched=unmatched, unlabeled=unlabeled)

def target_row(m, g):
    """Dòng báo bài nhận nhóm g: theo 'Tiết k' trong kế hoạch, mặc định dòng đầu."""
    rows = m['rows']
    if g.get('tiet') and len(rows) >= g['tiet']: return rows[g['tiet'] - 1]
    return rows[0]

def preview(plan, M, w, cats=None):
    r = match_week(plan, M, w)
    cats = set(cats or CAT_KEYS)
    out = []
    for m in r['matches']:
        gs = []
        for g in m['groups']:
            t = target_row(m, g)
            gs.append(dict(cat=g['cat'], cat_name=CAT_NAMES.get(g['cat'], g['cat']), on=g['cat'] in cats, tiet=g.get('tiet'),
                           row=t['key'], row_label=f"T{t['thu']} {t['mon']}: {t['ten']}",
                           items=[dict(iid=it['iid'], kind=it['kind'], label=it['label'], text=it['text'], urls=it['urls'], dest=it['dest'],
                                       hint=it.get('hint'), auto=bool(it.get('auto'))) for it in g['items']]))
        out.append(dict(pid=m['pid'], subject=m['subject'], lesson=m['lesson'], week=m['week'], tiet=m['tiet'], conf=m['conf'],
                        multi_week=m['multi_week'], rows=m['rows'], others=m['others'], groups=gs))
    return dict(week=w, matches=out, unmatched=[dict(pid=u['pid'], subject=u['subject'], lesson=u['lesson'], week=u['week'], reason=u['reason'],
                                                      n=sum(len(g['items']) for g in u['groups'])) for u in r['unmatched']],
                unlabeled=r['unlabeled'], cats=[dict(key=k, name=CAT_NAMES[k]) for k in CAT_KEYS])

# ---------------- dựng đoạn / run ----------------
def _sty_sz(G, sid, seen=None):
    st = {x.get(qn('w:styleId')): x for x in G.styles.element.findall(qn('w:style'))}
    seen = seen or set()
    while sid and sid in st and sid not in seen:
        seen.add(sid); e = st[sid]
        z = e.find(qn('w:rPr') + '/' + qn('w:sz'))
        if z is not None: return int(z.get(qn('w:val')))
        bo = e.find(qn('w:basedOn')); sid = bo.get(qn('w:val')) if bo is not None else None
    return None
def para_sz(G, p):
    for r in p.iter(qn('w:r')):
        z = r.find(qn('w:rPr') + '/' + qn('w:sz'))
        if z is not None and (''.join(t.text or '' for t in r.iter(qn('w:t')))).strip(): return int(z.get(qn('w:val')))
    ps = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
    s = _sty_sz(G, ps.get(qn('w:val'))) if ps is not None else None
    if s: return s
    dd = G.styles.element.find(qn('w:docDefaults'))
    z = dd.find('.//' + qn('w:sz')) if dd is not None else None
    if z is not None: return int(z.get(qn('w:val')))
    nrm = next((x for x in G.styles.element.findall(qn('w:style')) if x.get(qn('w:type')) == 'paragraph' and x.get(qn('w:default')) == '1'), None)
    s = _sty_sz(G, nrm.get(qn('w:styleId'))) if nrm is not None else None
    return s or 24

def _rpr(sz, color, bold=False, u=False):
    return (f'<w:rPr {WNS}><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>'
            f'<w:b w:val="{1 if bold else 0}"/><w:bCs w:val="{1 if bold else 0}"/><w:i w:val="0"/><w:iCs w:val="0"/>'
            f'<w:color w:val="{color}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>' + ('<w:u w:val="single"/>' if u else '') + '</w:rPr>')
def run(text, sz, color=RED, bold=False, u=False):
    r = parse_xml(f'<w:r {WNS}>{_rpr(sz, color, bold, u)}</w:r>')
    t = etree.SubElement(r, qn('w:t')); t.text = text; t.set(XML_SPACE, 'preserve'); return r
def norm_url(u):
    u = u.rstrip('.,;)')
    if not re.match(r'^https?://', u, re.I): u = 'https://' + u
    return u
def hyperlink(part, url, sz):
    rid = part.relate_to(norm_url(url), RT.HYPERLINK, is_external=True)
    h = parse_xml(f'<w:hyperlink {WNS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}" w:history="1"/>')
    h.append(run(url.rstrip('.,;)'), sz, BLUE, False, True)); return h
def text_runs(part, text, sz, bold=False):
    """Chữ đỏ, các đường link trong chữ -> link xanh."""
    out = []; pos = 0
    for m in URL_RX.finditer(text):
        if m.start() > pos: out.append(run(text[pos:m.start()], sz, RED, bold))
        out.append(hyperlink(part, m.group(0), sz)); pos = m.end()
    if pos < len(text): out.append(run(text[pos:], sz, RED, bold))
    return out
def new_p(ref, children):
    p = parse_xml(f'<w:p {WNS}/>')
    if ref is not None and ref.find(qn('w:pPr')) is not None:
        pp = copy.deepcopy(ref.find(qn('w:pPr')))
        for tg in ('w:numPr', 'w:pageBreakBefore', 'w:sectPr', 'w:framePr'):
            for x in pp.findall(qn(tg)): pp.remove(x)
        rp = pp.find(qn('w:rPr'))
        if rp is not None: pp.remove(rp)
        p.append(pp)
    for c in children: p.append(c)
    return p
def ptext(e): return re.sub(r'\s+', ' ', core.etext(e)).strip()
def nkey(s): return re.sub(r'[\W_]+', '', nfc(s).lower())

# ---------------- vị trí trong 1 bài soạn ----------------
TOP_RX = re.compile(r'^(hoạt động\s*\d+\s*[:.\-–]|\d+\s*[.)]\s*(?!\d)|[A-E]\s*[.)]\s+)', re.I)
SUB_RX = re.compile(r'^\*\s*\S')
def is_top(t): return len(t) <= 100 and bool(TOP_RX.match(t)) and not re.match(r'^bài\s*\d', t, re.I) and (letter(t) is not None or re.match(r'^hoạt động\s*\d', t, re.I))
def is_sub(t): return len(t) <= 100 and bool(SUB_RX.match(t))

def lesson_parts(els):
    """-> dict(I=(i,j), II=(i,j), III=i, end=len) chỉ số trong els (cấp thân văn bản)."""
    T = [ptext(e) if e.tag == qn('w:p') else None for e in els]
    def find(rx, start=0):
        for i in range(start, len(els)):
            if T[i] is not None and re.match(rx, T[i], re.I): return i
        return None
    i1 = find(r'^I\s*\.\s'); i2 = find(r'^II\s*\.\s', (i1 or 0)); i3 = None
    for i in range((i2 or i1 or 0), len(els)):
        if T[i] is not None and re.match(r'^III\s*\.\s', T[i], re.I): i3 = i
        if i3 is not None and T[i] is not None and re.match(r'^III\s*\.\s.*(hoạt động|tiến trình)', T[i], re.I): i3 = i; break
    return dict(I=i1, II=i2, III=i3, T=T)

def tbl_rows(tbl): return tbl.findall(qn('w:tr'))
def tr_tcs(tr): return tr.findall(qn('w:tc'))
def gv_col(tbl):
    rs = tbl_rows(tbl)
    if not rs: return 0
    for i, tc in enumerate(tr_tcs(rs[0])):
        t = ptext(tc).lower()
        if re.search(r'\bgv\b|giáo viên', t): return i
    return 0

def anchors(els, parts):
    """Danh sách tiêu đề hoạt động trong phần III (theo thứ tự)."""
    out = []
    s = (parts['III'] + 1) if parts['III'] is not None else len(els)
    for i in range(s, len(els)):
        e = els[i]
        if e.tag == qn('w:p'):
            t = parts['T'][i]
            if t and (is_top(t) or is_sub(t)): out.append(dict(kind='p', i=i, el=e, text=t, top=is_top(t)))
        elif e.tag == qn('w:tbl'):
            rs = tbl_rows(e)
            for ri, tr in enumerate(rs):
                tcs = tr_tcs(tr)
                txt = [ptext(tc) for tc in tcs]
                if ri > 0 and (len(tcs) == 1 or (len(set(txt)) == 1 and txt[0])):
                    t = txt[0].split('Mục tiêu')[0].strip()
                    t0 = core.etext(tcs[0].find(qn('w:p'))).strip() if tcs[0].find(qn('w:p')) is not None else t
                    if letter(t0 or t) or is_top(t0 or t): out.append(dict(kind='row', i=i, el=e, tr=tr, ri=ri, text=t0 or t, top=True)); continue
                for ci, tc in enumerate(tcs):
                    for pi, p in enumerate(tc.findall(qn('w:p'))):
                        t = ptext(p)
                        if t and (is_top(t) or is_sub(t)):
                            out.append(dict(kind='cell', i=i, el=e, tr=tr, ri=ri, ci=ci, p=p, text=t, top=is_top(t)))
    return out

def choose(anc, letter_, hint):
    """-> (tiêu đề, ghi chú) theo: tên mục trong kế hoạch -> chữ cái -> dự phòng."""
    if hint:
        h = nkey(re.sub(r'^(?:(?:đặt|tích hợp)\s+)?trong\s+(?:phần|hoạt động|hđ)?\s*(\d+\s*[-:–]?\s*)?', '', hint, flags=re.I))
        best = None
        for a in anc:
            k = nkey(re.sub(r'^(hoạt động\s*\d+\s*[:.\-–]|\d+\s*[.)]|[A-E]\s*[.)]|\*)\s*', '', a['text'], flags=re.I))
            if len(k) >= 5 and len(h) >= 5 and (k in h or h in k):
                if best is None or (a['top'] and not best['top']): best = a
        if best: return best, ''
    tops = [a for a in anc if a['top']]
    order = {'A': 'ABCD', 'B': 'BCDA', 'C': 'CBDA', 'D': 'DCBA'}[letter_ or 'B']
    for j, L in enumerate(order):
        for a in tops:
            if letter(a['text']) == L:
                return a, ('' if j == 0 else f'không thấy mục {NAMES[letter_]}, đặt vào “{a["text"][:40]}”')
    return None, f'không thấy tiêu đề hoạt động, đặt cuối phần III'
NAMES = {'A': 'Khởi động', 'B': 'Khám phá', 'C': 'Luyện tập', 'D': 'Vận dụng'}

def section_end(anc, a):
    """Phần tử/vị trí ngay sau mục a (trước tiêu đề kế tiếp cùng nơi chứa)."""
    same = [x for x in anc if x['kind'] == a['kind'] and x['el'] is a['el'] and (a['kind'] != 'cell' or (x['ci'] == a['ci'] and x['tr'] is a['tr']))]
    k = next(i for i, x in enumerate(same) if x is a)
    for x in same[k + 1:]:
        if a['top'] and not x['top']: continue
        return x
    return None

class Applier:
    def __init__(s, G):
        s.G = G; s.part = G.part; s.placed = []; s.notes = []
    def sz_of(s, ref): return para_sz(s.G, ref) if ref is not None else 28
    def line_children(s, label, text, sz, urls=()):
        ch = []
        if label: ch.append(run(label.rstrip(':') + ': ', sz, RED, True))
        ch += text_runs(s.part, text, sz)
        extra = [u for u in urls if u not in text]
        for u in extra: ch.append(run(' ', sz)); ch.append(hyperlink(s.part, u, sz))
        return ch

    # I. 4. Tích hợp
    def add_I4(s, els, parts, lines):
        if parts['I'] is None: return 'không thấy mục I. Yêu cầu cần đạt'
        end = parts['II'] if parts['II'] is not None else (parts['III'] if parts['III'] is not None else len(els))
        rng = list(range(parts['I'] + 1, end))
        T = parts['T']
        exist = {nkey(T[i]) for i in rng if T[i]}
        lines = [l for l in lines if nkey((l[0] + ': ' if l[0] else '') + l[1]) not in exist and nkey(l[1]) not in exist]
        if not lines: return 'đã có sẵn (bỏ trùng)'
        i4 = next((i for i in rng if T[i] and re.match(r'^4\s*\.\s*tích hợp', T[i], re.I)), None)
        bullets = [els[i] for i in rng if T[i] and re.match(r'^[-–+•]', T[i])]
        ref_b = bullets[-1] if bullets else els[rng[-1]] if rng else els[parts['I']]
        sz = s.sz_of(ref_b)
        if i4 is not None:
            j = i4
            for i in range(i4 + 1, end):
                if T[i] and re.match(r'^\d\s*\.\s', T[i]): break
                j = i
            anchor = els[j]; where = 'I. 4. Tích hợp (bổ sung)'
        else:
            i3 = next((i for i in rng if T[i] and re.match(r'^3\s*\.\s', T[i])), None)
            if i3 is not None:
                j = i3
                for i in range(i3 + 1, end):
                    if T[i] and re.match(r'^\d\s*\.\s', T[i]): break
                    if T[i]: j = i
                ref_h = els[i3]
                hb = any(x.get(qn('w:val')) in (None, '1', 'true') for x in ref_h.iter(qn('w:b')))
                h = new_p(ref_h, [run('4. Tích hợp', s.sz_of(ref_h), RED, hb)])
                els[j].addnext(h); anchor = h; where = 'I. 4. Tích hợp (mới)'
            else:
                j = rng[-1] if rng else parts['I']
                while j > parts['I'] and not T[j]: j -= 1
                h = new_p(ref_b, [run('* Tích hợp:', sz, RED, True)])
                els[j].addnext(h); anchor = h; where = 'I. Yêu cầu cần đạt – “* Tích hợp”'
        for lab, txt, urls in lines:
            p = new_p(ref_b, [run('- ', sz)] + s.line_children(lab, txt, sz, urls))
            anchor.addnext(p); anchor = p
        return where

    # II. Đồ dùng – dòng GV
    def add_II(s, els, parts, lines):
        T = parts['T']
        if parts['II'] is None:
            if parts['III'] is None: return None
            ref = els[parts['III']]; anchor = els[parts['III']].getprevious(); where = 'trước mục III (không thấy mục II)'
        else:
            end = parts['III'] if parts['III'] is not None else len(els)
            rng = list(range(parts['II'] + 1, end))
            gv = [i for i in rng if T[i] and re.match(r'^[-–+•]?\s*(gv|giáo viên)\s*:', T[i], re.I)]
            if gv:
                j = gv[-1]
                for i in range(gv[-1] + 1, end):
                    if T[i] and re.match(r'^[-–+•]?\s*(hs|học sinh)\s*:', T[i], re.I): break
                    if T[i]: j = i
                anchor = els[j]; ref = els[gv[-1]]; where = 'II. Đồ dùng dạy học – GV'
            else:
                nz = [i for i in rng if T[i]]
                anchor = els[nz[-1]] if nz else els[parts['II']]; ref = anchor if nz else els[parts['II']]
                where = 'II. Đồ dùng dạy học – GV'
        sz = s.sz_of(ref) if ref is not None else 28
        for lab, txt, urls in lines:
            p = new_p(ref, [run('- GV: ', sz)] + s.line_children(lab, txt, sz, urls))
            anchor.addnext(p); anchor = p
        return where

    # III. hoạt động
    def add_act(s, els, parts, letter_, hint, gv_text, hs_text, label):
        anc = anchors(els, parts)
        a, note = choose(anc, letter_, hint)
        if a is None:
            # cuối phần III
            tbls = [e for e in els[(parts['III'] or 0):] if e.tag == qn('w:tbl')]
            if tbls:
                t = tbls[-1]; rs = tbl_rows(t); gc = gv_col(t); tr = rs[-1]; tcs = tr_tcs(tr)
                s._cell_insert(tcs, gc, None, gv_text, hs_text, label)
            elif els:
                ref = next((e for e in reversed(els) if e.tag == qn('w:p') and ptext(e)), els[-1])
                s._para_insert(els[-1], ref, gv_text, hs_text, label)
            return 'cuối phần III', note
        nxt = section_end(anc, a)
        if a['kind'] == 'p':
            # trước tiêu đề kế tiếp (cấp thân), hoặc trước bảng/hết bài
            stop = nxt['i'] if nxt is not None and nxt['kind'] == 'p' else None
            if stop is None:
                stop = next((k for k in range(a['i'] + 1, len(els)) if els[k].tag == qn('w:tbl') or (parts['T'][k] and is_top(parts['T'][k]) and a['top'])), len(els))
            last = stop - 1
            while last > a['i'] and els[last].tag == qn('w:p') and not parts['T'][last]: last -= 1
            ref = next((els[k] for k in range(last, a['i'], -1) if els[k].tag == qn('w:p') and parts['T'][k] and not is_top(parts['T'][k])), a['el'])
            s._para_insert(els[last], ref, gv_text, hs_text, label)
        elif a['kind'] == 'row':
            rs = tbl_rows(a['el']); gc = gv_col(a['el'])
            before = nxt['tr'] if nxt is not None and nxt['kind'] == 'row' else None
            hdr_n = len(tr_tcs(rs[0]))
            tmpl = next((r for r in rs[1:] if len(tr_tcs(r)) == hdr_n and len({ptext(tc) for tc in tr_tcs(r)}) > 1), None)
            if tmpl is None: tmpl = rs[-1]
            nr = copy.deepcopy(tmpl)
            for tc in tr_tcs(nr):
                ps = tc.findall(qn('w:p'))
                for p in ps[1:]: tc.remove(p)
                p0 = ps[0]
                for x in list(p0):
                    if x.tag != qn('w:pPr'): p0.remove(x)
            tcs = tr_tcs(nr); ref = tr_tcs(tmpl)[min(gc, len(tr_tcs(tmpl)) - 1)].find(qn('w:p'))
            sz = s.sz_of(ref)
            for ci, tc in enumerate(tcs):
                txt = gv_text if ci == gc else (hs_text if ci == (1 - gc if len(tcs) == 2 else gc + 1) else None)
                if txt: 
                    for c in s.line_children(label if ci == gc else '', txt, sz): tc.find(qn('w:p')).append(c)
            if before is not None: before.addprevious(nr)
            else: a['el'].append(nr)
        else:
            tr = a['tr']; tcs = tr_tcs(tr); gc = gv_col(a['el'])
            hp = tcs[a['ci']].findall(qn('w:p'))
            idx = hp.index(nxt['p']) if nxt is not None and nxt['kind'] == 'cell' else len(hp)
            s._cell_insert(tcs, gc, idx, gv_text, hs_text, label, a['ci'])
        return f'III – “{a["text"][:50]}”', note

    def _cell_insert(s, tcs, gc, idx, gv_text, hs_text, label, hcol=None):
        if len(tcs) < 2: gc = 0
        hs_c = (1 - gc) if len(tcs) == 2 else None
        for ci, tc in enumerate(tcs):
            if ci not in (gc, hs_c): continue
            ps = tc.findall(qn('w:p'))
            refs = [p for p in ps if ptext(p) and not is_top(ptext(p)) and not is_sub(ptext(p))]
            ref = refs[-1] if refs else (ps[-1] if ps else None)
            sz = s.sz_of(ref)
            txt = gv_text if ci == gc else hs_text
            p = new_p(ref, s.line_children(label if ci == gc else '', txt, sz) if txt else [])
            i = len(ps) if idx is None else idx
            if i < len(ps): ps[i].addprevious(p)
            elif ps: ps[-1].addnext(p)
            else: tc.append(p)
    def _para_insert(s, anchor, ref, gv_text, hs_text, label):
        sz = s.sz_of(ref)
        p = new_p(ref, s.line_children(label, gv_text, sz)); anchor.addnext(p)
        if hs_text:
            p2 = new_p(ref, s.line_children('', hs_text, sz)); p.addnext(p2)

def short_label(g, it=None):
    lab = (g.get('label') or '').strip()
    if not lab: lab = CAT_NAMES.get(g['cat'], 'Tích hợp')
    lab = re.sub(r'^tích hợp năng lực số', 'Tích hợp NLS', lab, flags=re.I)
    return lab[:70]

# ---------------- AI viết hoạt động ----------------
AI_SYS = ('Bạn là giáo viên tiểu học Việt Nam giỏi, soạn kế hoạch bài dạy lớp 2 theo Chương trình GDPT 2018 (Công văn 2345). '
          'Với mỗi gợi ý tích hợp, viết 1 hoạt động ngắn gọn gồm: "gv" – việc giáo viên làm (bắt đầu bằng "- GV", 1–2 câu), '
          '"hs" – việc học sinh làm tương ứng (bắt đầu bằng "- HS", 1–2 câu). Phù hợp lứa tuổi lớp 2, cụ thể, an toàn, không nêu tên trường, '
          'không bịa đường link. Trả về JSON đúng thứ tự id.')
def ai_write(jobs, call):
    """jobs: [{id, mon, bai, muc, nhan, goi_y}] ; call(system, user_json) -> list[{id,gv,hs}]"""
    if not jobs: return {}
    res = call(AI_SYS, json.dumps(jobs, ensure_ascii=False))
    out = {}
    for x in res or []:
        if isinstance(x, dict) and x.get('id') is not None and (x.get('gv') or x.get('hs')):
            out[str(x['id'])] = (clean_line(str(x.get('gv') or ''))[:600], clean_line(str(x.get('hs') or ''))[:600])
    return out

def gemini_json(system, user):
    """Gọi Gemini (REST) trả về JSON [{id,gv,hs}]. Lỗi -> ValueError thân thiện. Không ghi khoá ra log."""
    import urllib.request, urllib.error
    key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not key: raise ValueError('Chưa bật AI')
    import gemini
    schema = {'type': 'ARRAY', 'items': {'type': 'OBJECT', 'properties': {'id': {'type': 'STRING'}, 'gv': {'type': 'STRING'}, 'hs': {'type': 'STRING'}}, 'required': ['id', 'gv', 'hs']}}
    body = {'system_instruction': {'parts': [{'text': system}]}, 'contents': [{'role': 'user', 'parts': [{'text': user}]}],
            'generationConfig': {'temperature': 0.4, 'responseMimeType': 'application/json', 'responseSchema': schema}}
    try:
        res = gemini.post(body, 120)
    except urllib.error.HTTPError as e:
        log.warning('tichhop gemini: HTTP %s', e.code)
        if gemini.special_msg(e.code): raise ValueError(gemini.special_msg(e.code) + '; đã chép nguyên văn theo kế hoạch')
        raise ValueError('AI đang bận hoặc lỗi, đã chép nguyên văn theo kế hoạch')
    except Exception as e:
        log.warning('tichhop gemini: lỗi kết nối %s', type(e).__name__); raise ValueError('Chưa kết nối được AI, đã chép nguyên văn theo kế hoạch')
    try:
        txt = ''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts'])
        return json.loads(txt)
    except Exception:
        raise ValueError('AI trả kết quả không đọc được, đã chép nguyên văn theo kế hoạch')

# ---------------- áp vào giáo án ----------------
def apply(G, M, w, plan, built, cats=None, sel=None, picks=None, mode='verbatim', ai_call=None):
    """built = dict(heads=[...], rows=[...]) từ core.step3_insert. Trả về báo cáo."""
    cats = set(cats or CAT_KEYS); sel = sel or {}; picks = picks or {}
    r = match_week(plan, M, w)
    rows = built.get('rows') or []
    key_of = {id(rw): rowkey(k) for k, rw in core.keyed(rows)}
    les_of = {}     # rowkey -> (lesson dict, head)
    for h in built.get('heads') or []:
        Ls = h.get('lessons') or []
        rws = [x for x in h.get('rows', []) if x['ten'].strip()]
        if not Ls: 
            for x in rws: les_of[key_of[id(x)]] = (None, h)
            continue
        for j, x in enumerate(rws):
            if len(Ls) == 1: L = Ls[0]
            elif len(Ls) == len(rws): L = Ls[j]
            else: L = max(Ls, key=lambda y: core.score([x['ten']], y['l']['title']))
            les_of[key_of[id(x)]] = (L['l'], h)
    rep = dict(mode=mode, placed=[], skipped=[], problems=[], unmatched=[dict(subject=u['subject'], lesson=u['lesson'], week=u['week'], reason=u['reason']) for u in r['unmatched']],
               unlabeled=r['unlabeled'], ai_note='')
    A = Applier(G)
    jobs = []   # gợi ý hoạt động (cần chọn cách viết)
    for m in r['matches']:
        keys = [x['key'] for x in m['rows']]
        lessons = []
        for k in keys:
            L, h = les_of.get(k, (None, None))
            if L is not None and all(L is not y for y in lessons): lessons.append(L)
        if len(lessons) > 1:
            rep['problems'].append(f"{m['subject']} – “{m['lesson']}” khớp nhiều bài soạn ({', '.join(x['title'][:40] for x in lessons)}): đặt vào bài đầu (theo 'Tiết' ghi trong kế hoạch nếu có)")
        per = {}   # id(lesson) -> {I4:[], II:[], act:[]}
        for g in m['groups']:
            k = picks.get(m['pid']) or target_row(m, g)['key']
            if g.get('tiet') and len(m['rows']) >= g['tiet']: k = picks.get(m['pid']) or m['rows'][g['tiet'] - 1]['key']
            L, h = les_of.get(k, (None, None))
            if L is None and lessons and not picks.get(m['pid']): L = lessons[0]
            row = next((x for x in m['rows'] if x['key'] == k), m['rows'][0])
            where_lbl = f"T{row['thu']} {row['mon']}: {row['ten'][:60]}"
            for it in g['items']:
                o = sel.get(it['iid'], {})
                on = o.get('on', g['cat'] in cats)
                if not on: rep['skipped'].append(dict(lesson=where_lbl, label=short_label(g), text=it['text'][:80], why='bỏ chọn')); continue
                if L is None:
                    msg = f"{where_lbl}: chưa có bài soạn trong giáo án phụ – chưa chèn tích hợp"
                    if msg not in rep['problems']: rep['problems'].append(msg)
                    rep['skipped'].append(dict(lesson=where_lbl, label=short_label(g), text=it['text'][:80], why='chưa có bài soạn')); continue
                dest = o.get('dest') or it['dest']
                hint = None if o.get('dest') else it.get('hint')
                d = per.setdefault(id(L), dict(L=L, I4=[], II=[], act=[], where=where_lbl, m=m))
                lab = it['label'] if it['kind'] != 'tichhop' else (it['label'] or '')
                if dest == 'I4':
                    d['I4'].append(((lab or ('' if it['kind'] == 'tichhop' else short_label(g))), it['text'], it['urls'], g, it))
                elif dest == 'II':
                    d['II'].append((lab or short_label(g), it['text'], it['urls'], g, it))
                else:
                    d['act'].append(dict(letter=dest if dest in 'ABCD' else 'B', hint=hint, text=it['text'], label=short_label(g), g=g, it=it))
        for d in per.values():
            L = d['L']; els = L.get('els') or []
            parts = lesson_parts(els)
            if d['I4']:
                where = A.add_I4(els, parts, [(a, b, c) for a, b, c, _, _ in d['I4']])
                for a, b, c, g, it in d['I4']: rep['placed'].append(dict(lesson=d['where'], label=a or short_label(g), text=b[:120], where=where, cat=g['cat']))
                parts = lesson_parts(els_refresh(L))
                els = L['els']
            if d['II']:
                where = A.add_II(els, parts, [(a, b, c) for a, b, c, _, _ in d['II']])
                for a, b, c, g, it in d['II']: rep['placed'].append(dict(lesson=d['where'], label=a, text=(b + ' ' + ' '.join(c)).strip()[:120], where=where or 'không thấy mục II', cat=g['cat']))
                parts = lesson_parts(els_refresh(L)); els = L['els']
            for j in d['act']:
                j['L'] = L; j['where'] = d['where']; j['subject'] = d['m']['subject']; jobs.append(j)
    # gợi ý hoạt động
    texts = {}
    if mode == 'ai' and jobs:
        try:
            texts = ai_write([dict(id=str(i), mon=j['subject'], bai=j['where'], muc=NAMES.get(j['letter'], ''), nhan=j['label'], goi_y=j['text']) for i, j in enumerate(jobs)],
                             ai_call or gemini_json)
            rep['ai_used'] = True
        except ValueError as e:
            rep['ai_note'] = str(e); rep['mode'] = 'verbatim'
    for i, j in enumerate(jobs):
        L = j['L']; els = els_refresh(L); parts = lesson_parts(els)
        if str(i) in texts: gv, hs = texts[str(i)]; lab = j['label']
        else: gv, hs = j['text'], ''; lab = j['label']
        where, note = A.add_act(els, parts, j['letter'], j['hint'], gv, hs, lab)
        rep['placed'].append(dict(lesson=j['where'], label=lab, text=(gv if str(i) in texts else j['text'])[:160], where=where, note=note, cat=j['g']['cat'], act=True))
        if note: rep['problems'].append(f"{j['where']}: {note}")
    rep['n_act'] = len(jobs); rep['n'] = len(rep['placed'])
    return rep

def els_refresh(L):
    """Danh sách phần tử cấp thân của bài (kể cả đoạn mới chèn ở cấp thân)."""
    els = L.get('els') or []
    if not els: return els
    first, last = els[0], els[-1]
    out = [first]; e = first
    while e is not last:
        e = e.getnext()
        if e is None: break
        out.append(e)
    L['els'] = out; return out

def save_plan_json(plan, path):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(plan, f, ensure_ascii=False)
    os.replace(tmp, path)
def load_plan_json(path):
    with open(path, encoding='utf-8') as f: return json.load(f)
