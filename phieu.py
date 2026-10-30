"""Phiếu bài tập & Trò chơi (lớp 2, GDPT 2018) cho 1 bài học trong tuần.
- AI (Gemini, 1 lượt/lần tạo, JSON có lược đồ) viết: phiếu bài tập (4–6 bài), trò chơi khởi động, bộ câu hỏi trò chơi tương tác + ô chữ.
- Kết quả AI được kiểm tra chặt (validate) – thiếu/sai thì báo lỗi, KHÔNG tự bịa nội dung thay AI.
- Xuất .docx A4 (chữ to, chỗ trống/dòng chấm/ô vuông nằm trong lề, trang đáp án riêng)."""
import os, re, json, logging, unicodedata, hashlib
log = logging.getLogger('phieu')
TYPES = ['phieu', 'khoidong', 'tuongtac']
TYPE_VI = {'phieu': 'Phiếu bài tập', 'khoidong': 'Trò chơi khởi động', 'tuongtac': 'Trò chơi tương tác'}
MUC = {'de': 'Dễ', 'vua': 'Vừa', 'nc': 'Nâng cao'}
KINDS = {'mcq': 'Khoanh vào chữ cái đặt trước câu trả lời đúng', 'fill': 'Điền vào chỗ chấm', 'tf': 'Đúng ghi Đ, sai ghi S vào ô trống',
         'match': 'Nối cho phù hợp', 'write': 'Trả lời câu hỏi / viết vào chỗ chấm'}
L = 'ABCD'

def _s(x, n=400):
    x = '' if x is None else str(x)
    return re.sub(r'\s+', ' ', x).strip()[:n]

def norm_letters(s):
    """Chữ in hoa không dấu, bỏ khoảng trắng (ô chữ)."""
    s = unicodedata.normalize('NFD', str(s or '')).replace('đ', 'd').replace('Đ', 'D')
    return re.sub(r'[^A-Z0-9]', '', ''.join(c for c in s if unicodedata.category(c) != 'Mn').upper())

def _ans_letter(a, n):
    a = _s(a, 10).upper().strip('. )')
    if a[:1] in L[:n] and (len(a) == 1 or not a[1:2].isalpha()): return a[:1]
    if a.isdigit() and 1 <= int(a) <= n: return L[int(a) - 1]
    return ''

def _ans_tf(a):
    a = _s(a, 10).upper()
    if a[:1] in ('Đ', 'D', 'T') or a in ('ĐÚNG', 'DUNG', 'TRUE'): return 'Đ'
    if a[:1] in ('S', 'F') or a in ('SAI', 'FALSE'): return 'S'
    return ''

def clean_item(kind, it):
    it = it if isinstance(it, dict) else {}
    o = dict(q=_s(it.get('q'), 300), a=_s(it.get('a'), 300))
    if kind == 'mcq':
        o['options'] = [_s(x, 120) for x in (it.get('options') or [])[:4] if _s(x, 120)] if isinstance(it.get('options'), list) else []
        o['a'] = _ans_letter(o['a'], len(o['options'])) or o['a'][:1].upper()
    elif kind == 'tf': o['a'] = _ans_tf(o['a']) or o['a'][:1]
    elif kind == 'match': o['left'] = _s(it.get('left') or it.get('q'), 120); o['right'] = _s(it.get('right') or it.get('a'), 120); o['q'] = o['a'] = ''
    elif kind == 'write':
        try: o['lines'] = max(1, min(6, int(it.get('lines') or 2)))
        except (TypeError, ValueError): o['lines'] = 2
    return o

def item_ok(kind, o):
    if kind == 'mcq': return bool(o['q'] and 2 <= len(o['options']) <= 4 and o['a'] in L[:len(o['options'])])
    if kind == 'tf': return bool(o['q'] and o['a'] in ('Đ', 'S'))
    if kind == 'match': return bool(o['left'] and o['right'])
    return bool(o['q'])

def clean_ex(e):
    e = e if isinstance(e, dict) else {}
    k = e.get('kind') if e.get('kind') in KINDS else 'write'
    items = [clean_item(k, x) for x in (e.get('items') or [])[:8]] if isinstance(e.get('items'), list) else []
    return dict(kind=k, title=_s(e.get('title'), 200) or KINDS[k], items=items)

def clean_gq(q):
    q = q if isinstance(q, dict) else {}
    opts = [_s(x, 120) for x in (q.get('options') or [])[:4]] if isinstance(q.get('options'), list) else []
    while len(opts) < 4: opts.append('')
    return dict(q=_s(q.get('q'), 300), options=opts, a=_ans_letter(q.get('a'), 4) or _s(q.get('a'), 1).upper(), explain=_s(q.get('explain'), 300))

def gq_ok(q): return bool(q['q'] and all(q['options']) and q['a'] in L)

def ochu_rows(oc):
    """Vị trí chữ của từ khoá trong từng hàng ngang (-1 = hàng không chứa chữ đó)."""
    kw = norm_letters(oc.get('keyword')); out = []
    for i, r in enumerate(oc.get('rows') or []):
        ans = norm_letters(r.get('answer')); ch = kw[i] if i < len(kw) else ''
        out.append(dict(clue=r.get('clue', ''), answer=ans, pos=ans.find(ch) if ch else -1))
    return kw, out

def ochu_ok(oc):
    kw, rows = ochu_rows(oc)
    return bool(3 <= len(kw) <= 12 and len(rows) == len(kw) and all(r['pos'] >= 0 and r['clue'] for r in rows))

def clean_state(d):
    d = d if isinstance(d, dict) else {}
    ls = d.get('lesson') if isinstance(d.get('lesson'), dict) else {}
    try: wk = int(ls.get('week'))
    except (TypeError, ValueError): wk = None
    o = dict(lesson=dict(week=wk if wk and 1 <= wk <= 60 else None, mon=_s(ls.get('mon'), 60), thu=_s(ls.get('thu'), 20), tiet=_s(ls.get('tiet'), 20), ten=_s(ls.get('ten'), 200)),
             muc=d.get('muc') if d.get('muc') in MUC else 'vua', types=[t for t in TYPES if t in (d.get('types') or [])] or ['phieu'],
             src=d.get('src') if d.get('src') in ('ai', 'tay') else 'tay')
    p = d.get('phieu') if isinstance(d.get('phieu'), dict) else {}
    o['phieu'] = dict(exercises=[clean_ex(e) for e in (p.get('exercises') or [])[:8]] if isinstance(p.get('exercises'), list) else [])
    k = d.get('khoidong') if isinstance(d.get('khoidong'), dict) else {}
    o['khoidong'] = dict(name=_s(k.get('name'), 120), luat=_s(k.get('luat'), 1500),
                         questions=[dict(q=_s(x.get('q'), 300), a=_s(x.get('a'), 300)) for x in (k.get('questions') or [])[:15] if isinstance(x, dict)] if isinstance(k.get('questions'), list) else [])
    g = d.get('game') if isinstance(d.get('game'), dict) else {}
    oc = g.get('ochu') if isinstance(g.get('ochu'), dict) else {}
    o['game'] = dict(questions=[clean_gq(q) for q in (g.get('questions') or [])[:20]] if isinstance(g.get('questions'), list) else [],
                     ochu=dict(keyword=_s(oc.get('keyword'), 40), clue=_s(oc.get('clue'), 200),
                               rows=[dict(clue=_s(r.get('clue'), 200), answer=_s(r.get('answer'), 40)) for r in (oc.get('rows') or [])[:12] if isinstance(r, dict)] if isinstance(oc.get('rows'), list) else []))
    return o

# ---------------- AI ----------------
_STR = {'type': 'STRING'}
_ITEM = {'type': 'OBJECT', 'properties': {'q': _STR, 'options': {'type': 'ARRAY', 'items': _STR}, 'a': _STR, 'left': _STR, 'right': _STR, 'lines': {'type': 'INTEGER'}}}
_GQ = {'type': 'OBJECT', 'properties': {'q': _STR, 'options': {'type': 'ARRAY', 'items': _STR}, 'a': _STR, 'explain': _STR}, 'required': ['q', 'options', 'a']}
def schema(types):
    pr, req = {}, []
    if 'phieu' in types:
        pr['phieu'] = {'type': 'OBJECT', 'properties': {'exercises': {'type': 'ARRAY', 'items': {'type': 'OBJECT', 'properties': {
            'kind': {'type': 'STRING', 'enum': list(KINDS)}, 'title': _STR, 'items': {'type': 'ARRAY', 'items': _ITEM}}, 'required': ['kind', 'title', 'items']}}}, 'required': ['exercises']}
        req.append('phieu')
    if 'khoidong' in types:
        pr['khoidong'] = {'type': 'OBJECT', 'properties': {'name': _STR, 'luat': _STR, 'questions': {'type': 'ARRAY', 'items': {'type': 'OBJECT', 'properties': {'q': _STR, 'a': _STR}, 'required': ['q', 'a']}}},
                          'required': ['name', 'luat', 'questions']}
        req.append('khoidong')
    if 'tuongtac' in types:
        pr['game'] = {'type': 'OBJECT', 'properties': {'questions': {'type': 'ARRAY', 'items': _GQ}, 'ochu': {'type': 'OBJECT', 'properties': {
            'keyword': _STR, 'clue': _STR, 'rows': {'type': 'ARRAY', 'items': {'type': 'OBJECT', 'properties': {'clue': _STR, 'answer': _STR}, 'required': ['clue', 'answer']}}}, 'required': ['keyword', 'rows']}},
                      'required': ['questions', 'ochu']}
        req.append('game')
    return {'type': 'OBJECT', 'properties': pr, 'required': req}

SYSTEM = ('Bạn là giáo viên tiểu học giỏi ở Việt Nam, soạn học liệu cho HỌC SINH LỚP 2 theo Chương trình GDPT 2018. '
          'Nội dung phải ĐÚNG KIẾN THỨC, đúng bài học được nêu, vừa sức lớp 2 (câu ngắn, từ ngữ gần gũi, số trong phạm vi chương trình lớp 2), '
          'tiếng Việt chuẩn chính tả, không nhắc tên trường, không đưa thông tin cá nhân. Chỉ trả về JSON đúng lược đồ.')
MUC_HD = {'de': 'Dễ: nhận biết, nhắc lại kiến thức cơ bản của bài.', 'vua': 'Vừa: nhận biết + thông hiểu, có 1–2 câu vận dụng đơn giản.',
          'nc': 'Nâng cao: thông hiểu + vận dụng, có câu suy luận/giải quyết tình huống gần gũi (vẫn vừa sức lớp 2).'}

def prompt(st, ctx=''):
    ls = st['lesson']; t = st['types']; parts = [
        f"Bài học: {ls['ten']}" + (f" – môn {ls['mon']}" if ls['mon'] else '') + (f" (tuần {ls['week']})" if ls['week'] else ''),
        'Mức độ: ' + MUC_HD[st['muc']]]
    if 'phieu' in t: parts.append(
        'phieu: 4 đến 6 bài tập (exercises) đa dạng dạng bài, mỗi bài 2–5 ý (items). kind: '
        '"mcq" (q + options 3–4 phương án + a là chữ cái A/B/C/D đúng), "fill" (q có chỗ trống viết là ___ , a là đáp án), '
        '"tf" (q là một nhận định, a là "Đ" hoặc "S"), "match" (left – right là cặp đúng; 3–5 cặp), "write" (q câu hỏi, a đáp án gợi ý, lines 1–4 số dòng để HS viết). '
        'title là lệnh của bài (VD "Khoanh vào chữ cái đặt trước câu trả lời đúng:"), KHÔNG ghi "Bài 1".')
    if 'khoidong' in t: parts.append('khoidong: trò chơi khởi động 3–5 phút đầu tiết, in ra giấy cho GV đọc: name (tên trò chơi vui), luat (luật chơi rõ ràng 3–5 câu), questions 5–8 câu hỏi ngắn ôn kiến thức liên quan bài (q, a).')
    if 'tuongtac' in t: parts.append(
        'game: questions 8–12 câu trắc nghiệm (q, đúng 4 options ngắn, a là chữ A/B/C/D đúng, explain 1 câu giải thích ngắn); vị trí đáp án đúng phân bố đều A/B/C/D. '
        'ochu: ô chữ – keyword là TỪ KHOÁ hàng dọc 4–8 chữ cái liên quan bài (viết liền, có dấu cũng được); rows có ĐÚNG số hàng bằng số chữ cái của keyword (không tính dấu, khoảng trắng); '
        'hàng thứ i có answer (từ 3–10 chữ cái) PHẢI chứa chữ cái thứ i của keyword (so sánh không dấu), clue là câu gợi ý cho học sinh lớp 2.')
    if ctx: parts.append('Trích nội dung giáo án của bài (tham khảo, bám sát mục tiêu/hoạt động):\n' + ctx[:4000])
    return '\n'.join(parts)

def validate_ai(raw, st):
    """Ghép kết quả AI vào trạng thái; thiếu/sai -> ValueError (tiếng Việt). Không bịa thêm nội dung."""
    if not isinstance(raw, dict): raise ValueError('AI trả kết quả không đúng dạng')
    t = st['types']; o = clean_state({**st, **{k: raw.get(k) for k in ('phieu', 'khoidong', 'game') if k in raw}, 'src': 'ai'}); bad = []; warn = []
    if 'phieu' in t:
        ex = []
        for e in o['phieu']['exercises']:
            e['items'] = [x for x in e['items'] if item_ok(e['kind'], x)]
            if e['items'] and not (e['kind'] == 'match' and len(e['items']) < 2): ex.append(e)
        o['phieu']['exercises'] = ex[:6]
        if len(ex) < 4: bad.append(f'phiếu bài tập chỉ có {len(ex)} bài hợp lệ (cần 4–6)')
    else: o['phieu'] = st['phieu']
    if 'khoidong' in t:
        o['khoidong']['questions'] = [q for q in o['khoidong']['questions'] if q['q'] and q['a']]
        if not (o['khoidong']['name'] and o['khoidong']['luat'] and len(o['khoidong']['questions']) >= 3): bad.append('trò chơi khởi động thiếu tên/luật chơi/câu hỏi')
    else: o['khoidong'] = st['khoidong']
    if 'tuongtac' in t:
        o['game']['questions'] = [q for q in o['game']['questions'] if gq_ok(q)]
        if len(o['game']['questions']) < 5: bad.append(f"bộ câu hỏi trò chơi chỉ có {len(o['game']['questions'])} câu hợp lệ (cần ít nhất 5)")
        if not ochu_ok(o['game']['ochu']):
            o['game']['ochu'] = dict(keyword='', clue='', rows=[]); warn.append('Ô chữ AI tạo chưa khớp từ khoá nên đã bỏ – thầy cô có thể tự nhập ô chữ.')
    else: o['game'] = st['game']
    if bad: raise ValueError('AI trả nội dung chưa đạt (' + '; '.join(bad) + '). Thầy cô bấm "Tạo bằng AI" lại hoặc tự nhập.')
    return o, warn

def gemini_call(system, user, sch):
    import urllib.request, urllib.error
    key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not key: raise ValueError('Chưa bật AI')
    import gemini
    body = {'system_instruction': {'parts': [{'text': system}]}, 'contents': [{'role': 'user', 'parts': [{'text': user}]}],
            'generationConfig': {'temperature': 0.7, 'responseMimeType': 'application/json', 'responseSchema': sch}}
    try:
        res = gemini.post(body, 120)
    except urllib.error.HTTPError as e:
        log.warning('phieu gemini: HTTP %s', e.code)
        if gemini.special_msg(e.code): raise ValueError(gemini.special_msg(e.code))
        raise ValueError('AI đang quá tải hoặc hết lượt miễn phí, thầy cô thử lại sau ít phút' if e.code == 429 else f'AI đang bận (lỗi {e.code}), thầy cô thử lại sau')
    except Exception as e: log.warning('phieu gemini: lỗi kết nối %s', type(e).__name__); raise ValueError('Không kết nối được AI, thầy cô thử lại sau')
    try:
        txt = ''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts'])
        return json.loads(re.sub(r'^```(?:json)?\s*|\s*```$', '', txt.strip()))
    except Exception: raise ValueError('AI trả kết quả không đọc được, thầy cô bấm tạo lại')

def generate(st, ctx='', call=None):
    raw = (call or gemini_call)(SYSTEM, prompt(st, ctx), schema(st['types']))
    return validate_ai(raw, st)

# ---------------- giáo án làm ngữ cảnh ----------------
def lesson_context(ga_path, ten, n=3500):
    """Trích đoạn giáo án quanh tên bài (không có -> '')."""
    try:
        import docx
        d = docx.Document(ga_path); txt = []
        for el in d.element.body.iter():
            if el.tag.endswith('}p'):
                s = ''.join(t.text or '' for t in el.iter() if t.tag.endswith('}t')).strip()
                if s: txt.append(s)
        del d
        key = norm_letters(ten)[:40]
        if len(key) < 4: return ''
        for i, s in enumerate(txt):
            if key in norm_letters(s):
                return '\n'.join(txt[i:i + 120])[:n]
    except Exception as e: log.warning('phieu: không đọc được giáo án (%s)', type(e).__name__)
    return ''

# ---------------- xuất Word ----------------
def _doc():
    import docx
    from docx.shared import Cm, Pt
    from docx.oxml.ns import qn
    d = docx.Document(); s = d.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7); s.left_margin = s.right_margin = Cm(2); s.top_margin = s.bottom_margin = Cm(1.8)
    st = d.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    pf = st.paragraph_format; pf.space_after = Pt(3); pf.space_before = Pt(0); pf.line_spacing = 1.2
    return d

TEXT_W = 17.0   # cm (21 - 2 - 2)
TBL_W = 16.4    # bảng hẹp hơn vùng chữ + thụt 0,2 cm -> viền bảng không vượt lề

def _par(d, text='', bold=False, size=None, align=None, italic=False, after=None):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = d.add_paragraph(); r = p.add_run(text); r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p

def _dots(p, *stops):
    """Tab có dòng chấm dẫn tới các mốc (cm) – luôn dừng trong lề phải."""
    from docx.shared import Cm
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    for x in stops: p.paragraph_format.tab_stops.add_tab_stop(Cm(min(x, TEXT_W - 0.05)), WD_TAB_ALIGNMENT.RIGHT if x >= TEXT_W - 0.1 else WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.DOTS)

def _line(d, prefix=''):
    p = d.add_paragraph(); p.add_run(prefix + '\t'); _dots(p, TEXT_W); p.paragraph_format.space_after = 0; p.paragraph_format.line_spacing = 1.5
    return p

def _table(d, widths, rows, border=True):
    from docx.shared import Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    t = d.add_table(rows=len(rows), cols=len(widths)); t.autofit = False
    if border: t.style = d.styles['Table Grid']
    tp = t._tbl.tblPr
    for tag, attrs in (('w:tblLayout', {'w:type': 'fixed'}), ('w:tblInd', {'w:w': str(int(0.2 / 2.54 * 1440)), 'w:type': 'dxa'}),
                       ('w:tblW', {'w:w': str(int(sum(widths) / 2.54 * 1440)), 'w:type': 'dxa'})):
        for old in tp.findall(qn(tag)): tp.remove(old)
        e = OxmlElement(tag)
        for k, v in attrs.items(): e.set(qn(k), v)
        tp.append(e)
    for j, w in enumerate(widths): t.columns[j].width = Cm(w)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j); c.width = Cm(widths[j]); c.text = ''
            par = c.paragraphs[0]; par.paragraph_format.space_after = 0
            if isinstance(val, tuple): r = par.add_run(val[0]); r.bold = val[1]
            else: par.add_run(str(val))
    return t

def _blank(q):
    q = re.sub(r'_{2,}|\.{4,}|…{2,}', '……………', q)
    return q if '……' in q else q + ' ……………'

def _shuffle(items, seed):
    h = int(hashlib.md5(seed.encode()).hexdigest(), 16); idx = list(range(len(items)))
    for i in range(len(idx) - 1, 0, -1): j = h % (i + 1); h //= (i + 1); idx[i], idx[j] = idx[j], idx[i]
    if len(idx) > 1 and all(i == v for i, v in enumerate(idx)): idx = idx[1:] + idx[:1]
    return idx

def export_phieu(st, out):
    ex = [e for e in st['phieu']['exercises'] if any(item_ok(e['kind'], x) or (e['kind'] in ('fill', 'write') and x['q']) for x in e['items'])]
    if not ex: raise ValueError('Phiếu chưa có bài tập nào có nội dung')
    d = _doc(); ls = st['lesson']
    p = d.add_paragraph(); p.add_run('Họ và tên:\t'); p.add_run('Lớp:\t'); _dots(p, 11.6, TEXT_W)
    _par(d, 'PHIẾU BÀI TẬP', bold=True, size=18, align='c', after=0)
    _par(d, ((ls['mon'] + ' – ') if ls['mon'] else '') + (ls['ten'] or 'Bài học'), bold=True, size=15, align='c', after=0)
    _par(d, ('Tuần ' + str(ls['week']) + ' · ' if ls['week'] else '') + 'Mức: ' + MUC[st['muc']], italic=True, size=12, align='c', after=6)
    answers = []
    for n, e in enumerate(ex, 1):
        _par(d, f'Bài {n}. {e["title"]}', bold=True, after=2); k = e['kind']; its = [x for x in e['items'] if item_ok(k, x) or (k in ('fill', 'write') and x['q'])]
        ab = []
        if k == 'mcq':
            for i, x in enumerate(its):
                _par(d, f'{chr(97 + i)}) {x["q"]}', after=1)
                o = x['options']; cells = [f'□ {L[j]}. {o[j]}' for j in range(len(o))]
                rows = [cells[j:j + 2] + [''] * (2 - len(cells[j:j + 2])) for j in range(0, len(cells), 2)]
                _table(d, [TBL_W / 2] * 2, rows, border=False); ab.append(f'{chr(97 + i)}) {x["a"]}')
        elif k == 'fill':
            for i, x in enumerate(its): _par(d, f'{chr(97 + i)}) {_blank(x["q"])}', after=2); ab.append(f'{chr(97 + i)}) {x["a"] or "(GV tự chấm)"}')
        elif k == 'tf':
            _table(d, [1.0, TBL_W - 2.6, 1.6], [[(f'{chr(97 + i)})', False), x['q'], ''] for i, x in enumerate(its)])
            ab += [f'{chr(97 + i)}) {x["a"]}' for i, x in enumerate(its)]
        elif k == 'match':
            order = _shuffle(its, ls['ten'] + e['title']); rows = []
            for i in range(len(its)): rows.append([f'{i + 1}. {its[i]["left"]}', '', f'{chr(97 + i)}. {its[order[i]]["right"]}'])
            t = _table(d, [6.7, TBL_W - 13.4, 6.7], rows)
            _no_mid_border(t)
            ab.append(', '.join(f'{i + 1} – {chr(97 + order.index(i))}' for i in range(len(its))))
        else:
            for i, x in enumerate(its):
                _par(d, f'{chr(97 + i)}) {x["q"]}', after=0)
                for _ in range(x.get('lines', 2)): _line(d)
                ab.append(f'{chr(97 + i)}) {x["a"] or "(HS tự trả lời – GV chấm theo bài làm)"}')
        _par(d, '', after=2); answers.append((n, e['title'], ab))
    from docx.enum.text import WD_BREAK
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    _par(d, 'ĐÁP ÁN – PHIẾU BÀI TẬP', bold=True, size=16, align='c', after=0)
    _par(d, ls['ten'] or '', italic=True, size=13, align='c', after=6)
    for n, title, ab in answers:
        _par(d, f'Bài {n}. ' + '; '.join(ab), after=3)
    d.save(out); del d

def _no_mid_border(t):
    """Cột giữa bảng nối: không kẻ viền trên/dưới (chỗ để HS nối)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    for row in t.rows:
        tcPr = row.cells[1]._tc.get_or_add_tcPr(); b = OxmlElement('w:tcBorders')
        for side in ('top', 'bottom', 'insideH'):
            e = OxmlElement(f'w:{side}'); e.set(qn('w:val'), 'nil'); b.append(e)
        tcPr.append(b)

def export_khoidong(st, out):
    k = st['khoidong']; qs = [q for q in k['questions'] if q['q']]
    if not (k['name'] or qs): raise ValueError('Trò chơi khởi động chưa có nội dung')
    d = _doc(); ls = st['lesson']
    _par(d, 'TRÒ CHƠI KHỞI ĐỘNG', bold=True, size=18, align='c', after=0)
    _par(d, k['name'] or 'Trò chơi', bold=True, size=16, align='c', after=0)
    _par(d, ((ls['mon'] + ' – ') if ls['mon'] else '') + (ls['ten'] or '') + (f' · Tuần {ls["week"]}' if ls['week'] else ''), italic=True, size=12, align='c', after=8)
    _par(d, 'Luật chơi:', bold=True, after=2)
    for s in [x.strip() for x in re.split(r'(?<=[.!?])\s+', k['luat']) if x.strip()] or ['(Thầy cô ghi luật chơi)']: _par(d, '• ' + s, after=2)
    _par(d, 'Câu hỏi:', bold=True, after=2)
    for i, q in enumerate(qs, 1): _par(d, f'Câu {i}. {q["q"]}', after=3)
    _par(d, 'Đáp án:', bold=True, after=2)
    _table(d, [1.6, TBL_W - 1.6], [[('Câu', True), ('Đáp án', True)]] + [[str(i), q['a'] or '(GV tự đánh giá)'] for i, q in enumerate(qs, 1)])
    d.save(out); del d
