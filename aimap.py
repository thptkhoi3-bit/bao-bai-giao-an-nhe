"""AI hiểu mẫu báo bài / giáo án khi bộ quy tắc không nhận diện được.

Luồng: dựng bản tóm tắt cấu trúc các bảng ứng viên (chỉ vài dòng đầu mỗi bảng,
không có dữ liệu cá nhân) -> Gemini trả sơ đồ cột dạng JSON -> kiểm chứng bằng
quy tắc trên dữ liệu thật (đọc thử theo sơ đồ: tỉ lệ dòng ra ngày/môn/tên bài)
-> đạt mới lưu vào colmap.json của người dùng theo chữ ký tiêu đề bảng (cơ chế
CMAP sẵn có của core) để mọi lần sau đi thẳng bằng quy tắc, không tốn AI nữa.
Chỉ chạy khi bộ quy tắc không đạt; mỗi mẫu lạ tốn đúng một lần gọi AI.
"""
import os, re, json, logging
import core, gemini

log = logging.getLogger('aimap')

PROMPT = """Bạn là chuyên gia đọc cấu trúc bảng trong văn bản Word của trường học Việt Nam.
Dưới đây là bản đồ cấu trúc một tài liệu: vài đoạn văn đầu và các bảng ứng viên.
Mỗi bảng được in theo dòng R0, R1, ...; trong một dòng, các ô ngăn nhau bằng dấu " | " và được đánh số cột từ 0 ở phía trái. Ô bị gộp sẽ lặp lại cùng một nội dung ở các cột nó chiếm.

Hãy tìm bảng KẾ HOẠCH TUẦN / THỜI KHÓA BIỂU (bảng liệt kê các tiết học theo thứ, buổi, tiết, môn, tên bài) và trả về DUY NHẤT một JSON đúng khuôn:
{"table": <số hiệu bảng>, "week": <số tuần dạng số nguyên hoặc null>, "hdr_rows": <số dòng đầu không phải dữ liệu bài học>, "cols": {"thu": <số cột hoặc null>, "buoi": <số cột hoặc null>, "tiet": <số cột hoặc null>, "mon": <số cột hoặc null>, "ten": <số cột hoặc null>, "tich": <số cột hoặc null>, "dd": <số cột hoặc null>}}

Định nghĩa các vai trò cột:
- thu: Thứ / ngày học (giá trị kiểu: Hai, Ba, Thứ 2, 27/10...)
- buoi: Buổi (Sáng/Chiều) - có trường ghi cột này là "Thời gian"
- tiet: Tiết học / số thứ tự tiết (1, 2, 3...)
- mon: Môn học (Toán, Tiếng Việt, SHDC...)
- ten: Tên bài học / tên bài dạy / nội dung bài
- tich: Tích hợp / điều chỉnh / ghi chú
- dd: Đồ dùng dạy học / thiết bị / phương tiện

hdr_rows gồm mọi dòng phía trên dòng dữ liệu bài học đầu tiên: dòng tựa đề (ví dụ "THỜI KHOÁ BIỂU LỚP 4B", "TUẦN 09"), dòng tiêu đề cột, và cả dòng tiêu đề cột bị lặp lại lần thứ hai nếu có.
week: số tuần nếu thấy ghi ở bất cứ đâu trong tài liệu (tựa đề bảng, đoạn văn, dòng đầu bảng); không thấy thì null.
Nếu không bảng nào là bảng kế hoạch tuần, trả {"table": null}.
Chỉ trả JSON, không giải thích."""


def _digest(doc):
    """-> (văn bản tóm tắt, {số hiệu bảng: table}) hoặc (None, {}) nếu không có bảng ứng viên."""
    cands = {}
    for i, t in enumerate(doc.tables):
        if len(t.rows) >= 5 and max((len(r.cells) for r in t.rows), default=0) >= 4:
            cands[i] = t
        if len(cands) >= 3:
            break
    if not cands:
        return None, {}
    lines = []
    paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()][:10]
    if paras:
        lines.append('CÁC ĐOẠN VĂN ĐẦU TÀI LIỆU:')
        lines += ['- ' + p[:120] for p in paras]
    for i, t in cands.items():
        w = max(len(r.cells) for r in t.rows)
        lines.append(f'BẢNG {i} ({len(t.rows)} dòng, {w} cột):')
        for ri, r in enumerate(t.rows[:9]):
            cells = [re.sub(r'\s+', ' ', core.ctext(c)).strip()[:45] for c in r.cells]
            cells += [''] * (w - len(cells))
            lines.append(f'R{ri}: ' + ' | '.join(cells))
    return '\n'.join(lines), cands


def _ask(digest):
    body = {'contents': [{'parts': [{'text': PROMPT + '\n\n' + digest}]}],
            'generationConfig': {'temperature': 0, 'responseMimeType': 'application/json'}}
    res = gemini.post(body, timeout=45)
    txt = ''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts'])
    try:
        return json.loads(txt)
    except Exception:
        m = re.search(r'\{.*\}', txt, re.S)
        return json.loads(m.group(0)) if m else None


def _verify(doc, cands, ai):
    """Kiểm chứng sơ đồ AI bằng quy tắc trên dữ liệu thật -> (table, mục colmap) hoặc None."""
    if not isinstance(ai, dict) or ai.get('table') is None:
        return None
    try:
        ti = int(ai.get('table'))
    except Exception:
        return None
    t = cands.get(ti)
    if t is None:
        return None
    w = max(len(r.cells) for r in t.rows)
    cols = {}
    for k, v in (ai.get('cols') or {}).items():
        if k in core.ROLES and isinstance(v, int) and 0 <= v < w:
            cols[k] = v
    if 'mon' not in cols or 'ten' not in cols:
        return None
    hdr = ai.get('hdr_rows')
    hdr = int(hdr) if isinstance(hdr, int) else 1
    if not (1 <= hdr <= 6) or hdr >= len(t.rows):
        return None
    hrow = t.rows[hdr - 1]
    body = [r for r in t.rows[hdr:] if not core._is_sep(r) and not core._is_hdr_dup(r, hrow, cols)]
    if len(body) < 5:
        return None
    rows = core.rows_of(core.Blk(t, dict(cols=cols, hdr=hdr), body))
    if len(rows) < 5:
        return None
    fm = sum(1 for r in rows if r['mon'].strip()) / len(rows)
    ft = sum(1 for r in rows if str(r['thu']).isdigit()) / len(rows)
    fn = sum(1 for r in rows if r['ten'].strip()) / len(rows)
    if fm < .5 or ft < .45 or fn < .3:
        log.info('aimap: sơ đồ bị kiểm chứng loại (môn %.2f, thứ %.2f, tên %.2f)', fm, ft, fn)
        return None
    sy = core.school_year(doc)
    week = ai.get('week')
    week = int(week) if isinstance(week, int) and 1 <= week <= 40 else None
    rng = None
    if not week:
        week, rng = core.table_title_week(t, hdr, sy)
    entry = dict(cols)
    entry['hdr'] = hdr
    if week:
        entry['week'] = week
    if rng:
        entry['range'] = list(rng)
    return t, entry


def learn(data_dir, doc):
    """Nhờ AI hiểu mẫu và lưu sơ đồ đã kiểm chứng vào colmap.json của người dùng.
    Trả về True nếu đã học được một mẫu mới đạt kiểm chứng."""
    if not gemini.key():
        return False
    digest, cands = _digest(doc)
    if not digest:
        return False
    ai = _ask(digest)
    if not ai:
        return False
    got = _verify(doc, cands, ai)
    if not got:
        return False
    t, entry = got
    sig = core.head_sig(t)
    pth = os.path.join(data_dir, 'colmap.json')
    try:
        with open(pth, encoding='utf-8') as f:
            d = json.load(f)
    except Exception:
        d = {}
    d[sig] = entry
    tmp = pth + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False)
    os.replace(tmp, pth)
    core.CMAP.set(d)
    log.info('aimap: đã học mẫu %s -> %s', sig[:40], entry)
    return True
