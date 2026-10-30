"""AI gỡ rối ghép GIÁO ÁN sau các lượt quy tắc của core.assign().

Triết lý (thầy Đạt chốt): AI hiểu — quy tắc chép. AI chỉ được chạm vào phần còn
tồn đọng: bài phụ chưa dùng, TIẾT còn trống, và các cặp mà quy tắc tự thấy "chưa
chắc" (core.uncertain: điểm thấp / ghép theo thứ tự / ngang ngửa nhiều TIẾT).
AI đề xuất cặp bài -> TIẾT; máy chủ kiểm chứng cứng mới áp dụng:
- có BẰNG CHỨNG tên bài với cột tên bài của TIẾT (core.title_evidence: điểm >= THRESH
  hoặc trùng số bài) — cặp lệch hẳn bị loại, TIẾT để trống cho GV gán tay,
- cùng họ môn (bài không ghi môn thì linh hoạt),
- không vượt sức chứa của TIẾT, mỗi bài đúng một chỗ,
- không bao giờ đụng cặp quy tắc đã khớp chắc,
- tổng số bài được đặt không được giảm so với quy tắc (giảm là hủy toàn bộ).
Cặp do AI đặt mang cờ 'ai' để báo cáo gắn nhãn "AI ghép" cho thầy cô kiểm tra.
Không có khóa Gemini / AI lỗi / trả lời không hợp lệ -> trả None, giữ nguyên
kết quả quy tắc; khâu tạo giáo án không bao giờ hỏng vì AI.
"""
import json, logging, urllib.error, urllib.request
import core, gemini


def _post(body, timeout=40):
    """Gọi Gemini đi hết chuỗi mô hình: gemini.post chỉ chuyển mô hình khi 404, còn khâu này
    cần đi tiếp cả khi 503/429 (mô hình preview quá tải) để tới được mô hình ổn định."""
    last = None
    for m in gemini.models():
        rq = urllib.request.Request(f'{gemini.base_url()}/models/{m}:generateContent',
                                    data=json.dumps(body).encode(), method='POST',
                                    headers={'Content-Type': 'application/json', 'x-goog-api-key': gemini.key()})
        try:
            with urllib.request.urlopen(rq, timeout=timeout) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            last = e
            if e.code in (404, 429, 500, 502, 503):
                log.warning('aiassign: mô hình %s lỗi HTTP %s – thử mô hình kế', m, e.code)
                continue
            raise
    raise last

log = logging.getLogger('aiassign')
LAST = {'note': None}   # mã chẩn đoán lần chạy gần nhất (đưa vào báo cáo build để truy vết)

PROMPT = """Bạn là trợ lý ghép giáo án tiểu học Việt Nam. Dưới đây là các TIẾT học còn đang thiếu bài hoặc đã ghép nhưng chưa chắc chắn, và các bài dạy (từ giáo án phụ) đang chờ xếp vào TIẾT.

Hãy xếp mỗi bài vào đúng một TIẾT phù hợp nhất, theo các nguyên tắc:
- Ưu tiên tuyệt đối: bài và TIẾT phải cùng môn học. Tên môn ở giáo án phụ có thể viết khác báo bài, ví dụ "TV (Đọc)", "TV (Viết)" thuộc Tiếng Việt; "SHL", "Sinh hoạt lớp" thuộc HĐTN; "Đọc TV" thuộc Thư viện.
- Khớp theo NGHĨA của tên bài: tên bài trong báo bài là tên rút gọn, tiêu đề ở phụ thường dài hơn (có số bài, phần sau dấu hai chấm...). Ví dụ báo bài "Ôn tập giữa học kì 1" khớp với phụ "Bài 18: Ôn tập giữa học kì I".
- Một bài chỉ vào một TIẾT; số bài xếp vào một TIẾT không vượt "suc_chua" của TIẾT đó.
- TIẾT có "bai_bao_bai" (đã ghi tên bài theo báo bài gốc) là TIẾT CHÍNH: ưu tiên lấp đầy các TIẾT này trước. Sau khi xếp xong, MỌI TIẾT chính đều phải có ít nhất một bài — nếu dời một bài ra khỏi TIẾT chính thì phải xếp một bài phù hợp khác vào thay.
- "dang_o" của mỗi bài cho biết bài đang nằm ở TIẾT nào (null = chưa được xếp); "dang_co" của mỗi TIẾT liệt kê các bài đang ở đó mà bạn được quyền xếp lại. Bài không có trong danh sách BÀI thì không được đụng vào.
- Bài nào không có TIẾT phù hợp thì bỏ qua, đừng gán bừa. Chỉ xếp khi tên bài THẬT SỰ khớp với "bai_bao_bai" của TIẾT (trùng số bài hoặc cùng chủ đề rõ ràng); thà để trống còn hơn xếp một bài khác hẳn vào TIẾT đã có tên bài.

Dữ liệu:
TIẾT (slot = số hiệu TIẾT, thu = thứ, mon = môn, bai_bao_bai = tên bài theo báo bài gốc, suc_chua = số bài tối đa, dang_co = các bài đang ở TIẾT này và được quyền xếp lại):
{slots}

BÀI (bai = số hiệu bài, mon = môn ghi ở giáo án phụ, tieu_de = tiêu đề bài, dang_o = số hiệu TIẾT đang chứa bài này hoặc null):
{lessons}

Chỉ trả DUY NHẤT một JSON đúng khuôn: {{"pairs": [{{"bai": <số hiệu bài>, "slot": <số hiệu TIẾT>}}, ...]}}
Không giải thích thêm."""


def _cap(h):
    return max(1, len(h.get('ten') or []))


def _fam(mon):
    return core._fam_of(mon) if mon else None


def arbitrate(heads, secs, taken):
    """Cho AI gỡ phần tồn đọng sau assign(). Sửa heads/taken tại chỗ khi áp dụng được.
    Trả {'ai': [(hi, li)...], 'new': n, 'moved': m} hoặc None."""
    LAST['note'] = 'start'
    if not gemini.key():
        LAST['note'] = 'no_key'
        return None
    pool = set()
    slot_of = {}
    for hi, h in enumerate(heads):
        for x in h['lessons']:
            slot_of[x['li']] = hi
            if core.uncertain(x):
                pool.add(x['li'])
    for li, l in enumerate(secs):
        if li not in taken and l['title'].strip():
            pool.add(li)
    if not pool:
        LAST['note'] = 'no_pool'
        return None
    rel = [hi for hi, h in enumerate(heads)
           if len(h['lessons']) < _cap(h) or any(x['li'] in pool for x in h['lessons'])]
    if not rel:
        LAST['note'] = 'no_rel'
        return None
    slots = [dict(slot=hi, thu=heads[hi]['thu'], mon=heads[hi]['mon'],
                  bai_bao_bai=heads[hi]['ten'], suc_chua=_cap(heads[hi]),
                  dang_co=[x['li'] for x in heads[hi]['lessons'] if x['li'] in pool]) for hi in rel]
    lessons = [dict(bai=li, mon=secs[li]['subject'], tieu_de=secs[li]['title'][:130],
                    dang_o=slot_of.get(li)) for li in sorted(pool)][:48]
    if not lessons:
        LAST['note'] = 'no_lessons'
        return None
    body = {'contents': [{'parts': [{'text': PROMPT.format(
        slots=json.dumps(slots, ensure_ascii=False),
        lessons=json.dumps(lessons, ensure_ascii=False))}]}],
        'generationConfig': {'temperature': 0, 'responseMimeType': 'application/json'}}
    try:
        res = _post(body)
    except urllib.error.HTTPError as e:
        log.warning('aiassign: HTTP %s', e.code)
        LAST['note'] = f'ai_error:HTTP{e.code}'
        return None
    except Exception as e:
        log.warning('aiassign: %s', e)
        LAST['note'] = 'ai_error:' + type(e).__name__
        return None
    try:
        txt = ''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts'])
        pairs = [(int(p['bai']), int(p['slot'])) for p in json.loads(txt).get('pairs', [])]
    except Exception as e:
        log.warning('aiassign parse: %s', e)
        LAST['note'] = 'parse_error:' + type(e).__name__
        return None
    if not pairs:
        LAST['note'] = 'no_pairs'
        return None
    digest_ids = {l['bai'] for l in lessons}
    return _apply(heads, secs, taken, pairs, pool, slot_of, digest_ids)


def _apply(heads, secs, taken, pairs, pool, slot_of, digest_ids):
    snap = [list(h['lessons']) for h in heads]
    orig_x = {}
    for h in heads:
        for x in h['lessons']:
            if x['li'] in pool:
                orig_x[x['li']] = x
    orig_placed = sum(1 for li in pool if li in slot_of)
    # TIẾT CHÍNH (có tên bài theo báo bài) đang có bài: sau khi AI can thiệp vẫn phải có bài
    named_filled0 = {hi for hi, h in enumerate(heads) if h['ten'] and h['lessons']}
    for h in heads:
        h['lessons'] = [x for x in h['lessons'] if x['li'] not in pool]
    used = {}
    for li, hi in pairs:
        if li not in pool or li not in digest_ids or li in used:
            continue
        if not (0 <= hi < len(heads)):
            continue
        sf = _fam(secs[li]['subject'])
        if sf is not None and sf != _fam(heads[hi]['mon']):
            continue
        if len(heads[hi]['lessons']) >= _cap(heads[hi]):
            continue
        ev, num = core.title_evidence(heads[hi], secs[li]['title'])
        if ev < core.THRESH and not num:
            continue   # cổng bằng chứng tên bài: lệch hẳn với cột tên bài của TIẾT thì không đặt — để trống cho GV gán tay
        heads[hi]['lessons'].append(dict(l=secs[li], score=ev, li=li, hi=hi, ai=True, num=num))
        used[li] = hi
    restored = set()
    for li in sorted(pool):
        if li in used:
            continue
        hi = slot_of.get(li)
        if hi is not None and len(heads[hi]['lessons']) < _cap(heads[hi]):
            heads[hi]['lessons'].append(orig_x[li])
            restored.add(li)
    # Sửa riêng từng TIẾT chính bị làm trống (thay vì hủy toàn bộ): kéo bài gốc của TIẾT đó về,
    # hủy chỉ cặp AI đã dời bài đó đi nơi khác.
    for hi in sorted(named_filled0):
        if heads[hi]['lessons']:
            continue
        li0 = next((li for li in sorted(pool) if slot_of.get(li) == hi), None)
        if li0 is None:
            continue
        if li0 in used:
            cur = used.pop(li0)
            heads[cur]['lessons'] = [y for y in heads[cur]['lessons'] if y['li'] != li0]
        if li0 not in restored and len(heads[hi]['lessons']) < _cap(heads[hi]):
            heads[hi]['lessons'].append(orig_x[li0])
            restored.add(li0)
    if not used and not restored or len(used) + len(restored) < orig_placed or any(not heads[hi]['lessons'] for hi in named_filled0):
        for h, old in zip(heads, snap):
            h['lessons'] = old
        LAST['note'] = 'reverted'
        return None
    for li in pool:
        taken.discard(li)
    taken |= set(used) | restored
    LAST['note'] = f'applied:{len(used)}'
    return {'ai': sorted((hi, li) for li, hi in used.items()),
            'new': sum(1 for li in used if li not in slot_of),
            'moved': sum(1 for li, hi in used.items() if li in slot_of and slot_of[li] != hi)}
