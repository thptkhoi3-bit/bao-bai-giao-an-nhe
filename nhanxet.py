"""Nhận xét học sinh lớp 2 theo Thông tư 27/2020/TT-BGDĐT.
Môn học & HĐGD: T = Hoàn thành tốt, H = Hoàn thành, C = Chưa hoàn thành.
Phẩm chất chủ yếu, năng lực chung & đặc thù: T = Tốt, D = Đạt, C = Cần cố gắng.
Sinh nhận xét: Gemini (gom nhiều học sinh / 1 lượt, KHÔNG gửi tên học sinh) hoặc ngân hàng câu mẫu (luôn chạy được, mỗi em một câu khác)."""
import os, re, io, csv, json, hashlib, logging, itertools, unicodedata
log = logging.getLogger('nhanxet')
SUBJ = [('tv', 'Tiếng Việt'), ('toan', 'Toán'), ('daoduc', 'Đạo đức'), ('tnxh', 'Tự nhiên và Xã hội'), ('hdtn', 'Hoạt động trải nghiệm'),
        ('gdtc', 'Giáo dục thể chất'), ('amnhac', 'Âm nhạc'), ('mithuat', 'Mĩ thuật'), ('tienganh', 'Tiếng Anh')]
SUBJ_D = dict(SUBJ); DEFAULT_SUBJ = [k for k, _ in SUBJ if k != 'tienganh']
KI = ['Giữa học kì I', 'Cuối học kì I', 'Giữa học kì II', 'Cuối năm học']
MUC_MON = {'T': 'Hoàn thành tốt', 'H': 'Hoàn thành', 'C': 'Chưa hoàn thành'}
MUC_PN = {'T': 'Tốt', 'D': 'Đạt', 'C': 'Cần cố gắng'}
MAX_STUDENTS = 60; BATCH = 15

# ---------------- ngân hàng câu mẫu ----------------
B = {
 'tv': dict(T=['đọc to, rõ ràng, lưu loát', 'viết đúng chính tả, chữ viết đều và đẹp', 'hiểu nội dung bài đọc, trả lời câu hỏi đầy đủ', 'biết đặt câu, viết đoạn văn ngắn rõ ý',
               'kể chuyện tự nhiên, diễn cảm', 'dùng từ phong phú, đúng ngữ cảnh', 'mạnh dạn chia sẻ ý kiến trong giờ học'],
            H=['đọc rõ ràng, ngắt nghỉ đúng chỗ', 'viết đúng mẫu chữ, trình bày sạch sẽ', 'nắm được nội dung chính của bài đọc', 'biết đặt câu theo mẫu',
               'kể lại được câu chuyện theo tranh', 'viết được câu đúng yêu cầu', 'tích cực tham gia hoạt động đọc'],
            N=['luyện đọc thêm để đọc trôi chảy hơn', 'viết cẩn thận hơn, chú ý các lỗi chính tả', 'đọc kĩ câu hỏi trước khi trả lời', 'luyện viết chữ đều nét hơn',
               'mạnh dạn hơn khi trình bày ý kiến', 'tập đặt câu với từ mới học']),
 'toan': dict(T=['tính nhẩm nhanh, chính xác', 'nắm chắc phép cộng, phép trừ trong phạm vi đã học', 'giải toán có lời văn đúng, trình bày rõ ràng', 'nhận biết hình và đo độ dài thành thạo',
                 'vận dụng kiến thức vào bài toán thực tế linh hoạt', 'làm bài cẩn thận, ít sai sót', 'biết nêu nhiều cách làm khác nhau'],
              H=['thực hiện được các phép tính đã học', 'đặt tính và tính đúng', 'biết giải bài toán có lời văn đơn giản', 'nhận biết được các hình đã học',
                 'biết xem đồng hồ, xem lịch', 'hoàn thành các bài tập trên lớp', 'biết đo và ước lượng độ dài'],
              N=['luyện thêm bảng cộng, bảng trừ', 'đọc kĩ đề bài trước khi làm', 'kiểm tra lại kết quả sau khi tính', 'trình bày bài giải cẩn thận hơn',
                 'tập tính nhẩm hằng ngày', 'chú ý đặt tính thẳng cột']),
 'daoduc': dict(T=['biết quý trọng thời gian, giữ lời hứa', 'lễ phép với thầy cô, yêu quý bạn bè', 'biết nhận lỗi và sửa lỗi', 'tích cực làm việc nhà vừa sức',
                   'thực hiện tốt các chuẩn mực hành vi đã học', 'biết bày tỏ ý kiến đúng mực'],
                H=['thực hiện được các hành vi đạo đức đã học', 'biết chào hỏi lễ phép', 'biết giữ gìn đồ dùng cá nhân', 'biết chia sẻ với bạn bè',
                   'nêu được việc nên làm và không nên làm', 'biết giữ trật tự nơi công cộng'],
                N=['chủ động thực hiện điều đã học hằng ngày', 'giữ gìn đồ dùng cẩn thận hơn', 'mạnh dạn nhận lỗi khi làm sai', 'quan tâm, giúp đỡ bạn nhiều hơn']),
 'tnxh': dict(T=['quan sát kĩ, nêu được nhiều điều thú vị', 'hiểu biết tốt về gia đình, trường học, cộng đồng', 'biết giữ an toàn khi vui chơi, đi đường', 'yêu thiên nhiên, biết bảo vệ cây xanh, con vật',
                 'mạnh dạn chia sẻ hiểu biết với bạn', 'biết giữ gìn vệ sinh cơ thể và môi trường'],
              H=['nhận biết được các sự vật, hiện tượng đã học', 'biết một số việc giữ vệ sinh cá nhân', 'kể được các thành viên trong gia đình', 'biết quan sát tranh và nêu nội dung',
                 'nêu được cách giữ an toàn đơn giản', 'biết chăm sóc cây xanh'],
              N=['quan sát kĩ hơn khi học bài', 'mạnh dạn nêu câu hỏi khi chưa hiểu', 'ghi nhớ các việc giữ an toàn', 'tham gia thảo luận nhóm tích cực hơn']),
 'hdtn': dict(T=['tích cực, sôi nổi trong các hoạt động trải nghiệm', 'biết hợp tác tốt với bạn khi làm việc nhóm', 'tự tin thể hiện bản thân trước lớp', 'tham gia nhiệt tình hoạt động tập thể',
                 'biết tự phục vụ bản thân gọn gàng', 'có nhiều ý tưởng hay trong hoạt động'],
              H=['tham gia đầy đủ các hoạt động', 'biết làm việc cùng bạn trong nhóm', 'thực hiện được các nhiệm vụ được giao', 'biết giữ gìn trường lớp sạch đẹp',
                 'biết chia sẻ cảm xúc của mình', 'tham gia vui vẻ các trò chơi tập thể'],
              N=['mạnh dạn hơn khi tham gia hoạt động', 'chủ động phối hợp cùng bạn', 'tự tin chia sẻ trước lớp', 'hoàn thành nhiệm vụ đúng thời gian']),
 'gdtc': dict(T=['thực hiện động tác đúng, đẹp', 'tích cực tập luyện, có sức bền tốt', 'chơi trò chơi vận động khéo léo', 'giữ đúng đội hình, đội ngũ',
                 'chăm chỉ rèn luyện thân thể', 'biết giúp bạn tập luyện'],
              H=['thực hiện được các động tác đã học', 'tham gia đầy đủ giờ tập', 'biết chơi các trò chơi vận động', 'tập luyện đúng hướng dẫn',
                 'biết khởi động trước khi tập', 'giữ trật tự khi tập luyện'],
              N=['tập trung hơn khi tập luyện', 'chăm tập thêm ở nhà', 'chú ý thực hiện động tác cho đúng', 'mạnh dạn tham gia trò chơi']),
 'amnhac': dict(T=['hát đúng giai điệu, thuộc lời và diễn cảm', 'vỗ tay, gõ đệm đúng nhịp', 'biểu diễn tự tin trước lớp', 'cảm nhận âm nhạc tốt',
                   'vận động phụ hoạ nhịp nhàng', 'yêu thích và tích cực trong giờ học'],
                H=['hát đúng giai điệu bài hát', 'thuộc lời các bài hát đã học', 'biết vỗ tay theo nhịp', 'tham gia hát cùng tập thể',
                   'biết vận động theo bài hát', 'nghe nhạc chăm chú'],
                N=['luyện hát thuộc lời hơn', 'chú ý giữ đúng nhịp', 'mạnh dạn hát trước lớp', 'tập trung hơn trong giờ học']),
 'mithuat': dict(T=['vẽ đẹp, màu sắc tươi sáng, hài hoà', 'sáng tạo trong sản phẩm', 'biết giới thiệu sản phẩm của mình tự tin', 'tạo hình khéo léo, tỉ mỉ',
                    'có nhiều ý tưởng độc đáo', 'trình bày sản phẩm sạch đẹp'],
                 H=['hoàn thành sản phẩm theo yêu cầu', 'biết sử dụng màu sắc phù hợp', 'biết vẽ, xé dán theo chủ đề', 'giữ gìn đồ dùng học tập cẩn thận',
                    'biết nhận xét sản phẩm của bạn', 'tham gia tích cực giờ học'],
                 N=['tô màu cẩn thận hơn', 'hoàn thành sản phẩm đúng thời gian', 'mạnh dạn thể hiện ý tưởng', 'giữ sản phẩm sạch đẹp hơn']),
 'tienganh': dict(T=['phát âm chuẩn, nói tự tin', 'thuộc và dùng đúng nhiều từ vựng', 'nghe hiểu tốt các câu giao tiếp', 'tích cực phát biểu bằng tiếng Anh',
                     'đọc, viết từ đúng', 'hát tiếng Anh vui vẻ, đúng nhịp'],
                  H=['nhận biết được từ vựng đã học', 'biết chào hỏi bằng tiếng Anh', 'nghe hiểu câu lệnh đơn giản', 'tham gia trò chơi ngôn ngữ',
                     'đọc được các từ quen thuộc', 'biết giới thiệu bản thân đơn giản'],
                  N=['luyện phát âm thêm', 'ôn từ vựng thường xuyên', 'mạnh dạn nói tiếng Anh hơn', 'chú ý nghe cô hướng dẫn']),
 'pc': dict(T=['yêu thương, giúp đỡ bạn bè', 'lễ phép, kính trọng thầy cô', 'chăm chỉ, tự giác trong học tập', 'trung thực, thật thà',
               'có trách nhiệm với việc được giao', 'biết giữ gìn trường lớp sạch đẹp', 'yêu quê hương, gia đình'],
            D=['ngoan, lễ phép', 'biết giúp đỡ bạn', 'đi học đều, đúng giờ', 'thật thà trong học tập', 'biết giữ gìn đồ dùng', 'hoàn thành việc được giao'],
            N=['tự giác hơn trong học tập', 'giữ trật tự trong giờ học', 'quan tâm, chia sẻ với bạn nhiều hơn', 'cẩn thận giữ gìn đồ dùng', 'hoàn thành việc được giao đúng hẹn']),
 'nl': dict(T=['tự học tốt, biết tự chuẩn bị đồ dùng', 'giao tiếp mạch lạc, hợp tác tốt với bạn', 'biết tự giải quyết vấn đề trong học tập', 'diễn đạt rõ ràng, tự tin',
               'tính toán nhanh, sáng tạo', 'quan sát tinh, ham tìm hiểu', 'có năng khiếu thẩm mĩ, vận động tốt'],
            D=['biết tự phục vụ bản thân', 'biết hợp tác cùng bạn trong nhóm', 'biết hỏi khi chưa hiểu bài', 'diễn đạt được ý của mình', 'thực hiện được nhiệm vụ học tập', 'biết lắng nghe bạn'],
            N=['tự tin hơn khi trình bày', 'chủ động hơn trong học tập', 'tập trung hơn trong giờ học', 'mạnh dạn hợp tác cùng bạn', 'tự chuẩn bị đồ dùng học tập chu đáo hơn']),
}
LEAD = {'T': ['Em', 'Em luôn', 'Em đã', 'Trong học kì này em'], 'H': ['Em', 'Em đã', 'Em bước đầu', 'Em có tiến bộ:'],
        'C': ['Em có cố gắng', 'Em đã chăm đến lớp', 'Em có tiến bộ ở một số nội dung', 'Em đã tham gia học tập']}
TAIL = {'T': ['Cô khen em!', 'Em hãy phát huy nhé!', 'Tiếp tục cố gắng em nhé!', 'Cô rất vui vì em!', 'Em thật đáng khen!'],
        'H': ['Cố gắng hơn nữa em nhé!', 'Cô tin em sẽ còn tiến bộ!', 'Em cố gắng phát huy nhé!', 'Cô khen em nhé!', 'Em tiếp tục cố gắng nhé!'],
        'C': ['Cô tin em sẽ tiến bộ!', 'Em cố lên nhé!', 'Cô sẽ luôn giúp em!', 'Cô mong em tiến bộ hơn!', 'Em cố gắng thêm nhé!']}
TAIL_KI = {'Cuối năm học': {'T': ['Chúc em vững vàng lên lớp 3!'], 'H': ['Chúc em lên lớp 3 thật vui!'], 'C': ['Hè này em ôn thêm nhé!']}}
MAP_PN = {'T': 'T', 'D': 'H', 'C': 'C'}

def _seed(*a): return int(hashlib.md5('|'.join(map(str, a)).encode()).hexdigest()[:8], 16)

def _cands(key, lv, ki):
    """Mọi câu có thể cho (môn/khía cạnh, mức). Thứ tự cố định."""
    b = B[key]; L = MAP_PN.get(lv, lv) if key in ('pc', 'nl') else lv
    pos = b['T'] if L == 'T' else (b.get('H') or b.get('D'))
    tails = TAIL[L] + TAIL_KI.get(ki, {}).get(L, [])
    out = []
    if L in ('T', 'H'):
        for (a, c), lead, tail in itertools.product(itertools.permutations(pos, 2), LEAD[L], tails):
            s = f'{lead} {a}, {c}.' if not lead.endswith(':') else f'{lead} {a}, {c}.'
            out.append(s + ' ' + tail)
    else:
        for (a, c), lead, tail in itertools.product(itertools.permutations(b['N'], 2), LEAD['C'], tails):
            out.append(f'{lead}; cần {a} và {c}. {tail}')
    return out

def note_sentence(note):
    n = ' '.join(str(note or '').split()).strip(' .;')
    if not n: return ''
    n = n[0].upper() + n[1:]
    if not re.match(r'(?i)(em|con|bạn)\b', n): n = 'Em ' + n[0].lower() + n[1:]
    return n + '.'

def template_comments(students, subjects, ki, salt=''):
    """-> list[dict] mỗi em: {mon_key: câu, 'pc': câu, 'nl': câu}. Trong 1 lớp, cùng môn không em nào trùng câu."""
    res = [dict() for _ in students]
    for key in list(subjects) + ['pc', 'nl']:
        used = set(); cache = {}
        for i, s in enumerate(students):
            lv = (s.get('pn', {}) if key in ('pc', 'nl') else s.get('muc', {})).get(key) or ('D' if key in ('pc', 'nl') else 'H')
            if (key, lv) not in cache: cache[(key, lv)] = _cands(key, lv, ki)
            cs = cache[(key, lv)]; j = _seed(salt, s.get('name', ''), i, key, ki) % len(cs)
            for k in range(len(cs)):
                c = cs[(j + k * 7919) % len(cs)]
                if c not in used: break
            used.add(c); res[i][key] = c
        if key == 'pc':
            for i, s in enumerate(students):
                ns = note_sentence(s.get('note'))
                if ns: res[i]['pc'] = res[i]['pc'].split('. ')[0].rstrip('.') + '. ' + ns + ' ' + res[i]['pc'].split('. ', 1)[-1] if '. ' in res[i]['pc'] else res[i]['pc'] + ' ' + ns
    return res

# ---------------- Gemini ----------------
SYSTEM = ('Bạn là giáo viên chủ nhiệm lớp 2 ở Việt Nam, viết nhận xét học sinh theo Thông tư 27/2020/TT-BGDĐT. '
          'Mỗi nhận xét 1–2 câu ngắn (tối đa 35 từ), tiếng Việt chuẩn, xưng "em" với học sinh, tích cực, động viên, cụ thể theo môn, '
          'đúng với mức đạt: Hoàn thành tốt/Tốt -> khen nổi bật; Hoàn thành/Đạt -> ghi nhận và gợi ý nhẹ; Chưa hoàn thành/Cần cố gắng -> ghi nhận cố gắng, nêu 1–2 việc cần rèn, động viên. '
          'Không nêu tên học sinh, không dùng từ tiêu cực nặng nề. Mỗi học sinh một cách diễn đạt khác nhau, không lặp câu giữa các em. '
          'Nếu có ghi chú của giáo viên thì lồng ghép tự nhiên. Trả về JSON đúng lược đồ.')
def ai_comments(batch, subjects, ki, call):
    """batch: list[(i, student)]. call(system, user, schema) -> JSON. Trả về {i: {key: câu}}."""
    keys = list(subjects) + ['pc', 'nl']
    rows = [dict(i=i, ghi_chu=str(s.get('note') or '')[:200],
                 muc={SUBJ_D[k]: MUC_MON.get((s.get('muc') or {}).get(k) or 'H') for k in subjects} |
                     {'Phẩm chất': MUC_PN.get((s.get('pn') or {}).get('pc') or 'D'), 'Năng lực': MUC_PN.get((s.get('pn') or {}).get('nl') or 'D')}) for i, s in batch]
    user = (f'Kì đánh giá: {ki}. Lớp 2. Viết nhận xét cho từng học sinh (theo số thứ tự i). Khoá JSON: '
            + ', '.join(f'"{k}" = {SUBJ_D.get(k, "Phẩm chất" if k == "pc" else "Năng lực")}' for k in keys)
            + '.\nDanh sách:\n' + json.dumps(rows, ensure_ascii=False))
    schema = {'type': 'ARRAY', 'items': {'type': 'OBJECT', 'properties': {'i': {'type': 'INTEGER'}, **{k: {'type': 'STRING'} for k in keys}}, 'required': ['i'] + keys}}
    out = call(SYSTEM, user, schema); res = {}
    ids = {i for i, _ in batch}
    for x in out if isinstance(out, list) else []:
        if not isinstance(x, dict) or x.get('i') not in ids: continue
        res[x['i']] = {k: ' '.join(str(x.get(k) or '').split())[:400] for k in keys if str(x.get(k) or '').strip()}
    return res

def gemini_call(system, user, schema):
    import urllib.request, urllib.error
    key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not key: raise ValueError('Chưa bật AI')
    import gemini
    body = {'system_instruction': {'parts': [{'text': system}]}, 'contents': [{'role': 'user', 'parts': [{'text': user}]}],
            'generationConfig': {'temperature': 0.8, 'responseMimeType': 'application/json', 'responseSchema': schema}}
    try:
        res = gemini.post(body, 90)
    except urllib.error.HTTPError as e:
        log.warning('nhanxet gemini: HTTP %s', e.code); raise ValueError(gemini.special_msg(e.code) or 'AI đang bận')
    except Exception as e: log.warning('nhanxet gemini: lỗi kết nối %s', type(e).__name__); raise ValueError('Không kết nối được AI')
    try: return json.loads(''.join(p.get('text', '') for p in res['candidates'][0]['content']['parts']))
    except Exception: raise ValueError('AI trả kết quả không đọc được')

def generate(students, subjects, ki, use_ai, take_quota, call=None):
    """take_quota() -> True nếu còn lượt AI (và trừ 1 lượt). Trả về (comments, mode, n_ai_batches)."""
    base = template_comments(students, subjects, ki)
    if not use_ai or not students: return base, 'mau', 0
    call = call or gemini_call; n = 0; got = 0
    for b0 in range(0, len(students), BATCH):
        if not take_quota(): break
        batch = list(enumerate(students))[b0:b0 + BATCH]
        try: r = ai_comments(batch, subjects, ki, call); n += 1
        except ValueError as e: log.warning('nhanxet: AI lỗi, dùng câu mẫu (%s)', e); break
        for i, d in r.items():
            for k, v in d.items(): base[i][k] = v; got += 1
    # câu AI trùng nhau giữa các em -> thay câu mẫu (vẫn khác nhau)
    tpl = template_comments(students, subjects, ki, salt='dup')
    for k in list(subjects) + ['pc', 'nl']:
        seen = set()
        for i in range(len(students)):
            if base[i].get(k) in seen: base[i][k] = tpl[i][k]
            seen.add(base[i].get(k))
    total = len(students) * (len(subjects) + 2)
    return base, ('ai' if got >= total else 'mix' if got else 'mau'), n

# ---------------- đọc danh sách lớp ----------------
def _nfc(s): return unicodedata.normalize('NFC', str(s or '')).strip()
def _looks_name(s):
    s = _nfc(s); w = s.split()
    return 2 <= len(w) <= 6 and len(s) <= 60 and all(re.fullmatch(r"[^\W\d_]+", x) for x in w)
def names_from_rows(rows):
    rows = [[_nfc(c) for c in r] for r in rows if any(_nfc(c) for c in r)]
    if not rows: return []
    hi, ci, cj = None, None, None
    for ri, r in enumerate(rows[:6]):
        low = [c.lower() for c in r]
        for k, c in enumerate(low):
            if re.fullmatch(r'họ(\s*(và\s*)?tên\s*đệm|\s*đệm)?', c) and k + 1 < len(low) and low[k + 1] == 'tên': hi, ci, cj = ri, k, k + 1; break
            if re.search(r'họ\s*(và|&)?\s*tên|tên\s*học\s*sinh|học\s*sinh', c): hi, ci = ri, k; break
        if hi is not None: break
    if hi is None:   # không có tiêu đề: cột có nhiều "tên người" nhất
        w = max(range(max(len(r) for r in rows)), key=lambda k: sum(_looks_name(r[k]) for r in rows if k < len(r)))
        data, ci = rows, w
    else: data = rows[hi + 1:]
    out = []
    for r in data:
        ne = [c for c in r if c]
        nm = (r[ci] if ci < len(r) and r[ci] else (ne[-1] if len(ne) == 1 else '')) + (' ' + r[cj] if cj is not None and cj < len(r) else '')
        nm = ' '.join(re.sub(r'^\s*\d+\s*[.)\-–]\s*', '', nm).split())
        if _looks_name(nm): out.append(nm)
    return out[:MAX_STUDENTS]
def names_from_text(txt):
    rows = [re.split(r'\t|;', l) for l in str(txt or '').splitlines()]
    rows = [[re.sub(r'^\s*\d+\s*[.)\-–]?\s*', '', c) for c in r] for r in rows]
    return names_from_rows(rows)
def names_from_file(path, filename):
    ext = os.path.splitext(filename.lower())[1]
    if ext == '.csv' or ext == '.txt':
        raw = open(path, 'rb').read()
        for enc in ('utf-8-sig', 'utf-16', 'cp1258', 'latin-1'):
            try: txt = raw.decode(enc); break
            except UnicodeDecodeError: continue
        if ext == '.txt': return names_from_text(txt)
        try: dialect = csv.Sniffer().sniff(txt[:2000], delimiters=',;\t')
        except csv.Error: dialect = csv.excel
        return names_from_rows(list(csv.reader(io.StringIO(txt), dialect)))
    if ext == '.xlsx':
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            best = []
            for ws in wb.worksheets[:5]:
                n = names_from_rows([['' if v is None else str(v) for v in r] for r, _ in zip(ws.iter_rows(values_only=True), range(300))])
                if len(n) > len(best): best = n
            return best
        finally: wb.close()
    if ext == '.docx':
        import docx
        d = docx.Document(path); best = []
        for t in d.tables[:5]:
            rows = []
            for r in t.rows:
                seen, cells = set(), []
                for c in r.cells:
                    if id(c._tc) in seen: continue
                    seen.add(id(c._tc)); cells.append(c.text)
                rows.append(cells)
            n = names_from_rows(rows)
            if len(n) > len(best): best = n
        if not best: best = names_from_text('\n'.join(p.text for p in d.paragraphs))
        return best
    raise ValueError('Chỉ nhận tệp .xlsx, .docx, .csv hoặc .txt')

# ---------------- xuất tệp ----------------
def clean_state(st):
    st = st if isinstance(st, dict) else {}
    subs = [k for k in (st.get('subjects') or DEFAULT_SUBJ) if k in SUBJ_D] or DEFAULT_SUBJ
    ki = st.get('ki') if st.get('ki') in KI else KI[0]
    out = []
    for s in (st.get('students') or [])[:MAX_STUDENTS]:
        if not isinstance(s, dict): continue
        nm = ' '.join(str(s.get('name') or '').split())[:80]
        if not nm: continue
        muc = {k: v for k, v in (s.get('muc') or {}).items() if k in SUBJ_D and v in MUC_MON}
        pn = {k: v for k, v in (s.get('pn') or {}).items() if k in ('pc', 'nl') and v in MUC_PN}
        nx = {k: str(v)[:600] for k, v in (s.get('nx') or {}).items() if (k in SUBJ_D or k in ('pc', 'nl')) and isinstance(v, str)}
        out.append(dict(name=nm, note=str(s.get('note') or '')[:200], muc=muc, pn=pn, nx=nx))
    return dict(ki=ki, lop=str(st.get('lop') or '2')[:20], truong=str(st.get('truong') or '')[:120], subjects=subs, students=out)

def export_xlsx(st, path):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Nhận xét'
    cols = [(SUBJ_D[k], k, 'mon') for k in st['subjects']] + [('Phẩm chất', 'pc', 'pn'), ('Năng lực', 'nl', 'pn')]
    ncol = 2 + 2 * len(cols)
    ws.cell(1, 1, f"BẢNG NHẬN XÉT HỌC SINH – {st['ki'].upper()} – LỚP {st['lop']}" + (f" – {st['truong']}" if st['truong'] else '')).font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    th = Side(style='thin', color='999999'); bd = Border(left=th, right=th, top=th, bottom=th); hf = PatternFill('solid', fgColor='DCE6F7')
    for c, t in ((1, 'STT'), (2, 'Họ và tên')):
        ws.cell(2, c, t); ws.merge_cells(start_row=2, start_column=c, end_row=3, end_column=c)
    for j, (lab, k, kind) in enumerate(cols):
        c = 3 + 2 * j; ws.cell(2, c, lab); ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c + 1)
        ws.cell(3, c, 'Mức'); ws.cell(3, c + 1, 'Nhận xét')
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13; ws.column_dimensions[openpyxl.utils.get_column_letter(c + 1)].width = 42
    for r in (2, 3):
        for c in range(1, ncol + 1):
            x = ws.cell(r, c); x.font = Font(bold=True); x.fill = hf; x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); x.border = bd
    ws.column_dimensions['A'].width = 5; ws.column_dimensions['B'].width = 24
    for i, s in enumerate(st['students']):
        r = 4 + i; ws.cell(r, 1, i + 1); ws.cell(r, 2, s['name'])
        for j, (lab, k, kind) in enumerate(cols):
            m = (s['muc'].get(k) or 'H') if kind == 'mon' else (s['pn'].get(k) or 'D')
            ws.cell(r, 3 + 2 * j, (MUC_MON if kind == 'mon' else MUC_PN)[m]); ws.cell(r, 4 + 2 * j, s['nx'].get(k, ''))
        for c in range(1, ncol + 1):
            x = ws.cell(r, c); x.border = bd; x.alignment = Alignment(vertical='top', wrap_text=True, horizontal='center' if c == 1 else 'left')
    ws.freeze_panes = 'C4'; wb.save(path); wb.close()

def export_docx(st, path):
    import docx
    from docx.shared import Pt, Cm
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    d = docx.Document(); sec = d.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = Cm(29.7), Cm(21)
    for a in ('left_margin', 'right_margin'): setattr(sec, a, Cm(1.5))
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    stl = d.styles['Normal']; stl.font.name = 'Times New Roman'; stl.font.size = Pt(12)
    stl.element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', 'Times New Roman')
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"BẢNG NHẬN XÉT HỌC SINH – {st['ki'].upper()}"); r.bold = True; r.font.size = Pt(14)
    p = d.add_paragraph(f"Lớp {st['lop']}" + (f" – {st['truong']}" if st['truong'] else '')); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t = d.add_table(rows=1, cols=5); t.style = 'Table Grid'
    W = [Cm(1.1), Cm(4.2), Cm(12.6), Cm(4.9), Cm(4.9)]
    for c, (w, lab) in enumerate(zip(W, ['STT', 'Họ và tên', 'Môn học và hoạt động giáo dục', 'Phẩm chất', 'Năng lực'])):
        cell = t.rows[0].cells[c]; cell.width = w; cell.text = ''; rr = cell.paragraphs[0].add_run(lab); rr.bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, s in enumerate(st['students']):
        cs = t.add_row().cells
        for c, w in enumerate(W): cs[c].width = w
        cs[0].text = str(i + 1); cs[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER; cs[1].text = s['name']
        cs[2].text = ''
        for j, k in enumerate(st['subjects']):
            pp = cs[2].paragraphs[0] if j == 0 else cs[2].add_paragraph()
            rr = pp.add_run(f"{SUBJ_D[k]} ({MUC_MON[s['muc'].get(k) or 'H']}): "); rr.bold = True; pp.add_run(s['nx'].get(k, ''))
        for c, k in ((3, 'pc'), (4, 'nl')):
            cs[c].text = ''; pp = cs[c].paragraphs[0]; rr = pp.add_run(MUC_PN[s['pn'].get(k) or 'D'] + ': '); rr.bold = True; pp.add_run(s['nx'].get(k, ''))
    d.save(path)
