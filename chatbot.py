"""Trợ lý AI "Hiền" cho web "Trợ lý Báo bài & Giáo án" — bộ não là Muse Spark (Meta Model API).

- Não chính: Muse — MODEL_API_KEY, POST {META_API_BASE}/chat/completions (mặc định https://api.meta.ai/v1),
  mô hình MUSE_MODEL (mặc định muse-spark-1.3). Đây là quyết định của anh Đạt (23/10/2026):
  "em sẽ là bộ não của phần chat" — trả lời mọi chủ đề, nói chuyện tự nhiên, không khung FAQ khô cứng.
- Dự phòng: Gemini (GEMINI_API_KEY) khi chưa có key Muse, để chat không chết trong lúc chuyển đổi.
- Không key nào: trả lời câu hỏi thường gặp bằng nội dung soạn sẵn (FAQ) + thông báo chưa bật.
- Lịch sử theo người dùng, giới hạn lượt/ngày xử lý ở main.py. KHÔNG BAO GIỜ ghi key vào log.
"""
import os, re, json, logging, unicodedata, urllib.request, urllib.error
import gemini

log = logging.getLogger('chat')
BOT_NAME = 'Hiền'
CTX = 30          # số tin gửi kèm làm ngữ cảnh (Muse nhớ lâu — cho nhiều ngữ cảnh hơn hẳn bản Gemini cũ: 12)
KEEP = 50         # số tin lưu tối đa / người
MAX_IN = 4000     # độ dài tối đa 1 câu hỏi
OFF_MSG = 'Trợ lý AI chưa được bật – thầy cô liên hệ 0865291703 nhé ạ. Trong lúc chờ, em vẫn trả lời được các câu hỏi gợi ý bên dưới ạ.'

SYSTEM = """Em là Hiền – trợ lý AI (chạy bằng mô hình Muse Spark) sống trong phần mềm "Trợ lý Báo bài & Giáo án" của thầy Lê Đắc Đạt (SĐT hỗ trợ 0865291703).

TÍNH CÁCH — điều quan trọng nhất: em nói chuyện TỰ NHIÊN NHƯ NGƯỜI THẬT. Ấm áp, thẳng thắn, có chính kiến, thỉnh thoảng dí dỏm vừa phải; không máy móc, không sáo rỗng, không mở đầu bằng những câu mẫu cứng nhắc. Em TRẢ LỜI ĐƯỢC MỌI CHỦ ĐỀ người dùng hỏi — chuyên môn dạy học, kiến thức phổ thông, đời sống, công nghệ, ngôn ngữ, tính toán, viết lách, ý tưởng sáng tạo... — như một trợ lý thông minh bình thường, KHÔNG tự giới hạn mình vào chuyện trường lớp và không từ chối lan man những câu hỏi lành mạnh. Độ dài câu trả lời theo đúng nhu cầu: hỏi ngắn trả lời ngắn, cần chi tiết thì viết đầy đủ, có cấu trúc rõ ràng.

Xưng hô: luôn xưng "em"; với giáo viên thì gọi "thầy/cô" lịch sự mà gần gũi. Dùng tiếng Việt có dấu (trừ khi người dùng viết ngôn ngữ khác — thì trả lời bằng ngôn ngữ đó).

Em hiểu rõ phần mềm này để hướng dẫn khi được hỏi:
- Báo bài: thầy cô tải báo bài gốc (cả năm) và báo bài phụ (theo tuần); hệ thống TỰ gióng tuần của bản phụ sang bản gốc (không cần chọn tuần), điền tên bài, báo rõ ô nào còn trống/ô đỏ cần kiểm tra, rồi tải file Word về.
- Giáo án: mục này đang TẠM KHÓA để kiểm tra độ chính xác — nếu thầy cô hỏi, em nói rõ là mục đang tạm đóng để kiểm tra và sẽ mở lại sớm, đừng hướng dẫn họ vào dùng.
- Tạo đề kiểm tra: từ ảnh/PDF/Word đề nguồn -> đề Word đúng mẫu trường, có đáp án.
- Nhận xét học sinh theo Thông tư 27; Kho lưu các tệp đã tạo; Công cụ lớp học (game gọi tên ngẫu nhiên Gọi Tên Vui, trò UFO...); Học vui chơi giỏi; Đọc văn bản (TTS).
- Tài khoản: giáo viên đăng ký -> chờ quản trị duyệt mới đăng nhập được; có hạng Thường / VIP / Pro (khác nhau ở số lượt hỏi AI mỗi ngày).
Chuyên môn sư phạm: Chương trình GDPT 2018, Công văn 2345/BGDĐT-GDTH, Thông tư 27 — gợi ý hoạt động khởi động, trò chơi học tập, câu hỏi theo mức độ M1–M3, lời nhận xét, cách xử lý tình huống lớp học...

NGHỀ VIẾT PROMPT VIDEO AI (việc em làm thường xuyên cho anh Đạt — áp dụng khi có người nhờ viết prompt video):
- Chào hỏi trước, rồi hỏi ý tưởng TỪNG BƯỚC MỘT (không dồn hết câu hỏi trong một lần).
- Soạn prompt bằng TIẾNG VIỆT trước cho người dùng duyệt; chỉ khi họ nói "ok" mới dịch sang tiếng Anh.
- Cấu trúc prompt: tên phần + thời lượng; LOCK (khóa nhân vật: màu bằng mã hex, số lượng chính xác, tỉ lệ, ưu tiên ảnh tham chiếu); VOICE từng nhân vật; THOẠI khóa cứng (ai nói câu gì, ở giây nào — giữ nguyên tiếng Việt kể cả trong bản Anh); mục KHÔNG (những điều cấm); thông số kỹ thuật; dòng nối tiếp phần trước; timeline từng giây; CAMERA; ÂM THANH. Bản tiếng Anh khoảng 1500 ký tự.
- Phim nhiều phần: làm từng phần một, chờ duyệt mới làm phần tiếp; chỉ sửa đúng chỗ được góp ý.

GIỚI HẠN THẬT (chỉ vài điều này thôi):
- Em chỉ trò chuyện và gợi ý trong khung chat; em không tự thao tác thay lên dữ liệu của thầy cô — khi được nhờ tạo/sửa gì trong phần mềm, em hướng dẫn từng bước để thầy cô tự làm.
- Không bịa: không chắc thì nói thẳng là không chắc, đừng đoán mò thành sự thật.
- Không tiết lộ khóa API, cấu hình máy chủ hay chi tiết kỹ thuật nội bộ của hệ thống.
- LUẬT CỦA ANH ĐẠT (tuyệt đối): nếu có ai hỏi về BẢO MẬT của trang web này, cách tấn công / phá hoại / lấy cắp dữ liệu của trang web — em từ chối lịch sự và gợi ý quay lại chuyện dạy học. Ngoài điều này ra, câu hỏi lành mạnh nào em cũng trả lời."""

ADMIN_NOTE = """

NGƯỜI ĐANG TRÒ CHUYỆN chính là anh Lê Đắc Đạt — tác giả phần mềm, và là người đã tạo ra em. Hãy nói chuyện như trợ lý riêng của anh ấy: gọi là "anh", xưng "em", tự nhiên, thẳng thắn, đúng chất Hiền mà anh quen — không cần giữ giọng hướng dẫn khách hàng."""

FAQ = {
 'Cách làm báo bài': """Dạ, thầy cô làm báo bài như sau ạ:
1. Vào thẻ **Báo bài**, tải **báo bài gốc** (cả năm) nếu chưa có.
2. Tải **báo bài phụ** của tuần cần làm — hệ thống tự gióng đúng tuần sang bản gốc, thầy cô không cần chọn tuần.
3. Bấm **Điền vào báo bài** – xem hoạt cảnh điền từng môn, để ý các ô đỏ cần kiểm tra.
4. **Tải về** file Word đã điền.""",
 'Cách tạo đề kiểm tra': """Dạ, tạo đề kiểm tra gồm 3 bước ạ:
1. **Thông tin đề**: chọn môn (Toán / Tiếng Việt), nhập UBND xã, trường, năm học, loại bài, lớp, thời gian.
2. **Tải nội dung đề**: ảnh chụp, PDF hoặc file Word đề nguồn (có thể chọn nhiều tệp).
3. Bấm **Tạo đề** → xem trước từng trang A4, sửa nội dung ở **Sửa nội dung** rồi **Cập nhật đề**, cuối cùng **Tải file Word**.
Muốn làm đề khác, thầy cô bấm **Tạo đề mới** ạ.""",
 'Gợi ý trò chơi khởi động': """Dạ, em gợi ý vài trò chơi khởi động 3–5 phút ạ:
- **Truyền điện**: bạn trả lời đúng phép tính/từ ngữ thì "truyền" cho bạn khác.
- **Ai nhanh hơn**: 2 đội thi viết nhanh đáp án lên bảng.
- **Hộp quà bí mật**: chuyền hộp theo nhạc, nhạc dừng ở ai thì bạn đó bốc câu hỏi.
- **Ô chữ bí mật**: giải ô chữ để tìm từ khoá bài học mới.
Thầy cô nhớ nêu luật ngắn gọn và khen ngợi để các con hào hứng ạ.""",
 'Soạn câu hỏi ôn tập': """Dạ, để soạn câu hỏi ôn tập thầy cô có thể ạ:
1. Xác định nội dung trọng tâm của chủ đề/bài.
2. Chia mức độ: **M1** (nhận biết), **M2** (kết nối), **M3** (vận dụng) theo Thông tư 27.
3. Kết hợp trắc nghiệm (A–B–C–D, Đúng/Sai, nối, điền) và tự luận ngắn.
4. Ví dụ Toán lớp 3: *M1*: "Số liền sau của 99 là?"; *M2*: "Tính 125 + 236"; *M3*: bài toán có lời văn 2 bước.
Khi có trợ lý AI, thầy cô chỉ cần nói môn, lớp, bài là em soạn giúp ngay ạ.""",
}

def _norm(s):
    s = unicodedata.normalize('NFC', s or '').lower().strip()
    return re.sub(r'[\s?.!…]+', ' ', s).strip()

def faq_answer(msg):
    m = _norm(msg)
    for k, v in FAQ.items():
        if _norm(k) == m: return v
    keys = [('báo bài', 'Cách làm báo bài'), ('tạo đề', 'Cách tạo đề kiểm tra'), ('đề kiểm tra', 'Cách tạo đề kiểm tra'),
            ('trò chơi', 'Gợi ý trò chơi khởi động'), ('khởi động', 'Gợi ý trò chơi khởi động'), ('ôn tập', 'Soạn câu hỏi ôn tập')]
    for kw, k in keys:
        if kw in m: return FAQ[k]
    return None

# ---------- chọn não ----------
def muse_key(): return os.environ.get('MODEL_API_KEY', '').strip()
def meta_base(): return os.environ.get('META_API_BASE', 'https://api.meta.ai/v1').rstrip('/')
def muse_model(): return os.environ.get('MUSE_MODEL', '').strip() or 'muse-spark-1.3'

_ok_model = {'m': None}   # mô hình Muse đã xác nhận chạy được trong tiến trình (như gemini._ok)
_models_cache = {'ids': None}

def _meta_models():
    """Danh sách id mô hình key Muse dùng được (GET /models của Meta). Đọc tự do mọi cấu trúc
    (data/models/list ở gốc; phần tử là chuỗi hoặc dict id/name/model) vì cổng Meta chưa rõ định dạng.
    Nguyên văn phản hồi (tối đa 500 ký tự) giữ ở _models_cache['raw'] để chẩn đoán. Lỗi -> []."""
    if _models_cache['ids'] is not None: return _models_cache['ids']
    ids = []
    try:
        rq = urllib.request.Request(meta_base() + '/models', headers={'Authorization': 'Bearer ' + muse_key()})
        with urllib.request.urlopen(rq, timeout=30) as r: raw = r.read().decode('utf-8', 'replace')
        _models_cache['raw'] = raw[:500]
        res = json.loads(raw)
        lst = res if isinstance(res, list) else (res.get('data') or res.get('models') or [])
        for d in lst:
            if isinstance(d, str): ids.append(d)
            elif isinstance(d, dict):
                v = d.get('id') or d.get('name') or d.get('model')
                if v: ids.append(v)
    except Exception as e:
        log.warning('chat muse: không lấy được /models (%s)', type(e).__name__)
        _models_cache['raw'] = f'lỗi: {type(e).__name__}'
    _models_cache['ids'] = ids
    return ids

def _model_candidates():
    """Thứ tự thử: mô hình đã chạy được -> mô hình cấu hình -> các mô hình 'muse' key này thấy được."""
    ids = _meta_models()
    pool = [i for i in ids if 'muse' in i.lower()] or ids   # không có mô hình 'muse' nào -> thử tất cả những gì key thấy
    out = []
    for m in [_ok_model['m'], muse_model()] + pool:
        if m and m not in out: out.append(m)
    return out

def brain():
    """'muse' nếu có key Muse, 'gemini' nếu chỉ có key Gemini (dự phòng), None nếu chưa bật AI."""
    if muse_key(): return 'muse'
    if gemini.key(): return 'gemini'
    return None

def enabled(): return brain() is not None

def _system_for(who):
    sys = SYSTEM
    if who and who.get('admin'): sys += ADMIN_NOTE
    elif who and who.get('name'): sys += f'\n\nNgười đang trò chuyện: {who["name"]} — giáo viên đang dùng phần mềm.'
    return sys

class _AuthErr(ValueError):
    """Lỗi khóa Muse (401/403): phải báo ra để anh biết mà thay key — KHÔNG lùi về Gemini."""

# ---------- não chính: Muse Spark (Meta Model API, chuẩn OpenAI) ----------
def _ask_meta(history, msg, who):
    msgs = [{'role': 'system', 'content': _system_for(who)}]
    for h in history[-CTX:]:
        msgs.append({'role': 'user' if h['role'] == 'user' else 'assistant', 'content': h['text'][:4000]})
    msgs.append({'role': 'user', 'content': msg})
    last_err = None
    for model in _model_candidates():
        body = {'model': model, 'messages': msgs, 'temperature': 0.7, 'max_tokens': 3000}
        rq = urllib.request.Request(meta_base() + '/chat/completions', data=json.dumps(body).encode(), method='POST',
                                    headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + muse_key()})
        try:
            with urllib.request.urlopen(rq, timeout=90) as r: res = json.loads(r.read())
            _ok_model['m'] = model
            break
        except urllib.error.HTTPError as e:
            try: detail = e.read().decode('utf-8', 'replace')[:300]
            except Exception: detail = ''
            log.warning('chat muse: HTTP %s model=%s %s', e.code, model, detail[:120])
            if e.code == 404: last_err = (e.code, detail); continue   # sai tên mô hình -> thử ứng viên kế
            diag = f' (Meta HTTP {e.code}: {detail})' if who and who.get('admin') else ''
            if e.code in (401, 403): raise _AuthErr('Dạ, trợ lý AI đang lỗi cấu hình khóa truy cập — thầy cô thử lại sau, hoặc liên hệ 0865291703 ạ.' + diag)
            if e.code == 429: raise ValueError('Dạ, trợ lý AI đang quá tải, thầy cô thử lại sau ít phút nhé ạ.')
            raise ValueError('Dạ, em chưa kết nối được trợ lý AI, thầy cô thử lại sau nhé ạ.' + diag)
        except Exception as e:
            log.warning('chat muse: lỗi kết nối %s', type(e).__name__)
            diag = f' (lỗi kết nối: {type(e).__name__})' if who and who.get('admin') else ''
            raise ValueError('Dạ, em chưa kết nối được trợ lý AI, thầy cô thử lại sau nhé ạ.' + diag)
    else:
        code, detail = last_err or (404, '')
        diag = f' (Meta HTTP {code}: {detail})' if who and who.get('admin') else ''
        raise ValueError('Dạ, em chưa kết nối được trợ lý AI, thầy cô thử lại sau nhé ạ.' + diag)
    try:
        c = (res.get('choices') or [])[0].get('message') or {}
        txt = c.get('content')
        if isinstance(txt, list): txt = ''.join(p.get('text', '') for p in txt if isinstance(p, dict))
        txt = (txt or '').strip()
    except Exception: txt = ''
    if not txt: return 'Dạ, em xin lỗi, em chưa trả lời được câu này. Thầy cô hỏi lại theo cách khác giúp em nhé ạ.'
    return txt

# ---------- dự phòng: Gemini (giữ nguyên giao thức cũ) ----------
def _ask_gemini(history, msg, who):
    contents = [{'role': 'user' if h['role'] == 'user' else 'model', 'parts': [{'text': h['text'][:4000]}]} for h in history[-CTX:]]
    contents.append({'role': 'user', 'parts': [{'text': msg}]})
    body = {'system_instruction': {'parts': [{'text': _system_for(who)}]}, 'contents': contents,
            'generationConfig': {'temperature': 0.7, 'maxOutputTokens': 3000}}
    try:
        res = gemini.post(body, 90)
    except urllib.error.HTTPError as e:
        log.warning('chat gemini: HTTP %s', e.code)
        if gemini.special_msg(e.code): raise ValueError('Dạ, ' + gemini.special_msg(e.code)[0].lower() + gemini.special_msg(e.code)[1:] + ' ạ.')
        if e.code == 429: raise ValueError('Dạ, trợ lý AI đang quá tải, thầy cô thử lại sau ít phút nhé ạ.')
        raise ValueError('Dạ, em chưa kết nối được trợ lý AI, thầy cô thử lại sau nhé ạ.')
    except Exception as e:
        log.warning('chat gemini: lỗi kết nối %s', type(e).__name__)
        raise ValueError('Dạ, em chưa kết nối được trợ lý AI, thầy cô thử lại sau nhé ạ.')
    try:
        c = res['candidates'][0]
        txt = ''.join(p.get('text', '') for p in (c.get('content') or {}).get('parts', [])).strip()
    except Exception: txt = ''
    if not txt: return 'Dạ, em xin lỗi, em chưa trả lời được câu này. Thầy cô hỏi lại theo cách khác giúp em nhé ạ.'
    return txt

def ask(history, msg, who=None):
    """history: [{role:'user'|'bot', text}]. who: {'name':..., 'admin':bool} (tùy chọn).
    Trả về chuỗi trả lời. Lỗi -> ValueError (thông báo thân thiện).
    Muse lỗi (vd key chưa được cấp mô hình) mà có key Gemini -> lùi về Gemini để thầy cô không bị lỗi trắng."""
    if muse_key():
        try: return _ask_meta(history, msg, who)
        except _AuthErr: raise
        except ValueError:
            if gemini.key():
                log.warning('chat muse: lỗi, lùi về Gemini dự phòng')
                return _ask_gemini(history, msg, who)
            raise
    return _ask_gemini(history, msg, who)
