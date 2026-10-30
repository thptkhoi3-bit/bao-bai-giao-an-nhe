"""Gọi Gemini (REST) dùng chung cho chatbot, đề, tích hợp, nhận xét, phiếu.
- Mô hình: GEMINI_MODEL (nếu đặt) trước, rồi chuỗi dự phòng CHAIN.
- HTTP 404 (mô hình không còn/không có cho khoá mới) và lỗi máy chủ 5xx (500/502/503/504, gồm quá tải) -> tự thử mô hình kế tiếp; nhớ mô hình chạy được trong tiến trình.
- KHÔNG ghi khoá ra log. Mỗi lần tạo vẫn chỉ tính 1 lượt AI (việc đếm lượt ở nơi gọi)."""
import os, json, logging, threading, urllib.request, urllib.error
log = logging.getLogger('gemini')
DEFAULT = 'gemini-3.6-flash'
CHAIN = ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3-flash-preview', 'gemini-2.5-flash']
_ok = {'model': None}; _lock = threading.Lock()
MSG_404 = 'Không gọi được AI (mô hình không khả dụng) – thầy cô báo quản trị'
MSG_KEY = 'Mã AI (GEMINI_API_KEY) không hợp lệ – thầy cô báo quản trị'

def key(): return os.environ.get('GEMINI_API_KEY', '').strip()
def base_url(): return os.environ.get('GEMINI_API_URL', 'https://generativelanguage.googleapis.com/v1beta').rstrip('/')
def models():
    env = os.environ.get('GEMINI_MODEL', '').strip(); out = []
    for m in [_ok['model'], env] + CHAIN:
        if m and m not in out: out.append(m)
    return out
def current(): return models()[0]
def reset(): _ok['model'] = None

def special_msg(code):
    """Thông báo tiếng Việt rõ ràng cho lỗi cấu hình (404 mô hình, 400/403 khoá); None nếu không phải."""
    return MSG_404 if code == 404 else MSG_KEY if code in (400, 401, 403) else None

def post(body, timeout=120):
    """Gửi generateContent; trả về JSON phản hồi. Lỗi HTTP 404 và 5xx -> thử mô hình kế tiếp; lỗi khác (429/400/401/403) -> ném HTTPError gốc ngay; hết mô hình -> HTTPError cuối."""
    k = key(); last = None
    for m in models():
        rq = urllib.request.Request(f'{base_url()}/models/{m}:generateContent', data=json.dumps(body).encode(), method='POST',
                                    headers={'Content-Type': 'application/json', 'x-goog-api-key': k})
        try:
            with urllib.request.urlopen(rq, timeout=timeout) as r: res = json.loads(r.read())
        except urllib.error.HTTPError as e:
            if e.code == 404:
                log.warning('gemini: mô hình %s không khả dụng (404) – thử mô hình khác', m); last = e
                with _lock:
                    if _ok['model'] == m: _ok['model'] = None
                continue
            if e.code in (500, 502, 503, 504):
                log.warning('gemini: mô hình %s lỗi máy chủ (%s) – thử mô hình khác', m, e.code); last = e
                continue
            raise
        with _lock:
            if _ok['model'] != m: _ok['model'] = m; log.info('gemini: dùng mô hình %s', m)
        return res
    raise last
