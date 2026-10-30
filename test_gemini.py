"""Chọn mô hình Gemini tập trung (gemini.py): GEMINI_MODEL trước, 404 -> thử mô hình kế tiếp, nhớ mô hình chạy được;
400/403 -> báo mã AI không hợp lệ; hết mô hình -> báo rõ; mỗi lần tạo vẫn chỉ 1 lượt AI."""
import os, sys, json, time, tempfile, threading, subprocess, atexit, requests, urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import gemini, phieu, nhanxet, chatbot
SEEN = []; CFG = {'ok': {'gemini-3-flash-preview'}, 'code': 404}
class G(BaseHTTPRequestHandler):
    def log_message(s, *a): pass
    def do_POST(s):
        body = json.loads(s.rfile.read(int(s.headers['Content-Length']))); m = s.path.split('/models/')[1].split(':')[0]; SEEN.append(m)
        if m not in CFG['ok']:
            b = json.dumps({'error': {'code': CFG['code'], 'message': f'models/{m} is no longer available to new users'}}).encode()
            s.send_response(CFG['code']); s.send_header('Content-Length', str(len(b))); s.end_headers(); s.wfile.write(b); return
        sch = body.get('generationConfig', {}).get('responseSchema')
        if sch and sch.get('type') == 'ARRAY':   # nhận xét
            rows = json.loads(body['contents'][0]['parts'][0]['text'].split('Danh sách:\n', 1)[1]); keys = list(sch['items']['properties'])
            txt = json.dumps([{'i': r['i'], **{k: f'AI {r["i"]} {k}' for k in keys if k != 'i'}} for r in rows], ensure_ascii=False)
        else: txt = json.dumps({'ok': 1})
        b = json.dumps({'candidates': [{'content': {'parts': [{'text': txt}]}}]}).encode()
        s.send_response(200); s.send_header('Content-Type', 'application/json'); s.send_header('Content-Length', str(len(b))); s.end_headers(); s.wfile.write(b)
gs = ThreadingHTTPServer(('127.0.0.1', 0), G); threading.Thread(target=gs.serve_forever, daemon=True).start()
KEY = 'test-key-' + 'z' * 20
os.environ.update(GEMINI_API_KEY=KEY, GEMINI_API_URL=f'http://127.0.0.1:{gs.server_port}'); os.environ.pop('GEMINI_MODEL', None)
assert gemini.models()[0] == 'gemini-3.6-flash' and gemini.models() == gemini.CHAIN
# 404 -> thử tiếp, nhớ mô hình chạy được
assert phieu.gemini_call('s', 'u', {'type': 'OBJECT'}) == {'ok': 1}
assert SEEN == ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3-flash-preview'], SEEN
SEEN.clear(); assert phieu.gemini_call('s', 'u', {'type': 'OBJECT'}) == {'ok': 1} and SEEN == ['gemini-3-flash-preview'], SEEN
# GEMINI_MODEL đặt -> thử trước (sau khi reset bộ nhớ), không thử trùng
gemini.reset(); os.environ['GEMINI_MODEL'] = 'gemini-2.5-flash'; CFG['ok'] = {'gemini-3.5-flash-lite'}; SEEN.clear()
assert chatbot.ask([], 'xin chào') and SEEN == ['gemini-2.5-flash', 'gemini-3.6-flash', 'gemini-3.5-flash-lite'], SEEN
assert gemini.models() == ['gemini-3.5-flash-lite', 'gemini-2.5-flash', 'gemini-3.6-flash', 'gemini-3-flash-preview']
# mô hình đã nhớ bị gỡ -> tự quay lại chuỗi
CFG['ok'] = {'gemini-3-flash-preview'}; SEEN.clear(); assert phieu.gemini_call('s', 'u', {'type': 'OBJECT'}) == {'ok': 1}
assert SEEN == ['gemini-3.5-flash-lite', 'gemini-2.5-flash', 'gemini-3.6-flash', 'gemini-3-flash-preview'], SEEN
os.environ.pop('GEMINI_MODEL'); gemini.reset()
# hết mô hình -> thông báo rõ ở mọi nơi gọi
CFG['ok'] = set(); SEEN.clear()
for f in (lambda: phieu.gemini_call('s', 'u', {'type': 'OBJECT'}), lambda: nhanxet.gemini_call('s', 'u', {'type': 'OBJECT'}), lambda: chatbot.ask([], 'hi')):
    try: f(); raise AssertionError
    except ValueError as e: assert 'mô hình không khả dụng' in str(e) and 'quản trị' in str(e), e
assert len(SEEN) == 12 and len(set(SEEN)) == 4
import exam, tichhop
for f in (lambda: exam.gemini_call([('text', 'x')], 'toan'), lambda: tichhop.gemini_json('s', 'u')):
    try: f(); raise AssertionError
    except ValueError as e: assert 'mô hình không khả dụng' in str(e), e
# 400/403 -> mã AI không hợp lệ, không thử mô hình khác
for code in (400, 403):
    CFG['code'] = code; SEEN.clear()
    try: phieu.gemini_call('s', 'u', {'type': 'OBJECT'}); raise AssertionError
    except ValueError as e: assert 'GEMINI_API_KEY' in str(e) and 'không hợp lệ' in str(e) and len(SEEN) == 1, (e, SEEN)
    try: exam.gemini_call([('text', 'x')], 'toan'); raise AssertionError
    except ValueError as e: assert 'GEMINI_API_KEY' in str(e)
CFG['code'] = 404
print('gemini.py: mặc định gemini-3.6-flash, GEMINI_MODEL thử trước, 404 -> mô hình kế tiếp (không trùng), nhớ mô hình chạy được, báo lỗi 404/400/403 rõ ràng: OK')

# máy chủ: 1 lần tạo nhận xét = 1 lượt AI dù phải thử 3 mô hình; khoá không lọt vào log
dd = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=dd, SECRET_KEY='t' * 32, AI_DAILY_LIMIT='5',
    TEACHER_PASSWORD_SHA256='feb9a081d082f4d8d31fa3c51cd960922c29d6fb9440d1ac1a00a8c59bde3cdb',
    GUEST_PASSWORD_SHA256='4fcacb897dec6fd766c0538c222a2f58cfa5e10f13d682facac89f81d02e7011')
for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_MODEL'): env.pop(k, None)
p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', '8096'], cwd=HERE, env=env, stdout=open(dd + '/log', 'w'), stderr=subprocess.STDOUT); atexit.register(p.terminate)
B = 'http://127.0.0.1:8096'
for _ in range(100):
    try: requests.get(B, timeout=1); break
    except Exception: time.sleep(.2)
s = requests.Session(); s.post(B + '/api/login', data=dict(teacher='Cô Lê Tâm', password='test-teacher-pw')).raise_for_status()
s.headers['Cookie'] = 'gv_session=' + s.cookies.get('gv_session')  # test-only: requests không gửi cookie Secure qua HTTP
CFG['ok'] = {'gemini-3-flash-preview'}; SEEN.clear()
st = {'ki': 'Cuối học kì I', 'lop': '2A', 'subjects': ['tv', 'toan'], 'students': [{'name': n} for n in ('Nguyễn Văn An', 'Trần Thị Bình', 'Lê Minh Châu')]}
g = s.post(B + '/api/nhanxet/generate', data=json.dumps({'state': st}), headers={'Content-Type': 'application/json'}).json()
assert g['mode'] == 'ai' and g['ai_used'] == 1 and SEEN == ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3-flash-preview'], (g.get('mode'), g.get('ai_used'), SEEN)
SEEN.clear(); g = s.post(B + '/api/nhanxet/generate', data=json.dumps({'state': st}), headers={'Content-Type': 'application/json'}).json()
assert g['ai_used'] == 2 and SEEN == ['gemini-3-flash-preview'], SEEN
time.sleep(.5); assert KEY not in open(dd + '/log').read()
p.terminate()
print('Máy chủ: thử mô hình dự phòng vẫn tính 1 lượt AI, lần sau dùng ngay mô hình đã chạy, khoá không vào log: OK')
print('TẤT CẢ OK')
