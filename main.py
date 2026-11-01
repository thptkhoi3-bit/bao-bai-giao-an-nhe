import os, copy, shutil, uuid, datetime, hashlib, hmac, base64, json, re, unicodedata, secrets, time, gc, fcntl, sqlite3
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request, Depends
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
import core, storage, threading, logging, collections
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

def _req_env(name, hint):
    """Đọc biến môi trường bắt buộc — thiếu thì fail-fast ngay khi khởi động với thông báo tiếng Việt rõ ràng."""
    v = os.environ.get(name, '').strip()
    if not v:
        raise RuntimeError(
            '[KHỞI ĐỘNG THẤT BẠI] Thiếu biến môi trường bắt buộc: %s.\n%s\n'
            'Xem README.md mục "Biến môi trường bắt buộc" để biết cách tạo và đặt trên Render.' % (name, hint))
    return v

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT = os.environ.get('DATA_DIR', os.path.join(BASE, 'data'))
SEED = os.path.join(BASE, 'seed')
MAU_GIAOAN = os.path.join(SEED, 'mau_giao_an.docx')  # mẫu tham chiếu để TỰ TẠO giáo án mẫu trống (bước 2)
# SECRET_KEY: khóa ký cookie phiên — BẮT BUỘC đặt cố định, nếu không mỗi lần restart là toàn bộ người dùng bị đăng xuất
SECRET = _req_env('SECRET_KEY',
    'Hãy đặt SECRET_KEY là một chuỗi ngẫu nhiên dài, cố định trên Render (tab Environment).')
H = lambda p: hashlib.sha256(p.encode()).hexdigest()
# sha256 của mật khẩu giáo viên/khách — BẮT BUỘC đặt qua biến môi trường, không còn mật khẩu mặc định trong code
TEACHER_PW = _req_env('TEACHER_PASSWORD_SHA256',
    'Tạo bằng: python3 -c "import hashlib;print(hashlib.sha256(b\'mat-khau-moi\').hexdigest())" rồi đặt giá trị vào Render.')
GUEST_PW = _req_env('GUEST_PASSWORD_SHA256',
    'Tạo tương tự TEACHER_PASSWORD_SHA256 rồi đặt giá trị vào Render.')
TEACHERS = []  # 20/10/2026: bo dang nhap giao vien co dinh (Cô Tâm, Cô Thủy, Cô Linh); admin dang nhap an qua o so dien thoai
ADMIN_PHONE = os.environ.get('ADMIN_PHONE', '0865291703').strip()
ADMIN_PW = os.environ.get('ADMIN_PASSWORD_SHA256', '').strip()
SEEDED = 'Anh Lê Đạt'
COOKIE = 'gv_session'
# key phiên của tài khoản quản trị — tách khỏi tên đăng nhập, đổi được qua biến môi trường
ADMIN_KEY = os.environ.get('ADMIN_KEY', 'gv-anh-le-dat').strip() or 'gv-anh-le-dat'
MAX_FILE = 20 * 1024 * 1024   # giới hạn upload: 20 MB / tệp
CONTACT = '0865291703'
ACCOUNTS = os.path.join(DATA_ROOT, 'accounts.json')
_alock = threading.Lock()
os.makedirs(DATA_ROOT, exist_ok=True)
SYNC = storage.from_env(DATA_ROOT)
try: SYNC.pull()
except Exception as e: logging.getLogger('storage').warning('storage: không kéo được dữ liệu lúc khởi động (%s), sẽ thử lại', e)
SYNC.start()
import usage as _usage
USAGE = _usage.Usage(DATA_ROOT, SYNC); USAGE.start()   # nhật ký sử dụng theo ngày (giờ VN) -> usage/YYYY-MM-DD.json

def _warm_lo():
    """Làm nóng LibreOffice lúc khởi động (chạy nền, không chặn server): lần tải Word đầu tiên
    sau mỗi lần deploy/khởi động lại sẽ không phải chờ tạo profile LibreOffice nữa."""
    try:
        import time as _t; _t.sleep(3)
        if not exam.soffice_bin(): return
        import tempfile as _tf
        from docx import Document as _Doc
        with _tf.TemporaryDirectory() as td:
            p = os.path.join(td, 'warm.docx')
            _d = _Doc(); _d.add_paragraph('Khởi động'); _d.save(p)
            exam.render_pages(p, os.path.join(td, 'pages'))
    except Exception as e:
        logging.getLogger('exam').warning('warmup LibreOffice bỏ qua (%s)', type(e).__name__)
threading.Thread(target=_warm_lo, daemon=True, name='lo-warmup').start()

# ---------- mật khẩu: scrypt + muối ----------
def kdf(pw, salt=None):
    salt = salt or secrets.token_bytes(16)
    return 'scrypt$16384$8$1$' + salt.hex() + '$' + hashlib.scrypt(pw.encode(), salt=salt, n=16384, r=8, p=1, dklen=32).hex()
def kdf_ok(pw, stored):
    try:
        _, n, r, p_, salt, h = stored.split('$')
        return hmac.compare_digest(hashlib.scrypt(pw.encode(), salt=bytes.fromhex(salt), n=int(n), r=int(r), p=int(p_), dklen=32).hex(), h)
    except Exception: return False

# ---------- tài khoản đăng ký ----------
def acc_load():
    try:
        with open(ACCOUNTS, encoding='utf-8') as f: d = json.load(f)
        return {k: v for k, v in d.items() if isinstance(v, dict) and v.get('phone')} if isinstance(d, dict) else {}
    except FileNotFoundError: return {}
def acc_save(d):
    tmp = ACCOUNTS + '.' + uuid.uuid4().hex + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tmp, ACCOUNTS); SYNC.schedule()
def norm_phone(p):
    p = re.sub(r'[\s.\-()]', '', p or '')
    if p.startswith('+84'): p = '0' + p[3:]
    elif p.startswith('84') and len(p) == 11: p = '0' + p[2:]
    return p if re.fullmatch(r'0(3|5|7|8|9)\d{8}', p) else None
def today(): return datetime.date.today()
def fmt_d(s):
    try: return datetime.date.fromisoformat(s).strftime('%d/%m/%Y')
    except Exception: return s or ''
def acc_problem(a):
    st = a.get('status')
    if st == 'pending': return f'Tài khoản đang chờ duyệt. Vui lòng liên hệ {CONTACT}.'
    if st == 'rejected': return f'Đăng ký đã bị từ chối. Vui lòng liên hệ {CONTACT}.'
    if st == 'locked': return f'Tài khoản đã bị khóa. Vui lòng liên hệ {CONTACT}.'
    if st != 'approved': return f'Tài khoản không hợp lệ. Vui lòng liên hệ {CONTACT}.'
    if a.get('expires') and a['expires'] < today().isoformat():
        return f'Tài khoản đã hết hạn dùng (đến {fmt_d(a["expires"])}). Vui lòng liên hệ {CONTACT} để gia hạn.'
    return None

# ---------- giới hạn số lần thử theo IP ----------
_rl = collections.defaultdict(collections.deque); _rllock = threading.Lock()
RL = {'login': (int(os.environ.get('RL_LOGIN', '10')), 600), 'register': (int(os.environ.get('RL_REGISTER', '5')), 3600),
      'api': (int(os.environ.get('RL_API', '180')), 60)}  # chống quét/hút API hàng loạt
def client_ip(request):
    xf = request.headers.get('x-forwarded-for', '')
    # X-Forwarded-For: phần tử ĐẦU tiên mới là IP client thật (các proxy nối thêm vào sau)
    return (xf.split(',')[0].strip() if xf else '') or (request.client.host if request.client else '?')
def rl_check(kind, ip):
    lim, win = RL[kind]; q = _rl[(kind, ip)]; now = time.time()
    with _rllock:
        while q and q[0] < now - win: q.popleft()
        if len(q) >= lim: raise HTTPException(429, f'Thử quá nhiều lần, vui lòng đợi {max(1, int((q[0] + win - now) // 60) + 1)} phút rồi thử lại.')
def rl_hit(kind, ip):
    with _rllock: _rl[(kind, ip)].append(time.time())

def slug(name):
    s = unicodedata.normalize('NFD', name.replace('đ', 'd').replace('Đ', 'D'))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()
    return re.sub(r'[^a-z0-9]+', '-', s).strip('-')[:60]

def sign(d):
    raw = base64.urlsafe_b64encode(json.dumps(d).encode()).decode()
    return raw + '.' + hmac.new(SECRET.encode(), raw.encode(), 'sha256').hexdigest()
def unsign(v):
    try:
        raw, sig = v.rsplit('.', 1)
        if not hmac.compare_digest(sig, hmac.new(SECRET.encode(), raw.encode(), 'sha256').hexdigest()): return None
        d = json.loads(base64.urlsafe_b64decode(raw))
        return d if d.get('exp', 0) > time.time() else None
    except Exception: return None

def prune(bk, keep=int(os.environ.get('KEEP_BACKUPS', '10'))):  # giữ N bản sao lưu mới nhất mỗi loại (đỡ nặng kho dữ liệu)
    try: fs = sorted(os.listdir(bk), reverse=True)
    except FileNotFoundError: return
    for pre in ('baobai_', 'giaoan_'):
        for f in [f for f in fs if f.startswith(pre)][keep:]:
            try: os.remove(os.path.join(bk, f))
            except OSError: pass

class U:
    def __init__(s, name, key):
        s.name, s.key = name, key
        s.DATA = os.path.join(DATA_ROOT, 'users', key)
        s.OUT = os.path.join(s.DATA, 'out'); s.BK = os.path.join(s.DATA, 'backup')
        for d in (s.DATA, s.OUT, s.BK): os.makedirs(d, exist_ok=True)
        prune(s.BK)
        s.MASTER = os.path.join(s.DATA, 'baobai_goc.docx'); s.TEMPLATE = os.path.join(s.DATA, 'giaoan_goc.docx')
        if name == SEEDED and not os.path.exists(os.path.join(s.DATA, '.seeded')):
            for src, dst in (('baobai_goc.docx', s.MASTER), ('giaoan_goc.docx', s.TEMPLATE)):
                if os.path.exists(os.path.join(SEED, src)) and not os.path.exists(dst): shutil.copy(os.path.join(SEED, src), dst)
            open(os.path.join(s.DATA, '.seeded'), 'w').close()

_HV_SEEN = {}
def user(request: Request):
    d = unsign(request.cookies.get(COOKIE, ''))
    if not d: raise HTTPException(401, 'Chưa đăng nhập')
    if d['key'].startswith('gv-') and d['key'] != ADMIN_KEY: raise HTTPException(401, 'Phiên đăng nhập đã hết hiệu lực')
    if d['key'].startswith('dk-'):
        a = acc_load().get(d['key'])
        if not a or acc_problem(a): raise HTTPException(401, (acc_problem(a) if a else 'Tài khoản không còn tồn tại'))
        u = U(a['name'], d['key'])
    else:
        u = U(d['name'], d['key'])
    _HV_SEEN[u.key] = time.time()
    return u
def admin(u=Depends(user)):
    if u.key != ADMIN_KEY: raise HTTPException(403, 'Chỉ quản trị viên được dùng chức năng này')
    return u
app = FastAPI(title='Trợ lý Báo bài & Giáo án')
app.mount('/static', StaticFiles(directory=os.path.join(BASE, 'static')), name='static')
DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

def ts(): return datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
async def save_upload(f: UploadFile, path):
    if not f.filename.lower().endswith('.docx'): raise HTTPException(400, f'Tệp {f.filename} không phải .docx')
    with open(path, 'wb') as o: o.write(await read_upload_limited(f, MAX_FILE))
    try: core.docx.Document(path)
    except Exception: os.remove(path); raise HTTPException(400, f'Không mở được tệp {f.filename}')

async def read_upload_limited(f: UploadFile, limit: int):
    """Đọc file upload theo chunk 1MB, từ chối sớm khi vượt giới hạn (chống cạn RAM/OOM trên gói Free).
    Kiểm tra Content-Length trước nếu client gửi, rồi vẫn giới hạn cứng trong lúc đọc."""
    try: cl = int(f.headers.get('content-length') or 0)
    except (ValueError, TypeError): cl = 0
    if cl > limit:
        raise HTTPException(400, f'Tệp {f.filename} quá lớn (tối đa {limit // (1024 * 1024)} MB)')
    chunks = []; total = 0
    while True:
        chunk = await f.read(1024 * 1024)
        if not chunk: break
        total += len(chunk)
        if total > limit:
            raise HTTPException(400, f'Tệp {f.filename} quá lớn (tối đa {limit // (1024 * 1024)} MB)')
        chunks.append(chunk)
    return b''.join(chunks)

@app.post('/api/login')
async def login(request: Request, teacher: str = Form(...), password: str = Form(...), name: str = Form(''), phone: str = Form('')):
    ip = client_ip(request); rl_check('login', ip)
    try: return _login(teacher, password, name, phone)
    except HTTPException as e:
        if e.status_code in (403, 400): rl_hit('login', ip)
        raise
def _wipe_drafts(key):
    shutil.rmtree(os.path.join(DATA_ROOT, 'users', key, 'draft'), ignore_errors=True)

def _login(teacher, password, name, phone):
    if teacher == 'registered':
        ph = norm_phone(phone)
        if not ph: raise HTTPException(400, 'Số điện thoại không hợp lệ')
        if ph == ADMIN_PHONE and hmac.compare_digest(H(password), ADMIN_PW or TEACHER_PW):
            nm, key = SEEDED, ADMIN_KEY
            _wipe_drafts(key); USAGE.login(key, nm)
            r = JSONResponse({'name': nm})
            r.set_cookie(COOKIE, sign({'name': nm, 'key': key, 'exp': time.time() + 30 * 86400}), max_age=30 * 86400, httponly=True, samesite='lax', secure=True)
            return r
        a = next((v for v in acc_load().values() if v['phone'] == ph), None)
        if not a or not kdf_ok(password, a['pw']): raise HTTPException(403, 'Sai số điện thoại hoặc mật khẩu')
        pr = acc_problem(a)
        if pr: raise HTTPException(403, pr)
        nm, key = a['name'], a['id']
    elif teacher in TEACHERS:
        if not hmac.compare_digest(H(password), TEACHER_PW): raise HTTPException(403, 'Sai mật khẩu')
        nm = teacher; key = 'gv-' + slug(teacher)
    elif teacher == 'other':
        nm = ' '.join(name.split())
        if len(nm) < 2: raise HTTPException(400, 'Hãy nhập họ tên')
        if not hmac.compare_digest(H(password), GUEST_PW): raise HTTPException(403, 'Sai mật khẩu')
        key = 'khach-' + slug(nm)
        if key == 'khach-': raise HTTPException(400, 'Tên không hợp lệ')
    else: raise HTTPException(400, 'Hãy chọn giáo viên')
    _wipe_drafts(key); USAGE.login(key, nm)
    r = JSONResponse({'name': nm})
    r.set_cookie(COOKIE, sign({'name': nm, 'key': key, 'exp': time.time() + 30 * 86400}), max_age=30 * 86400, httponly=True, samesite='lax', secure=True)
    return r

@app.post('/api/logout')
def logout():
    r = JSONResponse({'ok': True}); r.delete_cookie(COOKIE); return r

@app.get('/api/me')
def me(u: U = Depends(user)):
    r = {'name': u.name, 'admin': u.key == ADMIN_KEY}
    if r['admin']:
        r['pending'] = sum(1 for a in acc_load().values() if a['status'] == 'pending')
        r['tier'] = 'admin'; r['tier_label'] = 'Quản trị viên'
        r['icon'] = VIP_TEACHER_ICON; r['icon_key'] = 'thay-dat'
    if u.key.startswith('dk-'):
        _a = acc_load()[u.key]; r['expires'] = _a.get('expires')
        r['tier'] = tier_of(_a); r['tier_label'] = TIERS[r['tier']][0]
        _ik = _a.get('icon') or ''
        r['icon'] = VIP_ICONS.get(_ik, ''); r['icon_key'] = _ik if _ik in VIP_ICONS else ''
    return r

VISIT_SEEN = os.path.join(DATA_ROOT, 'visits_seen.json')

@app.get('/api/admin/visits')
def visits(u: U = Depends(admin)):
    """Bảng thông báo người truy cập (không tính quản trị). Dùng nhật ký usage sẵn có (usage/YYYY-MM-DD.json)."""
    try: USAGE.flush()
    except Exception: pass
    today = _usage.now_vn().date()
    days = [(today - datetime.timedelta(days=i)).isoformat() for i in range(7)]
    per_day = []
    for d in days:
        us = (USAGE.load(d).get('users') or {})
        others = []
        for k, v in us.items():
            if k == ADMIN_KEY or not v.get('logins'): continue
            others.append(dict(name=v.get('name') or k, logins=int(v.get('logins') or 0),
                               times=(v.get('login_times') or [])[-6:], last=v.get('last_login')))
        others.sort(key=lambda x: x['last'] or '', reverse=True)
        per_day.append(dict(date=d, visitors=len(others), logins=sum(o['logins'] for o in others),
                            list=others if d == days[0] else []))
    seen = (jload(VISIT_SEEN, {}) or {}).get('seen_at')
    new_count = 0
    for d in days:
        for k, v in ((USAGE.load(d).get('users') or {}).items()):
            if k == ADMIN_KEY: continue
            for t in v.get('login_times') or []:
                if not seen or f'{d}T{t}' > seen: new_count += 1
    return dict(days=per_day, new=new_count, seen_at=seen)

@app.post('/api/admin/visits/seen')
def visits_seen(u: U = Depends(admin)):
    jsave(VISIT_SEEN, dict(seen_at=_usage.now_vn().isoformat(timespec='seconds')))
    return dict(ok=True)

@app.get('/api/stats/me')
def stats_me(u: U = Depends(user)):
    """Thống kê sử dụng của chính thầy/cô: số lần dùng từng chức năng + hoạt động 14 ngày."""
    try: USAGE.flush()
    except Exception: pass
    today = _usage.now_vn().date()
    days = [(today - datetime.timedelta(days=i)).isoformat() for i in range(14)]
    per_action = {a: 0 for a in _usage.ACTIONS}
    per_day = []
    for d in days:
        me = ((USAGE.load(d).get('users') or {}).get(u.key)) or {}
        acts = me.get('actions') or {}
        for a in per_action:
            per_action[a] += int(acts.get(a) or 0)
        per_day.append(dict(date=d, total=sum(int(acts.get(a) or 0) for a in acts),
                            logins=int(me.get('logins') or 0)))
    per_day.reverse()
    bb_weeks = 0
    try:
        from docx import Document
        if os.path.exists(u.MASTER):
            weeks = set()
            for t in Document(u.MASTER).tables:
                for row in t.rows:
                    for c in row.cells:
                        m = core.WEEK_RE.search(c.text or '')
                        if m: weeks.add(int(m.group(1)))
            bb_weeks = len(weeks)
    except Exception: pass
    kho_n = 0
    try:
        import kho as _kho
        kho_n = len(_kho.listing(DATA_ROOT, u.key))
    except Exception: pass
    return dict(actions=per_action, days=per_day, bb_weeks=bb_weeks, kho_files=kho_n)

# ---------- Đọc văn bản thành giọng nói (tts.py, Gemini TTS) ----------
import tts as _tts
import gemini as _gemini

def _tts_dir(key):
    d = os.path.join(DATA_ROOT, 'tts', re.sub(r'[^a-z0-9\-_]', '', str(key).lower()))
    os.makedirs(d, exist_ok=True)
    return d

def _tts_cleanup(d):
    now = time.time()
    try:
        for f in os.listdir(d):
            p = os.path.join(d, f)
            if (f.endswith('.wav') or f.endswith('.mp3')) and now - os.path.getmtime(p) > 12 * 3600:
                try: os.remove(p)
                except OSError: pass
    except OSError: pass

@app.get('/api/vc/questions')
def vc_questions(u: U = Depends(user)):
    """Ngan hang cau hoi trac nghiem Luat Vien chuc cho muc Giao vien on VC."""
    from vc_questions import QUESTIONS
    return {'total': len(QUESTIONS), 'questions': QUESTIONS}


def _vn_now():
    return datetime.datetime.now(datetime.timezone.utc).astimezone(GAME_VN)

def _vn_date_str():
    return _vn_now().date().isoformat()

def _vn_week_key(dstr):
    try:
        y, m, dd = map(int, dstr.split('-'))
        iso = datetime.date(y, m, dd).isocalendar()
        return '%d-W%02d' % (iso[0], iso[1])
    except Exception:
        return ''

def _entry_vn_date(entry):
    try:
        t = datetime.datetime.fromisoformat(str(entry.get('t') or ''))
    except Exception:
        return None
    if t.tzinfo is None:
        t = t.replace(tzinfo=datetime.timezone.utc)
    return t.astimezone(GAME_VN).date().isoformat()

def _game_path(key):
    return os.path.join(DATA_ROOT, 'users', key, 'game.json')

def _game_load(key):
    d = jload(_game_path(key), {})
    if not isinstance(d, dict):
        d = {}
    d.setdefault('xp', 0)
    d.setdefault('total_right', 0)
    d.setdefault('total_attempts', 0)
    d.setdefault('streak_count', 0)
    d.setdefault('streak_last', '')
    d.setdefault('freezes', 0)
    d.setdefault('freeze_month', '')
    d.setdefault('badges', [])
    d.setdefault('quests_daily', {'date': '', 'done': []})
    d.setdefault('quests_weekly', {'week': '', 'done': []})
    d.setdefault('quest_hist', [])
    d.setdefault('room_wins', [])
    d.setdefault('room_hist', [])
    d.setdefault('rooms_created', 0)
    d.setdefault('hide_board', False)
    d.setdefault('recent_posts', [])
    d.setdefault('last_seen', 0)
    return d

def _game_level(xp):
    idx = 0
    for i, (th, _t) in enumerate(GAME_LEVELS):
        if xp >= th:
            idx = i
    title = GAME_LEVELS[idx][1]
    if idx + 1 < len(GAME_LEVELS):
        nxt_th, nxt_title = GAME_LEVELS[idx + 1]
        return dict(level=idx + 1, title=title, to_next=nxt_th - xp, next_title=nxt_title)
    return dict(level=idx + 1, title=title, to_next=0, next_title=None)

def _game_badge_label(bid):
    if bid in GAME_BADGE_LABELS:
        return GAME_BADGE_LABELS[bid]
    if bid.startswith('docmaster:'):
        try:
            from onthi_questions import DOCS_META
            suffix = bid.split(':', 1)[1]
            for doc in DOCS_META:
                if doc['id'] == suffix:
                    return 'Thạo ' + doc['name']
            return 'Thạo ' + suffix
        except Exception:
            pass
        return 'Thạo văn bản'
    return bid

def _game_update_streak(g, today):
    month = today[:7]
    if g['freeze_month'] != month:
        g['freeze_month'] = month
        g['freezes'] = 1
    last = g['streak_last'] or ''
    if last == today:
        return
    try:
        gap = (datetime.date(*map(int, today.split('-'))) - datetime.date(*map(int, last.split('-')))).days
    except Exception:
        gap = 99
    if gap == 1:
        g['streak_count'] += 1
    elif gap == 2 and g['freezes'] > 0:
        g['freezes'] -= 1
        g['streak_count'] += 1
    else:
        g['streak_count'] = 1
    g['streak_last'] = today

def _game_eval_quests(d_hist, g, today, week):
    done_new = []
    if g['quests_daily'].get('date') != today:
        g['quests_daily'] = {'date': today, 'done': []}
    if g['quests_weekly'].get('week') != week:
        g['quests_weekly'] = {'week': week, 'done': []}
    todays = [h for h in d_hist if _entry_vn_date(h) == today]
    weeks = [h for h in d_hist if _vn_week_key(_entry_vn_date(h) or '') == week]
    right_today = sum(int(h.get('right') or 0) for h in todays)
    docs_today = set(h.get('doc') or '' for h in todays if h.get('doc'))
    streak_best_today = max([int(h.get('streak_best') or 0) for h in todays] or [0])
    wins_today = sum(1 for w in g['room_wins'] if w.get('date') == today)
    cond_day = {
        'qd_first': len(todays) >= 1,
        'qd_10right': right_today >= 10,
        'qd_streak5': streak_best_today >= 5,
        'qd_2docs': len(docs_today) >= 2,
        'qd_roomwin': wins_today >= 1,
    }
    for qid, label, xp in GAME_QUESTS_DAY:
        if qid not in g['quests_daily']['done'] and cond_day.get(qid):
            g['quests_daily']['done'].append(qid)
            g['xp'] += xp
            g['quest_hist'] = (g['quest_hist'] + [dict(date=today, qid=qid, xp=xp)])[-200:]
            done_new.append(dict(id=qid, label=label, xp=xp))
    right_week = sum(int(h.get('right') or 0) for h in weeks)
    days_week = set(_entry_vn_date(h) for h in weeks if _entry_vn_date(h))
    cond_week = {
        'qw_5days': len(days_week) >= 5,
        'qw_200right': right_week >= 200,
        'qw_90': any(int(h.get('pct') or 0) >= 90 and int(h.get('total') or 0) >= 20 for h in weeks),
    }
    for qid, label, xp in GAME_QUESTS_WEEK:
        if qid not in g['quests_weekly']['done'] and cond_week.get(qid):
            g['quests_weekly']['done'].append(qid)
            g['xp'] += xp
            g['quest_hist'] = (g['quest_hist'] + [dict(date=today, qid=qid, xp=xp)])[-200:]
            done_new.append(dict(id=qid, label=label, xp=xp))
    return done_new

def _game_eval_badges(d_hist, g, stats):
    new = []

    def add(bid):
        if bid not in g['badges']:
            g['badges'].append(bid)
            new.append(bid)
    if g['total_attempts'] >= 1:
        add('first')
    if g['total_right'] >= 100:
        add('right100')
    if g['total_right'] >= 1000:
        add('right1000')
    if g['streak_count'] >= 7:
        add('streak7')
    if g['streak_count'] >= 30:
        add('streak30')
    if any(int(h.get('pct') or 0) == 100 and int(h.get('total') or 0) >= 20 for h in d_hist):
        add('perfect')
    for doc, best in (stats.get('best') or {}).items():
        if best == 100:
            add('docmaster:' + doc)
    if g['rooms_created'] >= 5:
        add('room_creator5')
    if len(g['room_wins']) >= 10:
        add('room_winner10')
    return new

def _game_window_range(window):
    today = _vn_date_str()
    if window == 'day':
        return (today, today)
    if window == 'week':
        d = datetime.date(*map(int, today.split('-')))
        mon = d - datetime.timedelta(days=d.weekday())
        sun = mon + datetime.timedelta(days=6)
        return (mon.isoformat(), sun.isoformat())
    if window == 'month':
        return (today[:7] + '-01', today[:7] + '-31')
    return (None, None)

def _game_user_score(key, window):
    d = jload(os.path.join(DATA_ROOT, 'users', key, 'onthi.json'), {})
    hist = d.get('history') or []
    g = _game_load(key)
    if window == 'all':
        right = sum(int(h.get('right') or 0) for h in hist)
        total = sum(int(h.get('total') or 0) for h in hist)
        return dict(score=g['xp'], right=right, total=total,
                    pct=round(100 * right / total) if total else 0, attempts=len(hist))
    lo, hi = _game_window_range(window)
    sel = [h for h in hist if lo <= (_entry_vn_date(h) or '') <= hi]
    right = sum(int(h.get('right') or 0) for h in sel)
    total = sum(int(h.get('total') or 0) for h in sel)
    score = right * GAME_XP_RIGHT + sum(GAME_XP_PERFECT for h in sel
                                        if int(h.get('pct') or 0) == 100 and int(h.get('total') or 0) >= 10)
    for q in g['quest_hist']:
        if lo <= (q.get('date') or '') <= hi:
            score += int(q.get('xp') or 0)
    _seen_rh = {(w.get('date'), w.get('code')) for w in g.get('room_hist', [])}
    for w in g['room_wins']:
        if lo <= (w.get('date') or '') <= hi and (w.get('date'), w.get('code')) not in _seen_rh:
            score += GAME_XP_ROOM[0]
    for rh in g.get('room_hist', []):
        if lo <= (rh.get('date') or '') <= hi:
            score += int(rh.get('xp') or 0)
    return dict(score=score, right=right, total=total,
                pct=round(100 * right / total) if total else 0, attempts=len(sel))

def _pdf_font_paths():
    reg = None; bold = None
    for p in ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
              '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'):
        if os.path.exists(p): reg = p; break
    for p in ('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
              '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'):
        if os.path.exists(p): bold = p; break
    return reg, bold

# ---------------- Đấu trường cờ vua (V124) ----------------
# Tài khoản người chơi = tài khoản đăng ký web (dep `user`); không có hệ
# tài khoản riêng. Trạng thái lưu 1 file JSON, mọi thay đổi dưới _alock,
# theo đúng mẫu phòng thi Ôn Thi (mã phòng 6 ký tự, jsave/jload).



# ---------- THÚ CƯNG MÀN HÌNH HIỀN (PWA tại /hien-pet/, cài "Thêm vào Màn hình chính" trên iPhone) ----------
@app.get('/hien-pet')
def hp_page_redirect():
    from fastapi.responses import RedirectResponse
    return RedirectResponse('/hien-pet/')

@app.get('/hien-pet/')
def hp_page():
    from fastapi.responses import HTMLResponse
    import hienpet_web
    return HTMLResponse(hienpet_web.HTML, headers={'Cache-Control': 'no-cache'})

@app.get('/hien-pet/sw.js')
def hp_sw():
    from fastapi.responses import Response
    import hienpet_web
    return Response(hienpet_web.SW, media_type='application/javascript', headers={'Cache-Control': 'no-cache'})

@app.get('/hien-pet/manifest.json')
def hp_manifest():
    from fastapi.responses import Response
    import hienpet_web
    return Response(hienpet_web.MANIFEST, media_type='application/manifest+json')

# ---------- ỨNG DỤNG ĐIỂM DANH PHÒNG 5 + BÁO CƠM (PWA tại /diem-danh/, cài "Thêm vào Màn hình chính" trên iPhone) ----------
try:
    import diemdanh_web
    app.mount('/diem-danh', diemdanh_web.app)
    logging.getLogger('app').info('diemdanh: đã gắn ứng dụng Điểm danh tại /diem-danh/')
except Exception:
    logging.getLogger('app').exception('diemdanh: KHÔNG gắn được ứng dụng Điểm danh (web chính vẫn chạy bình thường)')

# ---------- TRANG CHỦ V150: bản tin di động + tóm tắt bán trú + xóa tin chát (dữ liệu thật) ----------
BULLETIN_F = os.path.join(DATA_ROOT, 'bulletin.json')

def _hv_board_top(window):
    # Lấy người đứng đầu một bảng Ôn thi (cùng logic với /api/onthi/game/leaderboard).
    users_dir = os.path.join(DATA_ROOT, 'users')
    try:
        keys = [k for k in os.listdir(users_dir) if os.path.isdir(os.path.join(users_dir, k))]
    except Exception:
        keys = []
    try:
        accs = acc_load()
    except Exception:
        accs = {}
    rows = []
    for key in keys:
        try:
            g = _game_load(key)
            if g.get('hide_board'):
                continue
            sc = _game_user_score(key, window)
            if window == 'all':
                if g['xp'] <= 0:
                    continue
            elif sc['score'] <= 0:
                continue
            a = accs.get(key) or {}
            rows.append(dict(name=a.get('name') or key, score=sc['score']))
        except Exception:
            continue
    rows.sort(key=lambda r: (-r['score'], r['name']))
    return rows[0] if rows else None

def _hv_chess_top():
    try:
        d = _cv_state()
        players = []
        for p in d['profiles'].values():
            games = p['w'] + p['l'] + p['d']
            if games:
                players.append({'name': p['name'], 'points': p['w'] + 0.5 * p['d']})
        players.sort(key=lambda x: (-x['points'], x['name']))
        return players[0] if players else None
    except Exception:
        return None

@app.get('/api/bulletin')
def hv_bulletin(u: U = Depends(user)):
    d = jload(BULLETIN_F, {'items': []})
    anns = [dict(idx=i, text=a.get('text') or '', at=a.get('at') or '')
            for i, a in enumerate(d.get('items') or [])][:20]
    honors = []
    seen = set()
    wk = _hv_board_top('week')
    if wk:
        honors.append(dict(kind='onthi-week', text='🏆 %s — đứng đầu tuần Ôn thi viên chức (%s điểm)' % (wk['name'], wk['score'])))
        seen.add(wk['name'])
    al = _hv_board_top('all')
    if al and al['name'] not in seen:
        honors.append(dict(kind='onthi-all', text='⭐ %s — dẫn đầu mọi thời đại Ôn thi viên chức (%s điểm)' % (al['name'], al['score'])))
        seen.add(al['name'])
    cv = _hv_chess_top()
    if cv and cv['name'] not in seen:
        honors.append(dict(kind='chess', text='♟️ %s — dẫn đầu đấu trường cờ vua (%s điểm)' % (cv['name'], cv['points'])))
    is_admin = bool(getattr(u, 'admin', False) or u.key == ADMIN_KEY)
    ai = (d.get('ai_items') or [])[:2]
    return {'announcements': anns, 'honors': honors, 'ai': ai, 'is_admin': is_admin}

@app.post('/api/bulletin/announce')
async def hv_bulletin_post(request: Request, u: U = Depends(admin)):
    try:
        body = await request.json()
    except Exception:
        raise HTTPException(400, 'Dữ liệu không hợp lệ')
    text = str((body or {}).get('text') or '').strip()[:300]
    if not text:
        raise HTTPException(400, 'Chưa nhập nội dung thông báo')
    with _alock:
        d = jload(BULLETIN_F, {'items': []})
        items = d.get('items') or []
        items.insert(0, {'text': text, 'at': _vn_now().strftime('%H:%M %d/%m'), 'by': u.name})
        d['items'] = items[:50]
        jsave(BULLETIN_F, d)
    return {'ok': True}

@app.post('/api/bulletin/announce/delete')
async def hv_bulletin_del(request: Request, u: U = Depends(admin)):
    try:
        body = await request.json()
    except Exception:
        raise HTTPException(400, 'Dữ liệu không hợp lệ')
    try:
        idx = int((body or {}).get('idx'))
    except Exception:
        raise HTTPException(400, 'Thiếu vị trí')
    with _alock:
        d = jload(BULLETIN_F, {'items': []})
        items = d.get('items') or []
        if 0 <= idx < len(items):
            items.pop(idx)
            d['items'] = items
            jsave(BULLETIN_F, d)
    return {'ok': True}

@app.get('/api/bantru/summary')
def hv_bantru_summary(u: U = Depends(user)):
    # Đọc TRỰC TIẾP kho bán trú mà ứng dụng /bantru đang dùng (không cần phiên bán trú riêng).
    bt = globals().get('_bantru')
    if bt is None:
        return {'ok': False, 'reason': 'bantru_unavailable'}
    try:
        path = bt.DB_PATH
        if not path or not os.path.exists(path):
            return {'ok': False, 'reason': 'no_db'}
        con = sqlite3.connect('file:%s?mode=ro' % path, uri=True)
        con.row_factory = sqlite3.Row
        today_s = _vn_date_str()
        today = datetime.date.fromisoformat(today_s)
        info = bt.meal_day_info(con, today)
        tom = today + datetime.timedelta(days=1)
        menu_posted = False
        dishes = []
        if tom.isoweekday() <= 5:
            monday = tom - datetime.timedelta(days=tom.isoweekday() - 1)
            row = con.execute('SELECT main_dish, veg_dish, soup, dessert FROM menu_days WHERE week_start=? AND weekday=?',
                              (monday.isoformat(), tom.isoweekday())).fetchone()
            if row:
                dishes = [x for x in [(row[k] or '').strip() for k in ('main_dish', 'veg_dish', 'soup', 'dessert')] if x]
                menu_posted = bool(dishes)
        late = None
        is_admin = bool(getattr(u, 'admin', False) or u.key == ADMIN_KEY)
        if is_admin:
            try:
                st = bt.day_stats(con, today, None)
                late = int((st.get('totals') or {}).get('late_pending') or 0)
            except Exception:
                late = None
        con.close()
        wd = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'Chủ nhật'][today.isoweekday() - 1]
        return {'ok': True, 'date': today_s, 'weekday': wd, 'serve': bool(info.get('serve')),
                'serve_note': info.get('note') or '', 'menu_tomorrow': {'posted': menu_posted, 'dishes': dishes},
                'late_pending': late}
    except Exception:
        return {'ok': False, 'reason': 'error'}

# ---------- TRANG CHỦ V150 VÒNG 2: avatar + phòng chát + bản tin AI ----------
AV_DIR = os.path.join(DATA_ROOT, 'avatars'); os.makedirs(AV_DIR, exist_ok=True)
AV_EXTS = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.webp': 'image/webp', '.gif': 'image/gif'}

def _hv_avatar_path(key):
    for ext, mt in AV_EXTS.items():
        p = os.path.join(AV_DIR, key + ext)
        if os.path.isfile(p):
            return p, mt
    return None, None

@app.post('/api/me/avatar')
async def hv_avatar_post(file: UploadFile = File(...), u: U = Depends(user)):
    raw = await file.read()
    if len(raw) > 2 * 1024 * 1024:
        raise HTTPException(400, 'Ảnh quá lớn (tối đa 2MB)')
    ct = (file.content_type or '').lower()
    ext = None
    for e, m in AV_EXTS.items():
        if ct == m:
            ext = e
            break
    if ext is None:
        raise HTTPException(400, 'Chỉ nhận ảnh (jpg, png, webp, gif)')
    with _alock:
        for e in AV_EXTS:
            p = os.path.join(AV_DIR, u.key + e)
            if os.path.isfile(p):
                os.remove(p)
        with open(os.path.join(AV_DIR, u.key + ext), 'wb') as f:
            f.write(raw)
    return {'ok': True}

@app.get('/api/avatar/me')
def hv_avatar_me(u: U = Depends(user)):
    p, mt = _hv_avatar_path(u.key)
    if p:
        return FileResponse(p, media_type=mt)
    if u.key == ADMIN_KEY:
        hp = os.path.join(BASE, 'static', 'img', 'hv-avatar.jpeg')
        if os.path.isfile(hp):
            return FileResponse(hp, media_type='image/jpeg')
    raise HTTPException(404, 'Chưa có ảnh đại diện')

def hv_chat_online(u: U = Depends(user)):
    now = time.time()
    return {'online': sum(1 for t in _HV_SEEN.values() if now - t < 300)}

def _hv_room_code(rooms):
    import secrets
    alph = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
    for _ in range(50):
        c = ''.join(secrets.choice(alph) for _ in range(6))
        if not any((r or {}).get('code') == c for r in rooms.values()):
            return c
    return None

def _hv_ai_gather():
    parts = []
    try:
        wk = _hv_board_top('week')
        if wk:
            parts.append('Đứng đầu tuần Ôn thi viên chức: %s (%s điểm)' % (wk['name'], wk['score']))
        al = _hv_board_top('all')
        if al and (not wk or al['name'] != wk['name']):
            parts.append('Dẫn đầu mọi thời đại Ôn thi viên chức: %s (%s điểm)' % (al['name'], al['score']))
        cv = _hv_chess_top()
        if cv:
            parts.append('Dẫn đầu cờ vua: %s (%s điểm)' % (cv['name'], cv['points']))
    except Exception:
        pass
    try:
        bt = globals().get('_bantru')
        if bt is not None and os.path.exists(bt.DB_PATH):
            con = sqlite3.connect('file:%s?mode=ro' % bt.DB_PATH, uri=True)
            con.row_factory = sqlite3.Row
            tom = datetime.date.fromisoformat(_vn_date_str()) + datetime.timedelta(days=1)
            if tom.isoweekday() <= 5:
                monday = tom - datetime.timedelta(days=tom.isoweekday() - 1)
                row = con.execute('SELECT main_dish, veg_dish, soup, dessert FROM menu_days WHERE week_start=? AND weekday=?',
                                  (monday.isoformat(), tom.isoweekday())).fetchone()
                if row:
                    ds = [x for x in [(row[k] or '').strip() for k in ('main_dish', 'veg_dish', 'soup', 'dessert')] if x]
                    if ds:
                        parts.append('Thực đơn bán trú ngày mai: ' + ' · '.join(ds))
            con.close()
    except Exception:
        pass
    return parts

def _hv_ai_bulletin_once():
    import gemini
    if not gemini.key():
        return False
    with _alock:
        d = jload(BULLETIN_F, {'items': []})
        if time.time() - float(d.get('ai_last_run') or 0) < 7200:
            return False
    facts = _hv_ai_gather()
    sys = ('Bạn viết bản tin ngắn cho trang chủ một trường tiểu học Việt Nam. Yêu cầu: tiếng Việt, ngắn gọn '
           '(tối đa 3 dòng, mỗi dòng một ý), vui vẻ, thân thiện. CHỈ dùng các sự kiện thật dưới đây, không bịa '
           'thêm sự kiện, tên người, món ăn hay con số nào. Nếu có tên người đạt điểm cao thì vinh danh người đó. '
           'Nếu không có sự kiện nào thì chỉ viết 1 câu chúc thầy cô một ngày vui vẻ, không bịa tin tức.')
    user = 'Hôm nay: %s.\nSự kiện thật:\n%s' % (_vn_date_str(), '\n'.join('- ' + f for f in facts) if facts else '(không có sự kiện nào)')
    body = {'system_instruction': {'parts': [{'text': sys}]},
            'contents': [{'role': 'user', 'parts': [{'text': user}]}],
            'generationConfig': {'temperature': 0.7, 'maxOutputTokens': 400}}
    try:
        res = gemini.post(body, 90)
        texts = []
        for cand in (res.get('candidates') or []):
            for part in ((cand.get('content') or {}).get('parts') or []):
                if part.get('text'):
                    texts.append(part['text'])
        text = ' '.join(texts).strip()[:500]
    except Exception as e:
        logging.getLogger('app').warning('ai bulletin: gemini lỗi %s', type(e).__name__)
        return False
    if not text:
        return False
    with _alock:
        d = jload(BULLETIN_F, {'items': []})
        ai = d.get('ai_items') or []
        ai.insert(0, {'text': text, 'at': _vn_now().strftime('%H:%M %d/%m')})
        d['ai_items'] = ai[:6]
        d['ai_last_run'] = time.time()
        jsave(BULLETIN_F, d)
    return True

def _hv_ai_loop():
    while True:
        try:
            _hv_ai_bulletin_once()
        except Exception:
            pass
        time.sleep(600)

threading.Thread(target=_hv_ai_loop, daemon=True).start()

# ---------- ỨNG DỤNG BÁN TRÚ THÔNG MINH (gắn tại /bantru/, đăng nhập riêng theo vai trò) ----------
BT_DIR = os.path.join(DATA_ROOT, 'bantru'); os.makedirs(BT_DIR, exist_ok=True)
os.environ.setdefault('BT_DB_PATH', os.path.join(BT_DIR, 'ban-tru.db'))
os.environ.setdefault('BT_BACKUP_DIR', os.path.join(BT_DIR, 'backups'))
os.environ.setdefault('BT_SECURE_COOKIE', '1')   # web chính chạy HTTPS trên Render
try:
    import bantru_app as _bantru
    app.mount('/bantru', _bantru.app)
    logging.getLogger('app').info('bantru: đã gắn ứng dụng Bán trú tại /bantru/')
except Exception:
    logging.getLogger('app').exception('bantru: KHÔNG gắn được ứng dụng Bán trú (web chính vẫn chạy bình thường)')

# ---------- MỤC 'BÁN TRÚ' MỚI: giao diện lấy ảnh làm nền, chức năng thật qua /bantru/api/ ----------
try:
    import bantru2_app as _bantru2
    app.mount('/bantru2', _bantru2.app)
    logging.getLogger('app').info('bantru2: đã gắn mục Bán trú mới tại /bantru2/')
except Exception:
    logging.getLogger('app').exception('bantru2: không gắn được mục Bán trú mới')

@app.get('/bantru2')
def bantru2_redirect():
    from fastapi.responses import RedirectResponse
    return RedirectResponse('/bantru2/')

@app.get('/bantru')
def bt_redirect():
    from fastapi.responses import RedirectResponse
    return RedirectResponse('/bantru/')

@app.get('/api/tts/voices')
def tts_voices(u: U = Depends(user)):
    return dict(voices=_tts.voices_public(), enabled=bool(_gemini.key()))

@app.post('/api/tts/synth')
async def tts_synth(request: Request, u: U = Depends(user)):
    if not _gemini.key(): raise HTTPException(400, 'Chưa bật AI (GEMINI_API_KEY) — thầy/cô báo quản trị nhé ạ.')
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    mode = str(body.get('mode') or 'read')
    voice = str(body.get('voice') or 'co_giao')
    try:
        if mode == 'dictation':
            wav = await run_in_threadpool(_tts.synth_dictation, str(body.get('text') or ''), voice,
                                          body.get('pause', 5), body.get('repeat', 2))
        elif mode == 'dialogue':
            cast = body.get('cast') if isinstance(body.get('cast'), dict) else {}
            wav = await run_in_threadpool(_tts.synth_dialogue, str(body.get('text') or ''), cast,
                                          str(body.get('narrator') or 'co_giao'))
        else:
            wav = await run_in_threadpool(_tts.synth_read, str(body.get('text') or ''), voice,
                                          str(body.get('style') or ''), str(body.get('speed') or 'vua'))
    except ValueError as e:
        raise HTTPException(400, str(e))
    d = _tts_dir(u.key); _tts_cleanup(d)
    tid = uuid.uuid4().hex[:16]
    data, ext, secs = wav
    with open(os.path.join(d, tid + '.' + ext), 'wb') as f: f.write(data)
    USAGE.action(u.key, u.name, 'doc_van_ban')
    return dict(id=tid, ext=ext, seconds=secs)

@app.get('/api/tts/audio/{tid}')
def tts_audio(tid: str, u: U = Depends(user)):
    if not re.fullmatch(r'[a-f0-9]{16}', tid or ''): raise HTTPException(404, 'Không thấy tệp âm thanh')
    d = _tts_dir(u.key)
    for ext, mt in (('mp3', 'audio/mpeg'), ('wav', 'audio/wav')):
        p = os.path.join(d, tid + '.' + ext)
        if os.path.exists(p): return FileResponse(p, media_type=mt, filename='giong-doc.' + ext)
    raise HTTPException(404, 'Tệp âm thanh đã hết hạn, thầy/cô tạo lại nhé ạ.')

@app.post('/api/tts/save')
async def tts_save(request: Request, u: U = Depends(user)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    tid = str(body.get('id') or '')
    if not re.fullmatch(r'[a-f0-9]{16}', tid): raise HTTPException(404, 'Không thấy tệp âm thanh')
    d = _tts_dir(u.key); p = ext = None
    for e in ('mp3', 'wav'):
        q = os.path.join(d, tid + '.' + e)
        if os.path.exists(q): p, ext = q, e; break
    if not p: raise HTTPException(404, 'Tệp âm thanh đã hết hạn, thầy/cô tạo lại nhé ạ.')
    name = ' '.join(str(body.get('name') or 'Giọng đọc').split())[:80] or 'Giọng đọc'
    import kho as _kho
    it = _kho.put(DATA_ROOT, u.key, 0, 'am_thanh', p, name + '.' + ext)
    return dict(ok=True, name=it['name'])

# ---------- Danh sách lớp (dùng chung cho công cụ lớp học) ----------
def _lop_load(u):
    d = jload(os.path.join(u.DATA, 'lop.json'), {})
    if not isinstance(d, dict): d = {}
    sts = []
    for s in d.get('students') or []:
        if isinstance(s, dict) and str(s.get('name') or '').strip():
            sts.append(dict(name=str(s['name'])[:60], bd=str(s.get('bd') or '')[:10]))
    return dict(class_name=str(d.get('class_name') or '')[:60], students=sts)

@app.get('/api/lop')
def lop_get(u: U = Depends(user)):
    return _lop_load(u)

@app.post('/api/lop/save')
async def lop_save(request: Request, u: U = Depends(user)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    sts = []
    for s in (body or {}).get('students') or []:
        nm = ' '.join(str((s or {}).get('name') or '').split())[:60]
        if nm: sts.append(dict(name=nm, bd=str((s or {}).get('bd') or '')[:10]))
        if len(sts) >= 200: break
    d = dict(class_name=str((body or {}).get('class_name') or '')[:60], students=sts)
    os.makedirs(u.DATA, exist_ok=True); jsave(os.path.join(u.DATA, 'lop.json'), d)
    SYNC.schedule()
    return dict(ok=True, count=len(sts))

# ---------- Đổi mật khẩu (tài khoản đăng ký) ----------
@app.post('/api/account/password')
def change_password(old: str = Form(...), new: str = Form(''), u: U = Depends(user)):
    if not u.key.startswith('dk-'):
        raise HTTPException(400, 'Tài khoản này dùng mật khẩu chung của hệ thống — thầy/cô liên hệ quản trị (0865291703) khi cần đổi nhé.')
    if len(new) < 6: raise HTTPException(400, 'Mật khẩu mới cần tối thiểu 6 ký tự')
    if len(new) > 128: raise HTTPException(400, 'Mật khẩu quá dài')
    acc = acc_load(); a = acc.get(u.key)
    if not a: raise HTTPException(404, 'Không thấy tài khoản')
    if not kdf_ok(old, a['pw']): raise HTTPException(403, 'Mật khẩu hiện tại chưa đúng')
    a['pw'] = kdf(new); acc_save(acc)
    return dict(ok=True)

# ---------- Số lượt truy cập công khai (hiển thị cho mọi người dùng) ----------
@app.get('/api/visits/total')
def visits_total(u: U = Depends(user)):
    try: USAGE.flush()
    except Exception: pass
    tot = 0
    d = os.path.join(DATA_ROOT, 'usage')
    try:
        for f in os.listdir(d):
            if not f.endswith('.json'): continue
            users = (USAGE.load(f[:-5]) or {}).get('users') or {}
            for me in users.values():
                try: tot += int((me or {}).get('logins') or 0)
                except Exception: pass
    except OSError: pass
    return dict(total=tot)

# ---------- Xem giáo án dạng tài liệu (LibreOffice -> PDF -> ảnh từng trang) ----------
import exam as _exam

@app.get('/api/giaoan/view')
def ga_view(u: U = Depends(user)):
    st = ga_state(u)
    if not st.get('plain'): raise HTTPException(404, 'Chưa tạo giáo án')
    oid = st[st.get('current') or 'plain'] or st['plain']
    p = os.path.join(u.OUT, oid + '.docx')
    if not os.path.exists(p): raise HTTPException(404, 'Không thấy tệp giáo án')
    vdir = os.path.join(u.OUT, 'view_' + oid)
    pdf = os.path.join(vdir, 'ga.pdf')
    def _pages():
        if not os.path.isdir(vdir): return []
        ps = [f for f in os.listdir(vdir) if re.fullmatch(r'p\d+\.png', f)]
        return sorted(ps, key=lambda f: int(f[1:-4]))
    pages = _pages()
    fresh = os.path.exists(pdf) and os.path.getmtime(pdf) >= os.path.getmtime(p)
    if not pages or not fresh:
        rendered = _exam.render_pages(p, vdir, dpi=110)
        if rendered is None: raise HTTPException(400, 'Máy chủ chưa sẵn sàng để xem trước — thầy/cô tải file Word về xem nhé.')
        try:
            tmp_pdf = os.path.join(vdir, 'de.pdf')
            if os.path.exists(tmp_pdf): os.replace(tmp_pdf, pdf)
        except OSError: pass
        pages = _pages()
    return dict(oid=oid, pages=len(pages), week=st.get('week'),
                th=(st.get('current') == 'th'))

@app.get('/api/giaoan/view/page/{n}')
def ga_view_page(n: int, oid: str = '', u: U = Depends(user)):
    oid = os.path.basename(str(oid or ''))
    p = os.path.join(u.OUT, 'view_' + oid, f'p{int(n)}.png')
    if not os.path.exists(p): raise HTTPException(404, 'Không thấy trang này')
    return FileResponse(p, media_type='image/png')

@app.get('/api/giaoan/view.pdf')
def ga_view_pdf(oid: str = '', u: U = Depends(user)):
    oid = os.path.basename(str(oid or ''))
    p = os.path.join(u.OUT, 'view_' + oid, 'ga.pdf')
    if not os.path.exists(p): raise HTTPException(404, 'Không thấy tệp PDF')
    return FileResponse(p, media_type='application/pdf')

@app.post('/api/register')
def register(request: Request, name: str = Form(...), phone: str = Form(...), school: str = Form(''), password: str = Form(...), password2: str = Form(...)):
    ip = client_ip(request); rl_check('register', ip); rl_hit('register', ip)
    nm = ' '.join(name.split()); sc = ' '.join(school.split())
    if not (2 <= len(nm) <= 80): raise HTTPException(400, 'Hãy nhập họ và tên (2–80 ký tự)')
    ph = norm_phone(phone)
    if not ph: raise HTTPException(400, 'Số điện thoại không hợp lệ (VD: 0912345678)')
    if len(sc) > 120: raise HTTPException(400, 'Tên trường/lớp quá dài')
    if len(password) < 6: raise HTTPException(400, 'Mật khẩu tối thiểu 6 ký tự')
    if len(password) > 128: raise HTTPException(400, 'Mật khẩu quá dài')
    if password != password2: raise HTTPException(400, 'Hai lần nhập mật khẩu không khớp')
    pw = kdf(password)
    with _alock:
        d = acc_load()
        if any(a['phone'] == ph for a in d.values()): raise HTTPException(409, f'Số điện thoại này đã đăng ký. Liên hệ {CONTACT} nếu cần hỗ trợ.')
        aid = 'dk-' + secrets.token_hex(5)
        now = datetime.datetime.now().isoformat(timespec='seconds')
        d[aid] = dict(id=aid, name=nm, phone=ph, school=sc, pw=pw, status='approved', created=now, decided=now, expires=(today() + datetime.timedelta(days=365)).isoformat(), tier='thuong')
        acc_save(d)
    USAGE.registration(aid, nm, sc)
    return {'status': 'approved', 'message': 'Đăng ký thành công! Anh/chị đăng nhập để học ngay.'}

TIERS = {'thuong': ('Thường', 50), 'vip': ('VIP', 200), 'pro': ('Pro', 10 ** 9)}
def tier_of(a): return a.get('tier') if a and a.get('tier') in TIERS else 'thuong'
def chat_limit(u):
    if u.key == ADMIN_KEY: return 10 ** 9
    if u.key.startswith('dk-'):
        t = tier_of(acc_load().get(u.key) or {})
        return CHAT_LIMIT if t == 'thuong' else TIERS[t][1]
    return CHAT_LIMIT

def pub(a): return {k: a.get(k) for k in ('id', 'name', 'phone', 'school', 'status', 'created', 'decided', 'expires')} | {'expired': bool(a.get('expires') and a['status'] == 'approved' and a['expires'] < today().isoformat()), 'tier': a.get('tier') or 'thuong'}

# ---------- phân biệt VIP/Pro (bản demo v2 đã duyệt) ----------
VIP_ICONS = {'pro-xanh': '/static/img/vip-pro-xanh.png', 'pro-hong': '/static/img/vip-pro-hong.png'}
VIP_TEACHER_ICON = '/static/img/vip-thay-dat.png'
VIP_NOTI_F = os.path.join(DATA_ROOT, 'vip_notifications.json')

def _vip_profile_for(key, accs=None):
    if key == ADMIN_KEY:
        return dict(icon=VIP_TEACHER_ICON, tier='admin', tier_label='Quản trị viên', name_cls='teacher')
    if key and key.startswith('dk-'):
        a = (accs or {}).get(key) or (acc_load().get(key) or {})
        t = tier_of(a)
        if t == 'vip':
            return dict(icon=VIP_ICONS.get(a.get('icon') or '', ''), tier='vip', tier_label='VIP', name_cls='vip')
        if t == 'pro':
            ik = a.get('icon') or ''
            return dict(icon=VIP_ICONS.get(ik, ''), tier='pro', tier_label='Pro', name_cls='pro-' + ik if ik in VIP_ICONS else 'pro')
        return dict(icon='', tier='thuong', tier_label='Thường', name_cls='normal')
    return dict(icon='', tier='thuong', tier_label='Thường', name_cls='normal')

def _vip_noti_add(kind, key, name, text):
    # CHỈ gọi trong `with _alock` (không tự khóa — tránh deadlock).
    d = jload(VIP_NOTI_F, {'items': []})
    items = d.get('items') or []
    items.insert(0, dict(kind=kind, key=key, name=name, text=text,
                         at=datetime.datetime.now(GAME_VN).strftime('%H:%M %d/%m')))
    d['items'] = items[:50]
    jsave(VIP_NOTI_F, d)

@app.get('/api/vip/notifications')
def vip_notifications(u: U = Depends(user)):
    items = jload(VIP_NOTI_F, {'items': []}).get('items', [])[:20]
    accs = acc_load()
    out = []
    for it in items:
        v = _vip_profile_for(it.get('key'), accs)
        out.append(dict(kind=it.get('kind'), key=it.get('key'), name=it.get('name'), text=it.get('text'), at=it.get('at'),
                        icon=v['icon'], tier=v['tier'], tier_label=v['tier_label'], name_cls=v['name_cls']))
    return {'items': out}

@app.post('/api/me/icon')
async def vip_set_icon(request: Request, u: U = Depends(user)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    icon = str((body or {}).get('icon') or '')
    if u.key == ADMIN_KEY: raise HTTPException(400, 'Icon của quản trị viên là cố định, không đổi được')
    if not u.key.startswith('dk-'): raise HTTPException(403, 'Chỉ tài khoản đăng ký mới chọn icon')
    if icon not in VIP_ICONS: raise HTTPException(400, 'Icon chưa có (ô VIP đang chờ ảnh từ thầy)')
    with _alock:
        d = acc_load(); a = d.get(u.key)
        if not a: raise HTTPException(404, 'Không có tài khoản này')
        if tier_of(a) not in ('vip', 'pro'): raise HTTPException(403, 'Chỉ tài khoản VIP/Pro mới được chọn icon')
        a['icon'] = icon; acc_save(d)
    return {'ok': True, 'icon': VIP_ICONS[icon], 'icon_key': icon}

@app.post('/api/admin/praise')
async def vip_praise(request: Request, u: U = Depends(admin)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    key = str((body or {}).get('key') or '')
    note = str((body or {}).get('note') or '').strip()[:200]
    with _alock:
        if key == ADMIN_KEY:
            name = SEEDED
        else:
            a = acc_load().get(key)
            if not a: raise HTTPException(404, 'Không có tài khoản này')
            name = a['name']
        text = '🏆 %s được khen thưởng' % name + ((': ' + note) if note else '')
        _vip_noti_add('praise', key, name, text)
    return {'ok': True}

@app.get('/api/admin/accounts')
def admin_list(u: U = Depends(admin)):
    L = sorted(acc_load().values(), key=lambda a: a['created'], reverse=True)
    return {'accounts': [pub(a) for a in L], 'pending': sum(1 for a in L if a['status'] == 'pending')}

@app.post('/api/admin/accounts/{aid}')
def admin_act(aid: str, action: str = Form(...), expires: str = Form(''), tier: str = Form(''), u: U = Depends(admin)):
    exp = None
    if expires.strip():
        try: exp = datetime.date.fromisoformat(expires.strip()).isoformat()
        except ValueError: raise HTTPException(400, 'Ngày hạn dùng không hợp lệ')
    with _alock:
        d = acc_load(); a = d.get(aid)
        if not a: raise HTTPException(404, 'Không có tài khoản này')
        now = datetime.datetime.now().isoformat(timespec='seconds')
        if action == 'approve':
            a.update(status='approved', decided=now, expires=exp or (today() + datetime.timedelta(days=365)).isoformat())
        elif action == 'reject': a.update(status='rejected', decided=now)
        elif action == 'lock': a.update(status='locked', decided=now)
        elif action == 'unlock':
            a.update(status='approved', decided=now, expires=exp or a.get('expires') or (today() + datetime.timedelta(days=365)).isoformat())
        elif action == 'tier':
            if tier not in TIERS: raise HTTPException(400, 'Hạng tài khoản không hợp lệ')
            _old_tier = tier_of(a)
            a['tier'] = tier
            if tier in ('vip', 'pro') and _old_tier != tier:
                _vip_noti_add('tierup', aid, a['name'], '🎉 %s đã lên hạng %s' % (a['name'], TIERS[tier][0]))
        elif action == 'expiry':
            if not exp: raise HTTPException(400, 'Hãy chọn ngày hạn dùng')
            a['expires'] = exp
        elif action == 'delete':
            d.pop(aid); acc_save(d)
            shutil.rmtree(os.path.join(DATA_ROOT, 'users', aid), ignore_errors=True); SYNC.schedule()
            return {'ok': True, 'deleted': aid}
        else: raise HTTPException(400, 'Thao tác không hợp lệ')
        acc_save(d)
    return pub(a)

@app.middleware('http')
async def sync_after_write(request: Request, call_next):
    # chống quét/hút API hàng loạt (không áp cho /api/health của cron giữ máy)
    if request.url.path.startswith('/api/') and request.url.path != '/api/health':
        try: rl_check('api', client_ip(request))
        except HTTPException: return JSONResponse({'detail': 'Thầy cô thao tác quá nhanh, vui lòng đợi một chút rồi thử lại nhé.'}, status_code=429)
        rl_hit('api', client_ip(request))
    # sơ đồ cột do chính người dùng ghép (khi mẫu của trường không tự nhận ra được)
    d = None
    try:
        d = unsign(request.cookies.get(COOKIE, ''))
        core.CMAP.set(jload(os.path.join(DATA_ROOT, 'users', d['key'], 'colmap.json'), {}) if d else {})
    except Exception: core.CMAP.set({})
    t0 = time.monotonic()
    try: r = await call_next(request)
    except Exception:  # lỗi không lường trước: ghi traceback vào log, trả lời bằng tiếng Việt (giao diện không 'treo')
        logging.getLogger('app').exception('Lỗi %s %s', request.method, request.url.path)
        msg = ('Có lỗi khi tạo giáo án. Thầy cô kiểm tra lại giáo án gốc / giáo án phụ của tuần đã chọn rồi thử lại; nếu vẫn lỗi, gửi tệp cho anh Lê Đắc Đạt (0865291703) nhé.'
               if request.url.path.startswith('/api/giaoan') else 'Có lỗi khi xử lý tệp. Thầy cô thử lại; nếu vẫn lỗi, gửi tệp cho anh Lê Đắc Đạt (0865291703) nhé.')
        gc.collect(); return JSONResponse({'detail': msg}, status_code=500)
    # python-docx tạo vòng tham chiếu giữ cả cây XML; không dọn thì RAM tăng ~40 MB mỗi lần ghép -> Render (512 MB) tắt máy giữa chừng
    if request.url.path.startswith('/api/') and time.monotonic() - t0 > 0.2: gc.collect()
    if d and request.method == 'POST' and r.status_code < 400:
        act = usage_action(request.url.path)
        if act: USAGE.action(d.get('key'), d.get('name'), act)
    if request.method in ('POST', 'PUT', 'DELETE', 'PATCH') and r.status_code < 400 and (request.url.path.startswith('/api/') or request.url.path.startswith('/bantru/api/')): SYNC.schedule()
    return r

_UA = {'/api/lop/save': 'cong_cu_lop', '/api/nhanxet/generate': 'nhan_xet', '/api/baobai/fill': 'dien_bao_bai', '/api/giaoan/build': 'tao_giao_an', '/api/giaoan/tichhop': 'tich_hop', '/api/de/extract': 'tao_de', '/api/chat': 'chat'}
def usage_action(path): return _UA.get(path)

@app.on_event('shutdown')
def _flush():
    USAGE.stop(); SYNC.flush()

ASSET_V = str(int(max(os.path.getmtime(os.path.join(BASE, 'static', f)) for f in ('index.html', 'game.js', 'game.css'))))
APP_VER = 'V180-ban-nhe'
GA_LOCKED = True   # V75 (anh chọn phương án C, 23/10): ẩn tab Giáo án chờ kiểm chứng độ chính xác chép bài; mở lại = GA_LOCKED=False trong bản phát hành kèm phần sửa khớp tên bài
def _ga_lock(u):
    if GA_LOCKED and u.key != ADMIN_KEY:
        raise HTTPException(403, 'Mục Giáo án đang tạm khóa để kiểm tra độ chính xác — sẽ mở lại sớm. Thầy cô dùng tạm các mục khác nhé!')

@app.get('/', response_class=HTMLResponse)
def index():  # trang tĩnh, API mới cần đăng nhập
    html = open(os.path.join(BASE, 'static', 'index.html'), encoding='utf-8').read().replace('__V__', ASSET_V).replace('__VER__', APP_VER)
    return HTMLResponse(html, headers={'Cache-Control': 'no-cache, no-store, must-revalidate'})

@app.get('/api/status')
def status(u: U = Depends(user)):
    """Không bao giờ lỗi 500: báo bài gốc không đọc được -> weeks=[] + master_error (giao diện vẫn hiện ô chọn tuần + hướng dẫn)."""
    weeks, err = [], None
    if os.path.exists(u.MASTER):
        try: weeks = _weeks_cached(u.MASTER)
        except Exception as e:
            logging.getLogger('app').warning('status: không đọc được báo bài gốc của %s: %r', u.key, e)
            err = 'Không đọc được báo bài gốc – thầy cô tải lại tệp ở tab Báo bài'
        if not weeks and not err: err = 'Không thấy tiêu đề "Tuần n" trong báo bài gốc'
    try: dws = sorted(int(x[1:]) for x in os.listdir(os.path.join(u.DATA, 'draft')) if x[1:].isdigit() and draft.load(u, int(x[1:]))['files'])
    except OSError: dws = []
    return dict(master=os.path.exists(u.MASTER), template=os.path.exists(u.TEMPLATE), weeks=weeks, master_error=err, draft_weeks=dws, ga_locked=GA_LOCKED and u.key != ADMIN_KEY)

@app.get('/api/health')
def health():
    """Kiểm tra sức khỏe service — KHÔNG cần đăng nhập (cho Render health check).
    Chỉ báo TÊN các biến môi trường còn thiếu, không bao giờ lộ giá trị."""
    req = ('DATA_REPO', 'DATA_TOKEN', 'SECRET_KEY', 'TEACHER_PASSWORD_SHA256', 'GUEST_PASSWORD_SHA256')
    missing = [v for v in req if not os.environ.get(v, '').strip()]
    store = 'github' if isinstance(SYNC, storage.GitSync) else 'local'
    warn = None
    if store == 'local':
        warn = ('Dữ liệu đang lưu trên ổ đĩa TẠM và sẽ MẤT TRẮNG khi service khởi động lại/redeploy. '
                'Hãy đặt DATA_REPO và DATA_TOKEN trên Render (tab Environment), xem README mục "Lưu dữ liệu bền vững".')
    return {'ok': True, 'storage': store, 'missing_env': missing,
            'storage_warning': warn, 'sync_error': getattr(SYNC, 'last_error', None)}

_WC = {}
def _weeks_cached(path):
    st = os.stat(path); k = (path, st.st_mtime_ns, st.st_size, json.dumps(core.CMAP.get(), sort_keys=True))
    if k not in _WC:
        if len(_WC) > 200: _WC.clear()
        _WC[k] = core.master_weeks(path)
    return copy.deepcopy(_WC[k])

_WP = {}
@app.get('/api/baobai/week/{w}')
def week(w: int, u: U = Depends(user)):
    k = None
    try:
        st = os.stat(u.MASTER); k = (u.MASTER, st.st_mtime_ns, st.st_size, w, json.dumps(core.CMAP.get(), sort_keys=True))
    except Exception: k = None
    if k and k in _WP: return copy.deepcopy(_WP[k])
    try: r = core.master_week_preview(u.MASTER, w)
    except KeyError: raise HTTPException(404, 'Không có tuần này')
    if k:
        if len(_WP) > 200: _WP.clear()
        _WP[k] = copy.deepcopy(r)
    return r

@app.post('/api/baobai/edit')
def edit(week: int = Form(...), idx: int = Form(...), ten: str = Form(''), tich: str = Form(''), mon: str = Form(None), u: U = Depends(user)):
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc')
    if len(ten) > 500 or len(tich) > 500: raise HTTPException(400, 'Nội dung quá dài')
    out = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx')
    try:
        try: rows = core.edit_row(u.MASTER, week, idx, ten, tich, mon, out)
        except KeyError: raise HTTPException(404, 'Không có tuần này')
        except IndexError: raise HTTPException(404, 'Không có dòng này')
        except ValueError as e: raise HTTPException(409, str(e))
        shutil.copy(u.MASTER, os.path.join(u.BK, f'baobai_{ts()}.docx')); os.replace(out, u.MASTER)
        kho_put(u, week, 'baobai', u.MASTER, f'BaoBai_Tuan{week}.docx')
    finally:
        try: os.remove(out)   # dọn file tạm kể cả khi có exception lạ
        except OSError: pass
    return rows

@app.post('/api/baobai/master')
async def replace_master(file: UploadFile = File(...), u: U = Depends(user)):
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(file, tmp)
    if not weeks_or_ai(u, tmp): return need_map(u, 'master', tmp)
    return _install_master(u, tmp)

def _install_master(u, tmp):
    if os.path.exists(u.MASTER): shutil.copy(u.MASTER, os.path.join(u.BK, f'baobai_{ts()}.docx'))
    os.replace(tmp, u.MASTER); return status(u)

MAP_MSG = {'master': 'Chưa nhận ra bảng kế hoạch tuần trong báo bài gốc của thầy cô (mẫu của trường khác mẫu quen thuộc).',
           'phu': 'Chưa nhận ra bảng kế hoạch tuần trong báo bài phụ.', 'template': 'Chưa nhận ra bảng tuần trong giáo án gốc.'}
def need_map(u, kind, path, extra=None):
    """Không tự nhận ra cấu trúc -> giữ tệp chờ, trả về các bảng để người dùng tự chọn cột (không dùng mẫu của người khác)."""
    pend = os.path.join(u.OUT, f'pending_{kind}.docx'); os.replace(path, pend)
    jsave(os.path.join(u.OUT, f'pending_{kind}.json'), extra or {})
    tabs = core.inspect(pend)
    msg = MAP_MSG[kind] + (' Thầy cô chọn giúp cột nào là Thứ, Tiết, Môn, Tên bài, Tích hợp nhé – chỉ cần chọn 1 lần.' if tabs else ' Trong tệp không có bảng nào.')
    return dict(need_map=True, kind=kind, message=msg, tables=tabs, roles=core.ROLE_VI)

def _sane_weeks(wt):
    """Kết quả nhận diện 'đạt': mọi tuần có đủ cột Thứ/Môn/Tên bài và đọc ra được ngày + môn."""
    if not wt: return False
    for blk, _r in wt.values():
        if not all(k in blk.cols for k in ('thu', 'mon', 'ten')): return False
        rows = core.rows_of(blk)
        if rows:
            fm = sum(1 for r in rows if r['mon'].strip()) / len(rows)
            ft = sum(1 for r in rows if str(r['thu']).isdigit()) / len(rows)
            if fm < .5 or ft < .4: return False
    return True

def weeks_or_ai(u, path):
    """Đọc tuần bằng quy tắc; chưa đạt thì nhờ AI hiểu mẫu (aimap, đã kiểm chứng) rồi đọc lại."""
    wt = core.week_tables(core.docx.Document(path))
    if _sane_weeks(wt): return wt
    try:
        import aimap
        if aimap.learn(u.DATA, core.docx.Document(path)):
            wt2 = core.week_tables(core.docx.Document(path))
            if wt2: return wt2
    except Exception: pass
    return wt

@app.post('/api/colmap')
def colmap(kind: str = Form(...), sig: str = Form(...), cols: str = Form(...), hdr: int = Form(1), week_heading: str = Form(''), u: U = Depends(user)):
    """Lưu cách ghép cột của người dùng (theo chữ ký tiêu đề bảng) rồi xử lý lại tệp đang chờ."""
    if kind not in MAP_MSG: raise HTTPException(400, 'Sai loại tệp')
    try: c = {k: int(v) for k, v in json.loads(cols).items() if k in core.ROLES and str(v).lstrip('-').isdigit() and int(v) >= 0}
    except Exception: raise HTTPException(400, 'Sơ đồ cột không hợp lệ')
    if 'mon' not in c or 'ten' not in c: raise HTTPException(400, 'Cần chọn ít nhất cột Môn và cột Tên bài')
    pth = os.path.join(u.DATA, 'colmap.json'); d = jload(pth, {}); d[sig[:400]] = dict(c, hdr=max(1, min(5, hdr))); jsave(pth, d)
    core.CMAP.set(d)
    pend = os.path.join(u.OUT, f'pending_{kind}.docx')
    if not os.path.exists(pend): return dict(ok=True, saved=True)
    extra = jload(os.path.join(u.OUT, f'pending_{kind}.json'), {})
    if kind == 'master':
        if not core.week_tables(core.docx.Document(pend)): return need_map(u, kind, pend)
        return dict(_install_master(u, pend), ok=True)
    if kind == 'template':
        if core.first_table(core.docx.Document(pend)) is None: return need_map(u, kind, pend)
        if os.path.exists(u.TEMPLATE): shutil.copy(u.TEMPLATE, os.path.join(u.BK, f'giaoan_{ts()}.docx'))
        os.replace(pend, u.TEMPLATE); return dict(status(u), ok=True)
    rep = _fill_run(u, pend, extra.get('week', ''))
    return dict(rep, ok=True)

@app.post('/api/baobai/fill')
async def fill(file: UploadFile = File(...), week: str = Form(''), u: U = Depends(user)):
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc')
    phu = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(file, phu)
    if not weeks_or_ai(u, phu): return need_map(u, 'phu', phu, {'week': week})
    if str(week).strip().isdigit(): return _fill_run(u, phu, week)
    return _fill_auto(u, phu, core.week_tables(core.docx.Document(phu)))

def _ranges_overlap(a, b):
    """Hai khoảng ngày (d/m/yyyy) có giao nhau không — dùng gióng tuần khi số tuần hai bên lệch nhau."""
    from datetime import datetime as _dt
    try:
        a0, a1 = (_dt.strptime(x.strip(), '%d/%m/%Y') for x in a)
        b0, b1 = (_dt.strptime(x.strip(), '%d/%m/%Y') for x in b)
    except Exception: return False
    return a0 <= b1 and b0 <= a1

def resolve_targets(mw, sw):
    """Gióng tuần của báo bài PHỤ (sw) sang tuần của báo bài GỐC (mw) — người dùng không cần chọn tuần.
    Ưu tiên: (1) trùng số tuần; (2) báo bài gốc chỉ có đúng 1 tuần -> tuần đó (số tuần chỉ là nhãn,
    khung tiết mới là chuẩn); (3) khoảng ngày giao nhau. Trả (pairs, unresolved)."""
    pairs, unresolved, used = [], [], set()
    for src in sorted(sw):
        if src in mw:
            pairs.append((src, src, 'trùng số tuần')); used.add(src); continue
        if len(mw) == 1:
            dst = next(iter(mw))
            if dst not in used: pairs.append((src, dst, 'báo bài gốc chỉ có một tuần')); used.add(dst)
            else: unresolved.append(src)
            continue
        dst = None; sr = sw[src][1]
        if sr:
            for w in sorted(mw):
                if w not in used and mw[w][1] and _ranges_overlap(sr, mw[w][1]): dst = w; break
        if dst is not None: pairs.append((src, dst, 'khớp khoảng ngày')); used.add(dst)
        else: unresolved.append(src)
    return pairs, unresolved

def _fill_auto(u, phu, sw):
    """Điền báo bài KHÔNG chọn tuần: tự gióng mọi tuần của phụ sang gốc rồi chép lần lượt theo từng cặp."""
    mw = core.week_tables(core.docx.Document(u.MASTER))
    pairs, unresolved = resolve_targets(mw, sw)
    if not pairs:
        os.remove(phu)
        return dict(weeks=[], unresolved=unresolved, message='Không gióng được tuần nào của báo bài phụ sang báo bài gốc — thầy cô kiểm tra lại hai tệp nhé')
    cur = u.MASTER; agg = {'weeks': [], 'skipped_weeks': unresolved, 'auto': True}; tmp_outs = []
    try:
        for src, dst, _how in pairs:
            out = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx')
            rep = core.fill_baobai(cur, phu, out, src, to_week=dst)
            agg['weeks'] += rep['weeks']; tmp_outs.append(out); cur = out
    except ValueError as e:
        for f in tmp_outs:
            if os.path.exists(f): os.remove(f)
        os.remove(phu); raise HTTPException(400, str(e))
    shutil.copy(u.MASTER, os.path.join(u.BK, f'baobai_{ts()}.docx'))
    os.replace(cur, u.MASTER)
    for f in tmp_outs:
        if f != cur and os.path.exists(f): os.remove(f)
    os.remove(phu)
    for _s, dst, _h in pairs: kho_put(u, dst, 'baobai', u.MASTER, f'BaoBai_Tuan{dst}.docx')
    agg['download'] = '/api/baobai/download'
    return agg

def _fill_run(u, phu, week):
    if not str(week).strip():
        return _fill_auto(u, phu, core.week_tables(core.docx.Document(phu)))
    out = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx')
    try: rep = core.fill_baobai(u.MASTER, phu, out, int(week) if str(week).strip() else None)
    except ValueError as e: raise HTTPException(400, str(e))
    finally: os.remove(phu)
    if not rep['weeks']: os.remove(out); rep['message'] = 'Không có tuần nào của báo bài phụ trùng với báo bài gốc'; return rep
    shutil.copy(u.MASTER, os.path.join(u.BK, f'baobai_{ts()}.docx')); os.replace(out, u.MASTER)
    kho_put(u, week, 'baobai', u.MASTER, f'BaoBai_Tuan{week}.docx')
    rep['download'] = '/api/baobai/download'; return rep

@app.post('/api/baobai/phuweeks')
async def phu_weeks(file: UploadFile = File(...), u: U = Depends(user)):
    """Các tuần có trong cả báo bài phụ vừa chọn và báo bài gốc – để giáo viên chọn tuần cần điền."""
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc')
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(file, tmp)
    try:
        w = core.common_weeks(u.MASTER, tmp)
        if not w:
            try:
                import aimap; aimap.learn(u.DATA, core.docx.Document(tmp))
            except Exception: pass
            w = core.common_weeks(u.MASTER, tmp)
        return dict(weeks=w)
    finally: os.remove(tmp)

@app.post('/api/baobai/phuparse')
async def phu_parse(file: UploadFile = File(...), u: U = Depends(user)):
    """Nhận diện báo bài phụ NGAY khi chọn tệp: có tuần nào, bao nhiêu bài, sẽ gióng vào tuần nào của gốc."""
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc')
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(file, tmp)
    try:
        wt = weeks_or_ai(u, tmp)
        if not wt: return need_map(u, 'phu', tmp, {'week': ''})
        mw = core.week_tables(core.docx.Document(u.MASTER))
        pairs, unresolved = resolve_targets(mw, wt)
        out = []
        for src, dst, how in pairs:
            n = sum(1 for r in core.rows_of(wt[src][0]) if r['ten'].strip())
            out.append(dict(src=src, dst=dst, how=how, lessons=n))
        return dict(pairs=out, unresolved=unresolved)
    finally:
        if os.path.exists(tmp): os.remove(tmp)

@app.get('/api/baobai/download')
def dl_master(u: U = Depends(user)):
    if not os.path.exists(u.MASTER): raise HTTPException(404, 'Chưa có báo bài gốc')
    return FileResponse(u.MASTER, media_type=DOCX, filename='BaoBai_Goc_moinhat.docx')

@app.post('/api/giaoan/template')
async def replace_template(file: UploadFile = File(...), u: U = Depends(user)):
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(file, tmp)
    if core.first_table(core.docx.Document(tmp)) is None: return need_map(u, 'template', tmp)
    if os.path.exists(u.TEMPLATE): shutil.copy(u.TEMPLATE, os.path.join(u.BK, f'giaoan_{ts()}.docx'))
    os.replace(tmp, u.TEMPLATE); return status(u)

def _th_opts(u, th, cats, sel, picks, mode):
    if th != '1': return None
    pth = os.path.join(u.DATA, 'khdh.json')
    if not os.path.exists(pth): return None
    try: sel_d = json.loads(sel) if sel else {}
    except Exception: sel_d = {}
    try: picks_d = json.loads(picks) if picks else {}
    except Exception: picks_d = {}
    o = dict(plan=tichhop.load_plan_json(pth)['plan'], cats=[c for c in cats.split(',') if c] or None, sel=sel_d, picks=picks_d, mode='verbatim')
    if mode == 'ai':
        if not exam.gemini_enabled(): raise HTTPException(400, 'Chưa bật AI – liên hệ 0865291703')
        if ai_used(u) >= AI_LIMIT: raise HTTPException(429, f'Hôm nay thầy cô đã dùng hết {AI_LIMIT} lượt AI. Thầy cô chọn "Chép nguyên văn" nhé.')
        o['mode'] = 'ai'
    return o

def _fresh_template(u, week):
    """Khung giáo án dựng MỚI từ báo bài của CHÍNH tuần đang làm — không bao giờ dùng khung
    ghi nhớ của một tuần khác (nguồn gốc các TIẾT lệch môn/lệch thứ tự so với báo bài).
    Cache trong OUT theo dấu vân tay (báo bài gốc + mẫu đã lưu + thông tin giáo án)."""
    try:
        w = int(week)
        st = os.stat(u.MASTER)
        tfp = ''
        if os.path.exists(u.TEMPLATE):
            ts = os.stat(u.TEMPLATE); tfp = f'{ts.st_mtime_ns}:{ts.st_size}'
        try: ifp = hashlib.sha1(json.dumps(ga_info_full(u), sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:12]
        except Exception: ifp = ''
        fp = {'m': st.st_mtime_ns, 's': st.st_size, 'w': w, 't': tfp, 'i': ifp}
        cache = os.path.join(u.OUT, f'tpl_w{w}.docx'); meta = cache + '.json'
        if os.path.exists(cache) and jload(meta, {}) == fp: return cache
        sample = u.TEMPLATE if os.path.exists(u.TEMPLATE) else MAU_GIAOAN
        core.gen_blank_giaoan(u.MASTER, w, sample, cache)
        try: core.apply_ga_info_file(cache, ga_info_full(u))
        except Exception: pass
        jsave(meta, fp)
        return cache
    except Exception as e:
        logging.getLogger('app').warning('dựng khung từ báo bài tuần %s lỗi (%r) — dùng mẫu đã lưu', week, e)
        return u.TEMPLATE

def _ga_run(u, week, src, opts, manual=None):
    oid = uuid.uuid4().hex[:10]; out = os.path.join(u.OUT, f'{oid}.docx')
    paths = sorted(os.path.join(src, f) for f in os.listdir(src) if f.endswith('.docx')) if src else []
    ex = {os.path.join(src, k): set(v) for k, v in jload(os.path.join(src, 'ex.json'), {}).items()} if src else {}
    try: rep = core.build_giaoan(_fresh_template(u, week), u.MASTER, week, paths, out, tichhop=opts, exclude=ex, manual=manual)
    except ValueError as e: raise HTTPException(400, str(e))
    try: core.apply_ga_info_file(out, ga_info_full(u))
    except Exception: pass
    t = rep.get('tichhop')
    if t and t.get('ai_used'):
        pa = os.path.join(u.DATA, 'ai_usage.json'); d = jload(pa, {}); k = today().isoformat(); jsave(pa, {k: int(d.get(k, 0)) + 1})
    rep['download'] = f'/api/giaoan/download/{oid}?tuan={week}'; rep['oid'] = oid
    kho_put(u, week, 'giaoan_th' if opts else 'giaoan', out, f"GiaoAn_Tuan{week}{'_TichHop' if opts else ''}.docx")
    return rep

def ga_state(u, d=None):
    p = os.path.join(u.OUT, 'state.json')
    if d is None: return jload(p, {})
    jsave(p, d); return d

@app.get('/api/giaoan/baobai_info')
def ga_baobai_info(u: U = Depends(user)):
    """Thông tin báo bài gốc đang dùng cho mục Giáo án: các tuần + số ô tên bài đã điền mỗi tuần."""
    if not os.path.exists(u.MASTER):
        return {'has_master': False}
    try:
        M = core.docx.Document(u.MASTER)
        mw = core.week_tables(M)
        filled = {}
        for w, (blk, _rng) in mw.items():
            rows = core.rows_of(blk)
            filled[str(w)] = {'total': len(rows), 'filled': sum(1 for r in rows if r['ten'].strip())}
        return {'has_master': True, 'weeks': sorted(mw.keys()), 'filled': filled}
    except Exception as e:
        logging.getLogger('app').warning('baobai_info lỗi: %r', e)
        return {'has_master': True, 'error': 'Không đọc được báo bài gốc'}

def ga_info_saved(u):
    p = os.path.join(u.DATA, 'giaoan_info.json')
    d = jload(p, {}) if os.path.exists(p) else {}
    return {k: str(d.get(k) or '') for k in ('truong', 'gv', 'lop', 'namhoc')}

def ga_info_full(u):
    """Thông tin giáo án của người dùng: giá trị đã lưu; thiếu thì tự điền sẵn từ tài khoản/năm học hiện tại."""
    d = ga_info_saved(u)
    if not d['gv']: d['gv'] = u.name
    if not d['truong'] and u.key.startswith('dk-'):
        a = acc_load().get(u.key) or {}
        d['truong'] = str(a.get('school') or '')
    if not d['namhoc']:
        now = datetime.datetime.now()
        d['namhoc'] = f'{now.year} - {now.year + 1}' if now.month >= 8 else f'{now.year - 1} - {now.year}'
    return d

@app.get('/api/giaoan/info')
def ga_info_get(u: U = Depends(user)):
    return dict(info=ga_info_full(u), saved=ga_info_saved(u))

@app.post('/api/giaoan/info')
def ga_info_post(truong: str = Form(''), gv: str = Form(''), lop: str = Form(''), namhoc: str = Form(''), u: U = Depends(user)):
    d = {k: ' '.join(str(v).split())[:80] for k, v in dict(truong=truong, gv=gv, lop=lop, namhoc=namhoc).items()}
    jsave(os.path.join(u.DATA, 'giaoan_info.json'), d)
    return dict(info=ga_info_full(u), saved=d)

@app.post('/api/giaoan/gen_template')
def gen_template(week: str = Form(''), u: U = Depends(user)):
    """BƯỚC 2: tự tạo giáo án mẫu TRỐNG từ báo bài của tuần đang chọn (khung báo bài + thứ tự môn +
    ngày tháng đúng tuần) theo mẫu tham chiếu; lưu thành giáo án gốc (mẫu) của người dùng."""
    _ga_lock(u)
    m = re.search(r'\d{1,2}', str(week or ''))
    if not m: raise HTTPException(400, 'Thầy cô chọn tuần cần làm trước nhé')
    week = int(m.group())
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc – thầy cô tải báo bài gốc ở tab Báo bài trước nhé')
    if not os.path.exists(MAU_GIAOAN): raise HTTPException(400, 'Hệ thống thiếu tệp mẫu tham chiếu – thầy cô liên hệ quản trị nhé')
    if os.path.exists(u.TEMPLATE):
        os.makedirs(u.BK, exist_ok=True); shutil.copy(u.TEMPLATE, os.path.join(u.BK, f'giaoan_{ts()}.docx'))
    try: info = core.gen_blank_giaoan(u.MASTER, week, MAU_GIAOAN, u.TEMPLATE)
    except ValueError as e: raise HTTPException(400, str(e))
    try: core.apply_ga_info_file(u.TEMPLATE, ga_info_full(u))
    except Exception: pass
    return dict(gen=info, **status(u))

@app.post('/api/draft/drop_subject')
def draft_drop_subject(week: int = Form(...), subject: str = Form(...), u: U = Depends(user)):
    """Bỏ 1 môn tải nhầm khỏi bản nháp tuần (không xoá tệp; hoàn tác được)."""
    draft.drop_subject(u, week, subject); return draft.summary(u, week, _fresh_template(u, week), u.MASTER)

@app.post('/api/draft/undrop_subject')
def draft_undrop_subject(week: int = Form(...), subject: str = Form(...), u: U = Depends(user)):
    draft.undrop_subject(u, week, subject); return draft.summary(u, week, _fresh_template(u, week), u.MASTER)

@app.post('/api/giaoan/build')
async def build(week: str = Form(''), files: list[UploadFile] = File(default=[]), manual: str = Form(''), u: U = Depends(user)):
    """Tạo giáo án THƯỜNG từ toàn bộ BẢN NHÁP giáo án phụ của tuần (không xoá bản nháp). Tạo mới -> xoá bản tích hợp cũ.
    manual: JSON '[[sec_idx, tiet_idx], ...]' – gán tay bài phụ chưa dùng vào TIẾT do người dùng chọn."""
    _ga_lock(u)
    m = re.search(r'\d{1,2}', str(week or ''))  # nhận cả 7, '7 ', 'Tuần 7'
    if not m: raise HTTPException(400, 'Thầy cô chọn tuần cần làm trước nhé')
    week = int(m.group())
    man = []
    if manual and manual.strip():
        try:
            man = [[int(a), int(b)] for a, b in json.loads(manual)]
        except (ValueError, TypeError): raise HTTPException(400, 'Danh sách gán tay không hợp lệ')
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc – thầy cô tải báo bài gốc ở tab Báo bài trước nhé')
    gen_auto = None
    if not os.path.exists(u.TEMPLATE):
        if not os.path.exists(MAU_GIAOAN): raise HTTPException(400, 'Chưa có giáo án gốc (mẫu) – thầy cô tải ở mục "Thay giáo án gốc (mẫu)" trước nhé')
        try: gen_auto = core.gen_blank_giaoan(u.MASTER, week, MAU_GIAOAN, u.TEMPLATE)
        except ValueError as e: raise HTTPException(400, str(e))
    try: wk_ok = {w['week'] for w in _weeks_cached(u.MASTER)}
    except Exception as e:
        logging.getLogger('app').warning('build: không đọc được báo bài gốc của %s: %r', u.key, e)
        raise HTTPException(400, 'Không đọc được báo bài gốc – thầy cô tải lại tệp ở tab Báo bài nhé')
    if week not in wk_ok: raise HTTPException(400, f'Báo bài gốc không có tuần {week} – thầy cô chọn tuần khác hoặc tải lại báo bài gốc nhé')
    added = await _draft_add(u, week, files) if any(f.filename for f in files) else []
    _st0 = ga_state(u)
    _reuse = bool(man) and not draft.load(u, week)['files'] and str(_st0.get('week')) == str(week) and _st0.get('src') and os.path.isdir(os.path.join(u.OUT, 'src_' + str(_st0['src'])))
    if _reuse:
        sid = str(_st0['src']); src = os.path.join(u.OUT, 'src_' + sid)
    else:
        sid = uuid.uuid4().hex[:10]; src = os.path.join(u.OUT, 'src_' + sid); os.makedirs(src, exist_ok=True)
        olds = sorted((x for x in os.listdir(u.OUT) if x.startswith('src_') and x != 'src_' + sid), key=lambda x: os.path.getmtime(os.path.join(u.OUT, x)))
        for x in olds[:-3]: shutil.rmtree(os.path.join(u.OUT, x), ignore_errors=True)
        draft.snapshot(u, week, src)
    old = ga_state(u)
    if old.get('th'):
        try: os.remove(os.path.join(u.OUT, old['th'] + '.docx'))
        except OSError: pass
    rep = _ga_run(u, week, src, None, manual=man); rep['src'] = sid
    if gen_auto: rep['mau_tu_dong'] = gen_auto
    ga_state(u, dict(src=sid, week=week, plain=rep['oid'], th=None, current='plain'))
    draft.clear(u, week)
    rep['state'] = ga_state(u); rep['draft_added'] = added; rep['draft'] = draft.summary(u, week, _fresh_template(u, week), u.MASTER); rep['draft_wiped'] = True; return rep

# ---------- bản nháp giáo án phụ (theo tuần) ----------
async def _draft_add(u, week, files):
    out = []
    for f in files:
        if not f.filename: continue
        tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(f, tmp)
        try: out.append(draft.add(u, week, tmp, f.filename))
        finally:
            try: os.remove(tmp)
            except OSError: pass
    return out

@app.get('/api/draft')
def draft_get(week: int, u: U = Depends(user)):
    return draft.summary(u, week, _fresh_template(u, week), u.MASTER)

@app.post('/api/draft/add')
async def draft_add(week: int = Form(...), files: list[UploadFile] = File(...), u: U = Depends(user)):
    rep = await _draft_add(u, week, files)
    return dict(reports=rep, draft=draft.summary(u, week, _fresh_template(u, week), u.MASTER))

@app.delete('/api/draft/file')
def draft_rm(week: int, id: str, u: U = Depends(user)):
    if not draft.remove(u, week, id): raise HTTPException(404, 'Không thấy tệp này trong bản nháp')
    return draft.summary(u, week, _fresh_template(u, week), u.MASTER)

@app.delete('/api/draft')
def draft_clear(week: int, u: U = Depends(user)):
    draft.clear(u, week); return draft.summary(u, week, _fresh_template(u, week), u.MASTER)

@app.post('/api/draft/forget')
def draft_forget(u: U = Depends(user)):
    """Giáo án phụ KHÔNG ghi nhớ: xoá toàn bộ bản nháp phụ của người dùng (web gọi khi mở trang/đăng nhập)."""
    shutil.rmtree(os.path.join(u.DATA, 'draft'), ignore_errors=True)
    return {'ok': True}

@app.post('/api/giaoan/tichhop')
def ga_tichhop(src: str = Form(...), th_cats: str = Form(''), th_sel: str = Form(''), th_picks: str = Form(''),
               th_mode: str = Form('verbatim'), u: U = Depends(user)):
    """Tích hợp vào giáo án vừa tạo (chỉ khi đã tạo giáo án). Bản thường giữ nguyên từng byte."""
    _ga_lock(u)
    st = ga_state(u)
    if not st.get('plain') or st.get('src') != src: raise HTTPException(409, 'Tạo giáo án xong mới tích hợp được')
    d = os.path.join(u.OUT, 'src_' + re.sub(r'[^0-9a-f]', '', src))
    if not os.path.isdir(d): raise HTTPException(404, 'Phiên ghép đã hết hạn, thầy cô bấm "Tạo giáo án" lại nhé')
    opts = _th_opts(u, '1', th_cats, th_sel, th_picks, th_mode)
    if not opts: raise HTTPException(400, 'Chưa tải Khung KHDH')
    rep = _ga_run(u, st['week'], d, opts)
    if st.get('th'):
        try: os.remove(os.path.join(u.OUT, st['th'] + '.docx'))
        except OSError: pass
    st.update(th=rep['oid'], current='th'); ga_state(u, st)
    rep['src'] = src; rep['state'] = st; return rep

@app.post('/api/giaoan/use')
def ga_use(which: str = Form(...), u: U = Depends(user)):
    _ga_lock(u)
    st = ga_state(u)
    if which not in ('plain', 'th') or not st.get(which): raise HTTPException(400, 'Không có bản này')
    st['current'] = which; ga_state(u, st); return st

@app.get('/api/giaoan/current')
def ga_current(u: U = Depends(user)):
    st = ga_state(u)
    if not st.get('plain'): raise HTTPException(404, 'Chưa tạo giáo án')
    oid = st[st.get('current') or 'plain'] or st['plain']
    p = os.path.join(u.OUT, oid + '.docx')
    if not os.path.exists(p): raise HTTPException(404)
    return FileResponse(p, media_type=DOCX, filename=f"GiaoAn_Tuan{st.get('week', '')}{'_TichHop' if st.get('current') == 'th' else ''}.docx")

@app.get('/api/giaoan/download/{oid}')
def dl_ga(oid: str, tuan: int = 0, u: U = Depends(user)):
    p = os.path.join(u.OUT, os.path.basename(oid) + '.docx')
    if not os.path.exists(p): raise HTTPException(404)
    th = ga_state(u).get('th') == os.path.basename(oid)
    return FileResponse(p, media_type=DOCX, filename=f'GiaoAn_Tuan{tuan}{"_TichHop" if th else ""}.docx')

# ---- Tích hợp NLS, AI, Link… theo Khung KHDH ----
import tichhop, draft
def th_info(u):
    pth = os.path.join(u.DATA, 'khdh.json')
    meta = tichhop.load_plan_json(pth).get('meta', {}) if os.path.exists(pth) else None
    return dict(has_plan=bool(meta), meta=meta, cats=[dict(key=k, name=tichhop.CAT_NAMES[k]) for k in tichhop.CAT_KEYS],
                ai=exam.gemini_enabled(), ai_limit=AI_LIMIT, ai_used=ai_used(u), ai_off='Chưa bật AI – liên hệ 0865291703')

@app.get('/api/tichhop/info')
def th_get(u: U = Depends(user)): return th_info(u)

@app.post('/api/tichhop/plan')
async def th_plan(file: UploadFile = File(...), u: U = Depends(user)):
    ext = os.path.splitext(file.filename or '')[1].lower()
    if ext not in ('.odt', '.docx'): raise HTTPException(400, 'Khung KHDH cần là tệp .odt hoặc .docx')
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + ext)
    with open(tmp, 'wb') as o: o.write(await read_upload_limited(file, MAX_FILE))
    try: plan = tichhop.parse_plan(tmp)
    except ValueError as e: raise HTTPException(400, str(e))
    except Exception: raise HTTPException(400, 'Không đọc được tệp kế hoạch')
    finally:
        try: os.remove(tmp)
        except OSError: pass
    meta = dict(filename=os.path.basename(file.filename)[:120], uploaded=datetime.datetime.now().strftime('%d/%m/%Y %H:%M'), stats=tichhop.plan_stats(plan))
    tichhop.save_plan_json(dict(meta=meta, plan=plan), os.path.join(u.DATA, 'khdh.json'))
    return th_info(u)

@app.get('/api/tichhop/preview')
def th_preview(week: int, cats: str = '', u: U = Depends(user)):
    pth = os.path.join(u.DATA, 'khdh.json')
    if not os.path.exists(pth): raise HTTPException(400, 'Chưa tải Khung KHDH')
    if not os.path.exists(u.MASTER): raise HTTPException(400, 'Chưa có báo bài gốc')
    try: return tichhop.preview(tichhop.load_plan_json(pth)['plan'], core.docx.Document(u.MASTER), week, [c for c in cats.split(',') if c] or None)
    except ValueError as e: raise HTTPException(400, str(e))


# =====================================================================================
# Tạo đề kiểm tra (v16) – xem exam.py và EXAM_RULES.md
# =====================================================================================
import exam
DE_TPL_BUNDLED = os.path.join(BASE, 'de_templates')
DE_TPL_CUSTOM = os.path.join(DATA_ROOT, 'de_templates')
AI_LIMIT = int(os.environ.get('AI_DAILY_LIMIT', '30'))
DE_KEEP = 1   # chỉ giữ đề mới nhất
MONS = {'toan': 'Toán', 'tv': 'Tiếng Việt'}
LOAI = ['Giữa học kì I', 'Cuối học kì I', 'Giữa học kì II', 'Cuối học kì II']
DE_EXT = ('.docx', '.pdf', '.jpg', '.jpeg', '.png', '.webp', '.doc')

def de_tpl(mon, kind=None):
    # Tiếng Việt tách 2 loại đề (10/2026): tv_doc = đọc hiểu, tv_viet = viết
    name = {'doc': 'tv_doc', 'viet': 'tv_viet'}.get(kind, mon) if mon == 'tv' else mon
    c = os.path.join(DE_TPL_CUSTOM, name + '.docx')
    return c if os.path.exists(c) else os.path.join(DE_TPL_BUNDLED, name + '.docx')

def jload(p, default):
    try:
        with open(p, encoding='utf-8') as f: return json.load(f)
    except Exception: return default
def jsave(p, d):
    tmp = p + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(d, f, ensure_ascii=False)
    os.replace(tmp, p)

def ai_used(u):
    d = jload(os.path.join(u.DATA, 'ai_usage.json'), {})
    return int(d.get(today().isoformat(), 0))
def ai_take(u):
    """Đếm lượt AI theo ngày — có file-lock (fcntl) chống race khi 2 request đồng thời đọc-kiểm tra-ghi."""
    p = os.path.join(u.DATA, 'ai_usage.json'); k = today().isoformat()
    with open(p, 'a+') as fh:
        fcntl.flock(fh, fcntl.LOCK_EX)
        try:
            fh.seek(0)
            try: d = json.load(fh)
            except Exception: d = {}
            if not isinstance(d, dict): d = {}
            if int(d.get(k, 0)) >= AI_LIMIT: raise HTTPException(429, f'Hôm nay bạn đã dùng hết {AI_LIMIT} lượt đọc ảnh. Vui lòng thử lại vào ngày mai hoặc dùng file Word.')
            d[k] = int(d.get(k, 0)) + 1
            fh.seek(0); fh.truncate(); json.dump(d, fh, ensure_ascii=False); fh.flush()
            try: os.fsync(fh.fileno())
            except OSError: pass
        finally:
            fcntl.flock(fh, fcntl.LOCK_UN)

def de_dir(u, did=None):
    base = os.path.join(u.DATA, 'de'); os.makedirs(base, exist_ok=True)
    if did is None: return base
    if not re.fullmatch(r'[0-9a-f]{12}', did or ''): raise HTTPException(404, 'Không có đề này')
    d = os.path.join(base, did)
    if not os.path.isdir(d): raise HTTPException(404, 'Không có đề này')
    return d

def de_clear(u, keep=None):
    """Chỉ giữ đề mới nhất (keep); xoá mọi đề cũ của người dùng này."""
    base = de_dir(u)
    for x in os.listdir(base):
        if x != keep: shutil.rmtree(os.path.join(base, x), ignore_errors=True) if os.path.isdir(os.path.join(base, x)) else os.remove(os.path.join(base, x))

def de_fields(xa, truong, nam, lop, loai, thoigian):
    lop = str(lop).strip()
    if lop not in ('1', '2', '3', '4', '5'): raise HTTPException(400, 'Lớp phải từ 1 đến 5')
    if loai not in LOAI: raise HTTPException(400, 'Loại bài không hợp lệ')
    if not re.fullmatch(r'\s*20\d\d\s*-\s*20\d\d\s*', nam or ''): raise HTTPException(400, 'Năm học dạng 2026 - 2027')
    for v in (xa, truong, thoigian):
        if len(v or '') > 120: raise HTTPException(400, 'Nội dung quá dài')
    if not (xa or '').strip() or not (truong or '').strip(): raise HTTPException(400, 'Hãy điền UBND xã và tên trường để in lên đầu đề')
    return dict(xa=(xa or '').strip(), truong=(truong or '').strip(), nam=nam.strip(), lop=int(lop), loai=loai, thoigian=(thoigian or '').strip())

def public_exam(ex):
    e = {k: v for k, v in ex.items() if k not in ('raw', '_answers_orig')}
    e['answers'] = {k: {kk: vv for kk, vv in (a or {}).items() if kk != 'raw'} for k, a in (ex.get('answers') or {}).items()}
    e['figures'] = {k: {'kind': f.get('kind'), 'has_orig': bool(f.get('orig')), 'svg': bool(f.get('svg')), 'w_cm': f.get('w_cm')} for k, f in (ex.get('figures') or {}).items()}
    return e

def de_view(u, did):
    d = de_dir(u, did); st = jload(os.path.join(d, 'state.json'), {})
    ex = jload(os.path.join(d, 'exam.json'), {})
    pages = sorted((f for f in os.listdir(os.path.join(d, 'pages')) if re.fullmatch(r'p\d+\.png', f)), key=lambda f: int(f[1:-4])) if os.path.isdir(os.path.join(d, 'pages')) else []
    v = os.stat(os.path.join(d, 'de.docx')).st_mtime_ns if os.path.exists(os.path.join(d, 'de.docx')) else 0
    return dict(id=did, mon=ex.get('mon'), fields=st.get('fields'), dapan=st.get('dapan', True), exam=public_exam(ex),
                nq=exam.count_questions(ex) if ex else 0, warnings=st.get('warnings', []), labels=st.get('labels', []),
                pages=[f'/api/de/{did}/page/{f[1:-4]}?v={v}' for f in pages], html=None if pages else exam.html_preview(ex, st.get('fields') or {}),
                download=f'/api/de/{did}/download', created=st.get('created'),
                frame_bad=len(st.get('frame_bad') or []), frame_pages=sorted({v['page'] for v in st.get('frame_bad') or []}), frame_checked=st.get('frame_checked', False))

def de_build(u, did, ex, fields, dapan, render_png=True):
    d = de_dir(u, did); out = os.path.join(d, 'tmp_de.docx'); pdir = os.path.join(d, 'pages')
    try:
        # EXAM_RULES §6: dựng -> render -> kiểm tra khung trang -> tự sửa (ngắt trang / co hình) -> dựng lại
        warns, _, pages, bad = exam.build_checked(de_tpl(ex['mon'], ex.get('kind')), ex, fields, out, [os.path.join(d, 'media'), os.path.join(d, 'figs')],
                                                  with_answers=dapan, pages_dir=pdir if exam.soffice_bin() else None, pages_png=render_png)
    except (ValueError, KeyError, StopIteration) as e:
        logging.getLogger('exam').warning('build lỗi: %s', type(e).__name__)
        raise HTTPException(400, 'Không dựng được đề từ nội dung này. Hãy kiểm tra lại các câu hỏi.')
    os.replace(out, os.path.join(d, 'de.docx'))
    kho_put(u, current_week(u), 'de', os.path.join(d, 'de.docx'), f"DeKiemTra_{'Toan' if ex.get('mon') == 'toan' else 'TiengViet'}_Lop{(fields or {}).get('lop', '') if isinstance(fields, dict) else ''}.docx")
    if pages is None and os.path.isdir(pdir): shutil.rmtree(pdir, ignore_errors=True)
    labels = exam.page_labels(os.path.join(pdir, 'de.pdf'), ex['mon']) if pages else []
    if 'Trang trống!' in labels: warns.append('Có trang trống trong bản xem trước – hãy kiểm tra.')
    if bad:
        logging.getLogger('exam').warning('khung trang: còn %d chỗ chạm/tràn khung', len(bad))
        warns.append(exam.frame_summary(bad))
    st = jload(os.path.join(d, 'state.json'), {})
    st.update(fields=fields, dapan=dapan, warnings=(ex.get('warnings') or []) + warns, labels=labels,
              frame_bad=[dict(page=v['page'], side=v['side'], kind=v['kind']) for v in bad][:50], frame_checked=pages is not None)
    jsave(os.path.join(d, 'state.json'), st); jsave(os.path.join(d, 'exam.json'), ex)

@app.get('/api/de/info')
def de_info(u: U = Depends(user)):
    prefs = jload(os.path.join(u.DATA, 'de_prefs.json'), {})
    base = de_dir(u); drafts = []
    for x in sorted(os.listdir(base), key=lambda x: -os.path.getmtime(os.path.join(base, x)))[:DE_KEEP]:
        st = jload(os.path.join(base, x, 'state.json'), {})
        if st: drafts.append(dict(id=x, mon=st.get('mon'), created=st.get('created'), lop=(st.get('fields') or {}).get('lop'), loai=(st.get('fields') or {}).get('loai')))
    return dict(ai=exam.gemini_enabled(), ai_limit=AI_LIMIT, ai_used=ai_used(u), ai_off_msg=exam.AI_OFF_MSG, prefs=prefs, loai=LOAI,
                render=bool(exam.soffice_bin()), admin=u.key == ADMIN_KEY, drafts=drafts,
                templates={m: dict(custom=os.path.exists(os.path.join(DE_TPL_CUSTOM, m + '.docx'))) for m in MONS})

@app.post('/api/de/extract')
async def de_extract(mon: str = Form(...), xa: str = Form(''), truong: str = Form(''), nam: str = Form('2026 - 2027'), lop: str = Form('3'),
                     loai: str = Form('Giữa học kì I'), thoigian: str = Form('40 phút'), dapan: str = Form('1'), kind: str = Form(''),
                     files: list[UploadFile] = File(default=[]), doc_hieu: list[UploadFile] = File(default=[]),
                     de_viet: list[UploadFile] = File(default=[]), nghe_viet: list[UploadFile] = File(default=[]), u: U = Depends(user)):
    if mon not in MONS: raise HTTPException(400, 'Môn không hợp lệ')
    kind = kind if kind in ('doc', 'viet') else ''
    fields = de_fields(xa, truong, nam, lop, loai, thoigian)
    jsave(os.path.join(u.DATA, 'de_prefs.json'), {k: fields[k] for k in ('xa', 'truong', 'nam', 'lop', 'loai', 'thoigian')} | {'mon': mon, 'kind': kind})
    slots = [('all', files)] + ([('doc_hieu', doc_hieu), ('de_viet', de_viet), ('nghe_viet', nghe_viet)] if mon == 'tv' else [])
    ups = [(sl, f) for sl, fs in slots for f in fs if f and f.filename]
    if not ups: raise HTTPException(400, 'Chưa chọn tệp đề nào')
    if len(ups) > 12: raise HTTPException(400, 'Tối đa 12 tệp mỗi lần')
    de_clear(u)   # mỗi lần "Tạo đề" là một việc mới: xoá sạch đề cũ (tệp tải lên, JSON, hình, trang xem trước, docx)
    did = uuid.uuid4().hex[:12]; d = os.path.join(de_dir(u), did); src = os.path.join(d, 'src'); os.makedirs(src)
    try:
        paths = []
        for k, (sl, f) in enumerate(ups):
            ext = os.path.splitext(f.filename.lower())[1]
            if ext not in DE_EXT: raise HTTPException(400, f'Tệp {f.filename}: chỉ nhận ảnh JPG/PNG, PDF hoặc Word')
            data = await read_upload_limited(f, MAX_FILE)
            p = os.path.join(src, f'{k}_{sl}{ext}')
            with open(p, 'wb') as o: o.write(data)
            paths.append((sl, p))
        try: ex = exam.extract(paths, mon, d, ai_hook=lambda: ai_take(u))
        except HTTPException: raise
        except exam.NeedAI as e: raise HTTPException(400, str(e))
        except ValueError as e: raise HTTPException(400, str(e))
        except Exception as e:
            logging.getLogger('exam').warning('extract lỗi: %s', type(e).__name__)
            raise HTTPException(400, 'Không đọc được nội dung tệp. Hãy thử file Word (.docx) hoặc ảnh rõ nét hơn.')
        if exam.count_questions(ex) == 0 and not (mon == 'tv' and (ex['reading']['paragraphs'] or ex['writing']['topics'] or ex['dictation']['paragraphs'])):
            raise HTTPException(400, 'Không tìm thấy câu hỏi nào (dạng "Câu 1. …") trong tệp.')
        if mon == 'tv' and kind: ex['kind'] = kind
        ex['_answers_orig'] = {k: a.get('text') for k, a in ex['answers'].items()}
        jsave(os.path.join(d, 'state.json'), dict(mon=mon, created=datetime.datetime.now().strftime('%d/%m/%Y %H:%M')))
        de_build(u, did, ex, fields, dapan != '0')
    except HTTPException:
        shutil.rmtree(d, ignore_errors=True); raise
    de_clear(u, keep=did)
    return de_view(u, did)

@app.post('/api/de/reset')
def de_reset(u: U = Depends(user)):
    de_clear(u); return {'ok': True}

@app.get('/api/de/{did}')
def de_get(did: str, u: U = Depends(user)): return de_view(u, did)

def _clean_blocks(bs):
    out = []
    for b in bs or []:
        t = b.get('t')
        if t == 'text': out.append(dict(t='text', text=str(b.get('text') or '')[:2000], muc=b.get('muc'), diem=b.get('diem')))
        elif t == 'options': out.append(dict(t='options', items=[str(x)[:300] for x in (b.get('items') or [])][:8], layout='col' if b.get('layout') == 'col' else 'row'))
        elif t == 'lines': out.append(dict(t='lines', n=max(1, min(30, int(b.get('n') or 1)))))
        elif t == 'ds': out.append(dict(t='ds', items=[str(x)[:300] for x in (b.get('items') or [])][:10]))
        elif t in ('figure', 'raw'): out.append(dict(t=t, id=str(b.get('id') or ''), text=str(b.get('text') or '')[:200]))
        elif t == 'space': out.append(dict(t='space', n=max(1, min(20, int(b.get('n') or 1)))))
    return out

def _clean_tv(k, v):
    S = lambda x, n=3000: str(x or '')[:n]
    L = lambda x: [S(i) for i in (x or [])][:30] if isinstance(x, list) else []
    if k == 'reading': return dict(intro=S(v.get('intro')), title=S(v.get('title'), 300), paragraphs=L(v.get('paragraphs')), source=L(v.get('source')), instruction=S(v.get('instruction')))
    if k == 'writing': return dict(topics=[dict(label=S(t.get('label'), 30), text=S(t.get('text')), hint=S(t.get('hint'))) for t in (v.get('topics') or [])[:5] if isinstance(t, dict)])
    return dict(title=S(v.get('title'), 300), paragraphs=L(v.get('paragraphs')), source=S(v.get('source'), 300))

@app.post('/api/de/{did}/build')
async def de_rebuild(did: str, request: Request, u: U = Depends(user)):
    d = de_dir(u, did); ex = jload(os.path.join(d, 'exam.json'), None); st = jload(os.path.join(d, 'state.json'), {})
    if not ex: raise HTTPException(404, 'Không có đề này')
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    e = body.get('exam') or {}
    def opt(v): return (str(v).strip()[:20] or None) if v not in (None, '') else None
    secs = []
    for s_ in (e.get('sections') or [])[:10]:
        qs = [dict(so=int(Q.get('so') or 0), text=str(Q.get('text') or '')[:2000], muc=opt(Q.get('muc')), diem=opt(Q.get('diem')), blocks=_clean_blocks(Q.get('blocks')))
              for Q in (s_.get('questions') or [])[:60]]
        secs.append(dict(title=str(s_.get('title') or '')[:200], instruction=str(s_.get('instruction') or '')[:500], questions=qs))
    ex['sections'] = secs
    for k, a in (e.get('answers') or {}).items():
        if k in ex['answers']: ex['answers'][k].update(text=str(a.get('text') or '')[:2000], points=str(a.get('points') or '')[:20])
        else: ex['answers'][k] = dict(text=str(a.get('text') or '')[:2000], points=str(a.get('points') or '')[:20])
    if ex['mon'] == 'tv':
        for k in ('reading', 'writing', 'dictation'):
            if isinstance(e.get(k), dict): ex[k] = _clean_tv(k, e[k])
    f = body.get('fields') or st.get('fields') or {}
    fields = de_fields(f.get('xa', ''), f.get('truong', ''), f.get('nam', '2026 - 2027'), f.get('lop', 3), f.get('loai', LOAI[0]), f.get('thoigian', ''))
    de_build(u, did, ex, fields, bool(body.get('dapan', st.get('dapan', True))))
    return de_view(u, did)

@app.get('/api/de/{did}/page/{n}')
def de_page(did: str, n: int, u: U = Depends(user)):
    p = os.path.join(de_dir(u, did), 'pages', f'p{int(n)}.png')
    if not os.path.exists(p): raise HTTPException(404)
    return FileResponse(p, media_type='image/png', headers={'Cache-Control': 'private, max-age=600'})

@app.get('/api/de/{did}/fig/{fid}')
def de_fig(did: str, fid: str, orig: int = 0, u: U = Depends(user)):
    d = de_dir(u, did); ex = jload(os.path.join(d, 'exam.json'), {})
    f = (ex.get('figures') or {}).get(fid)
    if not f: raise HTTPException(404)
    p = f.get('orig') if orig else f.get('file')
    if not p: raise HTTPException(404)
    p = os.path.join(d, 'figs', os.path.basename(p))
    if not os.path.exists(p): raise HTTPException(404)
    return FileResponse(p, media_type='image/png')

@app.get('/api/de/{did}/download')
def de_download(did: str, force: int = 0, u: U = Depends(user)):
    d = de_dir(u, did); st = jload(os.path.join(d, 'state.json'), {}); f = st.get('fields') or {}
    p = os.path.join(d, 'de.docx')
    if not os.path.exists(p): raise HTTPException(404)
    if st.get('frame_bad') and not force:   # §6: không giao im lặng đề còn chạm/tràn khung
        raise HTTPException(409, 'Đề còn nội dung chạm/tràn khung viền trang – hãy sửa rồi “Cập nhật đề”, hoặc xác nhận tải về.')
    name = f"DeKiemTra_{'Toan' if st.get('mon') == 'toan' else 'TiengViet'}_Lop{f.get('lop', '')}.docx"
    return FileResponse(p, media_type=DOCX, filename=name)

@app.post('/api/admin/de/template')
async def de_template(mon: str = Form(...), kind: str = Form(''), reset: str = Form(''), file: UploadFile = File(None), u: U = Depends(admin)):
    if mon not in MONS: raise HTTPException(400, 'Môn không hợp lệ')
    kind = kind if kind in ('doc', 'viet') else ''
    name = {'doc': 'tv_doc', 'viet': 'tv_viet'}.get(kind, mon) if mon == 'tv' else mon
    os.makedirs(DE_TPL_CUSTOM, exist_ok=True); dst = os.path.join(DE_TPL_CUSTOM, name + '.docx')
    if reset == '1':
        if os.path.exists(dst): os.remove(dst)
        return de_info(u)
    if file is None or not file.filename: raise HTTPException(400, 'Chưa chọn tệp')
    tmp = os.path.join(DE_TPL_CUSTOM, 'tmp_' + uuid.uuid4().hex + '.docx'); await save_upload(file, tmp)
    try:
        ex0 = exam.extract_docx(tmp, mon, tempfile_dir())
        if kind != 'viet' and exam.count_questions(ex0) == 0: raise ValueError
        ex0['kind'] = kind or 'both'
        exam.build(tmp, ex0, {}, tmp + '.chk.docx')
    except Exception:
        for x in (tmp, tmp + '.chk.docx'):
            if os.path.exists(x): os.remove(x)
        raise HTTPException(400, 'Đề mẫu phải là đề kiểm tra có tiêu đề (UBND…) và câu hỏi dạng "Câu 1. … (M1) (0,5đ)"')
    os.remove(tmp + '.chk.docx'); os.replace(tmp, dst)
    return de_info(u)

def tempfile_dir():
    import tempfile
    return tempfile.mkdtemp(prefix='detpl_')


# =====================================================================================
# Soạn đề mới bằng AI (Tạo đề AI) – màn hình riêng full-screen, dựng Word qua luồng sẵn có
# =====================================================================================
import taode_toc
TAODE_LOAI = ['Phiếu bài tập theo chủ đề', 'Đề ôn tập', 'Đề kiểm tra giữa HK1', 'Đề kiểm tra cuối HK1', 'Đề kiểm tra giữa HK2', 'Đề kiểm tra cuối năm']
TAODE_LOAI_MAP = {'Đề kiểm tra giữa HK1': 'Giữa học kì I', 'Đề kiểm tra cuối HK1': 'Cuối học kì I',
                  'Đề kiểm tra giữa HK2': 'Giữa học kì II', 'Đề kiểm tra cuối năm': 'ĐỀ KIỂM TRA CUỐI NĂM',
                  'Phiếu bài tập theo chủ đề': 'Phiếu bài tập theo chủ đề', 'Đề ôn tập': 'Đề ôn tập'}
TAODE_MONS = {'toan': 'Toán', 'tv': 'Tiếng Việt', 'ta': 'Tiếng Anh'}

@app.get('/api/taode/toc')
def taode_toc_get(request: Request, u: U = Depends(user)):
    # Hợp đồng với khối etView: ?mon=<nhãn tiếng Việt hoặc mã>&lop=<1..5>[&tap=...]
    # → {topics: [chuỗi], verified: bool, tap: [..] | None}. Không đủ tham số → trả nguyên dict cũ.
    mon = {'Toán': 'toan', 'Tiếng Việt': 'tv', 'Tiếng Anh': 'ta',
           'toan': 'toan', 'tv': 'tv', 'ta': 'ta'}.get((request.query_params.get('mon') or '').strip())
    lop = (request.query_params.get('lop') or '').strip()
    if not mon or lop not in taode_toc.TAODE_TOC:
        return taode_toc.TAODE_TOC
    g = taode_toc.TAODE_TOC[lop]
    tap_q = (request.query_params.get('tap') or '').strip()
    topics = []
    taps = None
    if mon == 'toan':
        for ch in g['subjects']['toan']:
            for b in ch.get('bai') or []:
                topics.append(f"{ch.get('chu_de', '')} — {b}")
    elif mon == 'tv':
        tv = g['subjects']['tv']
        taps = sorted({e.get('tap') for e in tv if e.get('tap')})
        for e in tv:
            if tap_q and e.get('tap') != tap_q:
                continue
            pre = '' if tap_q else ((e.get('tap') or '') + ' — ' if e.get('tap') else '')
            for b in e.get('bai') or []:
                if e.get('chu_diem') and b != e['chu_diem']:
                    topics.append(f"{pre}{e['chu_diem']} — {b}")
                else:
                    topics.append(f"{pre}{b}")
    else:
        for e in g['subjects']['ta']:
            if e.get('unit'):
                topics.append(e['unit'])
    return dict(topics=topics, verified=bool(g.get('verified')), tap=taps)

@app.post('/api/taode/ai-draft')
async def taode_ai_draft(request: Request, u: U = Depends(user)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    body = body or {}
    mon = str(body.get('mon') or '').strip()
    if mon not in TAODE_MONS: raise HTTPException(400, 'Môn không hợp lệ')
    try: lop = int(body.get('lop'))
    except Exception: raise HTTPException(400, 'Lớp phải từ 1 đến 5')
    if lop not in (1, 2, 3, 4, 5): raise HTTPException(400, 'Lớp phải từ 1 đến 5')
    loai_de = str(body.get('loai_de') or '').strip()
    if loai_de not in TAODE_LOAI: raise HTTPException(400, 'Loại đề không hợp lệ')
    chu_de = str(body.get('chu_de') or '').strip()  # để trống = toàn bộ nội dung kiểm tra (không bắt buộc)
    ghi_chu = str(body.get('ghi_chu') or '').strip()
    nam_hoc = str(body.get('nam_hoc') or '').strip()
    thoi_gian = str(body.get('thoi_gian') or '').strip()
    hinh_thuc = str(body.get('hinh_thuc') or '').strip()
    ai_take(u)   # đếm lượt AI như de_extract
    mon_vn = TAODE_MONS[mon]
    extra = []
    if nam_hoc: extra.append(f'Năm học: {nam_hoc}.')
    if thoi_gian: extra.append(f'Thời gian làm bài: {thoi_gian}.')
    if hinh_thuc: extra.append(f'Hình thức: {hinh_thuc}.')
    if mon == 'ta': extra.append('Môn Tiếng Anh: nội dung câu hỏi và phương án bằng tiếng Anh, câu dẫn bằng tiếng Việt.')
    prompt = (f'Bạn là giáo viên tiểu học Việt Nam soạn đề theo chương trình GDPT 2018.\n'
              f'Soạn một đề "{loai_de}" môn {mon_vn} lớp {lop}, phạm vi: ' + (f'chủ đề / phạm vi: {chu_de}.' if chu_de else f'toàn bộ nội dung môn {mon_vn} lớp {lop} (không giới hạn chủ đề cụ thể).') + '\n'
              + (' '.join(extra) + '\n' if extra else '') +
              'Yêu cầu mặc định: 10 câu, tổng 10 điểm, gồm 6 câu trắc nghiệm + 4 câu tự luận.\n'
              + (f'Ghi chú của giáo viên (ưu tiên hơn mặc định trên): {ghi_chu}\n' if ghi_chu else '') +
              'Quy tắc:\n'
              '- Mỗi câu có mức M1/M2/M3 và điểm (dạng "0,5", "1"); tổng điểm các câu = 10.\n'
              '- Trắc nghiệm: đúng 4 phương án, đáp án đúng ghi trong answers_list.\n'
              '- Tự luận: kèm đáp án và hướng dẫn chấm ngắn (giai).\n'
              '- Môn Tiếng Việt: BẮT BUỘC kèm bài đọc (reading) gồm tiêu đề + 2–4 đoạn văn ngắn phù hợp với lớp, '
              'nguồn ghi rõ nếu có (không có thì để trống); các câu đọc hiểu hỏi về bài đọc này.\n'
              '- Trả về DUY NHẤT một đối tượng JSON: sections gồm 2 phần "I. TRẮC NGHIỆM" (câu dẫn: "Khoanh vào chữ cái đặt trước kết quả đúng.") '
              'và "II. TỰ LUẬN"; mỗi câu hỏi có so/text/muc/diem/blocks (trắc nghiệm: block options 4 items; tự luận: block lines n=3); '
              'answers_list gồm so/text/points/giai; môn Tiếng Việt thêm reading {intro, title, paragraphs, source, instruction}.\n'
              '- Giữ nguyên chữ tiếng Việt có dấu. Không bịa kiến thức ngoài chủ đề đã cho.')
    try: raw = exam.gemini_compose(prompt)
    except exam.NeedAI as e: raise HTTPException(400, str(e))
    except ValueError as e: raise HTTPException(400, str(e))
    ex = exam.normalize_compose(raw, mon)
    if exam.count_questions(ex) == 0: raise HTTPException(400, 'AI không soạn được câu hỏi nào, vui lòng thử lại.')
    def _diem(v):
        try: return float(str(v or '').replace(',', '.'))
        except Exception: return 0.0
    tn = tl = 0; tong = 0.0; thieu_diem = False
    for s_ in ex['sections']:
        for Q in s_['questions']:
            if any(b.get('t') == 'options' for b in Q.get('blocks') or []): tn += 1
            else: tl += 1
            if not Q.get('diem'): thieu_diem = True
            tong += _diem(Q.get('diem'))
    totals = dict(tn=tn, tl=tl, total=tn + tl, points=round(tong, 2))
    warnings = list(ex.get('warnings') or [])
    if loai_de.startswith('Đề kiểm tra') and abs(tong - 10.0) > 0.01:
        warnings.append(f"Tổng điểm đang là {('%g' % tong).replace('.', ',')} thay vì 10 – thầy cô rà soát lại trước khi xuất Word.")
    if thieu_diem: warnings.append('Có câu chưa ghi điểm – thầy cô bổ sung trước khi xuất Word.')
    thieu_da = [str(Q['so']) for s_ in ex['sections'] for Q in s_['questions'] if str(Q['so']) not in ex['answers']]
    if thieu_da: warnings.append('Có câu chưa có đáp án – thầy cô bổ sung trước khi xuất Word.')
    return dict(exam=public_exam(ex), totals=totals, warnings=warnings)

def _taode_matrix_rows(ex):
    """Ma trận xác định từ câu hỏi (cùng logic với xem trước): M1/M2/M3 x TN/TL + điểm."""
    lv = {'M1': [0, 0.0, 0, 0.0], 'M2': [0, 0.0, 0, 0.0], 'M3': [0, 0.0, 0, 0.0]}
    for s_ in ex.get('sections') or []:
        for q in s_.get('questions') or []:
            kind = 'tn' if any((b or {}).get('t') == 'options' for b in q.get('blocks') or []) else 'tl'
            level = str(q.get('muc') or 'M1').upper()
            if level not in lv:
                level = 'M1'
            try:
                pts = float(str(q.get('diem') or '0').replace(',', '.'))
            except Exception:
                pts = 0.0
            if kind == 'tn':
                lv[level][0] += 1; lv[level][1] += pts
            else:
                lv[level][2] += 1; lv[level][3] += pts
    return lv

def _taode_append_matrix(docx_path, ex):
    """Bản Word tải từ màn hình EduTest phải đủ 3 phần: đề + đáp án + ma trận.
    Chỉ chạy ở đường taode/word (luồng Tạo đề cũ không động vào).
    - Ngắt trang trước đoạn 'ĐÁP ÁN' để đáp án ở trang riêng.
    - Thêm 'MA TRẬN ĐỀ' ở trang riêng cuối tài liệu, cùng phông Times New Roman 14 như mẫu."""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    def fmt(v):
        return ('%g' % round(v, 2)).replace('.', ',')

    def cell_text(cell, text, bold=False):
        cell.text = ''
        par = cell.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for j, part in enumerate(str(text).split('\n')):
            if j:
                par.add_run().add_break()
            r = par.add_run(part)
            r.bold = bold
            r.font.name = 'Times New Roman'
            r.font.size = Pt(14)

    d = Document(docx_path)
    dap = None
    for par in d.paragraphs:
        if par.text.strip() == 'ĐÁP ÁN':
            dap = par
            break
    if dap is not None:
        pb = OxmlElement('w:p')
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        pb.append(r)
        dap._p.addprevious(pb)
    d.add_page_break()
    h = d.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = h.add_run('MA TRẬN ĐỀ')
    hr.bold = True
    hr.font.name = 'Times New Roman'
    hr.font.size = Pt(14)
    lv = _taode_matrix_rows(ex)
    table = d.add_table(rows=1, cols=5)
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), '000000')
        borders.append(e)
    tblPr.append(borders)
    for i, t in enumerate(['Mức độ', 'Trắc nghiệm', 'Tự luận', 'Tổng số câu', 'Tổng điểm']):
        cell_text(table.rows[0].cells[i], t, bold=True)
    tot = [0, 0.0, 0, 0.0]
    for level in ('M1', 'M2', 'M3'):
        tn_n, tn_p, tl_n, tl_p = lv[level]
        tot[0] += tn_n; tot[1] += tn_p; tot[2] += tl_n; tot[3] += tl_p
        row = table.add_row().cells
        for i, t in enumerate([level, f'{tn_n} câu\n{fmt(tn_p)} điểm', f'{tl_n} câu\n{fmt(tl_p)} điểm',
                               f'{tn_n + tl_n} câu', fmt(tn_p + tl_p)]):
            cell_text(row[i], t)
    row = table.add_row().cells
    for i, t in enumerate(['Tổng cộng', f'{tot[0]} câu\n{fmt(tot[1])} điểm', f'{tot[2]} câu\n{fmt(tot[3])} điểm',
                           f'{tot[0] + tot[2]} câu', fmt(tot[1] + tot[3])]):
        cell_text(row[i], t, bold=True)
    d.save(docx_path)

@app.post('/api/taode/word')
async def taode_word(request: Request, u: U = Depends(user)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    body = body or {}
    f = body.get('fields') or {}
    mon = str(f.get('mon') or body.get('mon') or '').strip()
    if mon not in TAODE_MONS: raise HTTPException(400, 'Môn không hợp lệ')
    if mon == 'ta': raise HTTPException(400, 'Chưa có file mẫu cho môn Tiếng Anh')
    loai_de = str(f.get('loai_de') or '').strip()
    if loai_de not in TAODE_LOAI: raise HTTPException(400, 'Loại đề không hợp lệ')
    mapped = TAODE_LOAI_MAP[loai_de]
    # de_fields chỉ chấp nhận 4 loại kiểm tra (luồng cũ không đổi); 2 loại mới đi qua giá trị giữ chỗ
    # rồi ghi đè lại — tiêu đề thật do header_rules (exam.py) xử lý cộng thêm.
    fields = de_fields(f.get('xa', ''), f.get('truong', ''), f.get('nam', '2026 - 2027'), f.get('lop', 3),
                       mapped if mapped in LOAI else LOAI[0], f.get('thoigian', ''))
    fields['loai'] = mapped
    e = body.get('exam') or {}
    ex = exam.new_exam(mon); ex['source_kind'] = 'ai'
    def opt(v): return (str(v).strip()[:20] or None) if v not in (None, '') else None
    secs = []
    for s_ in (e.get('sections') or [])[:10]:
        qs = [dict(so=int(Q.get('so') or 0), text=str(Q.get('text') or '')[:2000], muc=opt(Q.get('muc')), diem=opt(Q.get('diem')), blocks=_clean_blocks(Q.get('blocks')))
              for Q in (s_.get('questions') or [])[:60]]
        secs.append(dict(title=str(s_.get('title') or '')[:200], instruction=str(s_.get('instruction') or '')[:500], questions=qs))
    ex['sections'] = secs
    for k, a in (e.get('answers') or {}).items():
        a = a or {}
        ans = dict(text=str(a.get('text') or '')[:2000], points=str(a.get('points') or '')[:20])
        if str(a.get('giai') or '').strip(): ans['giai'] = str(a.get('giai'))[:2000]
        ex['answers'][str(k)] = ans
    if mon == 'tv':
        for k in ('reading', 'writing', 'dictation'):
            if isinstance(e.get(k), dict): ex[k] = _clean_tv(k, e[k])
    if exam.count_questions(ex) == 0: raise HTTPException(400, 'Đề chưa có câu hỏi nào')
    did = uuid.uuid4().hex[:12]; d = os.path.join(de_dir(u), did); os.makedirs(d, exist_ok=True)
    try:
        jsave(os.path.join(d, 'state.json'), dict(mon=mon, created=datetime.datetime.now().strftime('%d/%m/%Y %H:%M')))
        de_build(u, did, ex, fields, True, render_png=False)  # tải Word: kiểm tra khung bằng PDF, không render ảnh PNG
        _taode_append_matrix(os.path.join(d, 'de.docx'), ex)
    except HTTPException:
        shutil.rmtree(d, ignore_errors=True); raise
    except Exception:
        shutil.rmtree(d, ignore_errors=True)
        raise HTTPException(400, 'Không thêm được ma trận vào file Word, vui lòng thử lại.')
    de_clear(u, keep=did)
    return dict(download=f'/api/de/{did}/download', did=did)


# =====================================================================================
# Trợ lý AI "Hiền" (chat, não Muse Spark — V76) – giới hạn riêng, tách khỏi lượt đọc ảnh của Tạo đề
# =====================================================================================
import chatbot
CHAT_LIMIT = int(os.environ.get('CHAT_DAILY_LIMIT', '50'))

def chat_hist(u): return jload(os.path.join(u.DATA, 'chat.json'), {}).get('messages', [])
def chat_save(u, msgs): os.makedirs(u.DATA, exist_ok=True); jsave(os.path.join(u.DATA, 'chat.json'), {'messages': msgs[-chatbot.KEEP:]})
def chat_used(u): return int(jload(os.path.join(u.DATA, 'chat_usage.json'), {}).get(today().isoformat(), 0))

@app.get('/api/chat')
def chat_get(u: U = Depends(user)):
    return dict(messages=chat_hist(u), ai=chatbot.enabled(), brain=chatbot.brain(), limit=chat_limit(u), used=chat_used(u), name=chatbot.BOT_NAME, faq=list(chatbot.FAQ))

@app.post('/api/chat')
async def chat_post(request: Request, u: U = Depends(user)):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    msg = str((body or {}).get('message') or '').strip()[:chatbot.MAX_IN]
    if not msg: raise HTTPException(400, 'Thầy cô chưa nhập câu hỏi ạ')
    hist = chat_hist(u); now = datetime.datetime.now().strftime('%d/%m %H:%M')
    if not chatbot.enabled():
        reply = chatbot.faq_answer(msg); reply = (reply + '\n\n_(' + chatbot.OFF_MSG.split(' – ')[0] + '.)_') if reply else chatbot.OFF_MSG
        src = 'faq'
    else:
        p = os.path.join(u.DATA, 'chat_usage.json'); d = jload(p, {}); k = today().isoformat()
        lim = chat_limit(u)
        if int(d.get(k, 0)) >= lim:
            raise HTTPException(429, f'Dạ, hôm nay thầy cô đã dùng hết {lim} lượt hỏi trợ lý AI. Mai thầy cô hỏi tiếp em nhé ạ!')
        try: reply = chatbot.ask(hist, msg, dict(name=u.name, admin=(u.key == ADMIN_KEY)))
        except ValueError as e: raise HTTPException(502, str(e))
        os.makedirs(u.DATA, exist_ok=True); jsave(p, {k: int(d.get(k, 0)) + 1}); src = 'ai'
    hist += [dict(role='user', text=msg, t=now), dict(role='bot', text=reply, t=now)]
    chat_save(u, hist)
    return dict(reply=reply, src=src, used=chat_used(u), limit=chat_limit(u), ai=chatbot.enabled())

@app.get('/api/chat/diag')
def chat_diag(model: str = '', u: U = Depends(admin)):
    """Chẩn đoán não Muse (chỉ quản trị): key thấy những mô hình nào, gọi thử 1 câu tối thiểu ra sao.
    Không bao giờ trả giá trị key — chỉ mã HTTP + đoạn nội dung lỗi/tối đa 300 ký tự."""
    import urllib.request as _ur, urllib.error as _ue
    out = {'brain': chatbot.brain(), 'model_config': chatbot.muse_model(), 'models': None, 'models_raw': None, 'probe': None}
    if not chatbot.muse_key(): return out
    hdr = {'Authorization': 'Bearer ' + chatbot.muse_key(), 'Content-Type': 'application/json'}
    try:
        out['models'] = chatbot._meta_models(); out['models_raw'] = chatbot._models_cache.get('raw')
    except Exception as e: out['models'] = f'lỗi: {type(e).__name__}'
    try:
        body = json.dumps({'model': model or chatbot.muse_model(), 'messages': [{'role': 'user', 'content': 'hi'}], 'max_tokens': 5}).encode()
        rq = _ur.Request(chatbot.meta_base() + '/chat/completions', data=body, method='POST', headers=hdr)
        with _ur.urlopen(rq, timeout=60) as r: out['probe'] = {'http': r.status, 'body': r.read().decode('utf-8', 'replace')[:300]}
    except _ue.HTTPError as e:
        try: snip = e.read().decode('utf-8', 'replace')[:300]
        except Exception: snip = ''
        out['probe'] = {'http': e.code, 'body': snip}
    except Exception as e: out['probe'] = {'http': None, 'body': f'lỗi kết nối: {type(e).__name__}'}
    return out

@app.post('/api/chat/clear')
def chat_clear(u: U = Depends(user)):
    chat_save(u, []); return {'ok': True}


# =====================================================================================
# Kho lưu theo tuần (kho.py) – mọi tệp đã tạo, theo người dùng, theo tuần
# =====================================================================================
import kho
def kho_put(u, week, typ, src, name):
    """Không bao giờ làm hỏng việc chính: lỗi kho chỉ ghi cảnh báo."""
    try: kho.put(DATA_ROOT, u.key, week, typ, src, name); SYNC.schedule()
    except Exception as e: logging.getLogger('kho').warning('kho: không lưu được %s (%r)', typ, e)

def current_week(u):
    """Tuần hiện tại theo báo bài gốc (ngày hôm nay, giờ VN, nằm trong tuần nào) – không có thì 'khac'."""
    try:
        d = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))).date()
        for w in _weeks_cached(u.MASTER) if os.path.exists(u.MASTER) else []:
            r = w.get('range') or []
            if len(r) == 2:
                a, b = (datetime.datetime.strptime(x, '%d/%m/%Y').date() for x in r)
                a -= datetime.timedelta(days=a.weekday()); b = a + datetime.timedelta(days=6)
                if a <= d <= b: return w['week']
    except Exception as e: logging.getLogger('kho').warning('kho: không xác định được tuần hiện tại (%r)', e)
    return 'khac'

@app.get('/api/kho')
def kho_list(u: U = Depends(user)): return kho.listing(DATA_ROOT, u.key)

@app.get('/api/kho/file/{fid}')
def kho_file(fid: str, u: U = Depends(user)):
    f = kho.find(DATA_ROOT, u.key, fid)
    if not f: raise HTTPException(404, 'Không có tệp này trong kho của thầy cô')
    it, p = f
    return FileResponse(p, media_type=kho.MIME.get(os.path.splitext(p)[1], 'application/octet-stream'), filename=it['name'])

@app.delete('/api/kho/file/{fid}')
def kho_del(fid: str, u: U = Depends(user)):
    if not (kho.find(DATA_ROOT, u.key, fid) and kho.delete(DATA_ROOT, u.key, fid)): raise HTTPException(404, 'Không có tệp này trong kho của thầy cô')
    SYNC.schedule(); return kho.listing(DATA_ROOT, u.key)

# =====================================================================================
# Nhận xét học sinh lớp 2 theo Thông tư 27 (nhanxet.py)
# =====================================================================================
import nhanxet
def nx_path(u): return os.path.join(u.DATA, 'nhanxet.json')
def nx_info(u):
    return dict(subjects=[dict(key=k, name=v) for k, v in nhanxet.SUBJ], default_subjects=nhanxet.DEFAULT_SUBJ, ki=nhanxet.KI,
                muc_mon=nhanxet.MUC_MON, muc_pn=nhanxet.MUC_PN, max_students=nhanxet.MAX_STUDENTS,
                ai=exam.gemini_enabled(), ai_limit=AI_LIMIT, ai_used=ai_used(u))

@app.get('/api/nhanxet')
def nx_get(u: U = Depends(user)):
    st = jload(nx_path(u), None)
    return dict(info=nx_info(u), state=nhanxet.clean_state(st) if st else None)

async def _nx_body(request):
    try: body = await request.json()
    except Exception: raise HTTPException(400, 'Dữ liệu không hợp lệ')
    return body if isinstance(body, dict) else {}

@app.post('/api/nhanxet/state')
async def nx_save(request: Request, u: U = Depends(user)):
    st = nhanxet.clean_state((await _nx_body(request)).get('state')); jsave(nx_path(u), st); return {'ok': True, 'n': len(st['students'])}

@app.post('/api/nhanxet/parse')
async def nx_parse(file: UploadFile = File(None), text: str = Form(''), u: U = Depends(user)):
    if file is not None and file.filename:
        ext = os.path.splitext(file.filename.lower())[1]
        if ext not in ('.xlsx', '.docx', '.csv', '.txt'): raise HTTPException(400, 'Chỉ nhận tệp .xlsx, .docx, .csv hoặc dán danh sách')
        b = await read_upload_limited(file, 5 * 1024 * 1024)
        tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + ext)
        with open(tmp, 'wb') as f: f.write(b)
        try: names = nhanxet.names_from_file(tmp, file.filename)
        except ValueError as e: raise HTTPException(400, str(e))
        except Exception as e:
            logging.getLogger('nhanxet').warning('parse lỗi: %s', type(e).__name__); raise HTTPException(400, 'Không đọc được tệp danh sách – thầy cô thử dán danh sách vào ô bên dưới nhé')
        finally:
            try: os.remove(tmp)
            except OSError: pass
    else: names = nhanxet.names_from_text(text[:50000])
    if not names: raise HTTPException(400, 'Không thấy tên học sinh nào – mỗi dòng 1 em (VD: 1. Nguyễn Văn An)')
    return dict(names=names)

@app.post('/api/nhanxet/generate')
async def nx_generate(request: Request, u: U = Depends(user)):
    body = await _nx_body(request); st = nhanxet.clean_state(body.get('state'))
    if not st['students']: raise HTTPException(400, 'Chưa có học sinh nào trong danh sách lớp')
    only = body.get('only')
    idx = [i for i in only if isinstance(i, int) and 0 <= i < len(st['students'])] if isinstance(only, list) else list(range(len(st['students'])))
    if not idx: raise HTTPException(400, 'Chưa chọn học sinh')
    sub = [st['students'][i] for i in idx]
    def take():
        if ai_used(u) >= AI_LIMIT: return False
        p = os.path.join(u.DATA, 'ai_usage.json'); d = jload(p, {}); k = today().isoformat(); jsave(p, {k: int(d.get(k, 0)) + 1}); return True
    want_ai = bool(body.get('ai', True)) and exam.gemini_enabled()
    nx, mode, n = nhanxet.generate(sub, st['subjects'], st['ki'], want_ai, take)
    for i, c in zip(idx, nx): st['students'][i]['nx'] = c
    jsave(nx_path(u), st)
    msg = {'ai': f'Đã viết nhận xét bằng AI ({n} lượt).', 'mix': f'AI viết một phần ({n} lượt), phần còn lại dùng câu mẫu.',
           'mau': 'Đã viết nhận xét bằng câu mẫu' + (' (hết lượt AI hôm nay).' if want_ai and ai_used(u) >= AI_LIMIT else '.' if want_ai else ' (chưa bật AI).')}[mode]
    return dict(state=st, mode=mode, ai_calls=n, message=msg, ai_used=ai_used(u), ai_limit=AI_LIMIT)

@app.post('/api/nhanxet/export')
async def nx_export(request: Request, u: U = Depends(user)):
    body = await _nx_body(request); st = nhanxet.clean_state(body.get('state')); fmt = body.get('fmt')
    if fmt not in ('xlsx', 'docx'): raise HTTPException(400, 'Định dạng không hợp lệ')
    if not st['students']: raise HTTPException(400, 'Chưa có học sinh nào')
    jsave(nx_path(u), st)
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.' + fmt)
    try:
        (nhanxet.export_xlsx if fmt == 'xlsx' else nhanxet.export_docx)(st, tmp)
        name = f"NhanXet_Lop{re.sub(r'[^0-9A-Za-z]+', '', slug(st['lop']).upper()) or '2'}_{['GiuaKiI', 'CuoiKiI', 'GiuaKiII', 'CuoiNam'][nhanxet.KI.index(st['ki'])]}.{fmt}"
        it = kho.put(DATA_ROOT, u.key, current_week(u), 'nhanxet_' + fmt, tmp, name); SYNC.schedule()
    finally:
        try: os.remove(tmp)
        except OSError: pass
    return dict(download=f"/api/kho/file/{it['id']}", item=it, week=kho.wlabel(it['week']))

# =====================================================================================
# Phiếu bài tập & Trò chơi (phieu.py + static/game.js)
# =====================================================================================
import phieu
def pf_path(u): return os.path.join(u.DATA, 'phieu.json')
def pf_lessons(u, w):
    """Các bài của tuần w trong báo bài (đã điền) của giáo viên: môn, thứ, tiết, tên bài."""
    if not os.path.exists(u.MASTER): return []
    import docx as _docx
    M = _docx.Document(u.MASTER)
    try: t = core.week_tables(M)[w][0]
    except KeyError: return []
    out, thu, cnt = [], '', {}
    for r in core.rows_of(t):
        thu = (r.get('thu') or '').strip() or thu; buoi = (r.get('buoi') or '').strip()
        if not r.get('mon'): continue
        n = cnt[(thu, buoi)] = cnt.get((thu, buoi), 0) + 1
        ten = re.sub(r'\s+', ' ', r.get('ten') or '').strip()
        if ten: out.append(dict(mon=r['mon'].strip(), thu=thu, tiet=(r.get('tiet') or '').strip() or ((buoi + ' – ' if buoi else '') + f'tiết {n}'), ten=ten))
    del M
    return out

def pf_ga(u, w):
    """Giáo án mới nhất của tuần (trong Kho) để làm ngữ cảnh."""
    try:
        its = [x for x in kho.load(DATA_ROOT, u.key) if x['week'] == kho.wkey(w) and x['type'] in ('giaoan_th', 'giaoan')]
        if its:
            it = max(its, key=lambda x: x['created']); f = kho.find(DATA_ROOT, u.key, it['id'])
            return f[1] if f else None
    except Exception as e: logging.getLogger('phieu').warning('phieu: không tìm được giáo án (%r)', e)

def pf_info(u):
    return dict(ai=exam.gemini_enabled(), ai_limit=AI_LIMIT, ai_used=ai_used(u), types=phieu.TYPE_VI, muc=phieu.MUC, kinds=phieu.KINDS)

@app.get('/api/phieu')
def pf_get(u: U = Depends(user)):
    st = jload(pf_path(u), None)
    return dict(info=pf_info(u), state=phieu.clean_state(st) if st else None)

@app.get('/api/phieu/lessons')
def pf_lessons_api(week: int, u: U = Depends(user)):
    try: ls = pf_lessons(u, week)
    except Exception as e:
        logging.getLogger('phieu').warning('phieu: không đọc được báo bài (%s)', type(e).__name__); ls = []
    return dict(week=week, lessons=ls, has_ga=bool(pf_ga(u, week)))

@app.post('/api/phieu/state')
async def pf_save(request: Request, u: U = Depends(user)):
    st = phieu.clean_state((await _nx_body(request)).get('state')); jsave(pf_path(u), st); return {'ok': True}

def _pf_need(st):
    if not st['lesson']['week']: raise HTTPException(400, 'Thầy cô chọn tuần trước nhé')
    if not st['lesson']['ten']: raise HTTPException(400, 'Thầy cô chọn bài (hoặc gõ tên bài) trước nhé')

@app.post('/api/phieu/generate')
async def pf_generate(request: Request, u: U = Depends(user)):
    st = phieu.clean_state((await _nx_body(request)).get('state')); _pf_need(st)
    if not exam.gemini_enabled(): raise HTTPException(400, 'Chưa bật AI nên chưa tạo tự động được. Thầy cô có thể tự gõ câu hỏi vào ô soạn bên dưới.')
    if ai_used(u) >= AI_LIMIT: raise HTTPException(429, f'Hôm nay thầy cô đã dùng hết {AI_LIMIT} lượt AI. Thầy cô có thể tự gõ câu hỏi vào ô soạn, mai AI lại tạo được ạ.')
    ga = pf_ga(u, st['lesson']['week']); ctx = phieu.lesson_context(ga, st['lesson']['ten']) if ga else ''
    p = os.path.join(u.DATA, 'ai_usage.json'); d = jload(p, {}); k = today().isoformat(); jsave(p, {k: int(d.get(k, 0)) + 1})
    try: new, warn = await run_in_threadpool(phieu.generate, st, ctx)
    except ValueError as e: raise HTTPException(502, str(e))
    jsave(pf_path(u), new)
    if 'phieu' in new['types']: USAGE.action(u.key, u.name, 'phieu_bai_tap')
    if {'khoidong', 'tuongtac'} & set(new['types']): USAGE.action(u.key, u.name, 'tro_choi')
    return dict(state=new, warnings=warn, used_context=bool(ctx), ai_used=ai_used(u), ai_limit=AI_LIMIT,
                message='Đã tạo bằng AI' + (' (có dùng giáo án của bài)' if ctx else '') + '. Thầy cô xem lại và sửa trước khi dùng nhé.')

def _pf_name(st, pre, ext):
    t = re.sub(r'[^0-9A-Za-z]+', '_', slug(st['lesson']['ten'] or 'bai')).strip('_')[:40] or 'bai'
    return f"{pre}_Tuan{st['lesson']['week']}_{t}.{ext}", slug(st['lesson']['ten'] or 'bai')[:60]

@app.post('/api/phieu/export')
async def pf_export(request: Request, u: U = Depends(user)):
    body = await _nx_body(request); st = phieu.clean_state(body.get('state')); kind = body.get('kind'); _pf_need(st)
    if kind not in ('phieu', 'khoidong'): raise HTTPException(400, 'Loại tệp không hợp lệ')
    jsave(pf_path(u), st)
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.docx')
    try:
        try: await run_in_threadpool(phieu.export_phieu if kind == 'phieu' else phieu.export_khoidong, st, tmp)
        except ValueError as e: raise HTTPException(400, str(e))
        name, tag = _pf_name(st, 'PhieuBaiTap' if kind == 'phieu' else 'TroChoiKhoiDong', 'docx')
        it = kho.put(DATA_ROOT, u.key, st['lesson']['week'], 'phieu_bai_tap' if kind == 'phieu' else 'tro_choi', tmp, name, tag=tag); SYNC.schedule()
    finally:
        try: os.remove(tmp)
        except OSError: pass
    if st['src'] == 'tay': USAGE.action(u.key, u.name, 'phieu_bai_tap' if kind == 'phieu' else 'tro_choi')
    return dict(download=f"/api/kho/file/{it['id']}", item=it, week=kho.wlabel(it['week']))

def game_payload(st):
    g = st['game']
    return dict(v=1, title=st['lesson']['ten'], mon=st['lesson']['mon'], week=st['lesson']['week'],
                questions=[q for q in g['questions'] if phieu.gq_ok(q)], ochu=g['ochu'] if phieu.ochu_ok(g['ochu']) else None)

@app.post('/api/phieu/game/save')
async def pf_game_save(request: Request, u: U = Depends(user)):
    st = phieu.clean_state((await _nx_body(request)).get('state')); _pf_need(st)
    pl = game_payload(st)
    if not pl['questions']: raise HTTPException(400, 'Trò chơi cần ít nhất 1 câu hỏi đủ 4 phương án và đáp án')
    jsave(pf_path(u), st)
    tmp = os.path.join(u.DATA, 'tmp_' + uuid.uuid4().hex + '.json'); jsave(tmp, pl)
    try:
        name, tag = _pf_name(st, 'TroChoi', 'json')
        it = kho.put(DATA_ROOT, u.key, st['lesson']['week'], 'tro_choi_web', tmp, name, tag=tag); SYNC.schedule()
    finally:
        try: os.remove(tmp)
        except OSError: pass
    if st['src'] == 'tay': USAGE.action(u.key, u.name, 'tro_choi')
    return dict(id=it['id'], play=f"/choi/{it['id']}", item=it, week=kho.wlabel(it['week']))

@app.get('/api/game/{fid}')
def game_get(fid: str, u: U = Depends(user)):
    f = kho.find(DATA_ROOT, u.key, fid)
    if not f or f[0]['type'] != 'tro_choi_web': raise HTTPException(404, 'Không có trò chơi này trong kho của thầy cô')
    return jload(f[1], {})

@app.get('/choi/{fid}', response_class=HTMLResponse)
def game_page(fid: str, request: Request):
    """Trang chơi toàn màn hình – chỉ chủ trò chơi (đã đăng nhập) mở được; dữ liệu tải qua /api/game (kiểm tra quyền)."""
    d = unsign(request.cookies.get(COOKIE, ''))
    if not d: return HTMLResponse('<!doctype html><meta charset="utf-8"><title>Trò chơi</title><body style="font-family:sans-serif;padding:40px;font-size:22px">Thầy cô cần <a href="/">đăng nhập</a> rồi mở lại đường dẫn trò chơi này.</body>', status_code=401)
    if not re.fullmatch(r'[0-9a-f]{12}', fid): raise HTTPException(404, 'Không có trò chơi này')
    return HTMLResponse(f'''<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Trò chơi lớp học</title><link rel="stylesheet" href="/static/game.css?v={ASSET_V}"></head><body class="gpage">
<div id="gpload" style="font:24px sans-serif;padding:40px">Đang tải trò chơi…</div><script src="/static/game.js?v={ASSET_V}"></script>
<script>fetch('/api/game/{fid}').then(r=>r.ok?r.json():r.json().then(j=>Promise.reject(j.detail||'Lỗi'))).then(g=>{{document.getElementById('gpload').remove();
fetch('/api/nhanxet').then(r=>r.ok?r.json():null).then(n=>{{const names=(n&&n.state&&n.state.students||[]).map(s=>s.name).filter(Boolean);GamePlayer.open(g,{{names,standalone:true}})}}).catch(()=>GamePlayer.open(g,{{standalone:true}}))}})
.catch(e=>{{document.getElementById('gpload').textContent=String(e)}})</script></body></html>''')

# =====================================================================================
# PWA: manifest, service worker (chỉ lưu khung tĩnh – không bao giờ lưu API/tệp người dùng), assetlinks (Android TWA)
# =====================================================================================
MANIFEST = {'name': 'Trợ lý Báo bài & Giáo án', 'short_name': 'Báo bài', 'lang': 'vi', 'dir': 'ltr', 'start_url': '/?src=app', 'scope': '/', 'id': '/',
            'display': 'standalone', 'background_color': '#fffcf3', 'theme_color': '#2563eb',
            'description': 'Điền báo bài, ghép giáo án, tạo đề, nhận xét, phiếu bài tập & trò chơi cho giáo viên tiểu học.',
            'icons': [{'src': f'/static/icons/icon-{s}.png', 'sizes': f'{s}x{s}', 'type': 'image/png', 'purpose': 'any'} for s in (192, 512)] +
                     [{'src': f'/static/icons/maskable-{s}.png', 'sizes': f'{s}x{s}', 'type': 'image/png', 'purpose': 'maskable'} for s in (192, 512)]}

@app.get('/manifest.webmanifest')
def pwa_manifest():
    return JSONResponse(MANIFEST, media_type='application/manifest+json', headers={'Cache-Control': 'no-cache'})

SW_JS = """/* Service worker: chỉ lưu khung tĩnh (/static). KHÔNG lưu /api, /choi, tệp người dùng.
   HTML/JS/CSS: mạng trước (bản mới hiện ngay sau mỗi lần cập nhật), chỉ dùng bản lưu khi mất mạng. Ảnh/phông: lưu trước. */
const C='bb-shell-v2-__V__',SHELL=['/static/offline.html','/static/icons/icon-192.png','/manifest.webmanifest'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
const put=(r,res)=>{if(res.ok&&res.type==='basic'){const cl=res.clone();caches.open(C).then(c=>c.put(r,cl))}return res};
self.addEventListener('fetch',e=>{const r=e.request,u=new URL(r.url);if(r.method!=='GET'||u.origin!==location.origin)return;
 if(r.mode==='navigate'){if(u.pathname==='/')e.respondWith(fetch(r,{cache:'no-store'}).catch(()=>caches.match('/static/offline.html')));return}
 if(!u.pathname.startsWith('/static/'))return;
 if(/\\.(js|css|html|json|webmanifest)$/.test(u.pathname)){e.respondWith(fetch(r,{cache:'no-cache'}).then(res=>put(r,res)).catch(()=>caches.match(r)));return}
 e.respondWith(caches.match(r).then(m=>m||fetch(r).then(res=>put(r,res))))});
"""
@app.get('/sw.js')
def pwa_sw():
    return Response(SW_JS.replace('__V__', ASSET_V), media_type='application/javascript', headers={'Cache-Control': 'no-cache', 'Service-Worker-Allowed': '/'})

@app.get('/.well-known/assetlinks.json')
def assetlinks():
    fps = [x.strip().upper() for x in os.environ.get('ANDROID_SHA256', '').split(',') if x.strip()]
    if not fps: raise HTTPException(404, 'Chưa cấu hình')
    pkg = os.environ.get('ANDROID_PACKAGE', 'vn.baobai.trolygiaoan').strip()
    return JSONResponse([{'relation': ['delegate_permission/common.handle_all_urls'], 'target': {'namespace': 'android_app', 'package_name': pkg, 'sha256_cert_fingerprints': fps}}])
