"""Nhật ký sử dụng theo ngày (giờ Việt Nam) – cho bản tóm tắt hằng đêm gửi trợ lý 'Quản Gia'.
Tệp: <DATA_DIR>/usage/YYYY-MM-DD.json, được storage.py đồng bộ lên kho dữ liệu riêng như mọi tệp khác.
Chỉ lưu: tên hiển thị, mã tài khoản, giờ đăng nhập, số lần dùng từng chức năng, yêu cầu đăng ký mới (tên, trường).
KHÔNG lưu mật khẩu, số điện thoại, IP hay nội dung tệp.
Ghi nhận = thêm vào bộ đệm trong RAM (không chạm ổ đĩa) -> luồng nền gộp vào tệp mỗi vài giây (khoá 1 chỗ, không mất số liệu)."""
import os, json, time, threading, logging, datetime, collections
VN = datetime.timezone(datetime.timedelta(hours=7))
ACTIONS = {'dien_bao_bai': 'điền báo bài', 'tao_giao_an': 'tạo giáo án', 'tich_hop': 'tích hợp', 'tao_de': 'tạo đề', 'chat': 'tin nhắn trợ lý', 'nhan_xet': 'nhận xét học sinh', 'phieu_bai_tap': 'phiếu bài tập', 'tro_choi': 'trò chơi', 'doc_van_ban': 'đọc văn bản thành giọng nói', 'cong_cu_lop': 'công cụ lớp học'}
MAX_TIMES = 30          # giữ tối đa N giờ đăng nhập mỗi người mỗi ngày
log = logging.getLogger('usage')

def now_vn(): return datetime.datetime.now(VN)

class Usage:
    def __init__(s, root, sync=None, interval=None, clock=now_vn):
        s.dir = os.path.join(root, 'usage'); s.sync = sync; s.clock = clock
        s.interval = float(os.environ.get('USAGE_FLUSH', '5') if interval is None else interval)
        s._buf = collections.deque(maxlen=20000); s._lock = threading.Lock(); s._flock = threading.Lock()
        s._ev = threading.Event(); s._th = None; s._stop = False

    # ---------- ghi nhận (rất nhẹ, gọi trong request) ----------
    def _add(s, ev):
        try:
            t = s.clock(); ev['d'] = t.date().isoformat(); ev['t'] = t.strftime('%H:%M:%S')
            with s._lock: s._buf.append(ev)
            s._ev.set()
        except Exception as e: log.warning('usage: bỏ qua sự kiện (%r)', e)
    def login(s, key, name): s._add(dict(k='login', id=str(key)[:80], name=str(name)[:100]))
    def action(s, key, name, act):
        if act in ACTIONS: s._add(dict(k='act', id=str(key)[:80], name=str(name)[:100], a=act))
    def registration(s, aid, name, school=''): s._add(dict(k='reg', id=str(aid)[:80], name=str(name)[:100], school=str(school or '')[:120]))

    # ---------- gộp vào tệp ----------
    def path(s, d): return os.path.join(s.dir, f'{d}.json')
    def load(s, d):
        try:
            with open(s.path(d), encoding='utf-8') as f: j = json.load(f)
            if isinstance(j, dict): return j
        except FileNotFoundError: pass
        except Exception as e: log.warning('usage: tệp %s hỏng (%r), ghi tệp mới', d, e)
        return {}
    @staticmethod
    def merge(j, d, evs):
        j.setdefault('date', d); j['tz'] = '+07:00'; us = j.setdefault('users', {}); regs = j.setdefault('registrations', [])
        for e in evs:
            if e['k'] == 'reg':
                if not any(r.get('id') == e['id'] for r in regs): regs.append(dict(id=e['id'], name=e['name'], school=e['school'], time=e['t']))
                continue
            u = us.setdefault(e['id'], dict(name=e['name'], logins=0, first_login=None, last_login=None, login_times=[], actions={}, first_seen=e['t'], last_seen=e['t']))
            u['name'] = e['name'] or u.get('name'); u['last_seen'] = max(u.get('last_seen') or e['t'], e['t']); u['first_seen'] = min(u.get('first_seen') or e['t'], e['t'])
            if e['k'] == 'login':
                u['logins'] = int(u.get('logins', 0)) + 1
                u['first_login'] = min(u['first_login'] or e['t'], e['t']); u['last_login'] = max(u['last_login'] or e['t'], e['t'])
                lt = u.setdefault('login_times', []); lt.append(e['t']); del lt[:-MAX_TIMES]
            else:
                a = u.setdefault('actions', {}); a[e['a']] = int(a.get(e['a'], 0)) + 1
        j['updated'] = now_vn().isoformat(timespec='seconds')
        return j
    def flush(s):
        with s._flock:   # chỉ 1 nơi đọc-gộp-ghi -> không mất số liệu
            with s._lock: evs = list(s._buf); s._buf.clear()
            if not evs: return 0
            try:
                if s.sync is not None and getattr(s.sync, 'ready', True) is False:
                    try: s.sync.pull()   # lấy bản trên kho trước, tránh bị bản cũ đè lên số liệu mới
                    except Exception as e: log.warning('usage: chưa kéo được dữ liệu (%s)', e)
                by = collections.OrderedDict()
                for e in evs: by.setdefault(e['d'], []).append(e)
                os.makedirs(s.dir, exist_ok=True)
                for d, es in by.items():
                    j = s.merge(s.load(d), d, es); tmp = s.path(d) + '.tmp'
                    with open(tmp, 'w', encoding='utf-8') as f: json.dump(j, f, ensure_ascii=False, indent=1)
                    os.replace(tmp, s.path(d))
            except Exception as e:
                log.warning('usage: ghi lỗi (%r), giữ lại để ghi sau', e)
                with s._lock: s._buf.extendleft(reversed(evs))
                return 0
            if s.sync is not None:
                try: s.sync.schedule()
                except Exception: pass
            return len(evs)

    # ---------- luồng nền ----------
    def start(s):
        if s._th: return
        s._th = threading.Thread(target=s._loop, daemon=True, name='usage'); s._th.start()
    def _loop(s):
        while not s._stop:
            s._ev.wait(); s._ev.clear()
            if s._stop: break
            time.sleep(s.interval)
            s.flush()
    def stop(s):
        s._stop = True; s._ev.set(); s.flush()
