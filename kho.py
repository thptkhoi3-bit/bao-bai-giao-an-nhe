"""Kho lưu theo tuần: mọi tệp giáo viên tạo ra được lưu theo người dùng, theo tuần.
<DATA_DIR>/kho/<user_key>/<w07|khac>/<loại>.<đuôi> + <DATA_DIR>/kho/<user_key>/index.json (storage.py đồng bộ lên kho dữ liệu).
- Cùng loại + cùng tuần: bản mới thay bản cũ (chỉ giữ bản mới nhất).
- Giữ tối đa KHO_WEEKS (10) tuần gần nhất mỗi giáo viên; tổng dung lượng tối đa KHO_MAX_MB (50) MB – bỏ tệp cũ nhất trước.
- Chỉ đọc/ghi trong thư mục của chính người dùng (mã tệp được kiểm tra chặt)."""
import os, re, json, shutil, threading, datetime, logging, uuid
VN = datetime.timezone(datetime.timedelta(hours=7))
TYPES = {'baobai': 'Báo bài đã điền', 'giaoan': 'Giáo án', 'giaoan_th': 'Giáo án tích hợp', 'de': 'Đề kiểm tra',
         'nhanxet_xlsx': 'Nhận xét học sinh (Excel)', 'nhanxet_docx': 'Nhận xét học sinh (Word)',
         'phieu_bai_tap': 'Phiếu bài tập', 'tro_choi': 'Trò chơi khởi động', 'tro_choi_web': 'Trò chơi tương tác',
         'am_thanh': 'Âm thanh giọng đọc'}
EXT = {'nhanxet_xlsx': '.xlsx', 'tro_choi_web': '.json', 'am_thanh': '.wav'}
MIME = {'.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.json': 'application/json'}
log = logging.getLogger('kho'); _lock = threading.Lock()
def max_weeks(): return max(1, int(os.environ.get('KHO_WEEKS', '10')))
def max_bytes(): return int(float(os.environ.get('KHO_MAX_MB', '50')) * 1024 * 1024)
def _safe_key(k):
    k = str(k or '')
    if not re.fullmatch(r'[a-z0-9][a-z0-9\-]{0,80}', k): raise ValueError('mã người dùng không hợp lệ')
    return k
def udir(root, key): return os.path.join(root, 'kho', _safe_key(key))
def wkey(week):
    try: n = int(week)
    except (TypeError, ValueError): return 'khac'
    return f'w{n:02d}' if 1 <= n <= 60 else 'khac'
def wlabel(wk): return 'Khác' if wk == 'khac' else f'Tuần {int(wk[1:])}'
def _idx(root, key): return os.path.join(udir(root, key), 'index.json')
def load(root, key):
    try:
        with open(_idx(root, key), encoding='utf-8') as f: j = json.load(f)
        items = j.get('items', []) if isinstance(j, dict) else []
    except FileNotFoundError: items = []
    except Exception as e: log.warning('kho: index hỏng (%r) – dựng lại từ tệp', e); items = []
    d = udir(root, key)
    return [x for x in items if isinstance(x, dict) and x.get('id') and x.get('file') and os.path.exists(os.path.join(d, x['file']))]
def _save(root, key, items):
    os.makedirs(udir(root, key), exist_ok=True); p = _idx(root, key); tmp = p + '.' + uuid.uuid4().hex[:6] + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump({'items': items}, f, ensure_ascii=False, indent=1)
    os.replace(tmp, p)
def _rm(root, key, it):
    try: os.remove(os.path.join(udir(root, key), it['file']))
    except OSError: pass
    dd = os.path.join(udir(root, key), it['week'])
    try:
        if os.path.isdir(dd) and not os.listdir(dd): os.rmdir(dd)
    except OSError: pass

def put(root, key, week, typ, src, name, now=None, tag=''):
    """Lưu bản sao tệp vừa tạo vào kho. Trả về mục đã lưu (dict).
    tag: phân biệt nhiều tệp cùng loại trong 1 tuần (VD phiếu của từng bài) – chỉ thay bản cũ cùng loại + cùng tuần + cùng tag."""
    tag = re.sub(r'[^a-z0-9\-]+', '', str(tag or '').lower())[:60]
    if typ not in TYPES: raise ValueError('loại tệp không hợp lệ')
    wk = wkey(week); ext = EXT.get(typ, '.docx'); d = udir(root, key)
    now = now or datetime.datetime.now(VN)
    with _lock:
        items = load(root, key)
        for it in [x for x in items if x['week'] == wk and x['type'] == typ and x.get('tag', '') == tag]: items.remove(it); _rm(root, key, it)
        os.makedirs(os.path.join(d, wk), exist_ok=True)
        fid = uuid.uuid4().hex[:12]; rel = f'{wk}/{typ}_{fid}{ext}'
        tmp = os.path.join(d, rel + '.tmp'); shutil.copyfile(src, tmp); os.replace(tmp, os.path.join(d, rel))
        it = dict(id=fid, week=wk, type=typ, name=re.sub(r'[\\/:*?"<>|\r\n]+', '_', str(name))[:120] or (typ + ext), file=rel,
                  created=now.isoformat(timespec='seconds'), tag=tag, size=os.path.getsize(os.path.join(d, rel)))
        items.append(it); items = _retain(root, key, items, keep=fid); _save(root, key, items)
        return it

def _retain(root, key, items, keep=None):
    # 1) chỉ giữ N tuần gần nhất (theo lần tạo mới nhất trong tuần)
    last = {}
    for x in items: last[x['week']] = max(last.get(x['week'], ''), x['created'])
    weeks = sorted(last, key=lambda w: last[w], reverse=True)
    drop = set(weeks[max_weeks():])
    out = []
    for x in items:
        if x['week'] in drop: _rm(root, key, x)
        else: out.append(x)
    # 2) tổng dung lượng: bỏ tệp cũ nhất trước (không bỏ tệp vừa tạo)
    out.sort(key=lambda x: x['created']); tot = sum(int(x.get('size') or 0) for x in out)
    while tot > max_bytes() and len(out) > 1:
        victim = next((x for x in out if x['id'] != keep), None)
        if not victim: break
        out.remove(victim); _rm(root, key, victim); tot -= int(victim.get('size') or 0)
    return out

def find(root, key, fid):
    if not re.fullmatch(r'[0-9a-f]{12}', str(fid or '')): return None
    it = next((x for x in load(root, key) if x['id'] == fid), None)
    if not it: return None
    p = os.path.realpath(os.path.join(udir(root, key), it['file']))
    if not p.startswith(os.path.realpath(udir(root, key)) + os.sep): return None
    return it, p

def delete(root, key, fid):
    with _lock:
        items = load(root, key); it = next((x for x in items if x['id'] == fid), None)
        if not it: return False
        items.remove(it); _rm(root, key, it); _save(root, key, items); return True

def listing(root, key):
    items = load(root, key); groups = {}
    for x in sorted(items, key=lambda x: x['created'], reverse=True):
        groups.setdefault(x['week'], []).append(dict(id=x['id'], type=x['type'], type_vi=TYPES.get(x['type'], x['type']), name=x['name'],
                                                     created=x['created'], size=x.get('size', 0)))
    order = sorted(groups, key=lambda w: (w == 'khac', -int(w[1:]) if w != 'khac' else 0))
    return dict(groups=[dict(week=w, label=wlabel(w), items=groups[w]) for w in order], total=sum(int(x.get('size') or 0) for x in items),
                cap=max_bytes(), max_weeks=max_weeks())
