"""Đồng bộ thư mục data/ với một kho GitHub riêng tư (miễn phí, bền vững).

Biến môi trường: DATA_REPO (chu/kho), DATA_TOKEN (fine-grained PAT, Contents read/write),
DATA_BRANCH (mặc định main). Thiếu biến -> chỉ dùng ổ đĩa cục bộ.
Dùng Git Data API: mỗi lần đẩy = 1 commit (blob -> tree -> commit -> cập nhật ref không ép buộc),
nếu ref đã đổi (xung đột) thì làm lại trên commit mới nhất. Không bao giờ ghi token vào log.
"""
import os, json, base64, hashlib, threading, time, logging, urllib.request, urllib.error

log = logging.getLogger('storage')
PREFIX = 'data/'
SKIP_DIRS = {'out', '__pycache__', 'src', 'pages', 'pages_src'}  # src/pages: tệp gốc tải lên + ảnh xem trước (tạo lại được)

def git_sha(b):  # sha của blob git, so sánh không cần tải
    return hashlib.sha1(b'blob %d\0' % len(b) + b).hexdigest()

def wanted(rel):
    parts = rel.split('/')
    if any(p in SKIP_DIRS for p in parts[:-1]): return False
    n = parts[-1]
    return not (n.startswith('tmp_') or n.endswith('.tmp') or n.startswith('.~'))

class HTTPErr(Exception):
    def __init__(s, code, path): super().__init__(f'HTTP {code} {path}'); s.code = code

class GitSync:
    def __init__(s, root, repo, token, branch='main', api='https://api.github.com', debounce=3.0, max_wait=20.0):
        s.root, s.repo, s.branch, s.api = root, repo, branch, api.rstrip('/')
        s._tok = token; s.debounce, s.max_wait = debounce, max_wait
        s.remote = {}          # rel -> blob sha đã có trên kho
        s._known = {}          # rel -> (sha, mtime) tại lần đồng bộ gần nhất do chính ta ghi/đọc
        s.ready = False        # chỉ đẩy sau khi kéo về thành công (tránh xoá nhầm dữ liệu trên kho)
        s._lock = threading.Lock(); s._ev = threading.Event(); s._first = None; s._last = 0
        s._stop = False; s._th = None; s.pushes = 0; s.last_error = None

    def __repr__(s): return f'GitSync({s.repo}@{s.branch})'  # không lộ token

    # ---------- HTTP ----------
    def _req(s, method, path, body=None):
        url = f'{s.api}/repos/{s.repo}{path}'
        data = json.dumps(body).encode() if body is not None else None
        rq = urllib.request.Request(url, data=data, method=method, headers={
            'Authorization': 'Bearer ' + s._tok, 'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2022-11-28', 'User-Agent': 'bao-bai-giao-an',
            **({'Content-Type': 'application/json'} if data else {})})
        try:
            with urllib.request.urlopen(rq, timeout=60) as r:
                t = r.read(); return json.loads(t) if t else {}
        except urllib.error.HTTPError as e:
            raise HTTPErr(e.code, method + ' ' + path.split('?')[0]) from None

    def _head(s):
        try: return s._req('GET', f'/git/ref/heads/{s.branch}')['object']['sha']
        except HTTPErr as e:
            if e.code in (404, 409): return None   # kho/nhánh rỗng
            raise

    # ---------- kéo về ----------
    def pull(s):
        head = s._head(); remote = {}
        if head:
            tree = s._req('GET', f"/git/commits/{head}")['tree']['sha']
            for e in s._req('GET', f'/git/trees/{tree}?recursive=1').get('tree', []):
                if e.get('type') == 'blob' and e['path'].startswith(PREFIX):
                    remote[e['path'][len(PREFIX):]] = e['sha']
        n = 0; kept = 0
        for rel, sha in remote.items():
            if not wanted(rel) or '..' in rel.split('/'): continue
            p = os.path.join(s.root, *rel.split('/'))
            if os.path.exists(p):
                with open(p, 'rb') as f:
                    if git_sha(f.read()) == sha:
                        s._known[rel] = (sha, os.path.getmtime(p)); continue
                # File trên máy KHÁC kho: chỉ ghi đè khi file chưa hề bị đụng tới kể từ lần
                # đồng bộ gần nhất do chính ta ghi nhận. Ngược lại (không có dấu vết đồng bộ,
                # hoặc mtime mới hơn) = dữ liệu người dùng vừa ghi trên máy này mà kho chưa có
                # (vd lần kéo lúc khởi động bị lỗi, lần đẩy này mới kéo lại) -> TUYỆT ĐỐI giữ
                # bản trên máy; push() ngay sau sẽ tự đẩy nó lên kho vì sha khác kho.
                kn = s._known.get(rel)
                if not (kn and os.path.getmtime(p) <= kn[1]):
                    kept += 1
                    log.warning('storage: GIỮ bản trên máy của %s (mới hơn/chưa có trong kho) — không kéo đè', rel)
                    continue
            b = base64.b64decode(s._req('GET', f'/git/blobs/{sha}')['content'])
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p + '.tmp', 'wb') as f: f.write(b)
            os.replace(p + '.tmp', p); n += 1
            s._known[rel] = (sha, os.path.getmtime(p))
        s.remote = remote; s.ready = True
        log.info('storage: kéo %d tệp từ kho %s, giữ %d tệp mới trên máy', n, s.repo, kept)
        return n

    # ---------- đẩy lên ----------
    def local(s):
        out = {}
        for d, dirs, files in os.walk(s.root):
            dirs[:] = [x for x in dirs if x not in SKIP_DIRS]
            for fn in files:
                p = os.path.join(d, fn); rel = os.path.relpath(p, s.root).replace(os.sep, '/')
                if wanted(rel): out[rel] = p
        return out

    def push(s):
        if not s.ready: s.pull()
        with s._lock:
            loc = s.local(); blobs = {}; changes = []
            for rel, p in loc.items():
                try:
                    with open(p, 'rb') as f: b = f.read()
                except FileNotFoundError: continue
                sh = git_sha(b)
                if s.remote.get(rel) != sh: changes.append((rel, b, sh))
            gone = [r for r in s.remote if r not in loc]
            if not changes and not gone: return 0
            for rel, b, sh in changes:
                blobs[rel] = s._req('POST', '/git/blobs', {'content': base64.b64encode(b).decode(), 'encoding': 'base64'})['sha']
            for attempt in range(5):
                head = s._head()
                entries = [{'path': PREFIX + r, 'mode': '100644', 'type': 'blob', 'sha': blobs[r]} for r in blobs]
                entries += [{'path': PREFIX + r, 'mode': '100644', 'type': 'blob', 'sha': None} for r in gone]
                body = {'tree': entries}
                if head: body['base_tree'] = s._req('GET', f'/git/commits/{head}')['tree']['sha']
                tree = s._req('POST', '/git/trees', body)['sha']
                msg = f'data: {len(changes)} tệp đổi, {len(gone)} tệp xoá'
                commit = s._req('POST', '/git/commits', {'message': msg, 'tree': tree, 'parents': [head] if head else []})['sha']
                try:
                    if head: s._req('PATCH', f'/git/refs/heads/{s.branch}', {'sha': commit, 'force': False})
                    else: s._req('POST', '/git/refs', {'ref': f'refs/heads/{s.branch}', 'sha': commit})
                    break
                except HTTPErr as e:
                    if e.code in (409, 422) and attempt < 4: time.sleep(0.3 * (attempt + 1)); continue  # ref đã đổi -> làm lại
                    raise
            for rel, b, sh in changes:
                s.remote[rel] = blobs[rel]
                try: s._known[rel] = (blobs[rel], os.path.getmtime(loc[rel]))
                except OSError: pass
            for r in gone: s.remote.pop(r, None); s._known.pop(r, None)
            s.pushes += 1
            return len(changes) + len(gone)

    # ---------- nền: gom ghi trong vài giây ----------
    def schedule(s):
        now = time.time(); s._last = now
        if s._first is None: s._first = now
        s._ev.set()

    def start(s):
        if s._th: return
        s._th = threading.Thread(target=s._loop, daemon=True, name='gitsync'); s._th.start()

    def _loop(s):
        delay = 5
        while not s._stop:
            s._ev.wait(); 
            if s._stop: break
            while time.time() - s._last < s.debounce and time.time() - (s._first or 0) < s.max_wait: time.sleep(0.2)
            s._ev.clear(); s._first = None
            try: s.push(); s.last_error = None; delay = 5
            except Exception as e:
                s.last_error = str(e); log.warning('storage: đẩy lỗi (%s), thử lại sau %ss', e, delay)
                time.sleep(delay); delay = min(delay * 3, 300); s.schedule()

    def flush(s):
        try: s.push()
        except Exception as e: log.warning('storage: flush lỗi (%s)', e)

class LocalOnly:
    ready = True
    def schedule(s): pass
    def start(s): pass
    def flush(s): pass
    def pull(s): return 0

def from_env(root):
    repo, tok = os.environ.get('DATA_REPO', '').strip(), os.environ.get('DATA_TOKEN', '').strip()
    if not (repo and tok):
        log.warning('storage: ' + '=' * 70)
        log.warning('storage: CẢNH BÁO NGHIÊM TRỌNG: thiếu biến môi trường DATA_REPO hoặc DATA_TOKEN.')
        log.warning('storage: Toàn bộ dữ liệu (tài khoản, báo bài, giáo án, file tải lên...) hiện chỉ lưu trên')
        log.warning('storage: Ổ ĐĨA TẠM của Render và sẽ MẤT TRẮNG mỗi khi service khởi động lại / redeploy / sleep.')
        log.warning('storage: Cách khắc phục: Render Dashboard -> service web -> tab Environment -> thêm 2 biến')
        log.warning('storage:   DATA_REPO  = <chu/kho-du-lieu>   (VD: datle180693-byte/bao-bai-data)')
        log.warning('storage:   DATA_TOKEN = <github fine-grained PAT, quyền Contents Read/Write cho kho trên>')
        log.warning('storage: Xem README.md mục "Lưu dữ liệu bền vững miễn phí bằng kho GitHub riêng tư".')
        log.warning('storage: ' + '=' * 70)
        return LocalOnly()
    return GitSync(root, repo, tok, os.environ.get('DATA_BRANCH', 'main').strip() or 'main',
                   os.environ.get('DATA_API', 'https://api.github.com'),
                   float(os.environ.get('DATA_DEBOUNCE', '3')))
