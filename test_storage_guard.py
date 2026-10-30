"""Test hồi quy cho lỗi MẤT DỮ LIỆU báo bài (V73):
lần kéo dữ liệu lúc khởi động bị lỗi (ready=False) -> người dùng điền báo bài ->
lần đẩy đầu tiên kéo lại và ĐÈ mất bản gốc vừa điền bằng bản cũ trong kho.
Chạy: ../venv/bin/python test_storage_guard.py [đường-dẫn-storage.py]
Không truyền tham số = test bản storage.py hiện tại cạnh file này.
Truyền storage.py.bak-v73 = phải THẤT BẠI ở kịch bản 1 (chứng minh test bắt đúng lỗi).
"""
import os, sys, json, base64, hashlib, tempfile, importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'storage.py')
spec = importlib.util.spec_from_file_location('storage_under_test', target)
st = importlib.util.module_from_spec(spec); spec.loader.exec_module(st)


class FakeGitHub:
    """Mô hình tối thiểu của Git Data API: blobs/commits/trees/refs trong RAM."""
    def __init__(s):
        s.blobs = {}    # sha -> bytes
        s.trees = {}    # tree sha -> {path: blob sha}
        s.commits = {}  # commit sha -> tree sha
        s.ref = None    # commit sha
        s.fail_next = 0
        s._seq = 0

    def _sha(s, kind, payload):
        s._seq += 1
        return hashlib.sha1(f'{kind} {s._seq} '.encode() + payload).hexdigest()

    def seed_file(s, rel, data: bytes):
        bsha = st.git_sha(data); s.blobs[bsha] = data
        tree = dict(s.trees[s.commits[s.ref]]) if s.ref else {}
        tree['data/' + rel] = bsha
        tsha = s._sha('tree', json.dumps(tree, sort_keys=True).encode()); s.trees[tsha] = tree
        csha = s._sha('commit', tsha.encode()); s.commits[csha] = tsha; s.ref = csha

    def read_file(s, rel):
        tree = s.trees[s.commits[s.ref]]
        return s.blobs[tree['data/' + rel]]

    def req(s, method, path, body=None):
        if s.fail_next > 0:
            s.fail_next -= 1
            raise RuntimeError('mạng chập chờn (giả lập)')
        if method == 'GET' and path.startswith('/git/ref/heads/'):
            if s.ref is None: raise st.HTTPErr(404, path)
            return {'object': {'sha': s.ref}}
        if method == 'GET' and path.startswith('/git/commits/'):
            csha = path.split('/')[-1]
            return {'tree': {'sha': s.commits[csha]}}
        if method == 'GET' and path.startswith('/git/trees/'):
            tsha = path.split('/')[-1].split('?')[0]
            return {'tree': [{'path': p, 'type': 'blob', 'sha': b} for p, b in s.trees[tsha].items()]}
        if method == 'GET' and path.startswith('/git/blobs/'):
            bsha = path.split('/')[-1]
            return {'content': base64.b64encode(s.blobs[bsha]).decode()}
        if method == 'POST' and path == '/git/blobs':
            data = base64.b64decode(body['content']); bsha = st.git_sha(data)
            s.blobs[bsha] = data; return {'sha': bsha}
        if method == 'POST' and path == '/git/trees':
            tree = dict(s.trees[s.commits[s.ref]]) if (body.get('base_tree') and s.ref) else {}
            # base_tree là tree sha của commit head hiện tại
            if body.get('base_tree'):
                tree = dict(s.trees_by_sha(body['base_tree']))
            for e in body['tree']:
                if e['sha'] is None: tree.pop(e['path'], None)
                else: tree[e['path']] = e['sha']
            tsha = s._sha('tree', json.dumps(tree, sort_keys=True).encode()); s.trees[tsha] = tree
            return {'sha': tsha}
        if method == 'POST' and path == '/git/commits':
            csha = s._sha('commit', body['tree'].encode()); s.commits[csha] = body['tree']
            return {'sha': csha}
        if method == 'PATCH' and path.startswith('/git/refs/'):
            s.ref = body['sha']; return {}
        if method == 'POST' and path == '/git/refs':
            s.ref = body['sha']; return {}
        raise AssertionError(f'FakeGitHub chưa hỗ trợ: {method} {path}')

    def trees_by_sha(s, tsha): return s.trees[tsha]


def make_sync(root, fake):
    gs = st.GitSync(root, 'chu/kho', 'token-gia')
    gs._req = fake.req
    return gs


REL = 'users/u1/baobai_goc.docx'
fails = []

# ---- Kịch bản 1: kéo khởi động LỖI -> điền -> đẩy: bản đã điền phải SỐNG và được đẩy lên kho
root = tempfile.mkdtemp()
fake = FakeGitHub(); fake.seed_file(REL, b'BAN GOC CU (chua dien)')
gs = make_sync(root, fake)
fake.fail_next = 99                       # mạng sập lúc khởi động
try: gs.pull()
except Exception: pass                    # main.py nuốt lỗi này y như thật
fake.fail_next = 0
assert gs.ready is False, 'kịch bản yêu cầu ready=False sau lần kéo lỗi'
p = os.path.join(root, *REL.split('/')); os.makedirs(os.path.dirname(p), exist_ok=True)
with open(p, 'wb') as f: f.write(b'BAN GOC DA DIEN 27 BAI')   # _fill_auto os.replace vào MASTER
gs.push()                                  # lần đẩy đầu tiên (middleware gọi sau POST fill)
with open(p, 'rb') as f: local_now = f.read()
if local_now != b'BAN GOC DA DIEN 27 BAI':
    fails.append(f'Kịch bản 1: bản trên máy bị ĐÈ mất! còn lại: {local_now!r}')
elif fake.read_file(REL) != b'BAN GOC DA DIEN 27 BAI':
    fails.append('Kịch bản 1: bản trên máy còn nhưng kho chưa nhận được bản đã điền')
else:
    print('Kịch bản 1 OK: kéo lỗi lúc khởi động -> điền -> đẩy: dữ liệu sống, kho nhận đúng bản đã điền')

# ---- Kịch bản 2: khởi động bình thường, máy trắng -> kéo về phải khôi phục đủ
root2 = tempfile.mkdtemp()
fake2 = FakeGitHub(); fake2.seed_file(REL, b'BAN TRONG KHO')
fake2.seed_file('accounts.json', b'{"a": 1}')
gs2 = make_sync(root2, fake2)
gs2.pull()
ok = open(os.path.join(root2, *REL.split('/')), 'rb').read() == b'BAN TRONG KHO' \
    and open(os.path.join(root2, 'accounts.json'), 'rb').read() == b'{"a": 1}'
if ok: print('Kịch bản 2 OK: máy trắng kéo về khôi phục đủ tệp')
else: fails.append('Kịch bản 2: khôi phục thiếu tệp')

# ---- Kịch bản 3: file không đụng tới kể từ lần đồng bộ -> kho đổi (máy khác đẩy) -> kéo phải cập nhật
root3 = tempfile.mkdtemp()
fake3 = FakeGitHub(); fake3.seed_file(REL, b'V1')
gs3 = make_sync(root3, fake3); gs3.pull()
fake3.seed_file(REL, b'V2-moi-tu-phien-khac')
gs3.pull()
got = open(os.path.join(root3, *REL.split('/')), 'rb').read()
if got == b'V2-moi-tu-phien-khac': print('Kịch bản 3 OK: file không bị đụng -> kéo vẫn cập nhật từ kho')
else: fails.append(f'Kịch bản 3: file không được cập nhật, còn {got!r}')

# ---- Kịch bản 4: điền tiếp SAU khi đã đồng bộ khỏe -> đẩy bình thường (luồng chuẩn không được vỡ)
gs3_path = os.path.join(root3, *REL.split('/'))
with open(gs3_path, 'wb') as f: f.write(b'V3-nguoi-dung-dien-tiep')
gs3.push()
if fake3.read_file(REL) == b'V3-nguoi-dung-dien-tiep': print('Kịch bản 4 OK: luồng điền/đẩy chuẩn vẫn chạy')
else: fails.append('Kịch bản 4: đẩy sau khi đồng bộ khỏe bị hỏng')

print()
if fails:
    print('THẤT BẠI:'); [print(' -', x) for x in fails]; sys.exit(1)
print('TẤT CẢ OK trên', os.path.basename(target))
