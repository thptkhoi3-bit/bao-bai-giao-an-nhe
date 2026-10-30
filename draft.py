"""Bản nháp giáo án phụ theo từng tuần, từng người dùng.
Mỗi tệp phụ tải lên được lưu lại; bài trùng (cùng môn + cùng bài) -> bản mới nhất thắng, phần còn lại giữ nguyên."""
import os, re, json, uuid, shutil, datetime, hashlib, logging, copy
import core

NAME = dict(tv='Tiếng Việt', toan='Toán', ontv='Ôn TV', ontoan='Ôn Toán', hdtn='HĐTN', tnxh='TNXH', daoduc='Đạo đức', thuvien='Thư viện',
            chamsoc='Chăm sóc HS', gdtc='GDTC', amnhac='Âm nhạc', mithuat='Mĩ thuật', tinhoc='Tin học', tienganh='Tiếng Anh', clb='CLB')
def sname(k): return NAME.get(k, k or '?')

def ddir(u, w):
    d = os.path.join(u.DATA, 'draft', f'w{int(w)}'); os.makedirs(d, exist_ok=True); return d

def _jp(d): return os.path.join(d, 'draft.json')

def load(u, w):
    d = ddir(u, w)
    try:
        with open(_jp(d), encoding='utf-8') as f: j = json.load(f)
    except Exception: j = {'files': []}
    fs = []
    for x in j.get('files', []) if isinstance(j, dict) else []:
        if not (isinstance(x, dict) and x.get('id') and os.path.exists(os.path.join(d, x['id'] + '.docx'))): continue
        if not isinstance(x.get('lessons'), list):  # dữ liệu cũ / đồng bộ dở: đọc lại danh sách bài
            try: x['lessons'] = lessons_of(os.path.join(d, x['id'] + '.docx'))
            except Exception: continue
        x.setdefault('name', x['id'] + '.docx'); x.setdefault('added', ''); fs.append(x)
    return dict(j if isinstance(j, dict) else {}, files=fs)

def migrate_old(u, w):
    """Dữ liệu từ trước khi có bản nháp: giáo án phụ chỉ nằm trong out/src_<id>/ của lần tạo gần nhất (không có ex.json).
    Nếu tuần này chưa từng có bản nháp mà lần tạo cũ đúng tuần này -> chép các tệp đó vào bản nháp (1 lần)."""
    d = os.path.join(u.DATA, 'draft', f'w{int(w)}')
    if os.path.exists(_jp(d)): return 0
    try:
        with open(os.path.join(u.OUT, 'state.json'), encoding='utf-8') as f: st = json.load(f)
    except Exception: return 0
    if str(st.get('week')) != str(int(w)) or not st.get('src'): return 0
    src = os.path.join(u.OUT, 'src_' + re.sub(r'[^0-9a-f]', '', str(st['src'])))
    if not os.path.isdir(src) or os.path.exists(os.path.join(src, 'ex.json')): return 0
    n = 0
    for k, fn in enumerate(sorted(x for x in os.listdir(src) if x.endswith('.docx'))):
        try: add(u, w, os.path.join(src, fn), f'giáo án phụ đã tải trước đây ({k + 1}).docx'); n += 1
        except Exception as e: logging.getLogger('app').warning('migrate_old: bỏ qua %s: %r', fn, e)
    return n

def save(u, w, j):
    d = ddir(u, w); tmp = _jp(d) + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(j, f, ensure_ascii=False, indent=1)
    os.replace(tmp, _jp(d))

def _sig(title):
    a, t = core.toks(title); return a, t

def same(l1, l2):
    if (l1['subject'] or '') != (l2['subject'] or ''): return False
    a1, t1 = _sig(l1['title']); a2, t2 = _sig(l2['title'])
    if not a1 or not a2: return l1['title'].strip().lower() == l2['title'].strip().lower()
    if t1 != t2: return False
    return 2 * len(a1 & a2) / (len(a1) + len(a2)) >= 0.75

def lessons_of(path):
    _, ls = core.split_lessons(core.docx.Document(path))
    return [dict(i=l['idx'], subject=l['subject'], title=l['title'][:200]) for l in ls]

def resolve(j):
    """-> (active [(file, lesson)], exclude {fid: set(i)}, replaced_by {(fid,i): fid_moi})
    Bài thuộc MÔN ĐÃ BỎ (j['dropped'] – người dùng tải nhầm môn) không được tính là active."""
    dropped = set(j.get('dropped') or [])
    active = []; exclude = {}
    for f in j['files']:
        for l in f['lessons']:
            if (l['subject'] or '') in dropped: continue
            for k, (f0, l0) in enumerate(active):
                if f0['id'] != f['id'] and same(l0, l):
                    exclude.setdefault(f0['id'], set()).add(l0['i']); active[k] = None
            active = [x for x in active if x]
            active.append((f, l))
    return active, exclude

def drop_subject(u, w, key):
    """Bỏ 1 môn khỏi bản nháp (tải nhầm) – bài vẫn lưu trong tệp, chỉ không dùng khi ghép; hoàn tác được."""
    j = load(u, w); d = set(j.get('dropped') or [])
    d.add(key); j['dropped'] = sorted(d); save(u, w, j); return True

def undrop_subject(u, w, key):
    j = load(u, w); d = set(j.get('dropped') or [])
    d.discard(key); j['dropped'] = sorted(d); save(u, w, j); return True

def add(u, w, src_path, name):
    j = load(u, w); d = ddir(u, w); fid = uuid.uuid4().hex[:10]
    dst = os.path.join(d, fid + '.docx'); shutil.copy(src_path, dst)
    ls = lessons_of(dst)
    before, _ = resolve(j)
    added, replaced = [], []
    for l in ls:
        hit = next((l0 for f0, l0 in before if same(l0, l)), None)
        (replaced if hit else added).append(f"[{sname(l['subject'])}] {l['title']}")
    with open(dst, 'rb') as fh: sha = hashlib.sha1(fh.read()).hexdigest()
    for old in [x for x in j['files'] if x.get('sha') == sha]:  # tải lại đúng tệp cũ -> thay tệp cũ
        j['files'].remove(old)
        try: os.remove(os.path.join(d, old['id'] + '.docx'))
        except OSError: pass
    j['files'].append(dict(id=fid, name=os.path.basename(name or 'giao_an_phu.docx'), sha=sha, added=datetime.datetime.now().isoformat(timespec='seconds'), lessons=ls))
    save(u, w, j)
    return dict(file=fid, name=name, added=added, replaced=replaced, n=len(ls))

def remove(u, w, fid):
    j = load(u, w); d = ddir(u, w); fid = re.sub(r'[^0-9a-f]', '', fid)
    n = len(j['files']); j['files'] = [x for x in j['files'] if x['id'] != fid]
    if len(j['files']) == n: return False
    try: os.remove(os.path.join(d, fid + '.docx'))
    except OSError: pass
    save(u, w, j); return True

def clear(u, w):
    shutil.rmtree(ddir(u, w), ignore_errors=True)

def snapshot(u, w, dst):
    """Chép bản nháp sang thư mục phiên ghép (000.docx…) + ex.json."""
    ps, ex = paths(u, w); os.makedirs(dst, exist_ok=True); m = {}
    for i, p in enumerate(ps):
        n = f'{i:03d}.docx'; shutil.copy(p, os.path.join(dst, n))
        if p in ex: m[n] = sorted(ex[p])
    with open(os.path.join(dst, 'ex.json'), 'w') as f: json.dump(m, f)

def paths(u, w):
    """Đường dẫn tệp theo thứ tự thêm + exclude theo đường dẫn – dùng cho build_giaoan."""
    j = load(u, w); d = ddir(u, w); _, ex = resolve(j)
    ps = [os.path.join(d, f['id'] + '.docx') for f in j['files']]
    return ps, {os.path.join(d, k + '.docx'): v for k, v in ex.items()}

_SUM_CACHE = {}
def _sum_key(u, w, template, master):
    try:
        j = load(u, w)
        parts = [str(u.DATA), str(w), json.dumps(j, sort_keys=True, ensure_ascii=False)]
        for p in (template, master):
            if p and os.path.exists(p):
                st = os.stat(p); parts.append('%s:%d:%d' % (p, st.st_mtime_ns, st.st_size))
            else: parts.append(str(p))
        try: parts.append(json.dumps(core.CMAP.get(), sort_keys=True))
        except Exception: pass
        return hashlib.sha1('|'.join(parts).encode('utf-8')).hexdigest()
    except Exception:
        return None

def summary(u, w, template=None, master=None):
    key = _sum_key(u, w, template, master)
    if key and key in _SUM_CACHE: return copy.deepcopy(_SUM_CACHE[key])
    j = load(u, w); active, ex = resolve(j)
    dropped = set(j.get('dropped') or [])
    files = []
    for f in j['files']:
        gone = set(ex.get(f['id'], set())) | {l['i'] for l in f['lessons'] if (l['subject'] or '') in dropped}
        files.append(dict(id=f['id'], name=f['name'], added=f['added'], n=len(f['lessons']),
                          active=len(f['lessons']) - len(gone), replaced=len(gone)))
    lessons = [dict(subject=sname(l['subject']), title=l['title'], file=f['name']) for f, l in active]
    subjects = sorted({l['subject'] for l in lessons})
    from collections import Counter
    cnt = Counter((l['subject'] or '') for f, l in active)
    subjects2 = [dict(k=k2, name=sname(k2), n=cnt[k2]) for k2 in sorted(cnt, key=lambda x: sname(x))]
    dcnt = Counter()
    for f in j['files']:
        for l in f['lessons']:
            if (l['subject'] or '') in dropped: dcnt[l['subject'] or ''] += 1
    dropped_list = [dict(k=k2, name=sname(k2), n=dcnt[k2]) for k2 in sorted(dcnt, key=lambda x: sname(x))]
    tiet = None
    if template and master and os.path.exists(template) and os.path.exists(master):
        try: tiet = coverage(u, w, template, master)
        except Exception: tiet = None
    r = dict(week=int(w), files=files, lessons=lessons, subjects=subjects, subjects2=subjects2,
             dropped=dropped_list, tiet=tiet)
    if key:
        if len(_SUM_CACHE) > 128: _SUM_CACHE.clear()
        _SUM_CACHE[key] = copy.deepcopy(r)
    return r

def coverage(u, w, template, master):
    """TIẾT nào của tuần đã có bài trong bản nháp (chỉ so khớp, không ghi tệp)."""
    G = core.docx.Document(template); M = core.docx.Document(master)
    mw = core.week_tables(M, int(w))
    if w not in mw: return None
    core.step1_copy_table(G, M, w)
    ps, ex = paths(u, w)
    secs = core.load_secs(ps, ex) if ps else []
    heads, _, _ = core.assign(G, M, secs)
    out = []
    for h in heads:
        out.append(dict(thu=h['thu'], tiet=h['label'], ten=h['ten'], filled=bool(h['lessons']), need=bool(h['ten'])))
    return out
