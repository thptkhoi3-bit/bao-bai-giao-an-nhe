"""PWA: manifest, service worker (chỉ cache vỏ tĩnh, không bao giờ /api), assetlinks, nút Cài ứng dụng ẩn mặc định."""
import os, sys, json
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import tempfile, subprocess, atexit, time, requests
def server(port, extra):
    dd = tempfile.mkdtemp(); env = dict(os.environ, DATA_DIR=dd, SECRET_KEY='t' * 32,
        TEACHER_PASSWORD_SHA256='feb9a081d082f4d8d31fa3c51cd960922c29d6fb9440d1ac1a00a8c59bde3cdb',
        GUEST_PASSWORD_SHA256='4fcacb897dec6fd766c0538c222a2f58cfa5e10f13d682facac89f81d02e7011', **extra)
    for k in ('DATA_REPO', 'DATA_TOKEN', 'GEMINI_API_KEY', 'ANDROID_SHA256'): env.pop(k, None)
    env.update(extra)
    p = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'main:app', '--port', str(port)], cwd=HERE, env=env, stdout=open(os.path.join(dd, 'log'), 'w'), stderr=subprocess.STDOUT); atexit.register(p.terminate)
    for _ in range(100):
        try: requests.get(f'http://127.0.0.1:{port}', timeout=1); return p
        except Exception: time.sleep(.2)
class C:
    def __init__(s, b): s.b = b
    def get(s, path): return requests.get(s.b + path, timeout=10)
server(8097, {}); server(8098, {'ANDROID_SHA256': 'aa:bb:cc'})
c = C('http://127.0.0.1:8097'); c2 = C('http://127.0.0.1:8098')
r = c.get('/manifest.webmanifest'); assert r.status_code == 200 and 'manifest+json' in r.headers['content-type']
m = r.json(); assert m['name'] == 'Trợ lý Báo bài & Giáo án' and m['short_name'] == 'Báo bài' and m['lang'].startswith('vi')
assert m['display'] == 'standalone' and m['scope'] == '/'
sizes = {(i['sizes'], i.get('purpose', 'any')) for i in m['icons']}
assert ('192x192', 'any') in sizes and ('512x512', 'any') in sizes and any(p == 'maskable' for _, p in sizes)
for i in m['icons']: assert c.get(i['src']).status_code == 200, i['src']
r = c.get('/sw.js'); assert r.status_code == 200 and 'javascript' in r.headers['content-type']
assert r.headers.get('service-worker-allowed') == '/' and 'no-cache' in r.headers.get('cache-control', '')
sw = r.text; assert '__V__' not in sw and "startsWith('/static/')" in sw and "'/api" not in sw.split('*/',1)[1] and '/static/offline.html' in sw and 'bb-shell-v2-' in sw and "fetch(r,{cache:'no-cache'})" in sw
assert c.get('/static/offline.html').status_code == 200
assert c.get('/.well-known/assetlinks.json').status_code == 404
a = c2.get('/.well-known/assetlinks.json').json()
assert a[0]['target']['package_name'] == 'vn.baobai.trolygiaoan' and a[0]['target']['sha256_cert_fingerprints'] == ['AA:BB:CC']
html = c.get('/').text
assert 'rel="manifest"' in html and 'id="instbtn"' in html and 'class="thb hide" type="button" id="instbtn"' in html and 'BaoBaiApp' in html
print('test_pwa OK')
