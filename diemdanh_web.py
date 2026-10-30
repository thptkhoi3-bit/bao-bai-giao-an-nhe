"""Diem danh da phong + Bao com + Tong hop cuoi thang (PWA /diem-danh/).
Kien truc: du lieu theo tung phong (rooms collection). Moi phong co danh sach hoc sinh +
ngay phuc vu + marks + ngay da bao RIENG. Chi seed Phong 5 that; phong khac seed them
qua API (admin) ma khong sua code.

LUAT GIU NGUYEN (tu brief cu):
- 0 = vang (khong tinh com); chua danh = "Cho diem danh" (van tinh an); suat = tong - vang.
- Cham doi vong Cho -> An -> Vang -> Cho, POST luu ngay, tai lai van giu.
- Ngay phuc vu troi qua khong bao = nghi (0 suat). Bao xong: ngay thanh "Da bao",
  ngay chon nhay sang ngay phuc vu tiep theo.
- Nut Bao com / man bao cao / xuat file CHI tai khoan quan tri (admin).
- Ten do 6 hoc sinh dia phuong 3A3 (STT 48-53, chinh ta file goc); dong loc lop kem GVCN;
  mau trang thai lop: con cho = do, du = xanh (tieu de nhom + cham tren chip loc).

LUAT MOI (da phong + MK phong + tong hop):
- Sau dang nhap: quan tri thay tat ca phong; giao vien chi thay phong minh duoc gan
  (truong managers tren phong, quan tri sua duoc qua UI).
- Moi phong co 1 mat khau rieng, luu dang sha256 hex (khong plaintext, khong MK mac dinh).
  Vao khu diem danh/bao com cua phong nao phai nhap dung MK phong do. Token mo khoa
  luu o server, rang buoc theo (tai khoan + cookie phien dang nhap) -> dang xuat roi
  dang nhap lai phai nhap lai MK; dung token cua tai khoan khac bi tu choi.
- Phong chua co MK: lan dau vao hien man "Dat mat khau phong" (chi quan ly phong +
  quan tri duoc dat). Quan tri dat lai MK khong can biet MK cu; quan ly phong doi MK
  cu phai nhap dung MK cu.
- Quan tri xem TONG QUAN chi-doc cua moi phong (danh sach phong, tong hop) KHONG can MK;
  nhung muon VAO khu diem danh/bao com thi quan tri cung phai nhap MK phong do.
  Bao com ben trong phong da vao thi khong hoi lai MK.
- THEM HOC SINH / THEM DANH SACH (thay dan 09:36-09:37): chi quan ly phong + quan tri.
  Them le trong phong da mo khoa: ten + lop (chon lop cua phong hoac go lop moi) +
  GVCN tu lay theo lop; STT noi tiep; trang thai ban dau Cho.
  Them danh sach: (a) dan moi dong "Ho ten — Lop" kem xem truoc roi xac nhan;
  (b) tai file Excel dung dinh dang file goc (STT, Ho va ten, Lop, cac cot ngay, SDT,
  Ghi chu; bo qua hang tong ket "Vang"/"So luong bao com"). Tu man chon phong -> tao
  phong moi (sau do qua luong dat mat khau phong); tu trong phong -> them noi tiep.
  QUY TAC NGAY BAT DAU: hoc sinh moi co start_day (mac dinh ngay them, sua duoc);
  CHI tinh diem danh + tien an tu ngay do tro di — cac ngay truoc hoan toan khong tinh
  (khong tang suat ngay cu, khong tinh tien ngay cu). Trong trua la phi thang nen van
  tinh tron thang (thay muon giam thi sua rieng ve 0 nho tinh nang sua trong trua).
  BANG CO DUNG 4 COT: Tong so buoi an | Tien an | Trong trua | Tong tien.
  Tien an = buoi an x 25.000d/suat. Trong trua = phi thang co dinh theo hoc sinh.
  GIA CHUNG (thay chot 09:31-09:32): mot gia an + mot muc trong trua chung cho MOI phong,
  doi o mot cho la ap cho tat ca (khong gia rieng tung phong). Mac dinh: gia an 25.000d,
  trong trua chung 120.000d/thang/hoc sinh (muc thay chot, thay cho gia dinh 50.000d cu).
  TRONG TRUA RIENG TUNG HOC SINH: quan ly phong sua duoc rieng cho tung hoc sinh
  (mien/giam, ke ca ve 0); muc hieu dung = muc rieng (neu da dat) neu khong thi muc chung.
  Tong tien = Tien an + Trong trua (hieu dung). Cuoi moi lop 1 dong tong lop (cong du 4 cot),
  cuoi phong 1 dong tong toan phong. Xuat CSV kem theo.
"""
import base64, csv, hashlib, hmac, io, json, os, re, secrets, time
from datetime import datetime, timedelta, timezone
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, Response, StreamingResponse
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    HAVE_XLSX = True
except Exception:
    HAVE_XLSX = False

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT = os.environ.get('DATA_DIR', os.path.join(BASE, 'data'))
DD_DIR = os.path.join(DATA_ROOT, 'diemdanh'); os.makedirs(DD_DIR, exist_ok=True)
STATE_F = os.path.join(DD_DIR, 'state.json')
ACCOUNTS = os.path.join(DATA_ROOT, 'accounts.json')
SECRET = os.environ.get('SECRET_KEY', 'dev-secret')
ADMIN_KEY = os.environ.get('ADMIN_KEY', 'gv-anh-le-dat')
COOKIE = 'gv_session'
PRICE_DEFAULT = 25000
SETTINGS_DEFAULT = {'meal_price': 25000, 'lunch_fee': 120000}
# Hoc sinh dia phuong (3A3, STT 48-53): chi doi MAU TEN khi render.
# Chinh ta theo file Excel goc ("Tien", khong phai "Tieu" nhu anh thay gui).
RED_NAMES = {'Phạm Xuân Hùng', 'Ohamj Tiến Minh Quang', 'Trần Hải Đăng', 'Nguyễn Hữu An', 'Nguyễn Minh Nhật', 'Phạm Duy Thái'}
SEED = json.loads("""{"room":"Phòng 5","month":10,"year":2026,"days":[1,2,5,6,8,9,12,13,15,16,19,20,22,23,26,27,29,30],"students":[{"stt":1,"name":"Sùng A Hiếu","class":"3A2","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Hưng"},{"stt":2,"name":"Vừ A Lầu","class":"3A2","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":0,"30":null},"note":null},{"stt":3,"name":"Giàng A Cân","class":"3A2","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":0,"22":0,"23":0,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":4,"name":"Hạng Thanh Tùng","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Hoa"},{"stt":5,"name":"Tráng Hồng Hữu","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":6,"name":"Cứ A Mạnh","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":7,"name":"Sùng A Chua","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":0,"20":0,"22":0,"23":0,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":8,"name":"Mùa Gia Huy","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":9,"name":"Nguyễn Hữu Nhân Đức","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":10,"name":"Hoàng Văn Thánh","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":11,"name":"Hầu Minh Phúc","class":"4A2","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Hằng"},{"stt":12,"name":"Thào Dình Cống","class":"4A2","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":0,"16":0,"19":0,"20":0,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":13,"name":"Phan Xuân Khánh","class":"4A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Hiển"},{"stt":14,"name":"Giàng A Sảnh","class":"4A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":0,"22":0,"23":0,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":15,"name":"Vàng A Dơ","class":"4A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":16,"name":"Thào A Thiên","class":"4A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":0,"23":0,"26":0,"27":0,"29":0,"30":null},"note":null},{"stt":17,"name":"Hoàng Gia Bảo","class":"4A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":18,"name":"Vàng Xa Minh","class":"4A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":19,"name":"Chang A Hờ","class":"4A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":0,"20":0,"22":0,"23":0,"26":1,"27":1,"29":null,"30":null},"note":"Cô Phương"},{"stt":20,"name":"Giàng A Thương","class":"4A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":21,"name":"Lý A Sâm","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Duyên"},{"stt":22,"name":"Vừ A Công","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":23,"name":"Ly A Dinh","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":24,"name":"Cự A Tuấn","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":0,"23":0,"26":0,"27":1,"29":null,"30":null},"note":null},{"stt":25,"name":"Trần Lê Hữu Bảo","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":26,"name":"Mùa Ngọc Hoàng","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":27,"name":"Vàng A Khua","class":"4A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":28,"name":"Thào Mí Pó","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Ngà"},{"stt":29,"name":"Lý Khải Phong","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":30,"name":"Đinh Văn Quốc","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":31,"name":"Giàng A Sơn","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":0,"29":null,"30":null},"note":null},{"stt":32,"name":"Sần Phà Chang","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":33,"name":"Giàng Khánh Nam","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":0,"22":0,"23":0,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":34,"name":"Hạng A Trường","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":35,"name":"Vừ A Trọng","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":36,"name":"Lê Đình Hải Đăng","class":"5A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":37,"name":"Tráng A Sạnh","class":"5A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":0,"27":0,"29":0,"30":null},"note":"Cô Nguyệt"},{"stt":38,"name":"Lầu Anh Thanh","class":"5A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":39,"name":"S.Ng. Nguyên Khôi","class":"5A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":40,"name":"Lầu Xuân Thắng","class":"5A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":41,"name":"Vừ A Lự","class":"5A4","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":42,"name":"Sùng A Mạnh","class":"5A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":"Cô Thủy"},{"stt":43,"name":"Vừ A Hành","class":"5A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":44,"name":"Sùng A Bảy","class":"5A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":0,"23":0,"26":0,"27":1,"29":null,"30":null},"note":null},{"stt":45,"name":"Quan Anh Thiên Uy","class":"5A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":46,"name":"Hạng A Cử","class":"5A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":47,"name":"Thào A Sở","class":"5A5","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":48,"name":"Phạm Xuân Hùng","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":0,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":0,"30":null},"note":"Cô Hoa"},{"stt":49,"name":"Ohamj Tiến Minh Quang","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":50,"name":"Trần Hải Đăng","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":51,"name":"Nguyễn Hữu An","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":52,"name":"Nguyễn Minh Nhật","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":1,"19":1,"20":1,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null},{"stt":53,"name":"Phạm Duy Thái","class":"3A3","marks":{"1":1,"2":1,"5":1,"6":1,"8":1,"9":1,"12":1,"13":1,"15":1,"16":0,"19":0,"20":0,"22":1,"23":1,"26":1,"27":1,"29":null,"30":null},"note":null}]}""")

# ---------- auth: sao chep logic user() cua main.py (khong import main de tranh circular import) ----------
def _unsign(v):
    try:
        raw, sig = v.rsplit('.', 1)
        if not hmac.compare_digest(sig, hmac.new(SECRET.encode(), raw.encode(), 'sha256').hexdigest()):
            return None
        d = json.loads(base64.urlsafe_b64decode(raw))
        return d if d.get('exp', 0) > time.time() else None
    except Exception:
        return None

def _acc_load():
    try:
        return json.load(open(ACCOUNTS))
    except Exception:
        return {}

def _acc_problem(a):
    if not a:
        return 'Tài khoản không còn tồn tại'
    if a.get('status') != 'approved':
        return 'Tài khoản chưa được duyệt'
    try:
        if a.get('expires') and a['expires'] < datetime.now().date().isoformat():
            return 'Tài khoản đã hết hạn'
    except Exception:
        pass
    return None

class _U:
    def __init__(self, name, key):
        self.name, self.key = name, key
        self.admin = (key == ADMIN_KEY)

def current_user(request: Request):
    d = _unsign(request.cookies.get(COOKIE, ''))
    if not d:
        raise HTTPException(401, 'Chưa đăng nhập')
    if d['key'].startswith('gv-') and d['key'] != ADMIN_KEY:
        raise HTTPException(401, 'Phiên đăng nhập đã hết hiệu lực')
    if d['key'].startswith('dk-'):
        a = _acc_load().get(d['key'])
        pr = _acc_problem(a)
        if pr:
            raise HTTPException(401, pr)
        return _U(a['name'], d['key'])
    return _U(d.get('name', ''), d['key'])

def _chash(request: Request):
    return hashlib.sha256((request.cookies.get(COOKIE, '')).encode()).hexdigest()

def _pw_hash(pw: str):
    return hashlib.sha256(pw.encode()).hexdigest()

# ---------- state ----------
def _room_from_seed(seed):
    reported = [d for d in seed['days'] if all(s['marks'].get(str(d)) is not None for s in seed['students'])]
    return {'id': 'phong5', 'name': seed.get('room') or 'Phòng 5', 'month': seed.get('month', 10),
            'year': seed.get('year', 2026), 'days': list(seed['days']), 'students': seed['students'],
            'reported': reported, 'managers': [], 'pw_hash': None}

def _load():
    try:
        st = json.load(open(STATE_F))
    except Exception:
        st = None
    if not st:
        return {'rooms': {'phong5': _room_from_seed(SEED)}, 'tokens': {}, 'settings': dict(SETTINGS_DEFAULT)}
    if 'students' in st:  # schema cu don phong -> migrate sang phong5, giu nguyen marks/reported
        room = _room_from_seed(SEED)
        room['students'] = st['students']
        room['reported'] = st.get('reported', room['reported'])
        st = {'rooms': {'phong5': room}, 'tokens': {}, 'settings': dict(SETTINGS_DEFAULT)}
    st.setdefault('rooms', {})
    st.setdefault('tokens', {})
    st.setdefault('settings', dict(SETTINGS_DEFAULT))
    st['settings'].setdefault('meal_price', SETTINGS_DEFAULT['meal_price'])
    st['settings'].setdefault('lunch_fee', SETTINGS_DEFAULT['lunch_fee'])
    if 'phong5' not in st['rooms']:
        st['rooms']['phong5'] = _room_from_seed(SEED)
    return st

def _save(st):
    tmp = STATE_F + '.tmp'
    json.dump(st, open(tmp, 'w'), ensure_ascii=False)
    os.replace(tmp, STATE_F)
    try:  # dong bo GitSync nhu cac file data khac (lazy import, tranh circular import)
        import main as _m
        _m.SYNC.schedule()
    except Exception:
        pass

def vn_today():
    return (datetime.now(timezone.utc) + timedelta(hours=7)).date()

def _eligible(s, day):
    """Hoc sinh moi chi tinh tu start_day tro di (seed: start_day None = tu dau)."""
    sd = s.get('start_day')
    return sd is None or day >= sd

def _default_start_day(room):
    today = vn_today()
    if today.year == room['year'] and today.month == room['month']:
        for d in room['days']:
            if d >= today.day:
                return d
        return room['days'][-1] if room['days'] else None
    return room['days'][0] if room['days'] else None

def _day_counts(room, day, today):
    tot = 0; v = 0; c = 0
    for s in room['students']:
        if not _eligible(s, day):
            continue
        tot += 1
        if s['marks'].get(str(day)) == 0:
            v += 1
        elif s['marks'].get(str(day)) is None:
            c += 1
    a = tot - v - c
    if day in room['reported']:
        status = 'Đã báo'
    else:
        try:
            dd = datetime(room['year'], room['month'], day).date()
        except ValueError:
            dd = today
        if dd < today:
            status = 'Nghỉ'
        elif dd == today:
            status = 'Đang báo'
        else:
            status = 'Chưa tới'
    suat = 0 if status == 'Nghỉ' else tot - v
    return {'day': day, 'suat': suat, 'vang': v, 'cho': c, 'an': a, 'status': status}

def _summary(room):
    today = vn_today()
    return [_day_counts(room, d, today) for d in room['days']]

def _public_students(room):
    return [{'stt': s['stt'], 'name': s['name'], 'class': s['class'], 'marks': s['marks'],
             'note': s.get('note'), 'start_day': s.get('start_day'),
             'red': s['name'] in RED_NAMES} for s in room['students']]

def _acc_names():
    acc = _acc_load()
    return {k: v.get('name', k) for k, v in acc.items()}

def _room_meta(room, u):
    today = vn_today()
    tc = _day_counts(room, today.day, today) if (today.year == room['year'] and today.month == room['month'] and today.day in room['days']) else None
    names = _acc_names()
    return {'id': room['id'], 'name': room['name'], 'month': room['month'], 'year': room['year'],
            'n': len(room['students']),
            'has_password': bool(room.get('pw_hash')),
            'managers': room.get('managers', []),
            'manager_names': [names.get(k, k) for k in room.get('managers', [])],
            'reported_n': len(room['reported']), 'days_n': len(room['days']),
            'today': tc, 'can_view': bool(u.admin or u.key in room.get('managers', []))}

def _room_or_404(st, rid):
    r = st['rooms'].get(rid)
    if not r:
        raise HTTPException(404, 'Không có phòng này')
    return r

def _can_view(u, room):
    return bool(u.admin or u.key in room.get('managers', []))

def _require_view(u, room):
    if not _can_view(u, room):
        raise HTTPException(403, 'Bạn không được gắn vào phòng này')

def _valid_token(st, request, u, room):
    tok = request.headers.get('X-Room-Token', '')
    rec = st['tokens'].get(tok)
    return bool(rec and rec['room'] == room['id'] and rec['key'] == u.key and rec['chash'] == _chash(request))

def _require_token(st, request, u, room):
    if not _valid_token(st, request, u, room):
        raise HTTPException(403, 'Chưa nhập mật khẩu phòng')

def _eff_fee(st, room, settings):
    """Muc trong trua hieu dung: muc rieng (neu da dat, ke ca 0) neu khong thi muc chung."""
    v = st.get('lunch_fee')
    return v if v is not None else settings.get('lunch_fee', SETTINGS_DEFAULT['lunch_fee'])

def _month_summary(room, settings):
    rep = set(room['reported'])
    meal_price = settings.get('meal_price', SETTINGS_DEFAULT['meal_price'])
    classes, order = {}, []
    for s in room['students']:
        if s['class'] not in classes:
            classes[s['class']] = []
            order.append(s['class'])
        meals = sum(1 for d in rep if (s.get('start_day') is None or d >= s['start_day']) and s['marks'].get(str(d)) != 0)
        fee = _eff_fee(s, room, settings)
        classes[s['class']].append({'stt': s['stt'], 'name': s['name'], 'meals': meals,
                                    'meal_money': meals * meal_price, 'lunch_fee': fee,
                                    'total': meals * meal_price + fee,
                                    'fee_override': s.get('lunch_fee') is not None})
    out, tot_m, tot_meal, tot_fee, tot_money = [], 0, 0, 0, 0
    for c in sorted(order):
        arr = sorted(classes[c], key=lambda x: x['stt'])
        m = sum(x['meals'] for x in arr)
        mm = sum(x['meal_money'] for x in arr)
        lf = sum(x['lunch_fee'] for x in arr)
        tt = sum(x['total'] for x in arr)
        out.append({'class': c, 'students': arr, 'meals': m, 'meal_money': mm, 'lunch_fee': lf, 'total': tt})
        tot_m += m
        tot_meal += mm
        tot_fee += lf
        tot_money += tt
    today = vn_today()
    is_final = (room['year'], room['month']) < (today.year, today.month)
    return {'room': room['name'], 'room_id': room['id'], 'month': room['month'], 'year': room['year'],
            'meal_price': meal_price, 'lunch_fee_common': settings.get('lunch_fee', SETTINGS_DEFAULT['lunch_fee']),
            'reported_days': sorted(rep), 'classes': out,
            'room_meals': tot_m, 'room_meal_money': tot_meal, 'room_lunch_fee': tot_fee, 'room_money': tot_money,
            'is_final': is_final, 'last_reported_day': (max(rep) if rep else None)}

app = FastAPI(title='Điểm danh đa phòng')

def _clean_entries(entries, days, drop_unknown_days):
    """Validate toan bo entries TRUOC khi thay doi (sai 1 dong -> huy ca dot, khong nap nua voi).
    Tra ve list {name, class, note, marks} da loc. marks chi nhan 1/0/None."""
    if not isinstance(entries, list) or not entries:
        raise HTTPException(400, 'Danh sách trống')
    out = []
    for idx, e in enumerate(entries):
        if not isinstance(e, dict):
            raise HTTPException(400, 'Dòng %d không hợp lệ' % (idx + 1))
        name = (e.get('name') or '').strip()
        cls = (e.get('class') or '').strip()
        if not name or not cls:
            raise HTTPException(400, 'Dòng %d thiếu họ tên hoặc lớp' % (idx + 1))
        marks = {}
        raw = e.get('marks') or {}
        if not isinstance(raw, dict):
            raise HTTPException(400, 'Dòng %d: marks không hợp lệ' % (idx + 1))
        for k, v in raw.items():
            try:
                d = int(k)
            except Exception:
                raise HTTPException(400, 'Dòng %d: cột ngày không hợp lệ' % (idx + 1))
            if v is None:
                continue
            if v not in (0, 1):
                raise HTTPException(400, 'Dòng %d: điểm danh chỉ nhận 0/1/trống' % (idx + 1))
            if days is not None and d not in days:
                if drop_unknown_days:
                    continue
                raise HTTPException(400, 'Dòng %d: ngày %d không thuộc phòng' % (idx + 1, d))
            marks[str(d)] = v
        out.append({'name': name, 'class': cls, 'note': (e.get('note') or '').strip() or None, 'marks': marks})
    return out

def _assign_stts(entries):
    """Giu STT file neu tat ca la so nguyen duy nhat, neu khong danh lai 1..n."""
    stts = [e.get('stt') for e in entries]
    ok = all(isinstance(s, int) for s in stts) and len(set(stts)) == len(stts)
    if ok:
        return stts
    return list(range(1, len(entries) + 1))

@app.post('/api/rooms/{rid}/students')
async def add_student(rid: str, request: Request):
    u = current_user(request)
    body = await request.json()
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    _require_token(st, request, u, room)
    name = (body.get('name') or '').strip()
    cls = (body.get('class') or '').strip()
    note = (body.get('note') or '').strip() or None
    if not name or not cls:
        raise HTTPException(400, 'Thiếu họ tên hoặc lớp')
    sd = body.get('start_day')
    if sd is None:
        sd = _default_start_day(room)
    else:
        try:
            sd = int(sd)
        except Exception:
            raise HTTPException(400, 'Ngày bắt đầu không hợp lệ')
        if sd not in room['days']:
            raise HTTPException(400, 'Ngày bắt đầu phải là ngày phục vụ của phòng')
    stt = max([s['stt'] for s in room['students']], default=0) + 1
    room['students'].append({'stt': stt, 'name': name, 'class': cls, 'marks': {}, 'note': note, 'start_day': sd})
    _save(st)
    return {'ok': True, 'stt': stt, 'start_day': sd}

@app.post('/api/rooms/{rid}/bulk-append')
async def bulk_append(rid: str, request: Request):
    u = current_user(request)
    body = await request.json()
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    _require_token(st, request, u, room)
    sd = body.get('start_day')
    if sd is None:
        sd = _default_start_day(room)
    else:
        try:
            sd = int(sd)
        except Exception:
            raise HTTPException(400, 'Ngày bắt đầu không hợp lệ')
        if sd not in room['days']:
            raise HTTPException(400, 'Ngày bắt đầu phải là ngày phục vụ của phòng')
    cleaned = _clean_entries(body.get('entries'), room['days'], drop_unknown_days=True)
    stt = max([s['stt'] for s in room['students']], default=0) + 1
    for e in cleaned:
        marks = {k: v for k, v in e['marks'].items() if int(k) >= sd}
        room['students'].append({'stt': stt, 'name': e['name'], 'class': e['class'],
                                 'marks': marks, 'note': e['note'], 'start_day': sd})
        stt += 1
    _save(st)
    return {'ok': True, 'added': len(cleaned), 'start_day': sd, 'first_stt': stt - len(cleaned)}

@app.post('/api/rooms/bulk')
async def bulk_create(request: Request):
    u = current_user(request)
    body = await request.json()
    name = (body.get('name') or '').strip()
    days = body.get('days')
    if not name or not isinstance(days, list) or not days:
        raise HTTPException(400, 'Thiếu tên phòng / ngày phục vụ')
    try:
        days = sorted(set(int(d) for d in days))
    except Exception:
        raise HTTPException(400, 'Ngày phục vụ không hợp lệ')
    if any(d < 1 or d > 31 for d in days):
        raise HTTPException(400, 'Ngày phục vụ phải từ 1 đến 31')
    sd = body.get('start_day')
    if sd is None:
        sd = days[0]
    else:
        try:
            sd = int(sd)
        except Exception:
            raise HTTPException(400, 'Ngày bắt đầu không hợp lệ')
        if sd not in days:
            raise HTTPException(400, 'Ngày bắt đầu phải thuộc ngày phục vụ')
    cleaned = _clean_entries(body.get('entries'), days, drop_unknown_days=False)
    stts = _assign_stts(body.get('entries') or [])
    st = _load()
    rid = _slug(name) or 'phong'
    base, i = rid, 2
    while rid in st['rooms']:
        rid = '%s-%d' % (base, i); i += 1
    students = []
    for e, s in zip(cleaned, stts):
        marks = {k: v for k, v in e['marks'].items() if int(k) >= sd}
        students.append({'stt': s, 'name': e['name'], 'class': e['class'],
                         'marks': marks, 'note': e['note'], 'start_day': sd})
    room = {'id': rid, 'name': name, 'month': int(body.get('month') or vn_today().month),
            'year': int(body.get('year') or vn_today().year), 'days': days,
            'students': students, 'reported': [], 'managers': [] if u.admin else [u.key], 'pw_hash': None}
    room['reported'] = [d for d in days if d >= sd and all(
        s['marks'].get(str(d)) is not None for s in students if (s.get('start_day') is None or d >= s['start_day']))]
    st['rooms'][rid] = room
    _save(st)
    return {'ok': True, 'room': _room_meta(room, u)}

@app.post('/api/rooms/excel-preview')
async def excel_preview(request: Request):
    u = current_user(request)
    if not HAVE_XLSX:
        raise HTTPException(501, 'Máy chủ chưa có openpyxl')
    raw = await request.body()
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
    except Exception:
        raise HTTPException(400, 'Không đọc được file Excel (file phải đúng định dạng .xlsx)')
    sheets = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None; hdr_vals = None
        for i, r in enumerate(rows):
            vals = [(str(c).strip() if c is not None else '') for c in (r or [])]
            if 'Họ và tên' in vals and 'Lớp' in vals:
                hdr_i, hdr_vals = i, vals
                break
        if hdr_i is None:
            sheets.append({'name': ws.title, 'error': 'Không tìm thấy hàng tiêu đề (cần có cột Họ và tên và Lớp)'})
            continue
        name_i = hdr_vals.index('Họ và tên')
        class_i = hdr_vals.index('Lớp')
        note_i = hdr_vals.index('Ghi chú') if 'Ghi chú' in hdr_vals else None
        # Cot ngay: lay tu hang ngay sau tieu de neu co so, neu khong lay tu chinh hang tieu de
        day_row = rows[hdr_i + 1] if hdr_i + 1 < len(rows) else []
        day_vals = [(str(c).strip() if c is not None else '') for c in (day_row or [])]
        if note_i is None and 'Ghi chú' in day_vals:
            note_i = day_vals.index('Ghi chú')
        day_cols = {}
        for j in range(class_i + 1, max(len(hdr_vals), len(day_vals))):
            v = day_vals[j] if j < len(day_vals) else ''
            try:
                dv = int(v)
            except Exception:
                continue
            if 1 <= dv <= 31:
                day_cols[j] = dv
        if not day_cols:
            sheets.append({'name': ws.title, 'error': 'Không tìm thấy cột ngày (các số ngày sau cột Lớp)'})
            continue
        students, errors = [], []
        for r in rows[hdr_i + 2:]:
            if r is None:
                continue
            vals = [(str(c).strip() if c is not None else '') for c in r]
            nm = vals[name_i] if name_i < len(vals) else ''
            if nm == 'Vắng' or nm.startswith('Số lượng') or not nm:
                continue
            stt_raw = vals[0] if vals else ''
            try:
                stt = int(float(stt_raw)) if stt_raw != '' else None
            except Exception:
                stt = None
            if stt is None:
                errors.append('Hàng "%s": không có STT hợp lệ' % nm)
                continue
            cls = vals[class_i] if class_i < len(vals) else ''
            if not cls:
                errors.append('Hàng "%s": thiếu lớp' % nm)
                continue
            marks, bad = {}, False
            for j, d in day_cols.items():
                v = r[j] if j < len(r) else None
                if v is None or v == '':
                    continue
                try:
                    iv = int(v)
                except Exception:
                    errors.append('Hàng "%s" ngày %d: điểm danh lạ (%s)' % (nm, d, v))
                    bad = True
                    break
                if iv not in (0, 1):
                    errors.append('Hàng "%s" ngày %d: điểm danh chỉ nhận 0/1/trống' % (nm, d))
                    bad = True
                    break
                marks[str(d)] = iv
            if bad:
                continue
            note = None
            if note_i is not None and note_i < len(vals) and vals[note_i]:
                note = vals[note_i]
            students.append({'stt': stt, 'name': nm, 'class': cls, 'marks': marks, 'note': note})
        sheets.append({'name': ws.title, 'days': sorted(day_cols.values()), 'students': students, 'errors': errors})
    return {'ok': True, 'sheets': sheets}

def _slug(name):
    import unicodedata
    s = unicodedata.normalize('NFD', name.replace('đ', 'd').replace('Đ', 'D'))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()
    s = re.sub(r'[^a-z0-9]+', '-', s).strip('-')[:40]
    return s or 'phong'

@app.get('/')
def shell():
    return HTMLResponse(HTML, headers={'Cache-Control': 'no-cache'})

@app.get('/manifest.json')
def manifest():
    return Response(MANIFEST, media_type='application/manifest+json')

@app.get('/sw.js')
def sw():
    return Response(SW, media_type='application/javascript', headers={'Cache-Control': 'no-cache'})

@app.get('/api/rooms')
def list_rooms(request: Request):
    u = current_user(request)
    st = _load()
    rooms = [_room_meta(r, u) for r in st['rooms'].values()]
    if not u.admin:
        rooms = [r for r in rooms if r['can_view']]
    return {'ok': True, 'user': {'name': u.name, 'admin': u.admin}, 'rooms': rooms}

@app.get('/api/rooms/accounts')
def list_accounts(request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị')
    acc = _acc_load()
    return {'ok': True, 'accounts': [{'key': k, 'name': v.get('name', k)} for k, v in acc.items() if v.get('status') == 'approved']}

@app.post('/api/rooms')
async def create_room(request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị')
    body = await request.json()
    name = (body.get('name') or '').strip()
    days = body.get('days')
    students = body.get('students')
    if not name or not isinstance(days, list) or not days or not isinstance(students, list) or not students:
        raise HTTPException(400, 'Thiếu tên phòng / ngày / danh sách học sinh')
    for s in students:
        if not isinstance(s, dict) or 'stt' not in s or 'name' not in s or 'class' not in s or not isinstance(s.get('marks'), dict):
            raise HTTPException(400, 'Học sinh thiếu stt/name/class/marks')
    st = _load()
    rid = (body.get('id') or _slug(name)).strip() or 'phong'
    base, i = rid, 2
    while rid in st['rooms']:
        rid = '%s-%d' % (base, i); i += 1
    room = {'id': rid, 'name': name, 'month': int(body.get('month') or vn_today().month),
            'year': int(body.get('year') or vn_today().year), 'days': [int(d) for d in days],
            'students': students, 'reported': [], 'managers': [], 'pw_hash': None}
    room['reported'] = [d for d in room['days'] if all(s['marks'].get(str(d)) is not None for s in students)]
    st['rooms'][rid] = room
    _save(st)
    return {'ok': True, 'room': _room_meta(room, u)}

@app.post('/api/rooms/{rid}/managers')
async def set_managers(rid: str, request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị')
    body = await request.json()
    keys = body.get('keys') or []
    acc = _acc_load()
    keys = [k for k in keys if k in acc and acc[k].get('status') == 'approved']
    st = _load()
    room = _room_or_404(st, rid)
    room['managers'] = keys
    _save(st)
    return {'ok': True, 'managers': keys}

@app.post('/api/settings')
async def set_settings(request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị mới đổi giá chung')
    body = await request.json()
    st = _load()
    if body.get('meal_price') is not None:
        try:
            mp = int(body.get('meal_price'))
        except Exception:
            raise HTTPException(400, 'Giá ăn không hợp lệ')
        if mp <= 0:
            raise HTTPException(400, 'Giá ăn phải > 0')
        st['settings']['meal_price'] = mp
    if body.get('lunch_fee') is not None:
        try:
            lf = int(body.get('lunch_fee'))
        except Exception:
            raise HTTPException(400, 'Mức trông trưa không hợp lệ')
        if lf < 0:
            raise HTTPException(400, 'Mức trông trưa không được âm')
        st['settings']['lunch_fee'] = lf
    _save(st)
    return {'ok': True, 'settings': st['settings']}

@app.delete('/api/rooms/{rid}')
def delete_room(rid: str, request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị')
    st = _load()
    if rid not in st['rooms']:
        raise HTTPException(404, 'Không có phòng này')
    del st['rooms'][rid]
    st['tokens'] = {t: r for t, r in st['tokens'].items() if r['room'] != rid}
    _save(st)
    return {'ok': True}

@app.post('/api/rooms/{rid}/students/{stt}/lunch-fee')
async def set_student_fee(rid: str, stt: int, request: Request):
    u = current_user(request)
    body = await request.json()
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    target = next((s for s in room['students'] if s['stt'] == stt), None)
    if not target:
        raise HTTPException(404, 'Không có học sinh này')
    fee = body.get('fee')
    if fee is None:
        target.pop('lunch_fee', None)
    else:
        try:
            fee = int(fee)
        except Exception:
            raise HTTPException(400, 'Mức trông trưa không hợp lệ')
        if fee < 0:
            raise HTTPException(400, 'Mức trông trưa không được âm')
        target['lunch_fee'] = fee
    _save(st)
    return {'ok': True, 'effective': _eff_fee(target, room, st['settings'])}

@app.post('/api/rooms/{rid}/set-password')
async def set_password(rid: str, request: Request):
    u = current_user(request)
    body = await request.json()
    pw = (body.get('password') or '')
    if len(pw) < 4:
        raise HTTPException(400, 'Mật khẩu phòng cần ít nhất 4 ký tự')
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    if room.get('pw_hash'):
        if not u.admin:
            old = (body.get('old_password') or '')
            if _pw_hash(old) != room['pw_hash']:
                raise HTTPException(403, 'Sai mật khẩu phòng cũ')
    room['pw_hash'] = _pw_hash(pw)
    _save(st)
    return {'ok': True}

@app.post('/api/rooms/{rid}/unlock')
async def unlock(rid: str, request: Request):
    u = current_user(request)
    body = await request.json()
    pw = (body.get('password') or '')
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    if not room.get('pw_hash'):
        raise HTTPException(409, 'Phòng chưa đặt mật khẩu')
    if _pw_hash(pw) != room['pw_hash']:
        raise HTTPException(403, 'Sai mật khẩu phòng')
    tok = secrets.token_hex(16)
    st['tokens'][tok] = {'room': rid, 'key': u.key, 'chash': _chash(request)}
    _save(st)
    return {'ok': True, 'token': tok}

@app.get('/api/rooms/{rid}/state')
def room_state(rid: str, request: Request):
    u = current_user(request)
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    _require_token(st, request, u, room)
    return {'ok': True, 'user': {'name': u.name, 'admin': u.admin}, 'room': _room_meta(room, u),
            'days': room['days'], 'students': _public_students(room),
            'reported': room['reported'], 'summary': _summary(room)}

@app.post('/api/rooms/{rid}/mark')
async def room_mark(rid: str, request: Request):
    u = current_user(request)
    body = await request.json()
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    _require_token(st, request, u, room)
    try:
        day = int(body.get('day')); stt = int(body.get('stt'))
    except Exception:
        raise HTTPException(400, 'Thiếu ngày/STT')
    mark = body.get('mark')
    if mark not in (1, 0, None):
        raise HTTPException(400, 'mark chỉ nhận 1/0/null')
    if day not in room['days']:
        raise HTTPException(400, 'Ngày không phải ngày phục vụ')
    target = next((s for s in room['students'] if s['stt'] == stt), None)
    if not target:
        raise HTTPException(404, 'Không có học sinh này')
    if mark is None:
        target['marks'].pop(str(day), None)
    else:
        target['marks'][str(day)] = mark
    _save(st)
    c = _day_counts(room, day, vn_today())
    return {'ok': True, 'day': day, 'suat': c['suat'], 'vang': c['vang'], 'cho': c['cho'], 'an': c['an']}

@app.post('/api/rooms/{rid}/report')
async def room_report(rid: str, request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị mới báo cơm')
    body = await request.json()
    st = _load()
    room = _room_or_404(st, rid)
    _require_token(st, request, u, room)
    try:
        day = int(body.get('day'))
    except Exception:
        raise HTTPException(400, 'Thiếu ngày')
    if day not in room['days']:
        raise HTTPException(400, 'Ngày không phải ngày phục vụ')
    if day not in room['reported']:
        room['reported'].append(day)
        room['reported'].sort()
    nxt = day
    for d in room['days']:
        if d > day:
            nxt = d
            break
    _save(st)
    return {'ok': True, 'reported': room['reported'], 'next_day': nxt, 'summary': _summary(room)}

def _csv_bytes(room):
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['DANH SÁCH HỌC SINH ĂN BÁN TRÚ %s - THÁNG %d/%d' % (room['name'].upper(), room['month'], room['year'])])
    w.writerow(['STT', 'Họ và tên', 'Lớp'] + ['Ngày %d' % d for d in room['days']] + ['SDT', 'Ghi chú'])
    for s in room['students']:
        w.writerow([s['stt'], s['name'], s['class']] + [s['marks'].get(str(d)) if s['marks'].get(str(d)) is not None else '' for d in room['days']] + ['', s.get('note') or ''])
    w.writerow(['', 'Vắng', ''] + [sum(1 for s in room['students'] if s['marks'].get(str(d)) == 0) for d in room['days']] + ['', ''])
    w.writerow(['', 'Số lượng báo cơm', ''] + [len(room['students']) - sum(1 for s in room['students'] if s['marks'].get(str(d)) == 0) for d in room['days']] + ['', ''])
    return '\ufeff' + out.getvalue()

@app.get('/api/rooms/{rid}/export')
def room_export(rid: str, request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị mới xuất file')
    st = _load()
    room = _room_or_404(st, rid)
    _require_token(st, request, u, room)
    data = _csv_bytes(room).encode('utf-8')
    return Response(data, media_type='text/csv',
                    headers={'Content-Disposition': 'attachment; filename="diem-danh-%s.csv"' % room['id']})

@app.get('/api/rooms/{rid}/export-xlsx')
def room_export_xlsx(rid: str, request: Request):
    u = current_user(request)
    if not u.admin:
        raise HTTPException(403, 'Chỉ quản trị mới xuất file')
    if not HAVE_XLSX:
        raise HTTPException(501, 'Máy chủ chưa có openpyxl')
    st = _load()
    room = _room_or_404(st, rid)
    _require_token(st, request, u, room)
    wb = Workbook(); ws = wb.active
    ws.title = 'Thang %d' % room['month']
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + len(room['days']) + 2)
    c = ws.cell(1, 1, 'DANH SÁCH HỌC SINH ĂN BÁN TRÚ %s' % room['name'].upper()); c.font = Font(bold=True, size=14)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3 + len(room['days']) + 2)
    ws.cell(2, 1, 'Tháng %d/%d' % (room['month'], room['year'])).font = Font(bold=True)
    ws.cell(3, 1, 'STT').font = Font(bold=True); ws.cell(3, 2, 'Họ và tên').font = Font(bold=True); ws.cell(3, 3, 'Lớp').font = Font(bold=True)
    ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=3 + len(room['days']))
    mc = ws.cell(3, 4, 'Tháng %d' % room['month']); mc.font = Font(bold=True)
    for j, d in enumerate(room['days']):
        ws.cell(4, 4 + j, d).font = Font(bold=True)
    ws.cell(3, 4 + len(room['days']), 'SDT').font = Font(bold=True)
    ws.cell(3, 5 + len(room['days']), 'Ghi chú').font = Font(bold=True)
    r = 5
    for s in room['students']:
        ws.cell(r, 1, s['stt']); ws.cell(r, 2, s['name']); ws.cell(r, 3, s['class'])
        for j, d in enumerate(room['days']):
            v = s['marks'].get(str(d))
            if v is not None:
                ws.cell(r, 4 + j, v)
        ws.cell(r, 5 + len(room['days']), s.get('note') or '')
        r += 1
    ws.cell(r, 2, 'Vắng').font = Font(bold=True)
    for j, d in enumerate(room['days']):
        ws.cell(r, 4 + j, sum(1 for s in room['students'] if s['marks'].get(str(d)) == 0))
    r += 1
    ws.cell(r, 2, 'Số lượng báo cơm').font = Font(bold=True)
    for j, d in enumerate(room['days']):
        ws.cell(r, 4 + j, len(room['students']) - sum(1 for s in room['students'] if s['marks'].get(str(d)) == 0))
    for col in range(1, 6 + len(room['days'])):
        ws.column_dimensions[get_column_letter(col)].width = 16 if col == 2 else 11
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return Response(buf.read(), media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    headers={'Content-Disposition': 'attachment; filename="diem-danh-%s.xlsx"' % room['id']})

@app.get('/api/rooms/{rid}/summary')
def room_summary(rid: str, request: Request):
    u = current_user(request)
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    return {'ok': True, 'user': {'name': u.name, 'admin': u.admin}, 'summary': _month_summary(room, st['settings'])}

@app.get('/api/rooms/{rid}/summary.csv')
def room_summary_csv(rid: str, request: Request):
    u = current_user(request)
    st = _load()
    room = _room_or_404(st, rid)
    _require_view(u, room)
    sm = _month_summary(room, st['settings'])
    out = io.StringIO(); w = csv.writer(out)
    status_txt = 'TỔNG CHÍNH THỨC' if sm['is_final'] else 'TẠM TÍNH đến hết ngày đã báo'
    w.writerow(['TỔNG HỢP CUỐI THÁNG - %s - THÁNG %d/%d - %s' % (room['name'].upper(), room['month'], room['year'], status_txt)])
    if not sm['is_final']:
        w.writerow(['Sau cuối tháng mới chốt tổng chính thức.'])
    w.writerow(['Giá ăn chung: %dđ/suất · Trông trưa chung: %dđ/tháng/học sinh (sửa riêng từng học sinh được)' % (sm['meal_price'], sm['lunch_fee_common'])])
    w.writerow(['Quy tắc: chỉ tính các ngày đã "Đã báo" (%s). Ăn khi mark khác 0 (1 = ăn xác nhận, trống = vẫn ăn theo luật chờ điểm danh; 0 = vắng không tính). Ngày "Đang báo"/"Chưa tới" chưa tính. Tổng tiền = Tiền ăn + Trông trưa.' % ', '.join(str(d) for d in sm['reported_days'])])
    w.writerow(['Lớp', 'STT', 'Họ và tên', 'Tổng số buổi ăn', 'Tiền ăn', 'Trông trưa', 'Tổng tiền'])
    for cl in sm['classes']:
        for s in cl['students']:
            w.writerow([cl['class'], s['stt'], s['name'], s['meals'], s['meal_money'], s['lunch_fee'], s['total']])
        w.writerow([cl['class'], '', 'TỔNG LỚP', cl['meals'], cl['meal_money'], cl['lunch_fee'], cl['total']])
    w.writerow(['', '', 'TỔNG TOÀN PHÒNG', sm['room_meals'], sm['room_meal_money'], sm['room_lunch_fee'], sm['room_money']])
    return Response(('\ufeff' + out.getvalue()).encode('utf-8'), media_type='text/csv',
                    headers={'Content-Disposition': 'attachment; filename="tong-hop-%s.csv"' % room['id']})

HTML = r"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="#2563eb">
<link rel="manifest" href="manifest.json">
<title>Điểm danh bán trú</title>
<style>
*{box-sizing:border-box}
*{scrollbar-width:none!important;-ms-overflow-style:none!important}
*::-webkit-scrollbar{width:0!important;height:0!important;display:none!important}
html,body{margin:0;padding:0}
body{font-family:system-ui,"Segoe UI",Arial,sans-serif;background:#f1f5f9;color:#0f172a;-webkit-text-size-adjust:100%}
#app{max-width:480px;margin:0 auto;background:#fff;min-height:100vh;padding-bottom:110px}
.top{background:linear-gradient(135deg,#1d4ed8,#2563eb);color:#fff;padding:14px 14px 10px}
.top h1{margin:0;font-size:21px;line-height:1.25}
.top p{margin:4px 0 0;font-size:14px;opacity:.9}
.top button{background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.5);color:#fff;border-radius:10px;padding:6px 10px;font-size:14px;font-weight:700;cursor:pointer}
.days{display:flex;gap:8px;overflow-x:auto;padding:10px 12px;background:#fff;border-bottom:1px solid #e2e8f0;scrollbar-width:none}
.days::-webkit-scrollbar{display:none}
.chip{flex:0 0 auto;min-width:52px;padding:8px 6px;border:2px solid #cbd5e1;border-radius:12px;background:#fff;font-size:17px;font-weight:700;color:#0f172a;cursor:pointer}
.chip small{display:block;font-size:12px;font-weight:600;color:#64748b}
.chip.sel{background:#2563eb;border-color:#2563eb;color:#fff}
.chip.sel small{color:#dbeafe}
.counts{position:-webkit-sticky;position:sticky;top:0;z-index:9;background:#fff;border-bottom:1px solid #e2e8f0;padding:10px 12px 12px}
.hero{text-align:center;background:linear-gradient(135deg,#fff7ed,#ffedd5);border:2px solid #fdba74;border-radius:16px;padding:10px 6px;margin-bottom:10px}
.hero b{display:block;font-size:44px;line-height:1;color:#c2410c}
.hero span{font-size:15px;font-weight:800;color:#7c2d12}
.sub3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}
.c{border-radius:12px;padding:8px 4px;text-align:center}
.c b{display:block;font-size:24px;line-height:1.1}
.c span{font-size:13px;font-weight:700}
.c.eat{background:#dcfce7;color:#166534}
.c.abs{background:#fee2e2;color:#991b1b}
.c.none{background:#fef3c7;color:#92400e}
.group{padding:4px 0 0}
.group h2{margin:0;padding:14px 14px 6px;font-size:16px;color:#334155}
.group h2.red{color:#dc2626}
.group h2.green{color:#16a34a}
.row{display:flex;align-items:center;gap:10px;padding:10px 14px;border-bottom:1px solid #f1f5f9}
.who{flex:1;min-width:0}
.who b{display:block;font-size:18px;line-height:1.3;overflow-wrap:break-word}
.who b.red{color:#dc2626}
.who span{display:block;font-size:14px;color:#64748b;margin-top:2px}
.st{flex:0 0 auto;min-width:142px;min-height:52px;padding:8px 8px;border-radius:14px;border:2px solid #cbd5e1;background:#fff;font-size:16px;font-weight:800;cursor:pointer}
.st.eat{background:#16a34a;border-color:#16a34a;color:#fff}
.st.abs{background:#dc2626;border-color:#dc2626;color:#fff}
.st.none{background:#fffbeb;border-color:#d97706;color:#92400e}
.bar{position:fixed;left:0;right:0;bottom:0;z-index:20;background:#fff;border-top:1px solid #e2e8f0;padding:12px 14px calc(12px + env(safe-area-inset-bottom))}
.bar-inner{max-width:480px;margin:0 auto}
#btnBaoCom{width:100%;min-height:58px;border:0;border-radius:16px;background:linear-gradient(135deg,#ea580c,#f97316);color:#fff;font-size:20px;font-weight:800;cursor:pointer}
.admin-note{text-align:center;font-size:13px;color:#64748b;margin:8px 0 0}
#reportView{padding:14px}
#reportView h2{margin:2px 0;font-size:21px}
#reportView .sub{color:#475569;font-size:15px;margin:4px 0 14px}
.bigmeal{text-align:center;background:linear-gradient(135deg,#fff7ed,#ffedd5);border:2px solid #fdba74;border-radius:20px;padding:20px 12px;margin:6px 0 16px}
.bigmeal .num{font-size:76px;line-height:1;font-weight:900;color:#c2410c}
.bigmeal .lbl{font-size:19px;font-weight:700;color:#7c2d12;margin-top:6px}
.bigmeal .det{font-size:15px;color:#9a3412;margin-top:4px}
table{width:100%;border-collapse:collapse;font-size:16px}
th,td{padding:8px 6px;border-bottom:1px solid #f1f5f9;text-align:right}
th:first-child,td:first-child{text-align:left}
th{color:#475569;font-size:14px}
td b{font-size:17px}
h3{font-size:18px;margin:20px 0 8px}
#btnBack{width:100%;min-height:52px;border-radius:14px;border:2px solid #cbd5e1;background:#fff;font-size:18px;font-weight:700;cursor:pointer;margin-bottom:14px}
#btnCopy{width:100%;min-height:56px;border:0;border-radius:16px;background:#16a34a;color:#fff;font-size:19px;font-weight:800;cursor:pointer;margin:12px 0}
#reportText{width:100%;min-height:110px;padding:12px;border:2px solid #bbf7d0;border-radius:14px;background:#f0fdf4;font-size:17px;line-height:1.5;font-family:system-ui,"Segoe UI",Arial,sans-serif;resize:vertical}
.hint{font-size:14px;color:#64748b;line-height:1.5}
.filters{display:flex;gap:8px;overflow-x:auto;padding:10px 12px 10px;scrollbar-width:none;position:-webkit-sticky;position:sticky;top:112px;z-index:8;background:#fff;border-bottom:1px solid #e2e8f0}
.filters::-webkit-scrollbar{display:none}
.fchip{flex:0 0 auto;min-width:58px;padding:8px 10px;border:2px solid #cbd5e1;border-radius:999px;background:#fff;font-size:16px;font-weight:700;cursor:pointer}
.fchip.sel{background:#2563eb;border-color:#2563eb;color:#fff}
.filternote{margin:0;padding:12px 14px 0;font-size:14px;font-weight:700;color:#1d4ed8}
.chip .tick{font-size:13px;color:#16a34a}
.chip.sel .tick{color:#fef08a}
.repstatus{font-size:16px;font-weight:700;color:#1d4ed8;margin:0 0 12px}
.rule{font-size:14px;color:#475569;line-height:1.5;background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:10px 12px}
#btnConfirmBao{width:100%;min-height:58px;border:0;border-radius:16px;background:linear-gradient(135deg,#ea580c,#f97316);color:#fff;font-size:19px;font-weight:800;cursor:pointer;margin:4px 0 12px}
#reportView table{font-size:15px}
#sumTables table{font-size:13px}
#sumTables th,#sumTables td{padding:6px 2px}
#sumTables .feeinp{width:76px;font-size:13px;padding:5px}
#sumTables .feereset{padding:4px 5px}
#sumTables td:nth-child(2){overflow-wrap:break-word}
#sumTables{overflow-x:auto}
#sumTables table{min-width:560px}
.countstrip{display:flex;gap:6px}
.cc{flex:1;text-align:center;padding:8px 2px;border-radius:12px;border:2px solid}
.cc b{display:block;font-size:22px;line-height:1.2}
.cc span{display:block;font-size:12px;font-weight:700;white-space:nowrap}
.cc.orange{color:#c2410c;border-color:#fdba74;background:#fff7ed}
.cc.red{color:#dc2626;border-color:#fca5a5;background:#fef2f2}
.cc.dyellow{color:#a16207;border-color:#fde68a;background:#fefce8}
.cc.green{color:#15803d;border-color:#86efac;background:#f0fdf4}
.curday{font-size:15px;font-weight:700;color:#1e40af;text-align:center;margin:0 0 8px}
#menuBtn{position:fixed;top:4px;right:10px;z-index:202;width:44px;height:44px;border-radius:50%;border:2px solid #cbd5e1;background:#fff;font-size:22px;font-weight:700;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.18);line-height:1}
#menuOverlay{position:fixed;inset:0;z-index:200;background:transparent}
#menuPanel{position:fixed;top:56px;right:12px;z-index:201;width:min(320px,86vw);max-height:80dvh;overflow-y:auto;background:#fff;border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.25);padding:8px}
#menuPanel button.mitem{display:block;width:100%;text-align:left;min-height:52px;font-size:17px;font-weight:700;border:0;border-bottom:1px solid #e2e8f0;background:#fff;cursor:pointer;padding:10px 12px;border-radius:10px}
#menuPanel button.mitem:active{background:#dbeafe}
#menuDaysHead{font-size:14px;font-weight:800;color:#64748b;padding:12px 12px 4px}
#menuPanel button.mday{display:block;width:100%;text-align:left;min-height:44px;font-size:16px;border:0;border-bottom:1px solid #f1f5f9;background:#fff;cursor:pointer;padding:8px 12px;border-radius:10px}
#menuPanel button.mday.sel{background:#dbeafe;font-weight:800}
#toast{position:fixed;top:70px;left:50%;transform:translateX(-50%);z-index:203;background:#0f172a;color:#fff;font-size:16px;font-weight:700;padding:12px 18px;border-radius:999px;box-shadow:0 4px 20px rgba(0,0,0,.3);white-space:nowrap}
#app h1{padding-right:56px}
td.sttcell{white-space:nowrap;font-size:14px}
.expbtn{display:block;width:100%;min-height:52px;border-radius:14px;border:2px solid #2563eb;background:#eff6ff;color:#1d4ed8;font-size:17px;font-weight:800;cursor:pointer;margin:8px 0;text-decoration:none;text-align:center;line-height:48px}
#loginGate{max-width:480px;margin:0 auto;background:#fff;min-height:100vh;padding:40px 20px;text-align:center}
#loginGate h1{font-size:22px}
#loginGate p{font-size:16px;color:#475569;line-height:1.5}
#loginGate a{display:inline-block;margin-top:16px;padding:14px 28px;border-radius:99px;background:#2563eb;color:#fff;font-size:18px;font-weight:800;text-decoration:none}
#nonAdminNote{display:none;text-align:center;font-size:14px;font-weight:700;color:#991b1b;margin:8px 0 0}
.dot{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:6px;vertical-align:1px}
.dot.red{background:#dc2626}
.dot.green{background:#16a34a}
.roomcard{border:2px solid #e2e8f0;border-radius:16px;padding:14px;margin:12px}
.roomcard h2{margin:0 0 6px;font-size:20px}
.roomcard p{margin:4px 0;font-size:15px;color:#475569}
.roomcard .btnrow{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.rbtn{flex:1;min-width:120px;min-height:48px;border-radius:12px;border:0;font-size:16px;font-weight:800;cursor:pointer}
.rbtn.blue{background:#2563eb;color:#fff}
.rbtn.green{background:#16a34a;color:#fff}
.rbtn.gray{background:#f1f5f9;color:#0f172a;border:2px solid #cbd5e1}
.rbtn.orange{background:#ea580c;color:#fff}
.gatebox{max-width:480px;margin:0 auto;padding:30px 20px}
.gatebox h2{font-size:22px}
.gatebox p{font-size:15px;color:#475569;line-height:1.5}
.gatebox input{width:100%;min-height:54px;font-size:18px;border:2px solid #cbd5e1;border-radius:12px;padding:10px 12px;margin:8px 0}
.gatebox button{width:100%;min-height:54px;border:0;border-radius:14px;background:#2563eb;color:#fff;font-size:18px;font-weight:800;cursor:pointer;margin-top:6px}
.gatebox .err{color:#dc2626;font-weight:700;min-height:20px}
.picker{border:1px solid #e2e8f0;border-radius:12px;padding:10px;margin-top:10px}
.picker label{display:block;font-size:16px;padding:6px 0}
.picker input{margin-right:8px}
textarea.seed{width:100%;min-height:120px;font-size:14px;border:2px solid #cbd5e1;border-radius:12px;padding:10px}
.sumprice{font-size:17px;font-weight:800;color:#7c2d12}
.sumprice input{width:110px;font-size:17px;padding:6px;border:2px solid #cbd5e1;border-radius:10px}
.feeinp{width:76px;font-size:13px;padding:6px 4px;border:2px solid #cbd5e1;border-radius:8px;text-align:right}
.feereset{font-size:14px;padding:6px 8px;border-radius:8px;border:2px solid #cbd5e1;background:#fff;cursor:pointer}
.sumprice button{font-size:15px;font-weight:800;padding:8px 12px;border-radius:10px;border:0;background:#2563eb;color:#fff;cursor:pointer}
tr.total td{font-weight:800;background:#f8fafc}
tr.roomtotal td{font-weight:900;background:#ffedd5}
</style>
</head>
<body>
<div id="loginGate" hidden>
  <h1>Điểm danh bán trú</h1>
  <p>Bạn chưa đăng nhập. Hãy đăng nhập bằng tài khoản web trước rồi quay lại đây.</p>
  <a href="/">Đăng nhập</a>
</div>
<div id="app" hidden>
    <div id="menuOverlay" hidden></div>
    <button id="menuBtn" type="button" aria-label="Menu">☰</button>
    <div id="menuPanel" hidden>
      <button class="mitem" type="button" data-menu="rooms">Đổi phòng</button>
      <button class="mitem" type="button" data-menu="addone">Thêm học sinh</button>
      <button class="mitem" type="button" data-menu="bulk">Thêm danh sách</button>
      <button class="mitem" type="button" data-menu="summary">Tổng hợp</button>
      <div id="menuDaysHead">Ngày</div>
      <div id="menuDays"></div>
    </div>
    <div id="toast" hidden></div>
  <div id="roomsView">
    <div class="top">
      <h1>Phòng bán trú</h1>
      <p id="roomsSub"></p>
    </div>
    <div id="roomsList"></div>
    <div class="roomcard" id="bulkCreateCard">
      <h2>＋ Thêm phòng</h2>
      <p class="hint">Dán danh sách hoặc tải file Excel để tạo phòng mới từ danh sách.</p>
      <p><input id="bcName" placeholder="Tên phòng mới (VD: Phòng 6)" style="width:100%;min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></p>
      <p>Tháng: <input id="bcMonth" type="number" min="1" max="12" style="width:70px;min-height:44px;font-size:16px"> Năm: <input id="bcYear" type="number" min="2025" max="2027" style="width:90px;min-height:44px;font-size:16px"></p>
      <p>Ngày phục vụ (cách nhau bằng dấu phẩy, bắt buộc khi dán danh sách): <input id="bcDays" placeholder="VD: 1,2,5,6,8,9" style="width:100%;min-height:48px;font-size:16px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></p>
      <p>Ngày bắt đầu: <select id="bcStart" style="min-height:44px;font-size:16px"></select></p>
      <p><textarea id="bcPaste" class="seed" placeholder="Mỗi dòng một học sinh: Họ tên — Lớp"></textarea></p>
      <p>Hoặc tải file Excel đúng định dạng file gốc: <input id="bcFile" type="file" accept=".xlsx"></p>
      <div id="bcSheetSel" hidden><p>Chọn sheet: <select id="bcSheet" style="min-height:44px;font-size:16px"></select></p></div>
      <div id="bcPreview"></div>
      <div class="btnrow"><button class="rbtn blue" id="bcConfirm" type="button" disabled>Xác nhận tạo phòng</button></div>
      <p id="bcMsg" style="color:#dc2626;font-weight:700"></p>
    </div>
    <div id="seedBox" class="roomcard" hidden>
      <h2>Thêm phòng mới (quản trị)</h2>
      <p>Dán JSON danh sách phòng (định dạng như file Phòng 5: name, month, year, days, students với stt/name/class/marks). Không cần sửa code.</p>
      <textarea id="seedJson" class="seed" placeholder='{"name": "Phòng 6", "month": 10, "year": 2026, "days": [1,2], "students": [{"stt": 1, "name": "Nguyễn Văn A", "class": "3A1", "marks": {}}]}'></textarea>
      <div class="btnrow"><button class="rbtn blue" id="btnSeed" type="button">Tạo phòng</button></div>
      <p id="seedMsg" class="err" style="color:#dc2626;font-weight:700"></p>
    </div>
  </div>

  <div id="gateView" hidden>
    <div class="gatebox">
      <h2 id="gateTitle">Vào phòng</h2>
      <p id="gateMsg"></p>
      <div id="pwForm">
        <input id="gatePw" type="password" placeholder="Mật khẩu phòng" autocomplete="off">
        <p class="err" id="gateErr"></p>
        <button id="gateBtn" type="button">Vào phòng</button>
      </div>
      <div id="setForm" hidden>
        <p>Phòng này chưa có mật khẩu. Quản lý phòng / quản trị đặt mật khẩu để bắt đầu (từ lần sau, vào phòng phải nhập đúng mật khẩu này).</p>
        <input id="setPw" type="password" placeholder="Đặt mật khẩu phòng (ít nhất 4 ký tự)" autocomplete="off">
        <p class="err" id="setErr"></p>
        <button id="setBtn" type="button">Đặt mật khẩu và vào phòng</button>
      </div>
      <button id="gateBack" type="button" style="background:#f1f5f9;color:#0f172a;border:2px solid #cbd5e1;margin-top:12px">← Danh sách phòng</button>
    </div>
  </div>

  <div id="roomView" hidden>
    <div class="top">
      <h1 id="roomTitle" style="margin-top:10px"></h1>
      <p id="roomSub"></p>
    </div>
    <div id="addOnePanel" class="roomcard" hidden>
      <h2>Thêm học sinh</h2>
      <p><input id="aoName" placeholder="Họ và tên" style="width:100%;min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></p>
      <p><select id="aoClass" style="width:100%;min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></select>
      <input id="aoClassNew" placeholder="Lớp mới (nếu chọn Lớp mới...)" style="width:100%;min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px;margin-top:8px" hidden></p>
      <p><input id="aoNote" placeholder="GVCN (tự lấy theo lớp, sửa được)" style="width:100%;min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></p>
      <p>Ngày bắt đầu: <select id="aoStart" style="min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></select></p>
      <p class="hint">Học sinh mới chỉ tính điểm danh + tiền ăn từ ngày bắt đầu trở đi; các ngày trước không tính.</p>
      <div class="btnrow"><button class="rbtn blue" id="aoSave" type="button">Thêm</button></div>
      <p id="aoMsg" style="color:#dc2626;font-weight:700"></p>
    </div>
    <div id="bulkRoomPanel" class="roomcard" hidden>
      <h2>Thêm danh sách vào phòng này</h2>
      <p>Ngày bắt đầu (cho cả đợt): <select id="brStart" style="min-height:48px;font-size:17px;border:2px solid #cbd5e1;border-radius:12px;padding:8px"></select></p>
      <p><textarea id="brPaste" class="seed" placeholder="Mỗi dòng một học sinh: Họ tên — Lớp"></textarea></p>
      <p>Hoặc tải file Excel đúng định dạng file gốc: <input id="brFile" type="file" accept=".xlsx"></p>
      <div id="brSheetSel" hidden><p>Chọn sheet: <select id="brSheet" style="min-height:44px;font-size:16px"></select></p></div>
      <div id="brPreview"></div>
      <div class="btnrow"><button class="rbtn blue" id="brConfirm" type="button" disabled>Xác nhận thêm</button></div>
      <p id="brMsg" style="color:#dc2626;font-weight:700"></p>
    </div>
    <div id="listView">
      <div class="counts" id="stickyInfo">
        <p id="curDayLine" class="curday"></p>
        <div class="countstrip">
          <div class="cc orange"><b id="cSuat">0</b><span>Suất ăn</span></div>
          <div class="cc red"><b id="cVang">0</b><span>Vắng</span></div>
          <div class="cc dyellow"><b id="cCho">0</b><span>Chờ điểm danh</span></div>
          <div class="cc green"><b id="cDaAn">0</b><span>Xác nhận</span></div>
        </div>
      </div>
      <div class="filters" id="filters"></div>
      <div id="filterNote" class="filternote" hidden></div>
      <div id="list"></div>
      <div class="bar"><div class="bar-inner">
        <button id="btnBaoCom" type="button">🍱 Báo cơm</button>
        <p id="nonAdminNote">🔒 Chỉ tài khoản quản trị mới báo cơm.</p>
      </div></div>
    </div>
    <div id="reportView" hidden>
      <button id="btnBack" type="button">← Quay lại điểm danh</button>
      <h2 id="repTitle">Báo cơm</h2>
      <p class="sub" id="repSub"></p>
      <div class="bigmeal">
        <div class="num" id="bigNum">0</div>
        <div class="lbl">suất ăn</div>
        <div class="countstrip">
          <div class="cc orange"><b id="repSuat">0</b><span>Suất ăn</span></div>
          <div class="cc red"><b id="repVang">0</b><span>Vắng</span></div>
          <div class="cc dyellow"><b id="repCho">0</b><span>Chờ điểm danh</span></div>
          <div class="cc green"><b id="repAn">0</b><span>Xác nhận</span></div>
        </div>
      </div>
      <p class="repstatus" id="repStatus"></p>
      <button id="btnConfirmBao" type="button">✅ Xác nhận báo cơm ngày <span id="cfDay">29</span></button>
      <p class="rule">📌 Luật: ngày phục vụ trôi qua mà không bấm báo thì tính là ngày nghỉ (0 suất, không tính vào ngày phục vụ).</p>
      <h3>Tổng kết tháng (ngày → suất ăn / vắng)</h3>
      <table>
        <thead><tr><th>Ngày</th><th>Suất ăn</th><th>Vắng</th><th>Chờ</th><th>Trạng thái</th></tr></thead>
        <tbody id="sumBody"></tbody>
      </table>
      <h3>Theo lớp — ngày đang chọn</h3>
      <table>
        <thead><tr><th>Lớp</th><th>Suất ăn</th><th>Vắng</th><th>Chờ</th></tr></thead>
        <tbody id="clsBody"></tbody>
      </table>
      <button id="btnCopy" type="button">📋 Sao chép báo cáo</button>
      <textarea id="reportText" readonly></textarea>
      <button class="expbtn" id="btnExpCsv" type="button">⬇ Xuất file CSV</button>
      <button class="expbtn" id="btnExpXlsx" type="button">⬇ Xuất file Excel (.xlsx)</button>
      <p class="hint">Bấm “Sao chép báo cáo” để copy đoạn văn bản trên gửi bếp. Số liệu cập nhật theo đúng trạng thái điểm danh hiện tại.</p>
    </div>
  </div>

  <div id="summaryView" hidden>
    <div class="top">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px">
        <button id="btnSumBack" type="button">← Quay lại</button>
      </div>
      <h1 id="sumTitle" style="margin-top:10px">Tổng hợp cuối tháng</h1>
      <p id="sumSub"></p>
    </div>
    <div style="padding:14px">
      <p class="sumprice">Giá ăn: <span id="sumMealPrice"></span> · Trông trưa chung: <span id="sumLunchFee"></span></p>
      <div id="settingsBox" hidden>
        <p style="font-size:15px;font-weight:700">Cài đặt chung (quản trị — đổi ở đây áp cho mọi phòng):</p>
        <p class="sumprice">Giá ăn/suất: <input id="mealPriceInp" type="number" min="1" step="1000">đ · Trông trưa chung: <input id="lunchFeeInp" type="number" min="0" step="1000">đ/tháng <button id="btnSettings" type="button">Lưu giá chung</button></p>
      </div>
      <p class="rule" id="sumRule"></p>
      <div id="sumTables"></div>
      <button class="expbtn" id="btnSumCsv" type="button">⬇ Xuất bảng tổng hợp (CSV)</button>
    </div>
  </div>
</div>

<script>
var ME = null, ROOMS = [], CUR = null, S = null, selDay = null, selCls = 'all', ACCOUNTS = null;
var WD = ['CN','T2','T3','T4','T5','T6','T7'];
function tokKey(id){ return 'dd_tok_' + id; }
function getTok(id){ try { return localStorage.getItem(tokKey(id)) || ''; } catch (e) { return ''; } }
function setTok(id, t){ try { localStorage.setItem(tokKey(id), t); } catch (e) {} }
function dropTok(id){ try { localStorage.removeItem(tokKey(id)); } catch (e) {} }
function wd(day){ return WD[new Date(CUR.year, CUR.month - 1, day).getDay()]; }
function fmt(n){ return n.toLocaleString('vi-VN'); }
function markOf(s, day){
  var v = s.marks[String(day)];
  return (v === 1 || v === 0) ? v : null;
}
function statusOf(d){
  for (var i = 0; i < S.summary.length; i++) if (S.summary[i].day === d) return S.summary[i].status;
  return '';
}
function eligible(s, day){
  return s.start_day === null || s.start_day === undefined || day >= s.start_day;
}
function counts(day){
  var a = 0, v = 0, c = 0, tot = 0;
  for (var i = 0; i < S.students.length; i++){
    var s = S.students[i];
    if (!eligible(s, day)) continue;
    tot++;
    var m = markOf(s, day);
    if (m === 1) a++; else if (m === 0) v++; else c++;
  }
  var suat = (statusOf(day) === 'Nghỉ') ? 0 : tot - v;
  return {suat: suat, v: v, c: c, a: a};
}
function stateLabel(m){
  if (m === 1) return '✓ Ăn';
  if (m === 0) return '✕ Vắng';
  return 'Chờ điểm danh';
}
function stateClass(m){
  if (m === 1) return 'eat';
  if (m === 0) return 'abs';
  return 'none';
}
function waitingOf(cls){
  var w = 0;
  for (var i = 0; i < S.students.length; i++){
    var s = S.students[i];
    if (!eligible(s, selDay)) continue;
    if (cls !== 'all' && s.class !== cls) continue;
    if (markOf(s, selDay) === null) w++;
  }
  return w;
}
function showView(id){
  closeMenu();
  ['roomsView', 'gateView', 'roomView', 'summaryView'].forEach(function(v){
    document.getElementById(v).hidden = (v !== id);
  });
  window.scrollTo(0, 0);
}
async function loadRooms(){
  var r = await fetch('api/rooms');
  if (!r.ok) return false;
  var d = await r.json();
  ROOMS = d.rooms;
  document.getElementById('roomsSub').textContent = 'Xin chào ' + d.user.name + (d.user.admin ? ' (Quản trị)' : '');
  document.getElementById('seedBox').hidden = !d.user.admin;
  if (ROOMS.length === 0){
    document.getElementById('roomsList').innerHTML = '<div class="roomcard"><p>Bạn chưa được gắn vào phòng nào. Liên hệ quản trị để được gắn phòng.</p></div>';
    return true;
  }
  var html = '';
  for (var i = 0; i < ROOMS.length; i++){
    var rm = ROOMS[i];
    var t = rm.today ? (rm.today.suat + ' suất ăn · ' + rm.today.vang + ' vắng · ' + rm.today.cho + ' chờ (hôm nay)') : ('Đã báo ' + rm.reported_n + '/' + rm.days_n + ' ngày');
    html += '<div class="roomcard" data-room="' + rm.id + '"><h2>' + rm.name + '</h2>'
      + '<p>Tháng ' + rm.month + '/' + rm.year + ' · ' + rm.n + ' học sinh · ' + (rm.has_password ? 'Đã đặt mật khẩu' : 'Chưa đặt mật khẩu') + '</p>'
      + '<p>' + t + '</p>'
      + '<p>Quản lý: ' + (rm.manager_names.length ? rm.manager_names.join(', ') : 'chưa gắn') + '</p>'
      + '<div class="btnrow"><button class="rbtn blue" data-act="enter">Vào phòng</button></div>';
    if (ME.admin){
      html += '<div class="btnrow"><button class="rbtn gray" data-act="managers">Gắn tài khoản</button>'
        + '<button class="rbtn gray" data-act="resetpw">Đặt lại mật khẩu</button></div>'
        + '<div class="picker" data-picker="managers" hidden></div>'
        + '<div class="picker" data-picker="resetpw" hidden><input type="password" placeholder="Mật khẩu mới (ít nhất 4 ký tự)" style="width:100%;min-height:44px;font-size:16px;border:2px solid #cbd5e1;border-radius:10px;padding:8px"><button class="rbtn blue" data-act="resetpw-save" style="margin-top:8px">Lưu mật khẩu mới</button><p class="err" style="color:#dc2626;font-weight:700"></p></div>';
    }
    html += '</div>';
  }
  document.getElementById('roomsList').innerHTML = html;
  return true;
}
async function enterRoom(id){
  CUR = null;
  for (var i = 0; i < ROOMS.length; i++) if (ROOMS[i].id === id) CUR = ROOMS[i];
  if (!CUR) return;
  var tok = getTok(id);
  var r = await fetch('api/rooms/' + id + '/state', {headers: tok ? {'X-Room-Token': tok} : {}});
  if (r.ok){
    S = await r.json();
    selCls = 'all';
    var today = new Date();
    var td = (today.getFullYear() === CUR.year && today.getMonth() + 1 === CUR.month) ? today.getDate() : null;
    if (td !== null && S.days.indexOf(td) >= 0) selDay = td;
    else {
      selDay = S.days[S.days.length - 1];
      for (var j = 0; j < S.days.length; j++) if (S.reported.indexOf(S.days[j]) < 0){ selDay = S.days[j]; break; }
    }
    document.getElementById('roomTitle').textContent = CUR.name + ' · Tháng ' + CUR.month + '/' + CUR.year;
    document.getElementById('roomSub').textContent = S.students.length + ' học sinh';
    document.getElementById('btnBaoCom').style.display = ME.admin ? '' : 'none';
    document.getElementById('nonAdminNote').style.display = ME.admin ? 'none' : 'block';
    document.getElementById('listView').hidden = false;
    document.getElementById('reportView').hidden = true;
    document.getElementById('addOnePanel').hidden = true;
    document.getElementById('bulkRoomPanel').hidden = true;
    showView('roomView');
    renderAll();
    return;
  }
  dropTok(id);
  document.getElementById('gateTitle').textContent = 'Vào ' + CUR.name;
  if (!CUR.has_password){
    document.getElementById('gateMsg').textContent = 'Phòng này chưa có mật khẩu.';
    document.getElementById('pwForm').hidden = true;
    document.getElementById('setForm').hidden = false;
  } else {
    document.getElementById('gateMsg').textContent = 'Nhập mật khẩu của ' + CUR.name + ' để vào điểm danh / báo cơm. (Mật khẩu chỉ nhớ trong phiên đăng nhập này; đăng xuất rồi đăng nhập lại sẽ hỏi lại.)';
    document.getElementById('pwForm').hidden = false;
    document.getElementById('setForm').hidden = true;
    document.getElementById('gateErr').textContent = '';
    document.getElementById('gatePw').value = '';
  }
  document.getElementById('setErr').textContent = '';
  document.getElementById('setPw').value = '';
  showView('gateView');
}
async function doUnlock(pw){
  var r = await fetch('api/rooms/' + CUR.id + '/unlock', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({password: pw})});
  if (!r.ok){
    var d = await r.json().catch(function(){ return {}; });
    return d.detail || 'Không vào được phòng.';
  }
  var dd = await r.json();
  setTok(CUR.id, dd.token);
  CUR.has_password = true;
  await enterRoom(CUR.id);
  return null;
}
function dayStatusText(d){
  var s = statusOf(d);
  if (s === 'Đã báo') return '✓ Đã báo';
  return s;
}
function renderMenuDays(){
  var head = document.getElementById('menuDaysHead');
  var el = document.getElementById('menuDays');
  if (!S || !CUR){ head.style.display = 'none'; el.innerHTML = ''; return; }
  head.style.display = '';
  el.innerHTML = S.days.map(function(d){
    return '<button type="button" class="mday' + (d === selDay ? ' sel' : '') + '" data-day="' + d + '">Ngày ' + d + '/' + CUR.month + ' (' + wd(d) + ') · ' + dayStatusText(d) + '</button>';
  }).join('');
}
function renderCurDayLine(){
  var el = document.getElementById('curDayLine');
  if (!S || selDay === null){ el.textContent = ''; return; }
  el.textContent = 'Ngày ' + selDay + '/' + CUR.month + ' (' + wd(selDay) + ') · ' + statusOf(selDay);
}
function renderCounters(){
  var c = counts(selDay);
  document.getElementById('cSuat').textContent = c.suat;
  document.getElementById('cVang').textContent = c.v;
  document.getElementById('cCho').textContent = c.c;
  document.getElementById('cDaAn').textContent = c.a;
}
function renderFilters(){
  var classes = [];
  for (var i = 0; i < S.students.length; i++) if (classes.indexOf(S.students[i].class) < 0) classes.push(S.students[i].class);
  classes.sort();
  var chips = ['<button type="button" class="fchip' + (selCls === 'all' ? ' sel' : '') + '" data-cls="all"><span class="dot ' + (waitingOf('all') > 0 ? 'red' : 'green') + '"></span>Tất cả</button>'].concat(classes.map(function(c){
    return '<button type="button" class="fchip' + (selCls === c ? ' sel' : '') + '" data-cls="' + c + '"><span class="dot ' + (waitingOf(c) > 0 ? 'red' : 'green') + '"></span>' + c + '</button>';
  }));
  document.getElementById('filters').innerHTML = chips.join('');
}
function renderList(){
  var all = S.students.filter(function(s){ return eligible(s, selDay); });
  var shown = (selCls === 'all') ? all : all.filter(function(s){ return s.class === selCls; });
  var note = document.getElementById('filterNote');
  if (selCls === 'all'){ note.hidden = true; }
  else {
    var gv = '';
    for (var i = 0; i < S.students.length; i++){
      if (S.students[i].class === selCls && S.students[i].note){ gv = S.students[i].note; break; }
    }
    note.hidden = false;
    note.textContent = 'Đang xem lớp ' + selCls + ': ' + shown.length + ' học sinh' + (gv ? ' · GVCN: ' + gv : '');
  }
  var groups = [], seen = {};
  for (var i = 0; i < shown.length; i++){
    var s = shown[i];
    if (!seen[s.class]){ seen[s.class] = []; groups.push([s.class, seen[s.class]]); }
    seen[s.class].push(s);
  }
  var html = '';
  for (var g = 0; g < groups.length; g++){
    var cls = groups[g][0], arr = groups[g][1];
    var w = 0;
    for (var k = 0; k < arr.length; k++) if (markOf(arr[k], selDay) === null) w++;
    html += '<div class="group"><h2 class="' + (w > 0 ? 'red' : 'green') + '">Lớp ' + cls + ' · ' + arr.length + ' học sinh · ' + (w > 0 ? ('Còn ' + w + ' chờ') : 'Đủ') + '</h2>';
    for (var j = 0; j < arr.length; j++){
      var st = arr[j], m = markOf(st, selDay);
      html += '<div class="row"><div class="who"><b' + (st.red ? ' class="red"' : '') + '>' + st.name + '</b><span>STT ' + st.stt + (st.note ? ' · ' + st.note : '') + (st.start_day ? ' · Từ ngày ' + st.start_day : '') + '</span></div>'
        + '<button type="button" class="st ' + stateClass(m) + '" data-stt="' + st.stt + '">' + stateLabel(m) + '</button></div>';
    }
    html += '</div>';
  }
  document.getElementById('list').innerHTML = html;
}
function syncStickyTops(){
  var info = document.getElementById('stickyInfo');
  var fl = document.getElementById('filters');
  if (!info || !fl) return;
  var h = info.offsetHeight;
  if (h > 0) fl.style.top = h + 'px';
}
window.addEventListener('resize', syncStickyTops);
function renderAll(){
  renderMenuDays();
  renderCurDayLine();
  renderCounters();
  renderFilters();
  renderList();
  if (!document.getElementById('reportView').hidden) renderReport();
  syncStickyTops();
}
function reportText(){
  var c = counts(selDay);
  var t = 'Báo cơm ' + CUR.name + ' ngày ' + selDay + '/' + CUR.month + ': ' + c.suat + ' suất.';
  if (c.c > 0) t += '\n(Còn ' + c.c + ' học sinh chờ điểm danh.)';
  var byCls = {}, order = [];
  for (var i = 0; i < S.students.length; i++){
    var s = S.students[i];
    if (!byCls[s.class]){ byCls[s.class] = 0; order.push(s.class); }
    if (markOf(s, selDay) !== 0) byCls[s.class]++;
  }
  var parts = [];
  for (var k = 0; k < order.length; k++) if (byCls[order[k]] > 0) parts.push(order[k] + ': ' + byCls[order[k]] + ' suất');
  if (parts.length) t += '\nChi tiết theo lớp: ' + parts.join('; ') + '.';
  return t;
}
function renderReport(){
  var c = counts(selDay);
  document.getElementById('repTitle').textContent = 'Báo cơm — ' + CUR.name;
  document.getElementById('repSub').textContent = 'Ngày ' + selDay + '/' + CUR.month + '/' + CUR.year + ' (' + wd(selDay) + ') · ' + CUR.name + ' · ' + S.students.length + ' học sinh';
  document.getElementById('repStatus').textContent = 'Trạng thái ngày ' + selDay + '/' + CUR.month + ': ' + statusOf(selDay);
  document.getElementById('cfDay').textContent = selDay;
  document.getElementById('bigNum').textContent = c.suat;
  document.getElementById('repSuat').textContent = c.suat;
  document.getElementById('repVang').textContent = c.v;
  document.getElementById('repCho').textContent = c.c;
  document.getElementById('repAn').textContent = c.a;
  document.getElementById('sumBody').innerHTML = S.summary.map(function(r){
    return '<tr><td>Ngày ' + r.day + ' (' + wd(r.day) + ')</td><td><b>' + r.suat + '</b></td><td>' + r.vang + '</td><td>' + r.cho + '</td><td class="sttcell">' + r.status + '</td></tr>';
  }).join('');
  var groups = {}, order = [];
  for (var i = 0; i < S.students.length; i++){
    var s = S.students[i];
    if (!eligible(s, selDay)) continue;
    if (!groups[s.class]){ groups[s.class] = {tot:0,v:0,c:0}; order.push(s.class); }
    groups[s.class].tot++;
    var m = markOf(s, selDay);
    if (m === 0) groups[s.class].v++;
    else if (m === null) groups[s.class].c++;
  }
  document.getElementById('clsBody').innerHTML = order.map(function(k){
    var g = groups[k];
    return '<tr><td>Lớp ' + k + '</td><td><b>' + (g.tot - g.v) + '</b></td><td>' + g.v + '</td><td>' + g.c + '</td></tr>';
  }).join('');
  document.getElementById('reportText').value = reportText();
}
function showReport(){
  if (!ME.admin) return;
  document.getElementById('addOnePanel').hidden = true;
  document.getElementById('bulkRoomPanel').hidden = true;
  document.getElementById('listView').hidden = true;
  document.getElementById('reportView').hidden = false;
  renderReport();
  window.scrollTo(0, 0);
}
function showList(){
  document.getElementById('reportView').hidden = true;
  document.getElementById('listView').hidden = false;
  renderAll();
  window.scrollTo(0, 0);
}
async function downloadWithToken(url, filename){
  var r = await fetch(url, {headers: {'X-Room-Token': getTok(CUR.id)}});
  if (!r.ok){ alert('Không tải được file.'); return; }
  var blob = await r.blob();
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); }, 5000);
}
async function openSummary(id){
  var meta = null;
  for (var i = 0; i < ROOMS.length; i++) if (ROOMS[i].id === id) meta = ROOMS[i];
  if (!meta) return;
  CUR = meta;
  var r = await fetch('api/rooms/' + id + '/summary');
  if (!r.ok){ alert('Không xem được tổng hợp.'); return; }
  var d = await r.json();
  var sm = d.summary;
  document.getElementById('sumTitle').textContent = (sm.is_final ? '' : '⚠️ ') + 'Tổng hợp tháng ' + sm.month + '/' + sm.year + ' — ' + sm.room + ' — ' + (sm.is_final ? 'TỔNG CHÍNH THỨC' : 'TẠM TÍNH đến hết ngày đã báo');
  document.getElementById('sumSub').textContent = 'Ngày đã báo: ' + sm.reported_days.join(', ') + (sm.is_final ? '' : ' · Sau cuối tháng mới chốt tổng chính thức.');
  document.getElementById('sumMealPrice').textContent = fmt(sm.meal_price) + 'đ/suất';
  document.getElementById('sumLunchFee').textContent = fmt(sm.lunch_fee_common) + 'đ/tháng';
  document.getElementById('mealPriceInp').value = sm.meal_price;
  document.getElementById('lunchFeeInp').value = sm.lunch_fee_common;
  document.getElementById('settingsBox').hidden = !ME.admin;
  document.getElementById('sumRule').textContent = 'QUY TẮC TÍNH (thầy kiểm tra, báo lại nếu khác): chỉ tính các ngày đã "Đã báo" (' + sm.reported_days.join(', ') + '). Ngày "Đang báo"/"Chưa tới" chưa tính. Học sinh ăn ngày đó khi mark khác 0 (1 = ăn xác nhận, trống = vẫn ăn theo luật chờ điểm danh; 0 = vắng không tính). Tiền ăn = buổi × giá ăn chung. Trông trưa = mức chung, hoặc mức riêng đã sửa cho học sinh đó. Tổng tiền = Tiền ăn + Trông trưa. Sửa ô Trông trưa của từng học sinh để miễn/giảm riêng; bấm ↺ để về lại mức chung.';
  var html = '';
  for (var i = 0; i < sm.classes.length; i++){
    var cl = sm.classes[i];
    html += '<h3>Lớp ' + cl.class + '</h3><table><thead><tr><th>STT</th><th>Họ và tên</th><th>Tổng số buổi ăn</th><th>Tiền ăn</th><th>Trông trưa</th><th>Tổng tiền</th></tr></thead><tbody>';
    for (var j = 0; j < cl.students.length; j++){
      var s = cl.students[j];
      html += '<tr><td>' + s.stt + '</td><td>' + s.name + '</td><td><b>' + s.meals + '</b></td><td>' + fmt(s.meal_money) + 'đ</td>'
        + '<td><input class="feeinp" data-stt="' + s.stt + '" type="number" min="0" step="1000" value="' + s.lunch_fee + '">'
        + (s.fee_override ? ' <button class="feereset" data-stt="' + s.stt + '" type="button">↺</button>' : '') + '</td>'
        + '<td>' + fmt(s.total) + 'đ</td></tr>';
    }
    html += '<tr class="total"><td></td><td>TỔNG LỚP ' + cl.class + '</td><td><b>' + cl.meals + '</b></td><td>' + fmt(cl.meal_money) + 'đ</td><td>' + fmt(cl.lunch_fee) + 'đ</td><td>' + fmt(cl.total) + 'đ</td></tr></tbody></table>';
  }
  html += '<h3>Tổng toàn phòng</h3><table><tbody><tr class="roomtotal"><td></td><td>TỔNG TOÀN PHÒNG</td><td><b>' + sm.room_meals + '</b></td><td>' + fmt(sm.room_meal_money) + 'đ</td><td>' + fmt(sm.room_lunch_fee) + 'đ</td><td>' + fmt(sm.room_money) + 'đ</td></tr></tbody></table>';
  document.getElementById('sumTables').innerHTML = html;
  document.getElementById('btnSumCsv').onclick = function(){
    var a = document.createElement('a');
    a.href = 'api/rooms/' + id + '/summary.csv';
    a.download = 'tong-hop-' + id + '.csv';
    a.click();
  };
  showView('summaryView');
}
document.getElementById('roomsList').addEventListener('click', async function(e){
  var btn = e.target.closest('[data-act]');
  if (!btn) return;
  var card = e.target.closest('.roomcard');
  var id = card.getAttribute('data-room');
  var act = btn.getAttribute('data-act');
  if (act === 'enter'){ enterRoom(id); return; }
  if (act === 'managers'){
    var pk = card.querySelector('[data-picker="managers"]');
    pk.hidden = !pk.hidden;
    if (!pk.hidden && !ACCOUNTS){
      var r = await fetch('api/rooms/accounts');
      ACCOUNTS = (await r.json()).accounts;
    }
    if (!pk.hidden){
      var rm = null;
      for (var i = 0; i < ROOMS.length; i++) if (ROOMS[i].id === id) rm = ROOMS[i];
      var h = '';
      for (var j = 0; j < ACCOUNTS.length; j++){
        var a = ACCOUNTS[j];
        h += '<label><input type="checkbox" value="' + a.key + '"' + (rm.managers.indexOf(a.key) >= 0 ? ' checked' : '') + '> ' + a.name + '</label>';
      }
      h += '<button class="rbtn blue" data-act="managers-save" style="margin-top:8px">Lưu gắn tài khoản</button>';
      pk.innerHTML = h;
    }
    return;
  }
  if (act === 'managers-save'){
    var pk2 = card.querySelector('[data-picker="managers"]');
    var keys = [];
    pk2.querySelectorAll('input[type=checkbox]:checked').forEach(function(cb){ keys.push(cb.value); });
    var r = await fetch('api/rooms/' + id + '/managers', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({keys: keys})});
    if (r.ok){ alert('Đã lưu gắn tài khoản.'); await loadRooms(); }
    else alert('Không lưu được.');
    return;
  }
  if (act === 'resetpw'){
    var rp = card.querySelector('[data-picker="resetpw"]');
    rp.hidden = !rp.hidden;
    return;
  }
  if (act === 'resetpw-save'){
    var rp2 = card.querySelector('[data-picker="resetpw"]');
    var npw = rp2.querySelector('input').value;
    var r = await fetch('api/rooms/' + id + '/set-password', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({password: npw})});
    if (r.ok){ alert('Đã đặt lại mật khẩu phòng.'); rp2.hidden = true; await loadRooms(); }
    else {
      var d = await r.json().catch(function(){ return {}; });
      rp2.querySelector('.err').textContent = d.detail || 'Không đặt được mật khẩu.';
    }
    return;
  }
});
document.getElementById('btnSeed').addEventListener('click', async function(){
  var txt = document.getElementById('seedJson').value;
  var msg = document.getElementById('seedMsg');
  var body = null;
  try { body = JSON.parse(txt); } catch (err) { msg.textContent = 'JSON không hợp lệ.'; return; }
  var r = await fetch('api/rooms', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)});
  if (r.ok){ msg.textContent = ''; document.getElementById('seedJson').value = ''; await loadRooms(); alert('Đã tạo phòng mới.'); }
  else {
    var d = await r.json().catch(function(){ return {}; });
    msg.textContent = d.detail || 'Không tạo được phòng.';
  }
});
document.getElementById('gateBtn').addEventListener('click', async function(){
  var pw = document.getElementById('gatePw').value;
  var err = await doUnlock(pw);
  if (err) document.getElementById('gateErr').textContent = err;
});
document.getElementById('gatePw').addEventListener('keydown', function(e){
  if (e.key === 'Enter') document.getElementById('gateBtn').click();
});
document.getElementById('setBtn').addEventListener('click', async function(){
  var pw = document.getElementById('setPw').value;
  var r = await fetch('api/rooms/' + CUR.id + '/set-password', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({password: pw})});
  if (!r.ok){
    var d = await r.json().catch(function(){ return {}; });
    document.getElementById('setErr').textContent = d.detail || 'Không đặt được mật khẩu.';
    return;
  }
  var err = await doUnlock(pw);
  if (err) document.getElementById('setErr').textContent = err;
});
document.getElementById('gateBack').addEventListener('click', async function(){
  CUR = null;
  await loadRooms();
  showView('roomsView');
});
document.getElementById('btnSumBack').addEventListener('click', async function(){
  if (getTok(CUR.id)){ await enterRoom(CUR.id); }
  else { await loadRooms(); showView('roomsView'); }
});
document.getElementById('btnSettings').addEventListener('click', async function(){
  var mp = parseInt(document.getElementById('mealPriceInp').value, 10);
  var lf = parseInt(document.getElementById('lunchFeeInp').value, 10);
  if (!mp || mp <= 0 || isNaN(lf) || lf < 0){ alert('Giá ăn phải > 0, trông trưa không được âm.'); return; }
  var r = await fetch('api/settings', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({meal_price: mp, lunch_fee: lf})});
  if (r.ok){ alert('Đã lưu giá chung (áp cho mọi phòng).'); openSummary(CUR.id); }
  else alert('Không lưu được giá chung.');
});
document.getElementById('sumTables').addEventListener('change', async function(e){
  var inp = e.target.closest('.feeinp');
  if (!inp) return;
  var stt = parseInt(inp.getAttribute('data-stt'), 10);
  var fee = parseInt(inp.value, 10);
  if (isNaN(fee) || fee < 0){ alert('Mức trông trưa không được âm.'); openSummary(CUR.id); return; }
  var r = await fetch('api/rooms/' + CUR.id + '/students/' + stt + '/lunch-fee', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({fee: fee})});
  if (r.ok) openSummary(CUR.id);
  else alert('Không lưu được mức trông trưa.');
});
document.getElementById('sumTables').addEventListener('click', async function(e){
  var btn = e.target.closest('.feereset');
  if (!btn) return;
  var stt = parseInt(btn.getAttribute('data-stt'), 10);
  var r = await fetch('api/rooms/' + CUR.id + '/students/' + stt + '/lunch-fee', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({fee: null})});
  if (r.ok) openSummary(CUR.id);
  else alert('Không đặt lại được mức trông trưa.');
});
document.getElementById('filters').addEventListener('click', function(e){
  var btn = e.target.closest('[data-cls]');
  if (!btn) return;
  selCls = btn.getAttribute('data-cls');
  renderFilters();
  renderList();
});
document.getElementById('btnConfirmBao').addEventListener('click', async function(){
  var r = await fetch('api/rooms/' + CUR.id + '/report', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-Room-Token': getTok(CUR.id)},
    body: JSON.stringify({day: selDay})});
  if (!r.ok){
    if (r.status === 403){ dropTok(CUR.id); enterRoom(CUR.id); }
    else alert('Không báo được. Hãy thử lại.');
    return;
  }
  var d = await r.json();
  S.reported = d.reported; S.summary = d.summary;
  selDay = d.next_day;
  showList();
});
function openMenu(){ renderMenuDays(); document.getElementById('menuPanel').hidden = false; document.getElementById('menuOverlay').hidden = false; }
function closeMenu(){ document.getElementById('menuPanel').hidden = true; document.getElementById('menuOverlay').hidden = true; }
function toast(msg){ var t = document.getElementById('toast'); t.textContent = msg; t.hidden = false; clearTimeout(t._h); t._h = setTimeout(function(){ t.hidden = true; }, 2500); }
document.getElementById('menuBtn').addEventListener('click', function(){
  var p = document.getElementById('menuPanel');
  if (p.hidden) openMenu(); else closeMenu();
});
document.getElementById('menuOverlay').addEventListener('click', closeMenu);
document.getElementById('menuPanel').addEventListener('click', async function(e){
  var mbtn = e.target.closest('[data-menu]');
  if (mbtn){
    var act = mbtn.getAttribute('data-menu');
    closeMenu();
    if (act === 'rooms'){ await loadRooms(); showView('roomsView'); return; }
    if (!CUR){ toast('Chọn phòng trước'); await loadRooms(); showView('roomsView'); return; }
    if (act === 'summary'){ openSummary(CUR.id); return; }
    if (document.getElementById('roomView').hidden || document.getElementById('listView').hidden){
      if (!getTok(CUR.id)){ toast('Vào phòng trước (nhập mật khẩu phòng)'); enterRoom(CUR.id); return; }
      await enterRoom(CUR.id);
    }
    if (act === 'addone') openAddOne(); else openBulkRoom();
    return;
  }
  var db = e.target.closest('[data-day]');
  if (db){
    var picked = parseInt(db.getAttribute('data-day'), 10);
    closeMenu();
    if (!CUR || !S){ toast('Chọn phòng trước'); return; }
    if (document.getElementById('roomView').hidden){
      if (!getTok(CUR.id)){ toast('Vào phòng trước (nhập mật khẩu phòng)'); enterRoom(CUR.id); return; }
      await enterRoom(CUR.id);
      if (S.days.indexOf(picked) >= 0) selDay = picked;
    } else {
      selDay = picked;
    }
    showList();
  }
});
document.getElementById('list').addEventListener('click', async function(e){
  var btn = e.target.closest('[data-stt]');
  if (!btn) return;
  var stt = parseInt(btn.getAttribute('data-stt'), 10);
  var s = null;
  for (var i = 0; i < S.students.length; i++) if (S.students[i].stt === stt){ s = S.students[i]; break; }
  if (!s) return;
  var cur = markOf(s, selDay);
  var next = (cur === null) ? 1 : (cur === 1) ? 0 : null;
  var r = await fetch('api/rooms/' + CUR.id + '/mark', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-Room-Token': getTok(CUR.id)},
    body: JSON.stringify({day: selDay, stt: stt, mark: next})});
  if (!r.ok){
    if (r.status === 403){ dropTok(CUR.id); enterRoom(CUR.id); }
    else alert('Không lưu được. Hãy thử lại.');
    return;
  }
  s.marks[String(selDay)] = next;
  renderCounters();
  renderFilters();
  renderList();
});
document.getElementById('btnBaoCom').addEventListener('click', showReport);
document.getElementById('btnBack').addEventListener('click', showList);
document.getElementById('btnExpCsv').addEventListener('click', function(){ downloadWithToken('api/rooms/' + CUR.id + '/export', 'diem-danh-' + CUR.id + '.csv'); });
document.getElementById('btnExpXlsx').addEventListener('click', function(){ downloadWithToken('api/rooms/' + CUR.id + '/export-xlsx', 'diem-danh-' + CUR.id + '.xlsx'); });
document.getElementById('btnCopy').addEventListener('click', function(){
  var ta = document.getElementById('reportText');
  var txt = ta.value;
  var done = function(){ var b = document.getElementById('btnCopy'); b.textContent = '✓ Đã sao chép'; setTimeout(function(){ b.textContent = '📋 Sao chép báo cáo'; }, 2000); };
  if (navigator.clipboard && navigator.clipboard.writeText){
    navigator.clipboard.writeText(txt).then(done).catch(function(){ ta.select(); done(); });
  } else { ta.select(); done(); }
});
function gvOf(cls){
  for (var i = 0; i < S.students.length; i++) if (S.students[i].class === cls && S.students[i].note) return S.students[i].note;
  return '';
}
function roomClasses(){
  var out = [];
  for (var i = 0; i < S.students.length; i++) if (out.indexOf(S.students[i].class) < 0) out.push(S.students[i].class);
  return out.sort();
}
function defaultStartForRoom(){
  var t = new Date();
  if (CUR && t.getFullYear() === CUR.year && t.getMonth() + 1 === CUR.month){
    for (var i = 0; i < S.days.length; i++) if (S.days[i] >= t.getDate()) return S.days[i];
    return S.days[S.days.length - 1];
  }
  return S.days[0];
}
function fillStartSelect(sel, days, def){
  sel.innerHTML = days.map(function(d){
    return '<option value="' + d + '"' + (d === def ? ' selected' : '') + '>Ngày ' + d + '</option>';
  }).join('');
}
function parsePaste(text){
  var lines = text.split('\n');
  var entries = [], errors = [];
  for (var i = 0; i < lines.length; i++){
    var ln = lines[i].trim();
    if (!ln) continue;
    var parts = ln.split('—');
    if (parts.length < 2) parts = ln.split('–');
    if (parts.length < 2) parts = ln.split(' - ');
    if (parts.length < 2){ errors.push('Dòng ' + (i + 1) + ': không tách được tên và lớp (dạng "Họ tên — Lớp")'); continue; }
    var nm = parts[0].trim();
    var cl = parts.slice(1).join('—').trim();
    if (!nm || !cl){ errors.push('Dòng ' + (i + 1) + ': thiếu tên hoặc lớp'); continue; }
    entries.push({name: nm, class: cl, note: null, marks: {}});
  }
  return {entries: entries, errors: errors};
}
function bulkPreviewHTML(entries, errors, extra){
  var classes = [];
  entries.forEach(function(e){ if (classes.indexOf(e.class) < 0) classes.push(e.class); });
  var h = '<p><b>Xem trước:</b> ' + entries.length + ' học sinh · ' + classes.length + ' lớp (' + classes.join(', ') + ')' + (extra ? ' · ' + extra : '') + '</p>';
  if (errors.length) h += '<p style="color:#dc2626;font-weight:700">Lỗi (sửa trước khi xác nhận):<br>' + errors.join('<br>') + '</p>';
  if (entries.length){
    h += '<table><thead><tr><th>Họ và tên</th><th>Lớp</th></tr></thead><tbody>';
    var show = entries.slice(0, 10);
    for (var i = 0; i < show.length; i++) h += '<tr><td>' + show[i].name + '</td><td>' + show[i].class + '</td></tr>';
    if (entries.length > 10) h += '<tr><td colspan="2">… và ' + (entries.length - 10) + ' bạn nữa</td></tr>';
    h += '</tbody></table>';
  }
  return h;
}
var BR = {entries: [], errors: [], sheets: null};
var BC = {entries: [], errors: [], sheets: null, days: []};
function refreshBrPreview(){
  var p = parsePaste(document.getElementById('brPaste').value);
  BR.entries = p.entries; BR.errors = p.errors;
  document.getElementById('brPreview').innerHTML = bulkPreviewHTML(BR.entries, BR.errors, '');
  document.getElementById('brConfirm').disabled = (BR.entries.length === 0 || BR.errors.length > 0);
}
function refreshBcPreview(){
  var p = parsePaste(document.getElementById('bcPaste').value);
  BC.entries = p.entries; BC.errors = p.errors;
  var daysTxt = document.getElementById('bcDays').value.trim();
  document.getElementById('bcPreview').innerHTML = bulkPreviewHTML(BC.entries, BC.errors, '');
  var daysOk = daysTxt.length > 0;
  document.getElementById('bcConfirm').disabled = (BC.entries.length === 0 || BC.errors.length > 0 || !daysOk);
}
function openAddOne(){
  var p = document.getElementById('addOnePanel');
  p.hidden = false;
  var sel = document.getElementById('aoClass');
  sel.innerHTML = roomClasses().map(function(c){ return '<option value="' + c + '">' + c + '</option>'; }).join('') + '<option value="__new">＋ Lớp mới...</option>';
  document.getElementById('aoClassNew').hidden = true;
  document.getElementById('aoNote').value = gvOf(sel.value);
  fillStartSelect(document.getElementById('aoStart'), S.days, defaultStartForRoom());
  document.getElementById('aoMsg').textContent = '';
  p.scrollIntoView({block: 'start'});
}
document.getElementById('aoClass').addEventListener('change', function(){
  var v = this.value;
  document.getElementById('aoClassNew').hidden = (v !== '__new');
  document.getElementById('aoNote').value = (v === '__new') ? '' : gvOf(v);
});
document.getElementById('aoSave').addEventListener('click', async function(){
  var name = document.getElementById('aoName').value.trim();
  var cv = document.getElementById('aoClass').value;
  var cls = (cv === '__new') ? document.getElementById('aoClassNew').value.trim() : cv;
  var note = document.getElementById('aoNote').value.trim();
  var sd = parseInt(document.getElementById('aoStart').value, 10);
  var msg = document.getElementById('aoMsg');
  if (!name || !cls){ msg.textContent = 'Nhập đủ họ tên và lớp.'; return; }
  var r = await fetch('api/rooms/' + CUR.id + '/students', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-Room-Token': getTok(CUR.id)},
    body: JSON.stringify({name: name, class: cls, note: note, start_day: sd})});
  if (!r.ok){
    var d = await r.json().catch(function(){ return {}; });
    msg.textContent = d.detail || 'Không thêm được.';
    return;
  }
  var dd = await r.json();
  alert('Đã thêm ' + name + ' (STT ' + dd.stt + ', từ ngày ' + dd.start_day + ').');
  document.getElementById('addOnePanel').hidden = true;
  document.getElementById('aoName').value = '';
  await enterRoom(CUR.id);
});
function openBulkRoom(){
  var p = document.getElementById('bulkRoomPanel');
  p.hidden = false;
  fillStartSelect(document.getElementById('brStart'), S.days, defaultStartForRoom());
  refreshBrPreview();
  p.scrollIntoView({block: 'start'});
}
document.getElementById('brPaste').addEventListener('input', refreshBrPreview);
document.getElementById('brFile').addEventListener('change', async function(){
  var f = this.files[0];
  if (!f) return;
  var buf = await f.arrayBuffer();
  var r = await fetch('api/rooms/excel-preview', {method: 'POST', body: buf});
  if (!r.ok){ document.getElementById('brMsg').textContent = 'Không đọc được file Excel.'; return; }
  var d = await r.json();
  BR.sheets = d.sheets;
  var sel = document.getElementById('brSheet');
  sel.innerHTML = d.sheets.map(function(s, i){ return '<option value="' + i + '">' + s.name + (s.error ? ' (lỗi)' : ' (' + (s.students ? s.students.length : 0) + ' HS)') + '</option>'; }).join('');
  document.getElementById('brSheetSel').hidden = false;
  applyBrSheet();
});
function applyBrSheet(){
  var idx = parseInt(document.getElementById('brSheet').value, 10);
  var sh = BR.sheets[idx];
  if (sh.error){
    BR.entries = []; BR.errors = [sh.error];
  } else {
    BR.entries = sh.students; BR.errors = sh.errors || [];
  }
  document.getElementById('brPreview').innerHTML = bulkPreviewHTML(BR.entries, BR.errors, 'Ngày trong file: ' + (sh.days ? sh.days.join(', ') : ''));
  document.getElementById('brConfirm').disabled = (BR.entries.length === 0 || BR.errors.length > 0);
}
document.getElementById('brSheet').addEventListener('change', applyBrSheet);
document.getElementById('brConfirm').addEventListener('click', async function(){
  var sd = parseInt(document.getElementById('brStart').value, 10);
  var entries = BR.entries.map(function(e){
    var note = e.note;
    if (!note) note = gvOf(e.class) || null;
    return {name: e.name, class: e.class, note: note, marks: e.marks || {}};
  });
  var r = await fetch('api/rooms/' + CUR.id + '/bulk-append', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-Room-Token': getTok(CUR.id)},
    body: JSON.stringify({entries: entries, start_day: sd})});
  if (!r.ok){
    var d = await r.json().catch(function(){ return {}; });
    document.getElementById('brMsg').textContent = d.detail || 'Không thêm được danh sách.';
    return;
  }
  var dd = await r.json();
  alert('Đã thêm ' + dd.added + ' học sinh (từ STT ' + dd.first_stt + ', từ ngày ' + dd.start_day + ').');
  document.getElementById('bulkRoomPanel').hidden = true;
  document.getElementById('brPaste').value = '';
  document.getElementById('brFile').value = '';
  await enterRoom(CUR.id);
});
function bcParseDays(){
  var txt = document.getElementById('bcDays').value.trim();
  if (!txt) return [];
  var out = [];
  txt.split(',').forEach(function(x){
    var v = parseInt(x.trim(), 10);
    if (!isNaN(v) && v >= 1 && v <= 31 && out.indexOf(v) < 0) out.push(v);
  });
  return out.sort(function(a, b){ return a - b; });
}
document.getElementById('bcDays').addEventListener('input', function(){
  var days = bcParseDays();
  BC.days = days;
  if (days.length) fillStartSelect(document.getElementById('bcStart'), days, days[0]);
  refreshBcPreview();
});
document.getElementById('bcPaste').addEventListener('input', refreshBcPreview);
document.getElementById('bcFile').addEventListener('change', async function(){
  var f = this.files[0];
  if (!f) return;
  var buf = await f.arrayBuffer();
  var r = await fetch('api/rooms/excel-preview', {method: 'POST', body: buf});
  if (!r.ok){ document.getElementById('bcMsg').textContent = 'Không đọc được file Excel.'; return; }
  var d = await r.json();
  BC.sheets = d.sheets;
  var sel = document.getElementById('bcSheet');
  sel.innerHTML = d.sheets.map(function(s, i){ return '<option value="' + i + '">' + s.name + (s.error ? ' (lỗi)' : ' (' + (s.students ? s.students.length : 0) + ' HS)') + '</option>'; }).join('');
  document.getElementById('bcSheetSel').hidden = false;
  applyBcSheet();
});
function applyBcSheet(){
  var idx = parseInt(document.getElementById('bcSheet').value, 10);
  var sh = BC.sheets[idx];
  if (sh.error){
    BC.entries = []; BC.errors = [sh.error];
    document.getElementById('bcPreview').innerHTML = bulkPreviewHTML([], BC.errors, '');
  } else {
    BC.entries = sh.students; BC.errors = sh.errors || [];
    BC.days = sh.days || [];
    document.getElementById('bcDays').value = BC.days.join(',');
    fillStartSelect(document.getElementById('bcStart'), BC.days, BC.days[0]);
    document.getElementById('bcPreview').innerHTML = bulkPreviewHTML(BC.entries, BC.errors, 'Ngày trong file: ' + BC.days.join(', '));
  }
  document.getElementById('bcConfirm').disabled = (BC.entries.length === 0 || BC.errors.length > 0 || BC.days.length === 0);
}
document.getElementById('bcSheet').addEventListener('change', applyBcSheet);
document.getElementById('bcConfirm').addEventListener('click', async function(){
  var name = document.getElementById('bcName').value.trim();
  var month = parseInt(document.getElementById('bcMonth').value, 10);
  var year = parseInt(document.getElementById('bcYear').value, 10);
  var sd = parseInt(document.getElementById('bcStart').value, 10);
  var msg = document.getElementById('bcMsg');
  if (!name){ msg.textContent = 'Nhập tên phòng.'; return; }
  var days = BC.days.length ? BC.days : bcParseDays();
  if (!days.length){ msg.textContent = 'Nhập ngày phục vụ.'; return; }
  var r = await fetch('api/rooms/bulk', {method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name: name, month: month, year: year, days: days, entries: BC.entries, start_day: sd})});
  if (!r.ok){
    var d = await r.json().catch(function(){ return {}; });
    msg.textContent = d.detail || 'Không tạo được phòng.';
    return;
  }
  var dd = await r.json();
  alert('Đã tạo ' + dd.room.name + ' với ' + BC.entries.length + ' học sinh. Vào phòng để đặt mật khẩu.');
  await loadRooms();
  enterRoom(dd.room.id);
});
async function boot(){
  var me = null;
  try {
    var r = await fetch('/api/me');
    if (r.ok) me = await r.json();
  } catch (e) {}
  if (!me){
    document.getElementById('loginGate').hidden = false;
    return;
  }
  ME = me;
  document.getElementById('app').hidden = false;
  var _t = new Date();
  if (!document.getElementById('bcMonth').value) document.getElementById('bcMonth').value = _t.getMonth() + 1;
  if (!document.getElementById('bcYear').value) document.getElementById('bcYear').value = _t.getFullYear();
  await loadRooms();
  showView('roomsView');
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(function(){});
}
boot();
</script>
</body>
</html>
"""

SW = """self.addEventListener('install', (e) => {
  e.waitUntil(caches.open('diemdanh-v2').then((c) => c.addAll(['./', './manifest.json'])));
});
self.addEventListener('fetch', (e) => {
  var u = new URL(e.request.url);
  if (u.pathname.indexOf('/diem-danh/api/') >= 0) return;
  e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
});
"""

MANIFEST = """{"name": "Điểm danh bán trú", "short_name": "Điểm danh", "start_url": "/diem-danh/", "display": "standalone", "background_color": "#ffffff", "theme_color": "#2563eb", "icons": []}"""
