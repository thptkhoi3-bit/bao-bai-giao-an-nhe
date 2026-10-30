"""Tạo báo bài / giáo án mẫu của 'trường khác' (khác mẫu anh Đạt) để kiểm thử."""
import docx, copy
from docx.shared import Pt, Mm
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def _font(doc, name, size):
    st = doc.styles['Normal']; st.font.name = name; st.font.size = Pt(size)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), name)
def _put(cell, text, font=None, size=None, bold=False):
    p = cell.paragraphs[0]
    for r in list(p.runs): r._r.getparent().remove(r._r)
    r = p.add_run(text); r.bold = bold
    if font: r.font.name = font
    if size: r.font.size = Pt(size)
def _vmerge(t, col, r0, r1):
    if r1 > r0: t.cell(r0, col).merge(t.cell(r1, col))
def landscape(doc):
    s = doc.sections[0]; s.orientation = WD_ORIENT.LANDSCAPE; s.page_width, s.page_height = Mm(297), Mm(210)

DAYN = {2: 'Hai', 3: 'Ba', 4: 'Tư', 5: 'Năm', 6: 'Sáu'}
# lịch tuần (thứ, tiết, môn) của lớp ở trường khác
SCHED = [(2, 1, 'Chào cờ'), (2, 2, 'Tiếng Việt'), (2, 3, 'Tiếng Việt'), (2, 4, 'Toán'), (2, 5, 'Đạo đức'),
         (3, 1, 'Tiếng Việt'), (3, 2, 'Toán'), (3, 3, 'Tự nhiên và Xã hội'), (3, 4, 'Hoạt động trải nghiệm'),
         (4, 1, 'Tiếng Việt'), (4, 2, 'Tiếng Việt'), (4, 3, 'Toán'),
         (5, 1, 'Toán'), (5, 2, 'Tiếng Việt'),
         (6, 1, 'Tiếng Việt'), (6, 2, 'Sinh hoạt lớp')]

def table_rows(t, cols, sched, data=None, font=None, size=None, merge_day=True):
    """cols: danh sách vai trò theo thứ tự cột: thu/ngay/tiet/buoi/mon/ten/tich/dd."""
    data = data or {}; start = len(t.rows)
    for (thu, tiet, mon) in sched:
        c = t.add_row().cells; d = data.get((thu, tiet), {})
        for j, k in enumerate(cols):
            v = {'thu': str(thu), 'ngay': f'Thứ {DAYN[thu]}\n{18 + thu}/10', 'tiet': str(tiet), 'buoi': 'Sáng', 'mon': mon}.get(k, d.get(k, ''))
            _put(c[j], v, font, size)
    if merge_day:
        j = cols.index('thu') if 'thu' in cols else (cols.index('ngay') if 'ngay' in cols else None)
        if j is not None:
            i = start
            while i < len(t.rows):
                k = i
                while k + 1 < len(t.rows) and sched[k + 1 - start][0] == sched[i - start][0]: k += 1
                if k > i:
                    txt = t.cell(i, j).text; t.cell(i, j).merge(t.cell(k, j)); _put(t.cell(i, j), txt, font, size)
                i = k + 1

def baobai_A(path, weeks=(6, 7), filled=(6,), data=None):
    """Ngày | Tiết | Môn học | Tên bài dạy | ĐDDH (không có cột Tích hợp); 'Tuần 7 (từ 19/10 đến 23/10)'; ngang; Arial 13."""
    d = docx.Document(); _font(d, 'Arial', 13); landscape(d)
    d.add_paragraph('TRƯỜNG TIỂU HỌC HOA SEN – Năm học 2026-2027')
    for w in weeks:
        a = 19 + 7 * (w - 7)
        d.add_paragraph(f'Tuần {w} (từ {a}/10 đến {a + 4}/10)' if a > 0 else f'Tuần {w} (từ {a + 30}/9 đến {a + 34}/9)')
        cols = ['ngay', 'tiet', 'mon', 'ten', 'dd']
        t = d.add_table(rows=1, cols=5); t.style = 'Table Grid'
        for j, h in enumerate(['Ngày', 'Tiết', 'Môn học', 'Tên bài dạy', 'ĐDDH']): _put(t.rows[0].cells[j], h, 'Arial', 13, True)
        dat = (data or {}).get(w) or ({k: dict(ten=f'Bài cũ tuần {w} – {m}', dd='Tranh') for k, m in [((s[0], s[1]), s[2]) for s in SCHED]} if w in filled else {})
        table_rows(t, cols, SCHED, dat, 'Arial', 13)
    d.save(path); return path

def baobai_B(path, data=None):
    """MỘT bảng cho cả năm, dòng gộp 'TUẦN n …'; cột: Thứ | Buổi | Tiết | Phân môn | Nội dung | Điều chỉnh, tích hợp; Calibri 11."""
    d = docx.Document(); _font(d, 'Calibri', 11)
    d.add_paragraph('LỊCH BÁO GIẢNG CẢ NĂM')
    t = d.add_table(rows=1, cols=6); t.style = 'Table Grid'
    for j, h in enumerate(['Thứ', 'Buổi', 'Tiết', 'Phân môn', 'Nội dung', 'Điều chỉnh, tích hợp']): _put(t.rows[0].cells[j], h, 'Calibri', 11, True)
    for w in (6, 7):
        r = t.add_row(); c = r.cells[0].merge(r.cells[5]); a = 19 + 7 * (w - 7)
        _put(c, f'TUẦN {w}: {a if a > 0 else a + 30}/{10 if a > 0 else 9}/2026 - {a + 4 if a > 0 else a + 34}/{10 if a > 0 else 9}/2026', 'Calibri', 11, True)
        table_rows(t, ['thu', 'buoi', 'tiet', 'mon', 'ten', 'tich'], SCHED, (data or {}).get(w, {}), 'Calibri', 11)
    d.save(path); return path

def baobai_C(path, data=None):
    """Tiêu đề 2 dòng (gộp), thứ tự cột khác: Tiết | Thứ | Tên bài | Môn | Tích hợp; 'KẾ HOẠCH DẠY HỌC TUẦN 7 – Từ 19/10/2026 đến 23/10/2026'."""
    d = docx.Document(); _font(d, 'Times New Roman', 14)
    d.add_paragraph('KẾ HOẠCH DẠY HỌC TUẦN 7 – Từ 19/10/2026 đến 23/10/2026')
    t = d.add_table(rows=2, cols=5); t.style = 'Table Grid'
    h0, h1 = t.rows[0].cells, t.rows[1].cells
    _put(h0[0], 'Tiết'); _put(h0[1], 'Thứ'); _put(h0[2], 'Nội dung dạy học'); _put(h0[4], 'Tích hợp')
    _put(h1[2], 'Tên bài'); _put(h1[3], 'Môn')
    t.cell(0, 0).merge(t.cell(1, 0)); t.cell(0, 1).merge(t.cell(1, 1)); t.cell(0, 2).merge(t.cell(0, 3)); t.cell(0, 4).merge(t.cell(1, 4))
    table_rows(t, ['tiet', 'thu', 'ten', 'mon', 'tich'], SCHED, (data or {}).get(7, {}), 'Times New Roman', 14)
    d.save(path); return path

def baobai_D(path):
    """Tiêu đề cột không theo quy ước (A, B, C…) -> phải ghép cột thủ công."""
    d = docx.Document(); d.add_paragraph('TUẦN 7 (19/10/2026 - 23/10/2026)')
    t = d.add_table(rows=1, cols=4); t.style = 'Table Grid'
    for j, h in enumerate(['Cột A', 'Cột B', 'Cột C', 'Cột D']): _put(t.rows[0].cells[j], h)
    table_rows(t, ['thu', 'mon', 'ten', 'tich'], SCHED, {})
    d.save(path); return path

LESS = {('Tiếng Việt', i): n for i, n in enumerate(['Đọc: Ngôi trường mới (T1)', 'Đọc: Ngôi trường mới (T2)', 'Viết: Chữ hoa G', 'Nói và nghe: Kể chuyện',
                                                     'Đọc: Mẹ của em (T1)', 'Đọc: Mẹ của em (T2)', 'Luyện từ và câu'])}
LESS.update({('Toán', i): n for i, n in enumerate(['Bài 14: Phép cộng (T1)', 'Bài 14: Phép cộng (T2)', 'Bài 15: Luyện tập'])})  # phụ chỉ có 3 tiết Toán (gốc có 4)
LESS.update({('Đạo đức', 0): 'Bài 4: Kính yêu thầy cô', ('Tự nhiên và Xã hội', 0): 'Bài 8: Ngôi nhà của em', ('Hoạt động trải nghiệm', 0): 'Chủ đề 2: Em yêu trường',
             ('Sinh hoạt lớp', 0): 'Sơ kết tuần 7'})
ABB = {'Tiếng Việt': 'TV', 'Tự nhiên và Xã hội': 'TNXH', 'Hoạt động trải nghiệm': 'HĐTN', 'Đạo đức': 'ĐĐ', 'Sinh hoạt lớp': 'SHL'}

def phu_rows(sched):
    cnt = {}; out = []
    for thu, tiet, mon in sched:
        n = cnt.get(mon, 0); cnt[mon] = n + 1
        if (mon, n) in LESS: out.append((thu, tiet, ABB.get(mon, mon), LESS[(mon, n)], 'NLS 1.1' if mon == 'Tiếng Việt' and n == 0 else ''))
    return out

def phu_X(path, sched=None):
    """Báo bài PHỤ khác bố cục: Môn | Thứ | Tiết | Tên bài | Tích hợp (viết tắt môn: TV, TNXH, HĐTN, ĐĐ); lịch khác ngày một số môn."""
    sched = sched or [(3 if m == 'Đạo đức' else t_, ti, m) for t_, ti, m in SCHED]  # Đạo đức học thứ Ba ở lớp phụ (cùng số tiết -> vẫn chắc chắn)
    d = docx.Document(); d.add_paragraph('BÁO GIẢNG TUẦN 7'); d.add_paragraph('Từ ngày 19/10/2026 đến ngày 23/10/2026')
    t = d.add_table(rows=1, cols=5); t.style = 'Table Grid'
    for j, h in enumerate(['Môn', 'Thứ', 'Tiết', 'Tên bài', 'Tích hợp']): _put(t.rows[0].cells[j], h)
    for thu, tiet, mon, ten, tich in phu_rows(sched):
        c = t.add_row().cells
        for j, v in enumerate([mon, f'Thứ {DAYN[thu]}', str(tiet), ten, tich]): _put(c[j], v)
    d.save(path); return path

def giaoan_V(path, master, w=7, anh_template='/workspace/GiaoAn_Goc.docx'):
    """Giáo án gốc kiểu khác: bảng Tiết | Ngày | Môn học | Tên bài dạy | Ghi chú (không có Buổi); 'THỨ 2, ngày 01/01/2026'; 'Tiết 2: Tiếng Việt'; Arial 12."""
    import core, re
    rows = core.rows_of(core.week_tables(docx.Document(master))[w][0])
    heads = core.tiet_headings(docx.Document(anh_template))
    d = docx.Document(); _font(d, 'Arial', 12)
    d.add_paragraph('KẾ HOẠCH BÀI DẠY TUẦN 1 TỪ NGÀY 1/1/2026 - 5/1/2026')
    t = d.add_table(rows=1, cols=5); t.style = 'Table Grid'
    for j, h in enumerate(['Tiết', 'Ngày', 'Môn học', 'Tên bài dạy', 'Ghi chú']): _put(t.rows[0].cells[j], h, 'Arial', 12, True)
    n = {}
    for r in rows:
        if not r['mon']: continue
        n[r['thu']] = n.get(r['thu'], 0) + 1
        c = t.add_row().cells
        for j, v in enumerate([str(n[r['thu']]), f"Thứ {DAYN.get(int(r['thu']), r['thu'])}", r['mon'], '', '']): _put(c[j], v, 'Arial', 12)
    day = None
    for h in heads:
        if h['thu'] != day:
            day = h['thu']; d.add_paragraph(f'THỨ {day}, ngày 01/01/2026')
        nums = re.findall(r'\d+', h['label'].split(h['mon'])[0])
        p = d.add_paragraph(f"Tiết {' + '.join(nums)}: {h['mon']}"); p.runs[0].bold = True
        d.add_paragraph('(soạn bài ở đây)')
    d.save(path); return path
