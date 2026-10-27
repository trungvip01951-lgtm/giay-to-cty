#!/usr/bin/env python3
"""Tạo các mẫu giấy tờ Trung Kiên (.docx) từ du-lieu/cong_ty_trung_kien.json.
Chạy: python3 cong-cu/tao_mau.py   -> ghi vào mau-giay-to/
Mọi thông tin công ty lấy từ JSON, không gõ lại trong mã."""
import json
import pathlib
import sys
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "mau-giay-to"
OUT_DIR = None
TK = json.load(open(ROOT / "du-lieu/cong_ty_trung_kien.json", encoding="utf-8"))["cong_ty"]
DD = TK["dai_dien"]
BANK = TK["ngan_hang"]
FONT = "Times New Roman"
# Dấu + chữ ký (ảnh PNG ghép, KHÔNG lưu trong kho). Đặt ANH_DAU = đường dẫn để chèn vào khối ký của Giám đốc.
ANH_DAU = None
SEAL_CM = 4.0                 # đường kính dấu tròn ngoài thực tế (giả định 40 mm, chờ anh đo dấu thật)
SEAL_PX, GHEP_PX = 575, (998, 780)   # đường kính dấu và kích thước ảnh ghép dau_chu_ky_moi_ghep.png


def stamp_size():
    k = SEAL_CM / SEAL_PX
    return GHEP_PX[0] * k, GHEP_PX[1] * k   # (rộng cm, cao cm)
C, L, R, J = (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
              WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.JUSTIFY)


def dots(n=40):
    return "." * n


def new_doc():
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.top_margin = s.bottom_margin = Cm(2)
    s.left_margin, s.right_margin = Cm(3), Cm(2)
    st = d.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_after = Pt(3)
    st.paragraph_format.space_before = Pt(0)
    return d


def run(p, text, bold=False, italic=False, size=None):
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if size:
        r.font.size = Pt(size)
    return r


def para(d, text="", bold=False, italic=False, align=J, size=None, after=3, indent=None, keep=False):
    p = d.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    if keep:
        p.paragraph_format.keep_with_next = True
    if text:
        run(p, text, bold, italic, size)
    return p


def mixed(d, parts, align=J, after=3):
    """parts: list of (text, bold)"""
    p = para(d, align=align, after=after)
    for t, b in parts:
        run(p, t, bold=b)
    return p


def shade_borders(tbl, borders=True):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}")
        x.set(qn("w:val"), "single" if borders else "nil")
        x.set(qn("w:sz"), "4")
        x.set(qn("w:color"), "000000")
        b.append(x)
    tblPr.append(b)


def fix_layout(t, widths):
    """Cố định độ rộng cột (Word và LibreOffice đều tôn trọng) và không cho dòng bảng bị cắt giữa 2 trang."""
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    grid = t._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        for c, w in zip(row.cells, widths):
            c.width = Cm(w)


def table(d, headers, rows, widths, size=12, center_cols=()):
    t = d.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    shade_borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.width = Cm(widths[i])
        c.paragraphs[0].alignment = C
        run(c.paragraphs[0], h, bold=True, size=size)
    for r_ in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r_):
            cells[i].width = Cm(widths[i])
            p = cells[i].paragraphs[0]
            p.alignment = C if i in center_cols else L
            run(p, v, size=size, bold=str(v).startswith("Tổng") or str(v).startswith("Cộng"))
    fix_layout(t, widths)
    gap = d.add_paragraph()
    gap.paragraph_format.space_after = Pt(0)
    gap.paragraph_format.line_spacing = Pt(6)
    return t


def quoc_hieu(d):
    para(d, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True, align=C, after=0)
    para(d, "Độc lập - Tự do - Hạnh phúc", bold=True, align=C, after=0)
    para(d, "------------o0o------------", align=C, after=8)


def title(d, text, so=None, sub=None):
    para(d, text, bold=True, align=C, size=15, after=2, keep=True)
    if so:
        para(d, so, italic=True, align=C, after=2)
    if sub:
        para(d, sub, align=C, after=6)


def date_line(d, place="Hà Nội", align=R):
    para(d, f"{place}, ngày ..... tháng ..... năm 20....", italic=True, align=align, after=6)


def header_cty(d):
    para(d, TK["ten"], bold=True, align=L, after=0)
    para(d, f"Địa chỉ: {TK['dia_chi']}", align=L, after=0, size=12)
    para(d, f"Mã số thuế: {TK['mst']}   |   Điện thoại: {TK['dien_thoai']}", align=L, after=6, size=12)


def ben_tk(d, label, goi):
    mixed(d, [(f"{label}: ", True), (TK["ten"], True), (f" ({goi})", False)])
    para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    para(d, f"Điện thoại: {TK['dien_thoai']}", after=0)
    para(d, f"Mã số thuế: {TK['mst']}", after=0)
    para(d, f"Số tài khoản: {BANK['so_tk']} tại {BANK['ten']}", after=0)
    para(d, f"Đại diện: {DD['ho_ten']}          Chức vụ: {DD['chuc_vu']}", after=6)


def ben_kh(d, label, goi, extra_tk=True):
    mixed(d, [(f"{label}: ", True), (dots(60), False), (f" ({goi})", False)])
    para(d, f"Địa chỉ: {dots(70)}", after=0)
    para(d, f"Điện thoại: {dots(25)}   Mã số thuế: {dots(25)}", after=0)
    if extra_tk:
        para(d, f"Số tài khoản: {dots(30)} tại {dots(30)}", after=0)
    para(d, f"Đại diện: {dots(30)}          Chức vụ: {dots(20)}", after=6)


def sign_block(d, left, right, left_note="(Ký, ghi rõ họ tên và đóng dấu)", right_note=None,
               left_name="", right_name="", gap=4):
    right_note = right_note or left_note
    if d.paragraphs:
        d.paragraphs[-1].paragraph_format.keep_with_next = True
    t = d.add_table(rows=1, cols=2)
    t.autofit = False
    shade_borders(t, borders=False)
    for i, (title_, note, name) in enumerate(((left, left_note, left_name), (right, right_note, right_name))):
        c = t.rows[0].cells[i]
        c.width = Cm(8.1)
        p = c.paragraphs[0]
        p.alignment = C
        run(p, title_, bold=True)
        if i == 1 and ANH_DAU and name == DD["ho_ten"]:
            q = c.add_paragraph()
            q.alignment = C
            q.paragraph_format.space_after = Pt(0)
            q.add_run().add_picture(str(ANH_DAU), width=Cm(stamp_size()[0]))
            continue
        q = c.add_paragraph()
        q.alignment = C
        run(q, note, italic=True, size=12)
        for _ in range(gap):
            c.add_paragraph()
        if name:
            n = c.add_paragraph()
            n.alignment = C
            run(n, name, bold=True)
    fix_layout(t, [8.1, 8.1])
    for c in t.rows[0].cells:
        for q in c.paragraphs[:-1]:
            q.paragraph_format.keep_with_next = True
    tail = d.add_paragraph()
    tail.paragraph_format.space_after = Pt(0)
    tail.paragraph_format.line_spacing = Pt(6)


def giamdoc_block(d):
    """Khối ký của Trung Kiên: Giám đốc, chừa trống chỗ ký/dấu."""
    return ("GIÁM ĐỐC", DD["ho_ten"])


def save(d, name):
    out = pathlib.Path(OUT_DIR) if OUT_DIR else OUT
    out.mkdir(parents=True, exist_ok=True)
    p = out / (name.replace(".docx", "_CoDau.docx") if ANH_DAU else name)
    d.save(p)
    return p


CAN_CU = [
    "Căn cứ Bộ luật Dân sự số 91/2015/QH13 ngày 24/11/2015 và các văn bản hướng dẫn thi hành;",
    "Căn cứ Luật Thương mại số 36/2005/QH11 ngày 14/06/2005;",
]


# ---------------------------------------------------------------- 1. Thanh lý
def thanh_ly():
    d = new_doc()
    quoc_hieu(d)
    title(d, "BIÊN BẢN THANH LÝ HỢP ĐỒNG", "(Số: ....../20....../BBTL)")
    for c in CAN_CU:
        para(d, c, italic=True, after=1)
    para(d, f"Căn cứ Hợp đồng số {dots(25)} ngày ..... tháng ..... năm 20.... giữa hai bên;", italic=True, after=1)
    para(d, "Căn cứ kết quả thực hiện hợp đồng của hai bên.", italic=True, after=6)
    para(d, f"Hôm nay, ngày ..... tháng ..... năm 20...., tại {dots(40)}, chúng tôi gồm:")
    ben_kh(d, "BÊN MUA", "gọi tắt là bên A")
    ben_tk(d, "BÊN BÁN", "gọi tắt là bên B")
    para(d, "Hai bên cùng thống nhất thanh lý Hợp đồng nêu trên với các nội dung sau:")
    para(d, "Điều 1: Xác nhận việc thực hiện hợp đồng", bold=True, after=2, keep=True)
    para(d, f"1.1. Bên B đã giao đủ hàng hoá theo Hợp đồng, có Biên bản bàn giao số {dots(15)} ngày ...../...../20.....")
    para(d, f"1.2. Bên B đã xuất Hóa đơn GTGT số {dots(12)}, ký hiệu {dots(10)}, ngày ...../...../20..... và bàn giao các chứng từ kèm theo (CO, CQ) theo Hợp đồng.")
    para(d, "Điều 2: Giá trị hợp đồng và thanh toán", bold=True, after=2, keep=True)
    table(d, ["Nội dung", "Số tiền (VNĐ)"],
          [["Tổng giá trị hợp đồng (đã gồm VAT .....%)", ""],
           ["Số tiền bên A đã thanh toán", ""],
           ["Số tiền còn lại bên A phải thanh toán", ""]],
          [10.5, 4.7], center_cols=(1,))
    para(d, f"Bằng chữ (số tiền còn lại): {dots(70)}")
    para(d, "Điều 3: Bảo hành", bold=True, after=2, keep=True)
    para(d, f"Thời hạn bảo hành ..... tháng kể từ ngày hai bên ký Biên bản bàn giao hàng hóa. Việc thanh lý hợp đồng không làm chấm dứt trách nhiệm bảo hành của bên B theo Hợp đồng.")
    para(d, "Điều 4: Cam kết", bold=True, after=2, keep=True)
    para(d, "4.1. Hai bên xác nhận đã hoàn thành nghĩa vụ của mình theo Hợp đồng, không có tranh chấp, khiếu nại gì với nhau, trừ nghĩa vụ bảo hành nêu tại Điều 3 và số tiền còn lại (nếu có) nêu tại Điều 2.")
    para(d, "4.2. Hợp đồng chấm dứt hiệu lực kể từ ngày hai bên ký Biên bản thanh lý này.")
    para(d, "Biên bản được lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.", after=8)
    sign_block(d, "ĐẠI DIỆN BÊN MUA", "ĐẠI DIỆN BÊN BÁN", right_name=DD["ho_ten"])
    return save(d, "BienBanThanhLyHopDong_TrungKien.docx")


# ---------------------------------------------------------------- 2. Nghiệm thu
def nghiem_thu():
    d = new_doc()
    quoc_hieu(d)
    title(d, "BIÊN BẢN NGHIỆM THU VÀ CHẠY THỬ THIẾT BỊ", "(Số: ....../20....../BBNT)")
    para(d, f"Căn cứ Hợp đồng số {dots(25)} ngày ..... tháng ..... năm 20.... giữa hai bên;", italic=True, after=1)
    para(d, f"Căn cứ Biên bản bàn giao hàng hoá số {dots(15)} ngày ...../...../20....", italic=True, after=6)
    para(d, f"Hôm nay, ngày ..... tháng ..... năm 20...., tại {dots(50)}, chúng tôi gồm:")
    ben_kh(d, "BÊN MUA (bên nhận thiết bị)", "gọi tắt là bên A", extra_tk=False)
    mixed(d, [("BÊN BÁN: ", True), (TK["ten"], True), (" (gọi tắt là bên B)", False)])
    para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    para(d, f"Mã số thuế: {TK['mst']}   Điện thoại: {TK['dien_thoai']}", after=0)
    para(d, f"Đại diện bên B: {dots(30)}          Chức vụ: {dots(20)}", after=6)
    para(d, "I. Thiết bị nghiệm thu", bold=True, after=2, keep=True)
    table(d, ["STT", "Tên thiết bị / Model", "ĐVT", "SL", "Số máy (seri)", "Vị trí lắp đặt"],
          [[str(i), "", "", "", "", ""] for i in range(1, 6)],
          [1.3, 5.2, 1.4, 1.2, 3.3, 3.0], size=11, center_cols=(0, 2, 3))
    para(d, "II. Kết quả kiểm tra và chạy thử", bold=True, after=2, keep=True)
    para(d, f"Thời gian chạy thử: từ ..... giờ ..... đến ..... giờ ....., ngày ...../...../20.....", after=2)
    table(d, ["Hạng mục kiểm tra", "Kết quả đo / quan sát", "Đạt", "Không đạt"],
          [["Ngoại quan, tình trạng thiết bị khi nhận", "", "☐", "☐"],
           ["Đấu nối điện, chiều quay động cơ", "", "☐", "☐"],
           ["Điện áp (V) / Dòng điện (A) khi chạy", "", "☐", "☐"],
           ["Lưu lượng (m³/h) / Cột áp (m) thực tế", "", "☐", "☐"],
           ["Độ rung, tiếng ồn, nhiệt độ vỏ động cơ", "", "☐", "☐"],
           ["Rò rỉ nước, dầu tại phớt và mối nối", "", "☐", "☐"],
           ["Hoạt động của phao, tủ điều khiển (nếu có)", "", "☐", "☐"]],
          [6.0, 5.0, 1.6, 2.4], size=11, center_cols=(2, 3))
    para(d, "III. Kết luận", bold=True, after=2, keep=True)
    para(d, f"☐ Thiết bị đạt yêu cầu, hai bên đồng ý nghiệm thu và đưa vào sử dụng.", after=1)
    para(d, f"☐ Thiết bị chưa đạt, bên B có trách nhiệm khắc phục: {dots(50)}", after=1)
    para(d, f"Ý kiến khác: {dots(70)}", after=2)
    para(d, "Thời hạn bảo hành tính từ ngày hai bên ký biên bản nghiệm thu / bàn giao theo điều khoản bảo hành của Hợp đồng.")
    para(d, "Biên bản được lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.", after=8)
    sign_block(d, "ĐẠI DIỆN BÊN MUA", "ĐẠI DIỆN BÊN BÁN", left_note="(Ký, ghi rõ họ tên)", gap=3)
    return save(d, "BienBanNghiemThu_ChayThu_TrungKien.docx")


# ---------------------------------------------------------------- 3. Giấy ủy quyền
def uy_quyen():
    d = new_doc()
    quoc_hieu(d)
    title(d, "GIẤY ỦY QUYỀN", "(Số: ....../20....../GUQ-TK)")
    para(d, "Căn cứ Bộ luật Dân sự số 91/2015/QH13 ngày 24/11/2015;", italic=True, after=1)
    para(d, "Căn cứ Luật Doanh nghiệp và Điều lệ công ty;", italic=True, after=6)
    para(d, f"Hôm nay, ngày ..... tháng ..... năm 20...., tại {TK['dia_chi']}, chúng tôi gồm:")
    mixed(d, [("BÊN ỦY QUYỀN: ", True), (TK["ten"], True)])
    para(d, f"Mã số thuế: {TK['mst']}", after=0)
    para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    para(d, f"Người đại diện theo pháp luật: Ông {DD['ho_ten']}          Chức vụ: {DD['chuc_vu']}", after=0)
    para(d, f"Ngày sinh: ...../...../.......          Số định danh cá nhân/CCCD: {dots(20)}", after=0)
    para(d, f"Ngày cấp: ...../...../.......          Nơi cấp: {dots(30)}", after=6)
    mixed(d, [("BÊN ĐƯỢC ỦY QUYỀN:", True)])
    para(d, f"Họ và tên: {dots(40)}          Giới tính: {dots(8)}", after=0)
    para(d, f"Ngày sinh: ...../...../.......          Chức vụ tại công ty: {dots(25)}", after=0)
    para(d, f"Số định danh cá nhân/CCCD: {dots(20)}  Ngày cấp: ...../...../.......  Nơi cấp: {dots(20)}", after=0)
    para(d, f"Địa chỉ liên lạc: {dots(60)}", after=6)
    para(d, "Điều 1: Phạm vi ủy quyền", bold=True, after=2, keep=True)
    para(d, "Bên được ủy quyền thay mặt và nhân danh bên ủy quyền thực hiện các công việc sau (đánh dấu X vào mục được ủy quyền):", after=2)
    for t in ["Ký phiếu xuất kho, biên bản giao nhận, biên bản bàn giao và phiếu bảo hành hàng hoá;",
              "Nhận hàng, kiểm đếm, ký xác nhận chứng từ giao nhận với khách hàng, nhà cung cấp, đơn vị vận chuyển;",
              "Liên hệ, nộp và nhận hồ sơ, chứng từ với cơ quan nhà nước, hãng tàu, ngân hàng, kho bãi;",
              f"Công việc khác: {dots(55)}"]:
        para(d, "☐  " + t, after=1, indent=None)
    para(d, "Bên được ủy quyền không được ký hợp đồng kinh tế, phụ lục hợp đồng, đề nghị thanh toán, công văn và không được ủy quyền lại cho người khác, trừ khi có văn bản ủy quyền riêng của Giám đốc.", after=3)
    para(d, "Điều 2: Thời hạn ủy quyền", bold=True, after=2, keep=True)
    para(d, "Từ ngày ...../...../20..... đến hết ngày ...../...../20..... (hoặc đến khi bên ủy quyền có văn bản thu hồi).")
    para(d, "Điều 3: Cam kết", bold=True, after=2, keep=True)
    para(d, "Bên được ủy quyền thực hiện đúng phạm vi ủy quyền và chịu trách nhiệm trước bên ủy quyền về công việc được giao. Bên ủy quyền chịu trách nhiệm trước pháp luật về các cam kết do bên được ủy quyền thực hiện trong phạm vi ủy quyền.")
    para(d, "Giấy ủy quyền được lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.", after=8)
    sign_block(d, "BÊN ĐƯỢC ỦY QUYỀN", "BÊN ỦY QUYỀN", left_note="(Ký, ghi rõ họ tên)",
               right_note="(Ký, ghi rõ họ tên và đóng dấu)", right_name=DD["ho_ten"])
    return save(d, "GiayUyQuyen_TrungKien.docx")


# ---------------------------------------------------------------- 4. Giấy giới thiệu
def gioi_thieu():
    d = new_doc()
    t = d.add_table(rows=1, cols=2)
    t.autofit = False
    shade_borders(t, borders=False)
    a, b = t.rows[0].cells
    fix_layout(t, [7.6, 8.6])
    p = a.paragraphs[0]
    p.alignment = C
    run(p, "CÔNG TY TNHH THƯƠNG MẠI VÀ CÔNG NGHIỆP", bold=True, size=12)
    q = a.add_paragraph(); q.alignment = C
    run(q, "TRUNG KIÊN", bold=True, size=12)
    q = a.add_paragraph(); q.alignment = C
    run(q, "-------***-------", size=12)
    q = a.add_paragraph(); q.alignment = C
    run(q, "Số: ......../20....../GGT-TK", size=12)
    p = b.paragraphs[0]; p.alignment = C
    run(p, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True, size=12)
    q = b.add_paragraph(); q.alignment = C
    run(q, "Độc lập - Tự do - Hạnh phúc", bold=True, size=12)
    q = b.add_paragraph(); q.alignment = C
    run(q, "-------***-------", size=12)
    q = b.add_paragraph(); q.alignment = C
    run(q, "Hà Nội, ngày ..... tháng ..... năm 20....", italic=True, size=12)
    para(d, "", after=6)
    title(d, "GIẤY GIỚI THIỆU")
    para(d, f"Kính gửi: {dots(60)}", bold=True, after=6)
    para(d, f"CÔNG TY TNHH THƯƠNG MẠI VÀ CÔNG NGHIỆP TRUNG KIÊN, mã số thuế {TK['mst']}, địa chỉ {TK['dia_chi']}, trân trọng giới thiệu:")
    para(d, f"Ông/Bà: {dots(40)}          Chức vụ: {dots(20)}", after=1)
    para(d, f"Số CCCD: {dots(20)}  Ngày cấp: ...../...../.......  Nơi cấp: {dots(20)}", after=1)
    para(d, f"Số điện thoại: {dots(25)}", after=4)
    para(d, f"Được cử đến Quý đơn vị để: {dots(60)}")
    para(d, dots(100), after=4)
    para(d, "Đề nghị Quý đơn vị tạo điều kiện để Ông/Bà có tên trên hoàn thành nhiệm vụ.")
    para(d, "Giấy giới thiệu có giá trị đến hết ngày ...../...../20..... và không có giá trị nếu tẩy xoá, sửa chữa.", after=10)
    t2 = d.add_table(rows=1, cols=2)
    t2.autofit = False
    shade_borders(t2, borders=False)
    l, r = t2.rows[0].cells
    fix_layout(t2, [8.1, 8.1])
    run(l.paragraphs[0], "Nơi nhận:", bold=True, size=12)
    for s in ["- Như trên;", "- Lưu: VT."]:
        run(l.add_paragraph(), s, size=12)
    r.paragraphs[0].alignment = C
    run(r.paragraphs[0], TK["ten"], bold=True, size=12)
    if ANH_DAU:
        q = r.add_paragraph(); q.alignment = C
        q.add_run().add_picture(str(ANH_DAU), width=Cm(stamp_size()[0]))
    else:
        q = r.add_paragraph(); q.alignment = C
        run(q, "GIÁM ĐỐC", bold=True)
        q = r.add_paragraph(); q.alignment = C
        run(q, "(Ký, ghi rõ họ tên và đóng dấu)", italic=True, size=12)
        for _ in range(4):
            r.add_paragraph()
        q = r.add_paragraph(); q.alignment = C
        run(q, DD["ho_ten"], bold=True)
    return save(d, "GiayGioiThieu_TrungKien.docx")


# ---------------------------------------------------------------- 5. Phụ lục hợp đồng
def phu_luc():
    d = new_doc()
    quoc_hieu(d)
    title(d, "PHỤ LỤC HỢP ĐỒNG", "(Số: ...../PL-HĐ ..... của Hợp đồng số " + dots(20) + ")")
    for c in CAN_CU:
        para(d, c, italic=True, after=1)
    para(d, f"Căn cứ Hợp đồng số {dots(25)} ngày ..... tháng ..... năm 20.... giữa hai bên (sau đây gọi là Hợp đồng);", italic=True, after=1)
    para(d, "Căn cứ nhu cầu và sự thoả thuận của hai bên.", italic=True, after=6)
    para(d, f"Hôm nay, ngày ..... tháng ..... năm 20...., chúng tôi gồm:")
    ben_kh(d, "BÊN MUA", "gọi tắt là bên A")
    ben_tk(d, "BÊN BÁN", "gọi tắt là bên B")
    para(d, "Hai bên thống nhất sửa đổi, bổ sung Hợp đồng như sau:")
    para(d, "Điều 1: Nội dung sửa đổi, bổ sung", bold=True, after=2, keep=True)
    table(d, ["STT", "Điều khoản của Hợp đồng", "Nội dung cũ", "Nội dung mới"],
          [[str(i), "", "", ""] for i in range(1, 5)],
          [1.3, 3.3, 5.3, 5.3], size=11, center_cols=(0,))
    para(d, "Điều 2: Hiệu lực", bold=True, after=2, keep=True)
    para(d, "2.1. Phụ lục này có hiệu lực kể từ ngày hai bên cùng ký và là phần không thể tách rời của Hợp đồng.")
    para(d, "2.2. Các điều khoản khác của Hợp đồng không được nêu trong Phụ lục này vẫn giữ nguyên giá trị.")
    para(d, "Phụ lục được lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.", after=8)
    sign_block(d, "ĐẠI DIỆN BÊN MUA", "ĐẠI DIỆN BÊN BÁN", right_name=DD["ho_ten"])
    return save(d, "PhuLucHopDong_TrungKien.docx")


# ---------------------------------------------------------------- 6. Đối chiếu công nợ
def cong_no():
    d = new_doc()
    quoc_hieu(d)
    title(d, "BIÊN BẢN ĐỐI CHIẾU CÔNG NỢ", "(Số: ....../20....../BBĐCCN)",
          "Kỳ đối chiếu: từ ngày ...../...../20..... đến ngày ...../...../20.....")
    para(d, f"Hôm nay, ngày ..... tháng ..... năm 20...., hai bên cùng đối chiếu công nợ mua bán hàng hoá, gồm:")
    mixed(d, [("BÊN BÁN: ", True), (TK["ten"], True)])
    para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    para(d, f"Mã số thuế: {TK['mst']}", after=6)
    mixed(d, [("BÊN MUA: ", True), (dots(60), False)])
    para(d, f"Địa chỉ: {dots(70)}", after=0)
    para(d, f"Mã số thuế: {dots(25)}", after=6)
    para(d, "Chi tiết phát sinh trong kỳ (đơn vị tính: VNĐ):", bold=True, after=2, keep=True)
    table(d, ["STT", "Chứng từ (Hợp đồng / Hóa đơn số)", "Ngày", "Giá trị hàng (gồm VAT)", "Đã thanh toán", "Còn nợ"],
          [[str(i), "", "", "", "", ""] for i in range(1, 8)] + [["", "Tổng cộng", "", "", "", ""]],
          [1.3, 4.5, 2.2, 3.1, 2.7, 2.4], size=11, center_cols=(0, 2))
    table(d, ["Tổng hợp", "Số tiền (VNĐ)"],
          [["Số dư nợ đầu kỳ", ""], ["Phát sinh tăng trong kỳ (giá trị hàng đã giao)", ""],
           ["Đã thanh toán trong kỳ", ""], ["Số dư nợ cuối kỳ (bên mua còn phải trả bên bán)", ""]],
          [10.5, 4.7], center_cols=(1,))
    para(d, f"Bằng chữ (số dư cuối kỳ): {dots(70)}")
    para(d, "Hai bên xác nhận số liệu trên là đúng ☐   /   có chênh lệch ☐ (ghi rõ nội dung chênh lệch): " + dots(40))
    para(d, dots(100), after=2)
    para(d, f"Bên mua cam kết thanh toán số dư còn nợ trước ngày ...../...../20..... bằng chuyển khoản vào tài khoản {BANK['so_tk']} - {BANK['ten']}, chủ tài khoản {BANK['chu_tk']}.")
    para(d, "Biên bản được lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.", after=8)
    sign_block(d, "ĐẠI DIỆN BÊN MUA", "ĐẠI DIỆN BÊN BÁN", right_name=DD["ho_ten"])
    return save(d, "BienBanDoiChieuCongNo_TrungKien.docx")


# ---------------------------------------------------------------- 7. Đơn đặt hàng nhà cung cấp
def don_dat_hang():
    d = new_doc()
    header_cty(d)
    title(d, "ĐƠN ĐẶT HÀNG", "(Số: ....../20....../ĐH-TK)")
    date_line(d, align=C)
    para(d, f"Kính gửi: {dots(60)}", bold=True, after=0)
    para(d, f"Địa chỉ: {dots(70)}", after=0)
    para(d, f"Mã số thuế: {dots(25)}   Điện thoại: {dots(25)}", after=0)
    para(d, f"Người liên hệ: {dots(40)}", after=6)
    para(d, f"{TK['ten']} đề nghị Quý Công ty cung cấp các hàng hoá sau:")
    table(d, ["STT", "Tên hàng / Model / Thông số", "ĐVT", "SL", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)"],
          [[str(i), "", "", "", "", ""] for i in range(1, 7)] +
          [["", "Cộng", "", "", "", ""], ["", "VAT (.....%)", "", "", "", ""], ["", "Tổng cộng", "", "", "", ""]],
          [1.3, 6.0, 1.3, 1.2, 2.9, 3.0], size=11, center_cols=(0, 2, 3))
    para(d, f"Bằng chữ: {dots(80)}")
    para(d, "Điều kiện đặt hàng", bold=True, after=2, keep=True)
    para(d, f"1. Thời gian giao hàng: {dots(55)}", after=1)
    para(d, f"2. Địa điểm giao hàng: {dots(55)}", after=1)
    para(d, f"3. Thanh toán: {dots(60)}", after=1)
    para(d, "4. Hàng mới 100%, đúng chủng loại, model; kèm hóa đơn GTGT và chứng từ CO, CQ (nếu có).", after=1)
    para(d, f"5. Bảo hành: ..... tháng kể từ ngày giao hàng.", after=1)
    para(d, f"6. Ghi chú khác: {dots(55)}", after=6)
    para(d, "Kính đề nghị Quý Công ty xác nhận đơn hàng và thời gian giao hàng bằng văn bản hoặc email. Trân trọng cảm ơn!", after=8)
    sign_block(d, "ĐẠI DIỆN NHÀ CUNG CẤP", "ĐẠI DIỆN BÊN ĐẶT HÀNG", left_note="(Xác nhận, ký, đóng dấu)",
               right_note="(Ký, ghi rõ họ tên và đóng dấu)", right_name=DD["ho_ten"])
    return save(d, "DonDatHang_NhaCungCap_TrungKien.docx")


# ---------------------------------------------------------------- 8. Nhận hàng nhập khẩu
def nhan_hang_nk():
    d = new_doc()
    t = d.add_table(rows=1, cols=2)
    t.autofit = False
    shade_borders(t, borders=False)
    a, b = t.rows[0].cells
    fix_layout(t, [7.6, 8.6])
    run(a.paragraphs[0], "CÔNG TY TNHH THƯƠNG MẠI VÀ CÔNG NGHIỆP TRUNG KIÊN", bold=True, size=12)
    run(a.add_paragraph(), f"ĐT: {TK['dien_thoai']}", size=12)
    b.paragraphs[0].alignment = C
    run(b.paragraphs[0], "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True, size=12)
    q = b.add_paragraph(); q.alignment = C
    run(q, "Độc lập - Tự do - Hạnh phúc", bold=True, size=12)
    q = b.add_paragraph(); q.alignment = C
    run(q, "Hà Nội, ngày ..... tháng ..... năm 20....", italic=True, size=12)
    para(d, "", after=4)
    title(d, "BIÊN BẢN NHẬN HÀNG NHẬP KHẨU", "(Số: ....../20....../BBNH-TK)")
    para(d, f"Đơn vị nhận hàng: {TK['ten']}", bold=True, after=0)
    para(d, f"Địa chỉ: {TK['dia_chi']}   Mã số thuế: {TK['mst']}", after=4)
    para(d, "I. Thông tin lô hàng", bold=True, after=2, keep=True)
    for s in ["Số vận đơn (B/L): " + dots(30), "Số container / Seal: " + dots(20) + " / " + dots(15),
              "Tên tàu / Số chuyến: " + dots(20) + " / " + dots(12), "Nhà cung cấp (người bán): " + dots(45),
              "Hợp đồng / Invoice / Packing list số: " + dots(30), "Hãng tàu / đơn vị vận chuyển: " + dots(35),
              "Thời gian, địa điểm nhận hàng: ..... giờ ..... ngày ...../...../20....., tại " + dots(35)]:
        para(d, s, after=1)
    para(d, "II. Kiểm tra bên ngoài", bold=True, after=2, keep=True)
    table(d, ["Nội dung", "Kết quả"],
          [["Seal hãng tàu còn nguyên vẹn, số seal khớp vận đơn", "☐ Đúng   ☐ Không đúng"],
           ["Vỏ container, bao bì ngoài không móp, rách, ẩm ướt", "☐ Đúng   ☐ Không đúng"],
           ["Tình trạng khác (nếu có)", ""]],
          [8.2, 7.0], size=11)
    para(d, "III. Kiểm đếm hàng hoá", bold=True, after=2, keep=True)
    table(d, ["STT", "Tên hàng / Model", "ĐVT", "SL theo chứng từ", "SL thực nhận", "Tình trạng / Chênh lệch"],
          [[str(i), "", "", "", "", ""] for i in range(1, 7)],
          [1.3, 5.2, 1.3, 2.6, 2.3, 3.5], size=11, center_cols=(0, 2, 3, 4))
    para(d, "IV. Kết luận", bold=True, after=2, keep=True)
    para(d, "☐ Hàng nhận đủ số lượng, nguyên đai nguyên kiện, đúng chứng từ.", after=1)
    para(d, "☐ Hàng thiếu / hư hỏng / sai khác, đã lập ảnh và biên bản xác nhận với hãng tàu: " + dots(35), after=1)
    para(d, f"Hàng nhập kho: ☐ Hà Nội   ☐ {TK['dia_diem_kinh_doanh_hcm']['ten']} TP Hồ Chí Minh", after=2)
    para(d, "Biên bản được lập thành 02 bản, công ty giữ 01 bản lưu cùng bộ chứng từ nhập khẩu (Hóa đơn, Packing list, Vận đơn, CO, CQ).", after=8)
    sign_block(d, "NGƯỜI NHẬN HÀNG / THỦ KHO", "GIÁM ĐỐC", left_note="(Ký, ghi rõ họ tên)",
               right_note="(Ký, ghi rõ họ tên và đóng dấu)", right_name=DD["ho_ten"], gap=3)
    return save(d, "BienBanNhanHangNhapKhau_TrungKien.docx")


if __name__ == "__main__":
    # python3 cong-cu/tao_mau.py [anh_dau.png thu_muc_ra]   (có ảnh: tạo bản có dấu, chữ ký ở thu_muc_ra)
    if len(sys.argv) > 2:
        ANH_DAU, OUT_DIR = sys.argv[1], sys.argv[2]
    for f in (thanh_ly, nghiem_thu, uy_quyen, gioi_thieu, phu_luc, cong_no, don_dat_hang, nhan_hang_nk):
        print(f().name)
