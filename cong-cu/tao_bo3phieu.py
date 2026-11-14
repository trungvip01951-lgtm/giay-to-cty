#!/usr/bin/env python3
"""Bộ 3 phiếu bản Word để in: Phiếu xuất kho, Phiếu bảo hành, Biên bản bàn giao (mỗi phiếu 1 trang).
Bố cục theo bộ phiếu phần mềm báo giá (giay-to-khach-hang/Xcelwood/goc). Tên in sẵn: "Nguyễn Trung" dưới
đại diện bên giao/bên bán ở phiếu bảo hành và biên bản bàn giao (QUY_TAC 10/11/2026); phiếu xuất kho để trống tên.
Chạy: python3 cong-cu/tao_bo3phieu.py giay-to-khach-hang/<khach>/bo3phieu_x.json"""
import json, pathlib, sys
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402
from tao_denghi_fata import so, doc_so  # noqa: E402
from docx.enum.text import WD_BREAK  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

TK = m.TK
TEN_GIAO = "Nguyễn Trung"


def ky(d, cot, ten=None, gap=4):
    """cot: list tiêu đề; ten: {chỉ số cột: tên in sẵn}"""
    ten = ten or {}
    n = len(cot)
    w = 16.0 / n
    t = d.add_table(rows=1, cols=n)
    t.autofit = False
    m.shade_borders(t, borders=False)
    for i, tt in enumerate(cot):
        c = t.rows[0].cells[i]
        p = c.paragraphs[0]; p.alignment = m.C
        m.run(p, tt, bold=True)
        if not tt:
            continue
        q = c.add_paragraph(); q.alignment = m.C
        m.run(q, "(Ký, ghi rõ họ tên)", italic=True)
        for _ in range(gap):
            c.add_paragraph()
        if i in ten:
            r = c.add_paragraph(); r.alignment = m.C
            m.run(r, ten[i], bold=True)
    m.fix_layout(t, [w] * n)


def ngay_chu(nd):
    return f"Hà Nội, ngày {nd[0]} tháng {nd[1]} năm {nd[2]}"


def dau_trang_qh(d, nd):
    t = d.add_table(rows=1, cols=2); t.autofit = False
    m.shade_borders(t, borders=False)
    a, b = t.rows[0].cells
    p = a.paragraphs[0]; p.alignment = m.C; m.run(p, TK["ten"], bold=True, size=11)
    q = a.add_paragraph(); q.alignment = m.C; m.run(q, f"ĐT: {TK['dien_thoai']}", size=11)
    p = b.paragraphs[0]; p.alignment = m.C; m.run(p, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True, size=11)
    q = b.add_paragraph(); q.alignment = m.C; m.run(q, "Độc lập - Tự do - Hạnh phúc", bold=True, size=12)
    q = b.add_paragraph(); q.alignment = m.C; m.run(q, "---------------***---------------", size=11)
    q = b.add_paragraph(); q.alignment = m.C; m.run(q, ngay_chu(nd), italic=True)
    m.fix_layout(t, [7.0, 9.0])


def ngat_trang(d):
    p = d.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    r = p.add_run(); r.font.size = Pt(1)
    r.add_break(WD_BREAK.PAGE)


def build(don):
    k, nd, hang = don["khach"], don["ngay"], don["hang"]
    dt = k.get("dien_thoai") or m.dots(30)
    k = dict(k, dia_chi=k.get("dia_chi") or m.dots(70))
    cong = sum(h["sl"] * h["don_gia"] for h in hang)
    vat = cong * don["vat_phan_tram"] // 100
    tong = cong + vat
    kt = don.get("kiem_tra_hoa_don")
    if kt:
        assert (cong, vat, tong) == (kt["cong_tien_hang"], kt["tien_thue"], kt["tong"]), (cong, vat, tong)
    d = m.new_doc()

    # 1. PHIẾU XUẤT KHO
    m.para(d, TK["ten"], bold=True, align=m.L, after=0)
    m.para(d, f"Địa chỉ: {TK['dia_chi']}", align=m.L, after=0)
    m.para(d, f"MST: {TK['mst']}   |   ĐT: {TK['dien_thoai']}", align=m.L, after=10)
    m.para(d, "PHIẾU XUẤT KHO", bold=True, align=m.C, size=16, after=2)
    m.para(d, ngay_chu(nd), italic=True, align=m.C, after=8)
    m.mixed(d, [("Khách hàng: ", False), (k["ten"], True)], after=0)
    if k.get("mst"):
        m.para(d, f"Mã số thuế: {k['mst']}", after=0)
    m.para(d, f"Điện thoại: {dt}", after=0)
    m.para(d, f"Địa chỉ: {k['dia_chi']}", after=0)
    m.para(d, f"Nội dung: {don['can_cu']}", after=6)
    rows = [[str(i), "Kho Hà Nội", h["ten"], h["dvt"], str(h["sl"]), so(h["don_gia"]), so(h["sl"] * h["don_gia"])]
            for i, h in enumerate(hang, 1)]
    t = m.table(d, ["STT", "Mã kho", "Tên vật tư / Model", "ĐVT", "SL", "Đơn giá", "Thành tiền"], rows,
                [1.3, 2.0, 5.1, 1.3, 1.0, 2.6, 2.7], size=12, center_cols=(0, 1, 3, 4, 5, 6))
    for nhan, v in (("Cộng:", cong), (f"VAT {don['vat_phan_tram']}%:", vat), ("Tổng tiền thanh toán:", tong)):
        c = t.add_row().cells
        a = c[0].merge(c[5]); a.paragraphs[0].alignment = m.R; m.run(a.paragraphs[0], nhan, bold=True)
        c[6].paragraphs[0].alignment = m.C; m.run(c[6].paragraphs[0], so(v), bold=True)
    m.para(d, f"Bằng chữ: {doc_so(tong)}./.", italic=True, after=10)
    ky(d, ["Đại diện bên bán", "Người nhận hàng", "Thủ kho"])
    ngat_trang(d)

    # 2. PHIẾU BẢO HÀNH
    dau_trang_qh(d, nd)
    m.para(d, "PHIẾU BẢO HÀNH", bold=True, align=m.C, size=16, after=6)
    m.para(d, "Các bên có liên quan:", after=2)
    m.para(d, "I. BÊN MUA HÀNG:", bold=True, after=0)
    m.mixed(d, [("Khách hàng: ", False), (k["ten"], True)], after=0)
    m.para(d, f"Địa chỉ: {k['dia_chi']}", after=0)
    m.para(d, f"Điện thoại: {dt}", after=2)
    m.para(d, "II. BÊN BÁN HÀNG:", bold=True, after=0)
    m.para(d, TK["ten"], bold=True, after=0)
    m.para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    m.para(d, f"Điện thoại: {TK['dien_thoai']}", after=4)
    m.para(d, "Nội dung bảo hành:", after=2)
    m.table(d, ["STT", "Tên sản phẩm / Model", "Đơn vị", "Số lượng"],
            [[str(i), h["ten"], h["dvt"], str(h["sl"])] for i, h in enumerate(hang, 1)],
            [1.2, 10.0, 2.4, 2.4], size=11, center_cols=(0, 2, 3))
    for x in ["I./ Bên bán hàng cam kết bảo hành cho những hàng hoá đã giao theo Biên bản bàn giao hàng hoá.",
              f"II./ Thời gian bảo hành: {don['bao_hanh']} kể từ ngày ký biên bản bàn giao đưa vào sử dụng.",
              "III./ Điều kiện bảo hành: Bảo hành cho những lỗi của nhà sản xuất.",
              f"IV./ Địa điểm bảo hành: {TK['dia_chi']}",
              "V./ Những trường hợp không được bảo hành:"]:
        m.para(d, x, after=1, size=11)
    for x in ["1./ Khách hàng không có phiếu bảo hành hoặc phiếu bảo hành bị sửa chữa, rách nát, thiếu, sai lệch thông số.",
              "2./ Hư hỏng do những nguyên nhân khách quan, các trường hợp bất khả kháng.",
              "3./ Những hư hỏng do dùng sai điện áp quy định hoặc do nguồn điện không ổn định.",
              "4./ Những hư hỏng khác không do lỗi của nhà sản xuất."]:
        p = m.para(d, x, after=1, size=11); p.paragraph_format.left_indent = Cm(0.6)
    m.para(d, "Trân trọng cảm ơn Quý khách hàng đã sử dụng sản phẩm và dịch vụ của công ty chúng tôi.", italic=True, after=4)
    ky(d, ["", "ĐẠI DIỆN BÊN BÁN"], {1: TEN_GIAO}, gap=3)
    ngat_trang(d)

    # 3. BIÊN BẢN BÀN GIAO
    dau_trang_qh(d, nd)
    m.para(d, "", after=4)
    m.para(d, "BIÊN BẢN BÀN GIAO", bold=True, align=m.C, size=16, after=8)
    m.mixed(d, [("Bên giao: ", True), (TK["ten"], True)], after=0)
    m.para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    m.para(d, f"Người đại diện: {TK['dai_dien']['ho_ten']}          Chức vụ: {TK['dai_dien']['chuc_vu']}", after=4)
    m.mixed(d, [("Bên nhận: ", True), (k["ten"], True)], after=0)
    m.para(d, f"Địa chỉ: {k['dia_chi']}", after=0)
    m.para(d, f"Người đại diện: {m.dots(30)}          Chức vụ: {m.dots(20)}", after=4)
    m.para(d, f"Căn cứ: {don['can_cu'].replace('Xuất hàng theo ', '')}.", after=2)
    m.para(d, "Cùng tiến hành bàn giao hàng hóa sau:", after=2)
    m.table(d, ["STT", "Hàng hóa / Model", "Số lượng"],
            [[str(i), h["ten"], f"{h['sl']} {h['dvt'].lower()}"] for i, h in enumerate(hang, 1)],
            [1.2, 11.8, 3.0], center_cols=(0, 2))
    m.para(d, "Chất lượng: Hàng mới 100%", after=1)
    m.para(d, f"Bảo hành: {don['bao_hanh']}", after=1)
    m.para(d, "Hai bên tiến hành bàn giao với số lượng như trên. Biên bản được lập thành 02 bản, mỗi bên giữ 01 bản.", after=10)
    ky(d, ["Đại diện bên nhận", "Đại diện bên giao"], {1: TEN_GIAO})

    p = pathlib.Path(sys.argv[1]).parent / (don["ten_tep"] + ".docx")
    d.save(p)
    return p, cong, vat, tong


if __name__ == "__main__":
    print(build(json.load(open(sys.argv[1], encoding="utf-8"))))
