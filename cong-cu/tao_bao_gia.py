#!/usr/bin/env python3
"""Bảng chào giá thiết bị Trung Kiên (Word + PDF), bố cục theo báo giá của phần mềm báo giá
(giay-to-khach-hang/NghiaKen/goc/BG-09112026-009_goc.pdf). Không in website/email (QUY_TAC).
Chạy: python3 cong-cu/tao_bao_gia.py bao_gia.json"""
import datetime, json, pathlib, sys
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402
from tao_denghi_fata import so, doc_so  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

TK, BANK = m.TK, m.BANK


def bang_2cot(d, rows, w=(4.4, 11.6)):
    t = d.add_table(rows=0, cols=2); t.autofit = False
    m.shade_borders(t)
    for a, b in rows:
        c = t.add_row().cells
        for x in c:
            x.paragraphs[0].paragraph_format.space_after = Pt(0)
        m.run(c[0].paragraphs[0], a, bold=True, size=11)
        m.run(c[1].paragraphs[0], b, size=11)
    m.fix_layout(t, list(w))
    g = d.add_paragraph(); g.paragraph_format.space_after = Pt(0); g.paragraph_format.line_spacing = Pt(6)


def build(bg, js):
    k, hang = bg["khach"], bg["hang"]
    ngay = datetime.date(*bg["ngay"])
    cong = sum(h["sl"] * h["don_gia"] for h in hang)
    vat = cong * bg["vat_phan_tram"] // 100
    tong = cong + vat
    if bg.get("kiem_tra"):
        assert [cong, vat, tong] == bg["kiem_tra"], (cong, vat, tong)
    d = m.new_doc()
    m.para(d, TK["ten"], bold=True, align=m.L, after=0)
    m.para(d, f"HN: {TK['dia_chi']}", align=m.L, after=0)
    m.para(d, f"HCM: {TK['dia_diem_kinh_doanh_hcm']['dia_chi']}", align=m.L, after=0)
    m.para(d, f"Điện thoại: {TK['dien_thoai']}   |   MST: {TK['mst']}", align=m.L, after=10)
    m.para(d, "BẢNG CHÀO GIÁ THIẾT BỊ", bold=True, align=m.C, size=16, after=0)
    m.para(d, f"Số: {bg['so']}  –  Ngày {ngay.strftime('%d/%m/%Y')}", italic=True, align=m.C, after=6)
    m.para(d, "THÔNG TIN KHÁCH HÀNG", bold=True, after=2)
    dong2 = (("Mã số thuế", k["mst"] + (f" – ĐT: {k['dien_thoai']}" if k.get("dien_thoai") else ""))
             if k.get("mst") else ("Điện thoại", k.get("dien_thoai") or ""))
    bang_2cot(d, [("Khách hàng", k.get("ten") or ""), dong2, ("Địa chỉ", k.get("dia_chi") or "")], w=(2.6, 13.4))
    if not bg.get("gon"):
        m.para(d, "Kính gửi Quý khách hàng báo giá sản phẩm có thông số như sau:", after=2)
    rows = [[str(i), h["ten"], f"{h['sl']} {h['dvt']}", so(h["don_gia"]), so(h["sl"] * h["don_gia"])]
            for i, h in enumerate(hang, 1)]
    t = m.table(d, ["STT", "Thông số thiết bị", "Số lượng", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)"], rows,
                [1.2, 7.5, 2.1, 2.6, 2.6], size=11, center_cols=(0, 2, 3, 4))
    for nhan, v in (("Cộng tiền hàng:", cong), (f"VAT {bg['vat_phan_tram']}%:", vat), ("TỔNG CỘNG:", tong)):
        c = t.add_row().cells
        a = c[0].merge(c[3]); a.paragraphs[0].alignment = m.R; m.run(a.paragraphs[0], nhan, bold=True)
        c[4].paragraphs[0].alignment = m.C; m.run(c[4].paragraphs[0], so(v), bold=True)
    m.para(d, f"Bằng chữ: {doc_so(tong)}./.", bold=True, italic=True, after=6)
    for i, x in enumerate(bg["ghi_chu"]):
        q = m.para(d, after=0)
        if i == 0:
            m.run(q, "Ghi chú: ", bold=True, size=11)
        m.run(q, f"- {x}", size=11)
    m.para(d, "", after=0).paragraph_format.line_spacing = Pt(4)
    m.para(d, "PHƯƠNG THỨC THANH TOÁN VÀ GIAO HÀNG", bold=True, after=2)
    tt = [("Thanh toán", bg["thanh_toan"]), ("Giao hàng", bg["giao_hang"]), ("Địa điểm giao hàng", bg["dia_diem_giao"])]
    if len({v for _, v in tt}) == 1:
        tt = [("Thanh toán, giao hàng", tt[0][1])]
    bang_2cot(d, tt + [("Chuyển khoản", f"{BANK['chu_tk']}\nSố TK: {BANK['so_tk']} – {BANK['ten']}")])
    if not bg.get("gon"):   # gon: bỏ câu chào cuối để chừa chỗ đóng dấu trên cùng trang
      m.para(d, "Rất mong báo giá này đáp ứng được yêu cầu của Quý khách và sớm nhận được sự hợp tác.", italic=True, after=4)
    m.ANH_DAU = bg.get("anh_dau")   # có ảnh dấu: chèn dấu + chữ ký Giám đốc vào khối bên bán
    m.sign_block(d, "XÁC NHẬN ĐẶT HÀNG", "ĐẠI DIỆN BÊN BÁN", left_note="(Ký, họ tên, đóng dấu)", gap=3,
                 right_name=m.DD["ho_ten"] if m.ANH_DAU else "")
    cuoi = d.paragraphs[-1]   # đoạn trống sau khối ký: thu nhỏ để không đẩy sang trang mới
    cuoi.paragraph_format.line_spacing = Pt(1); cuoi.paragraph_format.space_after = Pt(0)
    p = pathlib.Path(js).parent / (bg["ten_tep"] + ("_CoDau" if m.ANH_DAU else "") + ".docx")
    d.save(p)
    return p, cong, vat, tong


if __name__ == "__main__":
    print(build(json.load(open(sys.argv[1], encoding="utf-8")), sys.argv[1]))
