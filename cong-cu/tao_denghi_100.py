#!/usr/bin/env python3
"""Giấy đề nghị thanh toán (100% hoặc một tỷ lệ) cho đơn hàng theo báo giá. Số liệu đọc từ tệp JSON của đơn.
Chạy: python3 cong-cu/tao_denghi_100.py don.json [anh_dau.png]   -> ghi cạnh tệp JSON: bản không dấu và (nếu có ảnh) bản _CoDau."""
import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402
from tao_denghi_fata import doc_so, so  # noqa: E402

TK, DD, BANK = m.TK, m.DD, m.BANK


def build(don, anh_dau, out_dir):
    vat = don["thanh_tien_chua_vat"] * don["vat_phan_tram"] // 100
    tong = don["thanh_tien_chua_vat"] + vat
    tt = tong * don["ty_le_thanh_toan"] // 100
    assert tong * don["ty_le_thanh_toan"] % 100 == 0, "Số tiền thanh toán không tròn đồng"
    d = m.new_doc()
    m.ANH_DAU = anh_dau
    m.quoc_hieu(d)
    m.header_cty(d)
    n = don["ngay_lap"]
    m.para(d, f"Hà Nội, ngày {n[0]} tháng {n[1]} năm {n[2]}", italic=True, align=m.R, after=6)
    m.para(d, "GIẤY ĐỀ NGHỊ THANH TOÁN", bold=True, align=m.C, size=15, after=6)
    m.mixed(d, [("Kính gửi: ", True), (don["khach"], True)])
    m.para(d, f"Địa chỉ: {don['dia_chi']}", after=0)
    m.para(d, f"Mã số thuế: {don['mst']}", after=6)
    m.para(d, f"Căn cứ Báo giá số {don['bao_gia_so']} ngày {don['bao_gia_ngay']} của {TK['ten']} về việc cung cấp {don['hang']};")
    m.para(d, f"Căn cứ phương thức thanh toán: {don['thoi_han']},")
    m.para(d, f"chúng tôi kính đề nghị Quý Công ty thanh toán {don['ty_le_thanh_toan']}% giá trị đơn hàng như sau:", after=4)
    m.table(d, ["Nội dung", "Số tiền (đồng)"],
            [["Giá trị hàng hóa (chưa VAT)", so(don["thanh_tien_chua_vat"])],
             [f"Thuế VAT {don['vat_phan_tram']}%", so(vat)],
             ["Tổng giá trị đơn hàng (đã gồm VAT)", so(tong)],
             [f"Số tiền đề nghị thanh toán ({don['ty_le_thanh_toan']}%)", so(tt)]],
            [11.0, 4.5], size=12, center_cols=(1,))
    m.para(d, f"Bằng chữ: {doc_so(tt)}.", bold=True)
    m.para(d, f"Kính đề nghị Quý Công ty chuyển khoản cho {BANK['chu_tk']}, số tài khoản {BANK['so_tk']} tại {BANK['ten']}; "
              f"nội dung: Thanh toán {don['ty_le_thanh_toan']}% đơn hàng theo báo giá {don['bao_gia_so']}. Trân trọng cảm ơn Quý Công ty!", after=4)
    m.sign_block(d, "", "BÊN ĐỀ NGHỊ THANH TOÁN", left_note="", right_name=DD["ho_ten"], gap=3)
    out = pathlib.Path(out_dir)
    p = out / f"DeNghiThanhToan_{don['ty_le_thanh_toan']}pct_{don['bao_gia_so']}{'_CoDau' if anh_dau else ''}.docx"
    d.save(p)
    return p, tt


if __name__ == "__main__":
    f = pathlib.Path(sys.argv[1])
    don = json.load(open(f, encoding="utf-8"))
    for anh in (None, *sys.argv[2:3]):
        p, tt = build(don, anh, f.parent)
        print(p, so(tt))
