#!/usr/bin/env python3
"""Giấy đề nghị thanh toán phần còn lại - Hợp đồng HĐKT102026-TK-H (Trung Kiên bán, Công ty CP Công nghệ FATA mua).
Số liệu lấy từ Hợp đồng ký 05/10/2026, Phụ lục số 01 ký 12/10/2026 và bảng công nợ của FATA (hóa đơn 00002771).
Chạy: python3 cong-cu/tao_denghi_fata.py"""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402
from docx.shared import Pt  # noqa: E402

TK, DD, BANK = m.TK, m.DD, m.BANK
OUT = m.ROOT / "giay-to-khach-hang" / "FATA"

HD_SO, HD_NGAY = "HĐKT102026-TK-H", "05/10/2026"
PL_SO, PL_NGAY = "01/PLHĐ-HĐKT102026-TK-H", "12/10/2026"
HOADON_SO, HOADON_NGAY = "00002771", "17/10/2026"
NGAY_LAP = (27, 10, 2026)

HD_GOC = 260_712_000          # Điều 2 hợp đồng gốc (đã gồm VAT 8%)
BO_SUNG = 30_456_000          # Phụ lục 01 (đã gồm VAT 8%)
TONG = HD_GOC + BO_SUNG
DA_TT = HD_GOC // 2           # bên A đã trả 50% hợp đồng gốc sau khi ký
CON_LAI = TONG - DA_TT
assert (TONG, DA_TT, CON_LAI) == (291_168_000, 130_356_000, 160_812_000)  # khớp bảng công nợ của FATA


def so(n):
    return f"{n:,}".replace(",", ".")


def doc_so(n):
    """Đọc số tiền (< 1 tỷ) bằng chữ."""
    d = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

    def ba(x, day_du):
        tram, chuc, dv = x // 100, x // 10 % 10, x % 10
        out = []
        if tram or day_du:
            out.append(d[tram] + " trăm")
        if chuc > 1:
            out.append(d[chuc] + " mươi")
        elif chuc == 1:
            out.append("mười")
        elif (tram or day_du) and dv:
            out.append("lẻ")
        if dv:
            if chuc > 1 and dv == 1:
                out.append("mốt")
            elif chuc >= 1 and dv == 5:
                out.append("lăm")
            else:
                out.append(d[dv])
        return " ".join(out)

    parts = []
    tr, ng, dn = n // 10**6, n // 1000 % 1000, n % 1000
    if tr:
        parts.append(ba(tr, False) + " triệu")
    if ng:
        parts.append(ba(ng, bool(tr)) + " nghìn")
    if dn:
        parts.append(ba(dn, bool(tr or ng)))
    s = " ".join(parts) + " đồng"
    return s[0].upper() + s[1:]


assert doc_so(CON_LAI) == "Một trăm sáu mươi triệu tám trăm mười hai nghìn đồng"
assert doc_so(130_356_000) == "Một trăm ba mươi triệu ba trăm năm mươi sáu nghìn đồng"


def build(anh_dau=None, out_dir=None):
    """anh_dau: đường dẫn ảnh PNG dấu+chữ ký (không lưu trong kho). Không có thì chừa trống để ký tay."""
    d = m.new_doc()
    m.quoc_hieu(d)
    m.header_cty(d)
    m.para(d, f"Hà Nội, ngày {NGAY_LAP[0]} tháng {NGAY_LAP[1]} năm {NGAY_LAP[2]}", italic=True, align=m.R, after=6)
    m.para(d, "GIẤY ĐỀ NGHỊ THANH TOÁN PHẦN CÒN LẠI", bold=True, align=m.C, size=15, after=6)
    m.mixed(d, [("Kính gửi: ", True), ("CÔNG TY CỔ PHẦN CÔNG NGHỆ FATA", True)])
    m.para(d, "Địa chỉ: Lô CN01, KCN Thụy Vân, xã Hy Cương, tỉnh Phú Thọ, Việt Nam", after=0)
    m.para(d, "Mã số thuế: 2601102666", after=6)
    m.para(d, f"Căn cứ Hợp đồng kinh tế số {HD_SO} ký ngày {HD_NGAY} và Phụ lục hợp đồng số {PL_SO} ký ngày {PL_NGAY} về việc cung cấp máy bơm nước AQUARIS (Italy): 04 máy CM65-160/15 và 04 máy AQR65-200/15 (sau bổ sung);")
    m.para(d, f"Căn cứ Hóa đơn GTGT số {HOADON_SO} ngày {HOADON_NGAY} và Điều 2 của Hợp đồng (giá trị còn lại thanh toán sau khi bên B giao hàng và hoàn thiện hồ sơ thanh toán, trong vòng 2-3 ngày),")
    m.para(d, "chúng tôi kính đề nghị Quý Công ty thanh toán phần giá trị còn lại của Hợp đồng như sau:", after=4)
    m.table(d, ["Nội dung", "Số tiền (đồng)"],
            [["Giá trị Hợp đồng gốc (đã gồm VAT 8%)", so(HD_GOC)],
             ["Giá trị bổ sung theo Phụ lục (đã gồm VAT 8%)", so(BO_SUNG)],
             ["Tổng giá trị Hợp đồng sau bổ sung", so(TONG)],
             ["Đã thanh toán (50% giá trị Hợp đồng gốc)", so(DA_TT)],
             ["Số tiền đề nghị thanh toán lần này", so(CON_LAI)]],
            [11.0, 4.5], size=12, center_cols=(1,))
    m.para(d, f"Bằng chữ: {doc_so(CON_LAI)}.", bold=True)
    m.para(d, f"Thời hạn thanh toán: theo Điều 2 của Hợp đồng và Phụ lục hợp đồng. Kính đề nghị Quý Công ty chuyển khoản cho {BANK['chu_tk']}, số tài khoản {BANK['so_tk']} tại {BANK['ten']}; nội dung: Thanh toán phần còn lại hợp đồng {HD_SO}. Trân trọng cảm ơn Quý Công ty!", after=4)
    m.ANH_DAU = anh_dau
    m.sign_block(d, "", "BÊN ĐỀ NGHỊ THANH TOÁN", left_note="", right_name=DD["ho_ten"], gap=3)
    o = pathlib.Path(out_dir) if out_dir else OUT
    o.mkdir(parents=True, exist_ok=True)
    p = o / (f"DeNghiThanhToan_ConLai_{HD_SO}" + ("_CoDau" if anh_dau else "") + ".docx")
    d.save(p)
    return p


if __name__ == "__main__":
    # python3 cong-cu/tao_denghi_fata.py [anh_dau.png thu_muc_ra]
    print(build(*(sys.argv[1:3])) if len(sys.argv) > 2 else build())
