#!/usr/bin/env python3
"""Hợp đồng nguyên tắc mua bán (Bên A = Trung Kiên bán, Bên B = khách), điều khoản lấy thẳng từ mục 3.4 của hồ sơ.
Chạy: python3 cong-cu/tao_hd_nguyen_tac.py hd.json [anh_dau.png]   -> ghi cạnh tệp JSON (bản in không tên; có ảnh thì thêm _CoDau)."""
import json
import pathlib
import re
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402

TK, DD, BANK = m.TK, m.DD, m.BANK
HOSO = m.ROOT / "ho-so" / "HoSo_ThongTinCongTy_VaMauGiayTo_TrungKien.md"


def dong_mau():
    t = HOSO.read_text(encoding="utf-8")
    k = t[t.index("### 3.4"):]
    k = k[k.index("```\n") + 4:]
    k = k[:k.index("```")]
    return [x.strip() for x in k.split("\n")]


def build(hd, anh, out):
    kh = hd["khach"]
    d0, m0, y0 = hd["ngay_hd"]
    dots = lambda v, n=25: v or m.dots(n)  # noqa: E731
    dong = dong_mau()
    can_cu = [x for x in dong if x.startswith("- Căn cứ")]
    i0 = next(i for i, x in enumerate(dong) if x.startswith("Hai bên cùng nhau thoả thuận"))
    i1 = next(i for i, x in enumerate(dong) if x.startswith("ĐẠI DIỆN BÊN A"))
    dieu = [x for x in dong[i0:i1] if x]
    d = m.new_doc()
    m.ANH_DAU = anh
    m.quoc_hieu(d)
    m.title(d, "HỢP ĐỒNG NGUYÊN TẮC", f"Số: {hd['so_hd']}", "(V/v mua bán các loại hàng hóa)")
    for c in can_cu:
        m.para(d, c, italic=True, after=1)
    m.para(d, f"Hôm nay, ngày {d0:02d} tháng {m0:02d} năm {y0}, tại Công ty TNHH Thương mại và Công nghiệp Trung Kiên, chúng tôi gồm:",
           after=4)
    m.mixed(d, [("I. BÊN A: ", True), (TK["ten"], True), (" (Bên bán)", False)])
    for t in (f"- Đại diện: Ông {DD['ho_ten']}          Chức vụ: {DD['chuc_vu']}", f"- Địa chỉ: {TK['dia_chi']}, Việt Nam",
              f"- Điện thoại: {TK['dien_thoai']}", f"- Mã số thuế: {TK['mst']}",
              f"- Tài khoản số: {BANK['so_tk']} tại {BANK['ten']}"):
        m.para(d, t, after=0)
    m.para(d, "", after=2)
    m.mixed(d, [("II. BÊN B: ", True), (kh["ten"], True), (" (Bên mua)", False)])
    for t in (f"- Đại diện: {dots(kh.get('dai_dien'), 30)}          Chức vụ: {dots(kh.get('chuc_vu'), 20)}",
              f"- Địa chỉ: {dots(kh.get('dia_chi'), 60)}", f"- Điện thoại: {dots(kh.get('dien_thoai'))}",
              f"- Mã số thuế: {dots(kh.get('mst'))}", f"- Tài khoản số: {dots(kh.get('tai_khoan'), 50)}"):
        m.para(d, t, after=0)
    m.para(d, "", after=2)
    for x in dieu:
        if re.match(r"^ĐIỀU [IVX]+", x):
            m.para(d, x, bold=True, after=2, keep=True)
            continue
        if x.startswith("Hợp đồng có giá trị từ ngày"):
            x = re.sub(r"từ ngày \S+ đến hết ngày \S+", f"từ ngày {d0:02d}/{m0:02d}/{y0} đến hết ngày {hd['hieu_luc_den']}", x)
        if x.startswith("Sau khi hai bên hoàn thành"):
            x = re.sub(r"sau ngày .*$", f"sau ngày {hd['hieu_luc_den']}.", x)
        # tranh chấp: theo quy tắc công ty, tòa án tại TP Hà Nội
        x = x.replace("đưa ra tòa án Tỉnh Quảng Ninh xét xử", "đưa ra Tòa án nhân dân có thẩm quyền tại TP Hà Nội giải quyết")
        assert "Quảng Ninh" not in x, x
        m.para(d, x, after=3)
    m.sign_block(d, "ĐẠI DIỆN BÊN A", "ĐẠI DIỆN BÊN B", left_note="(Ký, ghi rõ họ tên, đóng dấu)",
                 right_note="(Ký, ghi rõ họ tên, đóng dấu)")
    p = out / f"{hd['ten_tep']}{'_CoDau' if anh else ''}.docx"
    d.save(p)
    return p


if __name__ == "__main__":
    f = pathlib.Path(sys.argv[1])
    hd = json.load(open(f, encoding="utf-8"))
    for anh in (None, *sys.argv[2:3]):
        print(build(hd, anh, f.parent))
