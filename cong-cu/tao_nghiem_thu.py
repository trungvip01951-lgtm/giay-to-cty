#!/usr/bin/env python3
"""Biên bản nghiệm thu và chạy thử thiết bị (bản in), điền sẵn bên mua và danh mục hàng; số, ngày để trống nếu không có.
Dùng chung JSON hợp đồng (khach, hang, so_hd, ngay_hd_chu). Chạy: python3 cong-cu/tao_nghiem_thu.py don.json"""
import json, pathlib, sys
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402

TK = m.TK


def build(don, js):
    kh = don["khach"]
    g = lambda k, n=25: kh.get(k) or m.dots(n)  # noqa: E731
    d = m.new_doc()
    d.styles["Normal"].font.size = m.Pt(11.5)   # gọn 1 trang
    m.quoc_hieu(d)
    m.title(d, "BIÊN BẢN NGHIỆM THU VÀ BÀN GIAO THIẾT BỊ", "(Số: ....../2026/BBNT)")
    m.para(d, f"Căn cứ Hợp đồng kinh tế số {don.get('so_hd') or m.dots(25)} ngày {don.get('ngay_hd_chu') or '..... tháng ..... năm 2026'} giữa hai bên;", italic=True, after=6)
    m.para(d, f"Hôm nay, ngày ..... tháng ..... năm 2026, tại {m.dots(50)}, chúng tôi gồm:", after=4)
    m.mixed(d, [("BÊN MUA: ", True), (kh["ten"], True), (" (gọi tắt là bên A)", False)])
    m.para(d, f"Địa chỉ: {kh['dia_chi']}", after=0)
    m.para(d, f"Mã số thuế: {kh['mst']}   Điện thoại: {g('dien_thoai')}", after=0)
    m.para(d, f"Đại diện: {g('dai_dien', 30)}          Chức vụ: {g('chuc_vu', 20)}", after=6)
    m.mixed(d, [("BÊN BÁN: ", True), (TK["ten"], True), (" (gọi tắt là bên B)", False)])
    m.para(d, f"Địa chỉ: {TK['dia_chi']}", after=0)
    m.para(d, f"Mã số thuế: {TK['mst']}   Điện thoại: {TK['dien_thoai']}", after=0)
    m.para(d, f"Đại diện: {TK['dai_dien']['ho_ten']}          Chức vụ: {TK['dai_dien']['chuc_vu']}", after=6)
    m.para(d, "I. Thiết bị nghiệm thu", bold=True, after=2, keep=True)
    m.table(d, ["STT", "Tên thiết bị / Model", "ĐVT", "SL", "Tình trạng"],
            [[str(i), h["ten"].replace("\n", " – "), h["dvt"], str(h["sl"]), "Mới 100%"] for i, h in enumerate(don["hang"], 1)],
            [1.2, 9.0, 1.5, 1.3, 3.0], size=11, center_cols=(0, 2, 3, 4))
    m.para(d, "II. Kết quả kiểm tra và chạy thử", bold=True, after=2, keep=True)
    m.table(d, ["Hạng mục kiểm tra", "Đạt", "Không đạt"],
            [["Số lượng, chủng loại, model đúng theo Hợp đồng", "☐", "☐"],
             ["Đấu nối điện, chiều quay động cơ", "☐", "☐"],
             ["Thiết bị chạy thử ổn định, không rung, không ồn bất thường", "☐", "☐"],
             ["Không rò rỉ nước, dầu tại phớt và mối nối", "☐", "☐"]][:4],
            [11.0, 2.4, 2.6], size=11, center_cols=(1, 2))
    m.para(d, "III. Kết luận", bold=True, after=2, keep=True)
    m.para(d, "☐ Đạt yêu cầu, hai bên thống nhất nghiệm thu, bàn giao và đưa vào sử dụng.   ☐ Chưa đạt, bên B khắc phục.", after=1)
    m.para(d, f"Bảo hành {don.get('bao_hanh', '12 tháng')} kể từ ngày ký biên bản này. Biên bản là căn cứ để bên A thanh toán phần còn lại "
              "của Hợp đồng, lập thành 02 bản có giá trị như nhau, mỗi bên giữ 01 bản.", after=6)
    m.ANH_DAU = None
    m.sign_block(d, "ĐẠI DIỆN BÊN MUA", "ĐẠI DIỆN BÊN BÁN", left_note="(Ký, ghi rõ họ tên, đóng dấu)", gap=3)
    p = pathlib.Path(js).parent / ((don.get("ten_tep_nt") or "BienBanNghiemThu") + ".docx")
    d.save(p)
    return p


if __name__ == "__main__":
    print(build(json.load(open(sys.argv[1], encoding="utf-8")), sys.argv[1]))
