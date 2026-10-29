#!/usr/bin/env python3
"""File Excel có macro (.xlsm) theo dõi đơn hàng và công nợ cho MỘT khách hàng.
- Mỗi đơn một sheet (nhiều mã hàng, số lượng, đơn giá, thành tiền chưa/gồm VAT, các lần thu tiền).
- Nút "TẠO ĐƠN MỚI": chép sheet MAU_DON, đặt tên theo ngày (ĐH-29.10.26, trùng ngày thì -2, -3), tự thêm vào TongQuan.
- Nút "CẬP NHẬT TỔNG QUAN": ghi lại danh sách đơn từ các sheet; tự chạy khi mở file (Auto_Open).
Macro: cong-cu/vba/Module1.bas, đóng gói thành vbaProject.bin bằng LibreOffice (cong-cu/vba/tao_vba_bin.py).
Chạy: python3 cong-cu/tao_cong_no_mot_khach.py [ra.xlsm]   (mặc định cong-no/CongNo_ThuLien.xlsm)
Chỉ dùng để tạo file ban đầu; khi anh đã nhập dữ liệu thật thì file của anh là bản gốc, không chạy lại đè lên."""
import pathlib
import subprocess
import sys
import tempfile
import zipfile

import xlsxwriter

ROOT = pathlib.Path(__file__).resolve().parent.parent
VBA_DIR = ROOT / "cong-cu" / "vba"
KHACH = {"ten": "Công ty TNHH vật tư thiết bị Thu Liên", "dia_chi": "Hà Nội", "mst": "", "lien_he": ""}
DS_DON_MAU = [{"ten": "ĐH-2610-003", "noi_dung": "Đơn hàng Dương Kính",
               "ghi_chu": "VÍ DỤ thật từ phiếu xuất kho: chưa có đơn giá, ngày đơn",
               "hang": [("Máy bơm hỏa tiễn 6SA35/6", 2), ("Máy bơm chìm 200WQ600-12-37", 2)]}]
BANG_GIA = [  # nhóm, model, giá chưa VAT, ghi chú (giá anh Trung gửi 29/10/2026; nhóm theo dòng model)
    ("Bơm trục ngang CM", "CM40-160/4", 6200000, ""), ("Bơm trục ngang CM", "CM40-250/11", 13700000, ""),
    ("Bơm trục ngang CM", "CM50-160/5.5", 8500000, ""),
    ("Bơm trục đứng CDLF", "CDLF16-8", 20500000, ""), ("Bơm trục đứng CDLF", "CDLF32-12", 52000000, ""),
    ("Bơm giếng khoan (hỏa tiễn) 6SA/6SP", "6SA30/6", 10200000, ""), ("Bơm giếng khoan (hỏa tiễn) 6SA/6SP", "6SP30-5", 11800000, ""),
    ("Bơm giếng khoan (hỏa tiễn) 6SA/6SP", "6SP30-7", 13500000, ""),
    ("Bơm chìm nước thải WB/WC/WQ", "50WC15-15-1.5", 3800000, ""), ("Bơm chìm nước thải WB/WC/WQ", "80WB50-15-4", 6300000, ""),
    ("Bơm chìm nước thải WB/WC/WQ", "150WC150-15-11", 16000000, ""), ("Bơm chìm nước thải WB/WC/WQ", "150WC150-25-15", 18500000, ""),
    ("Bơm chìm nước thải WB/WC/WQ", "200WQ600-12-37", 69000000, ""),
    ("Khác (anh xác nhận nhóm)", "65DVS53.7", 22000000, ""), ("Khác (anh xác nhận nhóm)", "MB300", 4600000, ""),
    ("Khác (anh xác nhận nhóm)", "SFm6-18-0.55F", 1850000, ""),
    ("Phụ kiện", "Bình tích áp 200L", 5800000, "Trong danh sách anh gửi có 1 dòng 6.000.000"),
    ("Phụ kiện", "Bình tích áp 300L", 7500000, "Trong danh sách anh gửi có 1 dòng 7.300.000"),
    ("Phụ kiện", "Khớp nối nhanh DN80", 3000000, "")]
BG1 = 500                            # bảng giá dùng dòng 2-500
H0, H1, HD1 = 13, 32, 22             # dòng hàng hóa 13-32; bảng hóa đơn của đơn dòng 13-22
CONG, THUE, TONG, HDSUM = H1 + 1, H1 + 2, H1 + 3, HD1 + 1
TT0, TT1 = 13, 212                   # TongQuan: bảng các lần thanh toán (cột O-S), 200 dòng
TQ0, TQ1 = 13, 62                    # TongQuan: 50 đơn ở cột B dòng 13-62 (khớp hằng số trong Module1.bas)
NAVY, BLUE, SKY, PALE, GREY = "#1F3864", "#2E75B6", "#DDEBF7", "#F3F8FC", "#F2F2F2"
TIEN, NGAY = '#,##0;[Red]-#,##0;"-"', "dd/mm/yyyy"


class Fmt:
    """Tạo và dùng lại định dạng ô."""
    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def __call__(self, **k):
        key = tuple(sorted(k.items()))
        if key not in self.cache:
            base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter", "border": 1, "border_color": "#BFBFBF",
                    "text_wrap": True}
            base.update(k)
            self.cache[key] = self.wb.add_format(base)
        return self.cache[key]


def mau_to(ws, rng, formula, bg, f, chu="#000000", dam=False):
    ws.conditional_format(rng, {"type": "formula", "criteria": "=" + formula,
                                "format": f(bg_color=bg, font_color=chu, bold=dam)})


def o_cong_thuc(ws, rng, formula, fm):
    """Ghi công thức vào ô (gộp ô nếu là vùng)."""
    a, b = (rng.split(":") + [None])[:2]
    if b and b != a:
        ws.merge_range(rng, None, fm)
    ws.write_formula(a, formula, fm)


def sheet_don(ws, f, don=None):
    ws.hide_gridlines(2)
    for i, w in enumerate((6, 9, 40, 8, 9, 15, 16, 17, 2, 12, 15, 14, 22, 2, 20)):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    o_cong_thuc(ws, "A1:M1", '="ĐƠN HÀNG - "&UPPER(TongQuan!C3)',
                f(bold=True, font_size=15, font_color="white", bg_color=NAVY, align="center"))
    ws.merge_range("A2:M2", "Bên bán: Công ty TNHH Thương mại và Công nghiệp Trung Kiên  ·  ô trắng là ô nhập, ô xanh nhạt là ô tự tính",
                   f(italic=True, font_size=9, font_color="#595959", align="center", border=0))
    tieu = f(bold=True, font_color="white", bg_color=BLUE)
    ws.merge_range("A3:H3", "THÔNG TIN ĐƠN HÀNG", tieu)
    ws.merge_range("J3:M3", "TIỀN & CÔNG NỢ", tieu)
    nhan, nhap = f(bold=True, bg_color=SKY), f()
    ngay = f(num_format=NGAY, align="left")
    for r, t in ((4, "Mã đơn"), (5, "Ngày đơn"), (6, "Nội dung"), (7, "Số báo giá / HĐ"), (8, "Ghi chú"), (9, "Số mã hàng")):
        ws.merge_range(f"A{r}:B{r}", t, nhan)
    for r, t in ((4, "Giao hàng"), (5, "Ngày giao"), (6, "Xuất hóa đơn"), (7, "Số HĐ / ngày HĐ"), (8, "Hạn thanh toán")):
        ws.merge_range(f"D{r}:E{r}", t, nhan)
    for r in (4, 6, 7, 8):
        ws.write_blank(f"C{r}", None, nhap)
    ws.write_blank("C5", None, ngay)
    ws.write_formula("C9", f"=COUNTA(C{H0}:C{H1})", f(bold=True, bg_color=PALE, align="left"))
    ws.merge_range("F4:H4", None, nhap)
    ws.merge_range("F8:H8", None, ngay)
    ws.write_blank("F5", None, ngay)
    ws.merge_range("G5:H5", None, nhap)
    o_cong_thuc(ws, "F6:H6", f'=IF(COUNTA(J{H0}:J{HD1})=0,"Chưa xuất","Đã xuất")', f(bold=True, bg_color=PALE, align="left"))
    ws.write_formula("F7", f'=IF(COUNTA(J{H0}:J{HD1})=0,"",J{H0}&IF(COUNTA(J{H0}:J{HD1})>1," (+"&(COUNTA(J{H0}:J{HD1})-1)&" HĐ)",""))',
                     f(bg_color=PALE, align="left"))
    o_cong_thuc(ws, "G7:H7", f'=IF(K{H0}="","",K{H0})', f(bg_color=PALE, num_format=NGAY, align="left"))
    ws.data_validation("F4", {"validate": "list", "source": ["Chưa giao", "Giao một phần", "Đã giao"]})
    # thẻ tiền và công nợ
    for r, t, ct in ((4, "Tiền hàng (chưa VAT)", f"=G{CONG}"), (6, "TỔNG CỘNG (gồm VAT)", f"=G{TONG}"),
                     (7, "ĐÃ THU", f"=SUMIF(TongQuan!$P${TT0}:$P${TT1},$C$4,TongQuan!$Q${TT0}:$Q${TT1})"), (8, "CÒN NỢ", "=L6-L7")):
        ws.merge_range(f"J{r}:K{r}", t, f(bold=r in (6, 8), bg_color=SKY))
        o_cong_thuc(ws, f"L{r}:M{r}", ct, f(bold=True, font_size=12 if r in (6, 8) else 10, bg_color=PALE, num_format=TIEN, align="right"))
    ws.write("J5", "Thuế VAT", f(bg_color=SKY))
    ws.write_number("K5", 0.08, f(num_format="0%", align="center"))
    o_cong_thuc(ws, "L5:M5", f"=G{THUE}", f(bold=True, bg_color=PALE, num_format=TIEN, align="right"))
    ws.merge_range("J9:K9", "Tình trạng thu tiền", f(bold=True, bg_color=SKY))
    o_cong_thuc(ws, "L9:M9", '=IF(L6=0,"Chưa có giá",IF(L8<0,"Thu thừa",IF(L8=0,"Đã thu đủ",IF(AND(F8<>"",TODAY()>F8),'
                             '"QUÁ HẠN "&(TODAY()-F8)&" ngày",IF(L7>0,"Thu một phần","Chưa thu")))))',
                f(bold=True, bg_color=PALE, align="center"))
    ws.set_row(9, 26)
    o_cong_thuc(ws, "A10:M10", (f'="CẦN LÀM:  "&IF(C4="","Nhập mã đơn",IF(COUNTIFS(C{H0}:C{H1},"<>",F{H0}:F{H1},"")>0,'
                                '"Có mã hàng CHƯA CÓ GIÁ",IF(AND(F4="Đã giao",F6<>"Đã xuất"),"Đã giao - CHƯA XUẤT HÓA ĐƠN",'
                                'IF(AND(F6="Đã xuất",F4="Chưa giao"),"Đã xuất HĐ - chưa giao hàng",IF(F4="","Cập nhật tình trạng giao hàng",'
                                'IF(AND(L8>0,F8<>"",TODAY()>F8),"Quá hạn "&(TODAY()-F8)&" ngày - đòi tiền",'
                                'IF(L8>0,"Chờ thu tiền",IF(L8<0,"Kiểm tra thu thừa","Hoàn tất"))))))))'),
                f(bold=True, font_size=12, bg_color=GREY, align="center"))
    # bảng hàng hóa và bảng thu tiền
    ws.merge_range(f"A{H0 - 2}:H{H0 - 2}", "HÀNG HÓA TRONG ĐƠN", tieu)
    ws.merge_range(f"J{H0 - 2}:M{H0 - 2}", "HÓA ĐƠN ĐÃ XUẤT CHO ĐƠN NÀY (gõ số HĐ)", tieu)
    dau = f(bold=True, bg_color=SKY, align="center")
    for c, t in zip("ABCDEFGH", ("STT", "Mã kho", "Mã hàng / Tên hàng, model", "ĐVT", "Số lượng", "Đơn giá (chưa VAT)\ngợi ý theo bảng giá",
                                  "Thành tiền (chưa VAT)", "Thành tiền (gồm VAT)")):
        ws.write(f"{c}{H0 - 1}", t, dau)
    for c, t in zip("JKLM", ("Số hóa đơn", "Ngày HĐ", "Giá trị HĐ (gồm VAT)", "Ghi chú")):
        ws.write(f"{c}{H0 - 1}", t, dau)
    ws.set_row(H0 - 2, 44)
    for r in range(H0, H1 + 1):
        nen = "#FFFFFF" if (r - H0) % 2 == 0 else "#FAFCFE"
        ws.write_formula(f"A{r}", f'=IF(C{r}="","",ROW()-{H0 - 1})', f(bg_color=nen, align="center"))
        ws.write_formula(f"B{r}", f'=IF(C{r}="","","KHH")', f(bg_color=nen, align="center"))
        ws.write_blank(f"C{r}", None, f(bg_color=nen))
        ws.write_formula(f"D{r}", f'=IF(C{r}="","","Cái")', f(bg_color=nen, align="center"))
        ws.write_blank(f"E{r}", None, f(bg_color=nen, align="center", num_format="#,##0.##"))
        # giá gợi ý từ BangGia: khớp đúng model, nếu không thì tìm model nằm trong tên hàng; gõ đè để chiết khấu
        ws.write_formula(f"F{r}", (f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},BangGia!$A$2:$B${BG1},2,FALSE),'
                                   f'IFERROR(LOOKUP(2,1/((BangGia!$A$2:$A${BG1}<>"")*ISNUMBER(SEARCH(BangGia!$A$2:$A${BG1},C{r}))),'
                                   f'BangGia!$B$2:$B${BG1}),"")))'), f(bg_color=nen, align="right", num_format=TIEN))
        ws.write_formula(f"G{r}", f'=IF(OR(E{r}="",F{r}=""),"",ROUND(E{r}*F{r},0))', f(bg_color=PALE, align="right", num_format=TIEN))
        ws.write_formula(f"H{r}", f'=IF(G{r}="","",ROUND(G{r}*(1+$K$5),0))', f(bg_color=PALE, align="right", num_format=TIEN))
    nhan_r = f(bold=True, bg_color=SKY, align="right")
    for r, t, ct, sz in ((CONG, "Cộng tiền hàng (chưa VAT)", f"=SUM(G{H0}:G{H1})", 10),
                         (THUE, None, f"=ROUND(G{CONG}*$K$5,0)", 10),
                         (TONG, "TỔNG CỘNG THANH TOÁN (đã gồm VAT)", f"=G{CONG}+G{THUE}", 12)):
        if t is None:
            o_cong_thuc(ws, f"A{r}:F{r}", '="Thuế VAT "&TEXT($K$5,"0%")', nhan_r)
        else:
            ws.merge_range(f"A{r}:F{r}", t, f(bold=True, bg_color=SKY, align="right", font_size=sz))
        o_cong_thuc(ws, f"G{r}:H{r}", ct, f(bold=True, font_size=sz, bg_color=PALE, num_format=TIEN, align="right"))
    for r in range(H0, HD1 + 1):
        ws.write_blank(f"J{r}", None, f(num_format="@", align="center", bold=True))
        ws.write_blank(f"K{r}", None, f(num_format=NGAY, align="center"))
        ws.write_blank(f"L{r}", None, f(num_format=TIEN, align="right"))
        ws.write_blank(f"M{r}", None, f())
    ws.data_validation(f"C{H0}:C{H1}", {"validate": "list", "source": f"=BangGia!$A$2:$A${BG1}", "show_error": False,
                                         "input_title": "Model", "input_message": "Chọn model trong bảng giá hoặc gõ tự do"})
    ws.merge_range(f"J{HDSUM}:K{HDSUM}", "Tổng giá trị HĐ", nhan_r)
    ws.write_formula(f"L{HDSUM}", f"=SUM(L{H0}:L{HD1})", f(bold=True, bg_color=PALE, num_format=TIEN, align="right"))
    ws.write_formula(f"M{HDSUM}", f'=IF(L{HDSUM}=0,"",IF(L{HDSUM}=L6,"Khớp tổng đơn","LỆCH "&TEXT(L{HDSUM}-L6,"#,##0")))',
                     f(bold=True, bg_color=PALE, align="center"))
    ws.merge_range(f"J{HDSUM + 2}:M{HDSUM + 3}", "Các lần khách thanh toán: nhập ở sheet TongQuan (bảng bên phải, chọn đúng mã đơn). "
                   "Ô ĐÃ THU ở trên tự cộng.", f(italic=True, font_color="#595959", border=0))
    mau_to(ws, f"M{HDSUM}", f'ISNUMBER(SEARCH("LỆCH",$M${HDSUM}))', "#FF9999", f, "#7F0000", True)
    mau_to(ws, f"M{HDSUM}", f'$M${HDSUM}="Khớp tổng đơn"', "#C6E0B4", f, "#375623", True)
    # màu cảnh báo
    mau_to(ws, f"F{H0}:F{H1}", f'AND($C{H0}<>"",$F{H0}="")', "#FFE699", f)
    mau_to(ws, "L8:M8", "$L$8>0", "#FCE4E4", f, "#C00000", True)
    mau_to(ws, "L8:M8", "AND($L$6>0,$L$8<=0)", "#E2EFDA", f, "#375623", True)
    mau_to(ws, "L9:M9", 'ISNUMBER(SEARCH("QUÁ HẠN",$L$9))', "#FF9999", f, "#7F0000", True)
    mau_to(ws, "L9:M9", '$L$9="Đã thu đủ"', "#C6E0B4", f, "#375623", True)
    for txt, bg, chu in (("CHƯA XUẤT", "#F8CBAD", "#833C0B"), ("Quá hạn", "#FF9999", "#7F0000"), ("CHƯA CÓ GIÁ", "#FFE699", "#7F6000"),
                         ("Hoàn tất", "#C6E0B4", "#375623"), ("Chờ thu", "#DDEBF7", "#1F3864")):
        mau_to(ws, "A10:M10", f'ISNUMBER(SEARCH("{txt}",$A$10))', bg, f, chu, True)
    for rng, v, bg in (("F4:H4", "Đã giao", "#C6E0B4"), ("F4:H4", "Chưa giao", "#FFE699"), ("F6:H6", "Đã xuất", "#C6E0B4"),
                       ("F6:H6", "Chưa xuất", "#FFE699")):
        mau_to(ws, rng, f'${rng[0]}${rng[1]}="{v}"', bg, f)
    ws.insert_button("O3", {"macro": "VeTongQuan", "caption": "VỀ TỔNG QUAN", "width": 140, "height": 34})
    ws.insert_button("O6", {"macro": "TaoDonMoi", "caption": "+ TẠO ĐƠN MỚI", "width": 140, "height": 34})
    ws.freeze_panes(H0 - 1, 0)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.print_area(f"A1:M{TONG}")
    if don:
        ws.write("C4", don["ten"], nhap)
        ws.write("C6", don["noi_dung"], nhap)
        ws.write("C8", don["ghi_chu"], nhap)
        for i, (ten, sl) in enumerate(don["hang"]):
            nen = "#FFFFFF" if i % 2 == 0 else "#FAFCFE"
            ws.write(f"C{H0 + i}", ten, f(bg_color=nen))
            ws.write_number(f"E{H0 + i}", sl, f(bg_color=nen, align="center", num_format="#,##0.##"))


def tong_quan(wb, ws, f, ds):
    ws.hide_gridlines(2)
    for i, w in enumerate((5, 15, 12, 32, 9, 16, 16, 16, 13, 18, 12, 34, 10, 2, 12, 15, 15, 14, 24)):
        ws.set_column(i, i, w)
    ws.set_row(0, 34)
    ws.merge_range("A1:S1", "THEO DÕI ĐƠN HÀNG & CÔNG NỢ KHÁCH HÀNG",
                   f(bold=True, font_size=16, font_color="white", bg_color=NAVY, align="center"))
    o_cong_thuc(ws, "A2:S2", '="Công ty TNHH Thương mại và Công nghiệp Trung Kiên  ·  cập nhật "&TEXT(TODAY(),"dd/mm/yyyy")',
                f(italic=True, font_size=9, font_color="#595959", align="center", border=0))
    for r, t, v in ((3, "Khách hàng", KHACH["ten"]), (4, "Địa chỉ", KHACH["dia_chi"]), (5, "Mã số thuế", KHACH["mst"]),
                    (6, "Liên hệ / ĐT", KHACH["lien_he"])):
        ws.merge_range(f"A{r}:B{r}", t, f(bold=True, bg_color=SKY))
        ws.merge_range(f"C{r}:H{r}", v, f(bold=r == 3, font_size=12 if r == 3 else 10))
    ws.set_row(2, 22)
    a, b = TQ0, TQ1
    for r1, r2, t, ct, mau, nf in (("A8:C8", "A9:C9", "TỔNG GIÁ TRỊ ĐƠN", f"=SUM(F{a}:F{b})", "#1F3864", TIEN),
                                   ("D8", "D9", "ĐÃ THU", f"=SUM(G{a}:G{b})", "#375623", TIEN),
                                   ("E8:G8", "E9:G9", "CÒN NỢ", f"=SUM(H{a}:H{b})", "#C00000", TIEN),
                                   ("H8:J8", "H9:J9", "NỢ QUÁ HẠN", f'=SUMIF(M{a}:M{b},">0",H{a}:H{b})', "#7F0000", TIEN),
                                   ("K8:M8", "K9:M9", "SỐ ĐƠN", f'=COUNTIF(B{a}:B{b},"?*")', "#1F3864", "0")):
        hf = f(bold=True, font_color="white", bg_color=mau, align="center")
        if ":" in r1:
            ws.merge_range(r1, t, hf)
        else:
            ws.write(r1, t, hf)
        o_cong_thuc(ws, r2, ct, f(bold=True, font_size=18, font_color=mau, bg_color=PALE, align="center", num_format=nf))
    ws.set_row(8, 36)
    df = f(font_color="#404040", bg_color=GREY, align="center")
    for rng, ct in (("A10:C10", f'="Chưa giao: "&COUNTIF(I{a}:I{b},"Chưa giao")&"  ·  Giao một phần: "&COUNTIF(I{a}:I{b},"Giao một phần")'),
                    ("D10:G10", f'="Đã giao chưa xuất HĐ: "&COUNTIFS(I{a}:I{b},"Đã giao",J{a}:J{b},"Chưa xuất")'),
                    ("H10:J10", f'="Chưa thu đủ: "&COUNTIF(H{a}:H{b},">0")&" đơn"'),
                    ("K10:M10", f'="Thiếu giá: "&COUNTIF(L{a}:L{b},"*CHƯA CÓ GIÁ*")&" đơn"')):
        o_cong_thuc(ws, rng, ct, df)
    ws.merge_range("A11:M11", "DANH SÁCH ĐƠN HÀNG  (tự cập nhật khi bấm nút hoặc mở file)", f(bold=True, font_color="white", bg_color=BLUE))
    for i, t in enumerate(("STT", "Mã đơn", "Ngày đơn", "Nội dung đơn", "Số mã hàng", "Tổng tiền (gồm VAT)", "Đã thu", "Còn nợ",
                           "Giao hàng", "Xuất HĐ", "Hạn TT", "Cần làm", "Quá hạn (ngày)")):
        ws.write(11, i, t, f(bold=True, bg_color=SKY, align="center"))
    ws.set_row(11, 30)

    def lay(r, cell, so=False):
        rong = "0" if so else '""'
        return (f'=IF($B{r}="","",IFERROR(IF(INDIRECT("\'"&$B{r}&"\'!{cell}")="",{rong},'
                f'INDIRECT("\'"&$B{r}&"\'!{cell}")),"Không thấy sheet"))')

    for r in range(a, b + 1):
        nen = "#FFFFFF" if (r - a) % 2 == 0 else "#FAFCFE"
        ws.write_formula(f"A{r}", f'=IF(B{r}="","",ROW()-{a - 1})', f(bg_color=nen, align="center"))
        ws.write_blank(f"B{r}", None, f(bg_color=nen, align="center", bold=True))
        ws.write_formula(f"C{r}", lay(r, "C5"), f(bg_color=nen, align="center", num_format=NGAY))
        ws.write_formula(f"D{r}", lay(r, "C6"), f(bg_color=nen))
        ws.write_formula(f"E{r}", lay(r, "C9", True), f(bg_color=nen, align="center"))
        for c, cell in (("F", "L6"), ("G", "L7"), ("H", "L8")):
            ws.write_formula(f"{c}{r}", lay(r, cell, True), f(bg_color=nen, align="right", num_format=TIEN))
        ws.write_formula(f"I{r}", lay(r, "F4"), f(bg_color=nen, align="center"))
        ws.write_formula(f"J{r}", f'=IF($B{r}="","",IFERROR(IF(INDIRECT("\'"&$B{r}&"\'!F6")="Đã xuất","HĐ "&INDIRECT("\'"&$B{r}&"\'!F7"),'
                                  f'"Chưa xuất"),"Không thấy sheet"))', f(bg_color=nen, align="center"))
        ws.write_formula(f"K{r}", lay(r, "F8"), f(bg_color=nen, align="center", num_format=NGAY))
        ws.write_formula(f"L{r}", f'=IF(B{r}="","",IFERROR(SUBSTITUTE(INDIRECT("\'"&$B{r}&"\'!A10"),"CẦN LÀM:  ",""),"Không thấy sheet"))',
                         f(bg_color=nen))
        ws.write_formula(f"M{r}", f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(H{r})),IF(AND(H{r}>0,TODAY()>K{r}),TODAY()-K{r},""),"")',
                         f(bg_color=nen, align="center"))
    for i, ten in enumerate(ds):
        ws.write(f"B{a + i}", ten, f(bg_color="#FFFFFF" if i % 2 == 0 else "#FAFCFE", align="center", bold=True))
    for txt, bg, chu in (("CHƯA XUẤT", "#F8CBAD", "#833C0B"), ("Quá hạn", "#FF9999", "#7F0000"), ("CHƯA CÓ GIÁ", "#FFE699", "#7F6000"),
                         ("Không thấy sheet", "#D9D9D9", "#C00000")):
        mau_to(ws, f"L{a}:L{b}", f'ISNUMBER(SEARCH("{txt}",$L{a}))', bg, f, chu, True)
    mau_to(ws, f"L{a}:L{b}", f'$L{a}="Hoàn tất"', "#C6E0B4", f, "#375623", True)
    mau_to(ws, f"M{a}:M{b}", f"AND(ISNUMBER($M{a}),$M{a}>0)", "#FF9999", f, "#7F0000", True)
    for c, v, bg in (("I", "Đã giao", "#C6E0B4"), ("I", "Chưa giao", "#FFE699"), ("J", "Chưa xuất", "#FFE699")):
        mau_to(ws, f"{c}{a}:{c}{b}", f'${c}{a}="{v}"', bg, f)
    mau_to(ws, f"J{a}:J{b}", f'LEFT($J{a},2)="HĐ"', "#C6E0B4", f)
    # bảng các lần khách thanh toán (nhập ở đây; mỗi đơn tự cộng theo mã đơn)
    ws.merge_range("O11:S11", "CÁC LẦN KHÁCH THANH TOÁN (nhập ở đây)", f(bold=True, font_color="white", bg_color="#375623"))
    for c, t in zip("OPQRS", ("Ngày thu", "Mã đơn", "Số tiền", "Hình thức", "Ghi chú / chứng từ")):
        ws.write(f"{c}12", t, f(bold=True, bg_color="#E2EFDA", align="center"))
    for r in range(TT0, TT1 + 1):
        nen = "#FFFFFF" if (r - TT0) % 2 == 0 else "#F6FAF3"
        ws.write_blank(f"O{r}", None, f(bg_color=nen, num_format=NGAY, align="center"))
        ws.write_blank(f"P{r}", None, f(bg_color=nen, align="center", bold=True))
        ws.write_blank(f"Q{r}", None, f(bg_color=nen, num_format=TIEN, align="right"))
        ws.write_blank(f"R{r}", None, f(bg_color=nen, align="center"))
        ws.write_blank(f"S{r}", None, f(bg_color=nen))
    ws.data_validation(f"P{TT0}:P{TT1}", {"validate": "list", "source": f"=$B${a}:$B${b}", "error_title": "Mã đơn",
                                          "error_message": "Chọn mã đơn có trong danh sách đơn hàng bên trái"})
    ws.data_validation(f"R{TT0}:R{TT1}", {"validate": "list", "source": ["Chuyển khoản", "Tiền mặt", "Khác"]})
    o_cong_thuc(ws, "O10:S10", (f'="Tổng các lần thu: "&TEXT(SUM(Q{TT0}:Q{TT1}),"#,##0")&IF(SUM(Q{TT0}:Q{TT1})<>SUM(G{a}:G{b}),'
                                f'"  ·  "&TEXT(SUM(Q{TT0}:Q{TT1})-SUM(G{a}:G{b}),"#,##0")&" CHƯA KHỚP MÃ ĐƠN","")'), df)
    mau_to(ws, "O10:S10", 'ISNUMBER(SEARCH("CHƯA KHỚP",$O$10))', "#FF9999", f, "#7F0000", True)
    ws.conditional_format(f"H{a}:H{b}", {"type": "data_bar", "bar_color": "#F4B183", "bar_solid": True})
    ws.insert_button("J3", {"macro": "TaoDonMoi", "caption": "+ TẠO ĐƠN MỚI", "width": 230, "height": 40})
    ws.insert_button("J5", {"macro": "CapNhatTongQuan", "caption": "CẬP NHẬT TỔNG QUAN", "width": 230, "height": 32})
    ws.freeze_panes(a - 1, 0)
    ch = wb.add_chart({"type": "column"})
    for col, ten, mau in (("F", "Tổng tiền", "#2E75B6"), ("H", "Còn nợ", "#C00000")):
        ch.add_series({"name": ten, "categories": f"=TongQuan!$B${a}:$B${a + 19}",
                       "values": f"=TongQuan!${col}${a}:${col}${a + 19}", "fill": {"color": mau}})
    ch.set_title({"name": "Tổng tiền và còn nợ theo đơn (20 đơn đầu)"})
    ch.set_y_axis({"num_format": "#,##0"})
    ch.set_size({"width": 900, "height": 300})
    ws.insert_chart(f"A{b + 3}", ch)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)


def bang_gia(ws, f):
    ws.hide_gridlines(2)
    for i, w in enumerate((30, 18, 34, 44)):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    for c, t in enumerate(("Model / tên hàng", "Giá chưa VAT", "Nhóm hàng", "Ghi chú")):
        ws.write(0, c, t, f(bold=True, font_color="white", bg_color=NAVY, align="center"))
    nhom_truoc, dai = None, 0
    for r in range(1, BG1):
        if r - 1 < len(BANG_GIA):
            nhom, model, gia, gc = BANG_GIA[r - 1]
            if nhom != nhom_truoc:
                dai, nhom_truoc = dai + 1, nhom
            nen = "#FFFFFF" if dai % 2 else "#EAF1FB"
            ws.write(r, 0, model, f(bg_color=nen, bold=True))
            ws.write_number(r, 1, gia, f(bg_color=nen, num_format="#,##0", align="right"))
            ws.write(r, 2, nhom, f(bg_color=nen, font_color="#1F3864"))
            ws.write(r, 3, gc, f(bg_color=nen, font_color="#C00000", italic=True)) if gc else ws.write_blank(r, 3, None, f(bg_color=nen))
        else:
            for c, nf in ((0, None), (1, "#,##0"), (2, None), (3, None)):
                ws.write_blank(r, c, None, f(num_format=nf) if nf else f())
    ws.conditional_format(f"A2:A{BG1}", {"type": "duplicate", "format": f(bg_color="#FFC7CE", font_color="#9C0006")})
    ws.freeze_panes(1, 0)
    ws.autofilter(f"A1:D{BG1}")
    ws.write("F1", "Cách dùng: gõ thêm model, giá và nhóm vào dòng trống bên dưới. Các sheet đơn tự lấy giá gợi ý khi gõ đúng model "
                   "(hoặc tên hàng có chứa model). Model bị trùng sẽ tô đỏ - chỉ giữ một dòng. Muốn xếp lại theo nhóm: bấm mũi tên lọc ở cột "
                   "Nhóm hàng > Sắp xếp A-Z.", f(italic=True, border=0, text_wrap=False))


def huong_dan(ws, f):
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 125)
    dong = [("CÁCH DÙNG", True),
            ("0. Lần đầu mở file: nếu Excel báo chặn macro, đóng file > chuột phải vào file > Properties (Thuộc tính) > tích Unblock (Bỏ chặn) > OK,", False),
            ("    mở lại và bấm 'Enable Content' (Bật nội dung). Nút bấm chỉ chạy trên Excel máy tính (Windows/Mac), không chạy trên điện thoại.", False),
            ("1. Đơn mới: ở TongQuan bấm nút '+ TẠO ĐƠN MỚI'. File tự tạo sheet tên theo ngày (ví dụ ĐH-29.10.26; trùng ngày thì ĐH-29.10.26-2)", False),
            ("    và tự thêm vào bảng TongQuan. Sheet mới mở sẵn để nhập.", False),
            ("2. Trong sheet đơn: gõ nội dung, số báo giá/HĐ; mỗi mã hàng một dòng: chỉ cần gõ tên/model, số lượng, đơn giá CHƯA VAT.", False),
            ("    STT, Mã kho (KHH), ĐVT (Cái) tự điền; muốn khác thì gõ đè. Thành tiền chưa VAT, gồm VAT, thuế VAT và tổng cộng tự tính.", False),
            ("    Giá: gõ model (hoặc chọn trong danh sách) là đơn giá gợi ý tự hiện theo sheet BangGia; chiết khấu thì gõ đè giá mới.", False),
            ("    Bảng giá: sheet BangGia, gõ thêm model và giá vào dòng trống bên dưới, các đơn tự dùng.", False),
            ("3. Chọn Giao hàng trong ô có danh sách, gõ hạn thanh toán. VAT mặc định 8% (ô K5).", False),
            ("4. Hóa đơn: trong sheet đơn, gõ số hóa đơn (và ngày, giá trị) ở bảng 'Hóa đơn đã xuất' là đơn tự chuyển 'Đã xuất'.", False),
            ("    Mỗi lần khách trả tiền: ghi một dòng ở bảng 'Các lần khách thanh toán' bên phải sheet TongQuan, chọn đúng mã đơn. Còn nợ tự trừ.", False),
            ("5. Bấm 'VỀ TỔNG QUAN' để quay lại; tổng quan tự cập nhật. Nếu tự chép sheet bằng tay thì bấm 'CẬP NHẬT TỔNG QUAN'.", False),
            ("Không sửa, không xóa sheet MAU_DON (mẫu trống để tạo đơn). Xóa đơn: xóa sheet đơn đó rồi bấm 'CẬP NHẬT TỔNG QUAN'.", False),
            ("Màu: cam = đã giao chưa xuất HĐ · đỏ = quá hạn · vàng = thiếu giá / chưa giao · xanh lá = hoàn tất.", False)]
    for i, (t, b) in enumerate(dong):
        ws.write(i, 0, t, f(bold=b, font_size=12 if b else 11, font_color=NAVY if b else "#000000", border=0, text_wrap=False))


def dong_goi_xlsm(xlsx, bin_path, ra):
    """Thêm vbaProject.bin vào gói xlsx và đổi kiểu sang xlsm (giống cách Excel/LibreOffice lưu)."""
    with zipfile.ZipFile(xlsx) as zi, zipfile.ZipFile(ra, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == "[Content_Types].xml":
                s = data.decode()
                assert "spreadsheetml.sheet.main+xml" in s
                s = s.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                              "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                s = s.replace("</Types>", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
                data = s.encode()
            elif it.filename == "xl/_rels/workbook.xml.rels":
                data = data.decode().replace("</Relationships>", '<Relationship Id="rIdVBA1" Type="http://schemas.microsoft.com/office/2006/'
                                             'relationships/vbaProject" Target="vbaProject.bin"/></Relationships>').encode()
            zo.writestr(it, data)
        zo.write(bin_path, "xl/vbaProject.bin")


def tao(ra):
    ra = pathlib.Path(ra)
    with tempfile.TemporaryDirectory() as t:
        bin_path = pathlib.Path(t) / "vbaProject.bin"
        subprocess.run([sys.executable, str(VBA_DIR / "tao_vba_bin.py"), str(VBA_DIR / "Module1.bas"), str(bin_path)], check=True)
        tmp = pathlib.Path(t) / "tam.xlsx"
        wb = xlsxwriter.Workbook(str(tmp))
        f = Fmt(wb)
        tq = wb.add_worksheet("TongQuan")
        hd = wb.add_worksheet("HuongDan")
        bg = wb.add_worksheet("BangGia")
        bg.set_tab_color("#548235")
        bang_gia(bg, f)
        mau = wb.add_worksheet("MAU_DON")
        tq.set_tab_color(NAVY)
        mau.set_tab_color("#A6A6A6")
        sheet_don(mau, f)
        ds = []
        for d in DS_DON_MAU:
            ws = wb.add_worksheet(d["ten"])
            ws.set_tab_color(BLUE)
            sheet_don(ws, f, d)
            ds.append(d["ten"])
        tong_quan(wb, tq, f, ds)
        huong_dan(hd, f)
        tq.activate()
        wb.close()
        ra.parent.mkdir(parents=True, exist_ok=True)
        dong_goi_xlsm(tmp, bin_path, ra)
    print(ra)


if __name__ == "__main__":
    tao(sys.argv[1] if len(sys.argv) > 1 else ROOT / "cong-no" / "CongNo_ThuLien.xlsm")
