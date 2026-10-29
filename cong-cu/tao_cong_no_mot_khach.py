#!/usr/bin/env python3
"""File Excel theo dõi đơn hàng và công nợ cho MỘT khách hàng: mỗi đơn một sheet (nhiều mã hàng, số lượng, giá, các lần thu),
sheet TongQuan tổng hợp tất cả đơn bằng INDIRECT theo tên sheet, có thẻ số liệu, màu cảnh báo và biểu đồ.
Chạy: python3 cong-cu/tao_cong_no_mot_khach.py   (ra cong-no/CongNo_ThuLien.xlsx)
Chỉ dùng để tạo file ban đầu; khi anh đã nhập dữ liệu thật thì file .xlsx là bản gốc, không chạy lại đè lên."""
import datetime as dt
import pathlib
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = pathlib.Path(__file__).resolve().parent.parent
KHACH = {"ten": "Công ty TNHH vật tư thiết bị Thu Liên", "dia_chi": "Hà Nội", "mst": "", "lien_he": ""}
N_DON, HANG0, HANG1, THU1 = 50, 13, 32, 24          # 50 đơn; hàng hóa dòng 13-32 (20 dòng); thu tiền dòng 13-24

F = "Arial"
NAVY, BLUE, SKY, PALE, GREY, WHITE = "1F3864", "2E75B6", "DDEBF7", "F3F8FC", "F2F2F2", "FFFFFF"
fill = lambda c: PatternFill("solid", fgColor=c)                                   # noqa: E731
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
TIEN, NGAY = '#,##0;[Red]-#,##0;"-"', "dd/mm/yyyy"
C, LFT, R = (Alignment(horizontal=h, vertical="center", wrap_text=True) for h in ("center", "left", "right"))


def o(ws, ref, v=None, b=False, sz=10, col="000000", bg=None, al=None, nf=None, it=False, box=True):
    c = ws[ref]
    if v is not None:
        c.value = v
    c.font = Font(name=F, size=sz, bold=b, italic=it, color=col)
    if bg:
        c.fill = fill(bg)
    if al:
        c.alignment = al
    if nf:
        c.number_format = nf
    if box:
        c.border = BOX
    return c


def gop(ws, rng, v=None, **k):
    ws.merge_cells(rng)
    first = rng.split(":")[0]
    for row in ws[rng]:
        for c in row:
            c.border = BOX
            if k.get("bg"):
                c.fill = fill(k["bg"])
    return o(ws, first, v, **k)


def chon(ws, rng, items):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def to_mau(ws, rng, cong_thuc, bg, chu="000000", dam=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[cong_thuc], fill=PatternFill("solid", bgColor=bg, fgColor=bg),
                                                   font=Font(color=chu, bold=dam)))


def sheet_don(ws, mau=None):
    """Một sheet = một đơn hàng. Ô cố định để TongQuan đọc: C4 mã, C5 ngày, C6 nội dung, C7 BG/HĐ, C9 số mã hàng,
    F4 giao hàng, F5 ngày giao, F6 xuất HĐ, F7 số HĐ, G7 ngày HĐ, F8 hạn TT, K6 tổng, K7 đã thu, K8 còn nợ, K9 tình trạng, A10 cần làm."""
    for col, w in zip("ABCDEFGHIJKL", (6, 9, 42, 8, 9, 15, 17, 2, 12, 16, 15, 24)):
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    gop(ws, "A1:L1", '="ĐƠN HÀNG - "&UPPER(TongQuan!C3)', b=True, sz=15, col=WHITE, bg=NAVY, al=C)
    ws.row_dimensions[1].height = 30
    gop(ws, "A2:L2", "Bên bán: Công ty TNHH Thương mại và Công nghiệp Trung Kiên  ·  ô trắng là ô nhập, ô xanh nhạt là ô tự tính",
        sz=9, it=True, col="595959", al=C, box=False)
    # --- thông tin đơn
    gop(ws, "A3:G3", "THÔNG TIN ĐƠN HÀNG", b=True, col=WHITE, bg=BLUE, al=LFT)
    gop(ws, "I3:L3", "TIỀN & CÔNG NỢ", b=True, col=WHITE, bg=BLUE, al=LFT)
    nhan = [("A4:B4", "Mã đơn"), ("A5:B5", "Ngày đơn"), ("A6:B6", "Nội dung"), ("A7:B7", "Số báo giá / HĐ"), ("A8:B8", "Ghi chú"),
            ("A9:B9", "Số mã hàng"), ("D4:E4", "Giao hàng"), ("D5:E5", "Ngày giao"), ("D6:E6", "Xuất hóa đơn"),
            ("D7:E7", "Số HĐ / ngày"), ("D8:E8", "Hạn thanh toán")]
    for rng, t in nhan:
        gop(ws, rng, t, b=True, bg=SKY, al=LFT)
    for ref, nf in (("C4", None), ("C5", NGAY), ("C6", None), ("C7", None), ("C8", None), ("F5", NGAY), ("F7", None), ("G7", NGAY)):
        o(ws, ref, al=LFT, nf=nf)
    o(ws, "C9", f"=COUNTA(C{HANG0}:C{HANG1})", b=True, bg=PALE, al=LFT)
    for rng in ("F4:G4", "F6:G6", "F8:G8"):
        gop(ws, rng, al=LFT)
    ws["F8"].number_format = NGAY
    chon(ws, "F4", ["Chưa giao", "Giao một phần", "Đã giao"])
    chon(ws, "F6", ["Chưa xuất", "Đã xuất"])
    # --- thẻ tiền
    the = [(4, "Tiền hàng (chưa VAT)", f"=G{HANG1 + 1}"), (5, "Thuế VAT", "=ROUND(K4*J5,0)"), (6, "TỔNG CỘNG", "=K4+K5"),
           (7, "ĐÃ THU", f"=J{THU1 + 1}"), (8, "CÒN NỢ", "=K6-K7")]
    for r, t, ct in the:
        if r == 5:
            o(ws, "I5", t, bg=SKY, al=LFT)
        else:
            gop(ws, f"I{r}:J{r}", t, b=r in (6, 8), bg=SKY, al=LFT)
        gop(ws, f"K{r}:L{r}", ct, b=True, sz=12 if r in (6, 8) else 10, bg=PALE, al=R, nf=TIEN)
    o(ws, "J5", 0.08, nf="0%", al=C)                            # VAT % (ô nhập)
    gop(ws, "I9:J9", "Tình trạng thu tiền", b=True, bg=SKY, al=LFT)
    gop(ws, "K9:L9", '=IF(K6=0,"Chưa có giá",IF(K8<0,"Thu thừa",IF(K8=0,"Đã thu đủ",IF(AND(F8<>"",TODAY()>F8),'
                     '"QUÁ HẠN "&(TODAY()-F8)&" ngày",IF(K7>0,"Thu một phần","Chưa thu")))))', b=True, bg=PALE, al=C)
    gop(ws, "A10:L10", (f'="CẦN LÀM:  "&IF(C4="","Nhập mã đơn",IF(COUNTIFS(C{HANG0}:C{HANG1},"<>",F{HANG0}:F{HANG1},"")>0,'
                        '"Có mã hàng CHƯA CÓ GIÁ",IF(AND(F4="Đã giao",F6<>"Đã xuất"),"Đã giao - CHƯA XUẤT HÓA ĐƠN",'
                        'IF(AND(F6="Đã xuất",F4="Chưa giao"),"Đã xuất HĐ - chưa giao hàng",IF(F4="","Cập nhật tình trạng giao hàng",'
                        'IF(AND(K8>0,F8<>"",TODAY()>F8),"Quá hạn "&(TODAY()-F8)&" ngày - đòi tiền",'
                        'IF(K8>0,"Chờ thu tiền",IF(K8<0,"Kiểm tra thu thừa","Hoàn tất"))))))))'),
        b=True, sz=12, bg=GREY, al=C)
    ws.row_dimensions[10].height = 26
    # --- bảng hàng hóa
    gop(ws, f"A{HANG0 - 2}:G{HANG0 - 2}", "HÀNG HÓA TRONG ĐƠN", b=True, col=WHITE, bg=BLUE, al=LFT)
    gop(ws, f"I{HANG0 - 2}:L{HANG0 - 2}", "CÁC LẦN KHÁCH THANH TOÁN", b=True, col=WHITE, bg=BLUE, al=LFT)
    for col, t in zip("ABCDEFG", ("STT", "Mã kho", "Mã hàng / Tên hàng, model", "ĐVT", "Số lượng", "Đơn giá (chưa VAT)", "Thành tiền")):
        o(ws, f"{col}{HANG0 - 1}", t, b=True, bg=SKY, al=C)
    for col, t in zip("IJKL", ("Ngày thu", "Số tiền", "Hình thức", "Ghi chú / chứng từ")):
        o(ws, f"{col}{HANG0 - 1}", t, b=True, bg=SKY, al=C)
    for r in range(HANG0, HANG1 + 1):
        nen = WHITE if (r - HANG0) % 2 == 0 else "FAFCFE"
        o(ws, f"A{r}", f'=IF(C{r}="","",ROW()-{HANG0 - 1})', bg=nen, al=C)
        for col in "BCDE":
            o(ws, f"{col}{r}", bg=nen, al=LFT if col == "C" else C, nf="#,##0.##" if col == "E" else None)
        o(ws, f"F{r}", bg=nen, al=R, nf=TIEN)
        o(ws, f"G{r}", f'=IF(OR(E{r}="",F{r}=""),"",ROUND(E{r}*F{r},0))', bg=PALE, al=R, nf=TIEN)
    gop(ws, f"A{HANG1 + 1}:F{HANG1 + 1}", "Cộng tiền hàng (chưa VAT)", b=True, bg=SKY, al=R)
    o(ws, f"G{HANG1 + 1}", f"=SUM(G{HANG0}:G{HANG1})", b=True, bg=PALE, al=R, nf=TIEN)
    to_mau(ws, f"F{HANG0}:F{HANG1}", f'AND($C{HANG0}<>"",$F{HANG0}="")', "FFE699")
    for r in range(HANG0, THU1 + 1):
        o(ws, f"I{r}", al=C, nf=NGAY)
        o(ws, f"J{r}", al=R, nf=TIEN)
        o(ws, f"K{r}", al=C)
        o(ws, f"L{r}", al=LFT)
    chon(ws, f"K{HANG0}:K{THU1}", ["Chuyển khoản", "Tiền mặt", "Khác"])
    o(ws, f"I{THU1 + 1}", "Đã thu", b=True, bg=SKY, al=R)
    o(ws, f"J{THU1 + 1}", f"=SUM(J{HANG0}:J{THU1})", b=True, bg=PALE, al=R, nf=TIEN)
    # --- màu cảnh báo
    to_mau(ws, "K8:L8", "$K$8>0", "FCE4E4", "C00000", True)
    to_mau(ws, "K8:L8", "AND($K$6>0,$K$8<=0)", "E2EFDA", "375623", True)
    for txt, bg, chu in (("CHƯA XUẤT", "F8CBAD", "833C0B"), ("Quá hạn", "FF9999", "7F0000"), ("CHƯA CÓ GIÁ", "FFE699", "7F6000"),
                         ("Hoàn tất", "C6E0B4", "375623"), ("Chờ thu", "DDEBF7", "1F3864")):
        to_mau(ws, "A10:L10", f'ISNUMBER(SEARCH("{txt}",$A$10))', bg, chu, True)
    to_mau(ws, "K9:L9", 'ISNUMBER(SEARCH("QUÁ HẠN",$K$9))', "FF9999", "7F0000", True)
    to_mau(ws, "K9:L9", '$K$9="Đã thu đủ"', "C6E0B4", "375623", True)
    to_mau(ws, "F4:G4", '$F$4="Đã giao"', "C6E0B4")
    to_mau(ws, "F4:G4", '$F$4="Chưa giao"', "FFE699")
    to_mau(ws, "F6:G6", '$F$6="Đã xuất"', "C6E0B4")
    to_mau(ws, "F6:G6", '$F$6="Chưa xuất"', "FFE699")
    ws.freeze_panes = f"A{HANG0}"
    ws.page_setup.orientation, ws.page_setup.paperSize = "landscape", ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    if mau:
        for ref, v in mau["o"].items():
            ws[ref] = v
        for i, h in enumerate(mau["hang"]):
            for col, v in zip("BCDEF", h):
                if v is not None:
                    ws[f"{col}{HANG0 + i}"] = v


def tong_quan(ws, ds_don):
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJKLM", (5, 15, 12, 32, 9, 16, 16, 16, 13, 12, 12, 34, 10)):
        ws.column_dimensions[col].width = w
    gop(ws, "A1:M1", "THEO DÕI ĐƠN HÀNG & CÔNG NỢ KHÁCH HÀNG", b=True, sz=16, col=WHITE, bg=NAVY, al=C)
    ws.row_dimensions[1].height = 34
    gop(ws, "A2:M2", '="Công ty TNHH Thương mại và Công nghiệp Trung Kiên  ·  cập nhật "&TEXT(TODAY(),"dd/mm/yyyy")',
        sz=9, it=True, col="595959", al=C, box=False)
    for r, (t, v) in enumerate((("Khách hàng", KHACH["ten"]), ("Địa chỉ", KHACH["dia_chi"]), ("Mã số thuế", KHACH["mst"]),
                                ("Liên hệ / ĐT", KHACH["lien_he"])), 3):
        gop(ws, f"A{r}:B{r}", t, b=True, bg=SKY, al=LFT)
        gop(ws, f"C{r}:H{r}", v, b=(r == 3), sz=12 if r == 3 else 10, al=LFT)
    ws.row_dimensions[3].height = 22
    a, b = 13, 12 + N_DON
    the = [("A8:C8", "A9:C9", "TỔNG GIÁ TRỊ ĐƠN", f"=SUM(F{a}:F{b})", "1F3864"),
           ("D8:D8", "D9:D9", "ĐÃ THU", f"=SUM(G{a}:G{b})", "375623"),
           ("E8:G8", "E9:G9", "CÒN NỢ", f"=SUM(H{a}:H{b})", "C00000"),
           ("H8:J8", "H9:J9", "NỢ QUÁ HẠN", f'=SUMIF(M{a}:M{b},">0",H{a}:H{b})', "7F0000"),
           ("K8:M8", "K9:M9", "SỐ ĐƠN", f'=COUNTIF(B{a}:B{b},"?*")', "1F3864")]
    for r1, r2, t, ct, mau in the:
        gop(ws, r1, t, b=True, sz=10, col=WHITE, bg=mau, al=C)
        gop(ws, r2, ct, b=True, sz=18, col=mau, bg=PALE, al=C, nf=TIEN if t != "SỐ ĐƠN" else "0")
    ws.row_dimensions[9].height = 36
    dem = [("A10:C10", f'="Chưa giao: "&COUNTIF(I{a}:I{b},"Chưa giao")&"  ·  Giao một phần: "&COUNTIF(I{a}:I{b},"Giao một phần")'),
           ("D10:G10", f'="Đã giao chưa xuất HĐ: "&COUNTIFS(I{a}:I{b},"Đã giao",J{a}:J{b},"<>Đã xuất")'),
           ("H10:J10", f'="Chưa thu đủ: "&COUNTIF(H{a}:H{b},">0")&" đơn"'),
           ("K10:M10", f'="Thiếu giá: "&COUNTIF(L{a}:L{b},"*CHƯA CÓ GIÁ*")&" đơn"')]
    for rng, ct in dem:
        gop(ws, rng, ct, sz=10, col="404040", bg=GREY, al=C)
    gop(ws, "A11:M11", "DANH SÁCH ĐƠN HÀNG  (cột Mã đơn = đúng tên sheet của đơn)", b=True, col=WHITE, bg=BLUE, al=LFT)
    tieu = ("STT", "Mã đơn", "Ngày đơn", "Nội dung đơn", "Số mã hàng", "Tổng tiền", "Đã thu", "Còn nợ", "Giao hàng",
            "Xuất HĐ", "Hạn TT", "Cần làm", "Quá hạn (ngày)")
    for col, t in zip("ABCDEFGHIJKLM", tieu):
        o(ws, f"{col}12", t, b=True, bg=SKY, al=C)
    lay = lambda r, cell, so=False: (f'=IF($B{r}="","",IFERROR(IF(INDIRECT("\'"&$B{r}&"\'!{cell}")="",'      # noqa: E731
                                     f'{"0" if so else chr(34) * 2},INDIRECT("\'"&$B{r}&"\'!{cell}")),"Không thấy sheet"))')
    for r in range(a, b + 1):
        nen = WHITE if (r - a) % 2 == 0 else "FAFCFE"
        o(ws, f"A{r}", f'=IF(B{r}="","",ROW()-{a - 1})', bg=nen, al=C)
        o(ws, f"B{r}", bg=nen, al=C, b=True)
        o(ws, f"C{r}", lay(r, "C5"), bg=nen, al=C, nf=NGAY)
        o(ws, f"D{r}", lay(r, "C6"), bg=nen, al=LFT)
        o(ws, f"E{r}", lay(r, "C9", True), bg=nen, al=C)
        for col, cell in (("F", "K6"), ("G", "K7"), ("H", "K8")):
            o(ws, f"{col}{r}", lay(r, cell, True), bg=nen, al=R, nf=TIEN)
        o(ws, f"I{r}", lay(r, "F4"), bg=nen, al=C)
        o(ws, f"J{r}", lay(r, "F6"), bg=nen, al=C)
        o(ws, f"K{r}", lay(r, "F8"), bg=nen, al=C, nf=NGAY)
        o(ws, f"L{r}", f'=IF(B{r}="","",IFERROR(SUBSTITUTE(INDIRECT("\'"&$B{r}&"\'!A10"),"CẦN LÀM:  ",""),"Không thấy sheet"))',
          bg=nen, al=LFT)
        o(ws, f"M{r}", f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(H{r})),IF(AND(H{r}>0,TODAY()>K{r}),TODAY()-K{r},""),"")', bg=nen, al=C)
    for i, ten in enumerate(ds_don):
        ws[f"B{a + i}"] = ten
    rng = f"A{a}:L{b}"
    for txt, bg, chu in (("CHƯA XUẤT", "F8CBAD", "833C0B"), ("Quá hạn", "FF9999", "7F0000"), ("CHƯA CÓ GIÁ", "FFE699", "7F6000"),
                         ("Không thấy sheet", "D9D9D9", "C00000")):
        to_mau(ws, f"L{a}:L{b}", f'ISNUMBER(SEARCH("{txt}",$L{a}))', bg, chu, True)
    to_mau(ws, f"L{a}:L{b}", f'$L{a}="Hoàn tất"', "C6E0B4", "375623", True)
    to_mau(ws, f"I{a}:I{b}", f'$I{a}="Đã giao"', "C6E0B4")
    to_mau(ws, f"I{a}:I{b}", f'$I{a}="Chưa giao"', "FFE699")
    to_mau(ws, f"J{a}:J{b}", f'$J{a}="Đã xuất"', "C6E0B4")
    to_mau(ws, f"J{a}:J{b}", f'$J{a}="Chưa xuất"', "FFE699")
    to_mau(ws, f"M{a}:M{b}", f'AND(ISNUMBER($M{a}),$M{a}>0)', "FF9999", "7F0000", True)
    ws.conditional_formatting.add(f"H{a}:H{b}", DataBarRule(start_type="num", start_value=0, end_type="max", color="F4B183"))
    ws.freeze_panes = f"A{a}"
    # biểu đồ: tổng tiền và còn nợ từng đơn (20 đơn đầu)
    ch = BarChart()
    ch.type, ch.style, ch.title = "col", 10, "Tổng tiền và còn nợ theo đơn"
    ch.y_axis.title, ch.y_axis.numFmt, ch.height, ch.width = "đồng", "#,##0", 8, 22
    ch.add_data(Reference(ws, min_col=6, min_row=12, max_row=32), titles_from_data=True)
    ch.add_data(Reference(ws, min_col=8, min_row=12, max_row=32), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=13, max_row=32))
    ws.add_chart(ch, f"A{b + 3}")
    ws.page_setup.orientation, ws.page_setup.paperSize = "landscape", ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def huong_dan(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 120
    dong = [("CÁCH DÙNG", True),
            ("1. Mỗi đơn hàng là MỘT sheet. Thêm đơn mới: chuột phải vào sheet MAU_DON > Di chuyển hoặc Sao chép > tích 'Tạo bản sao' > OK,", False),
            ("    rồi đổi tên sheet mới thành mã đơn, ví dụ ĐH-2611-001 (không dùng dấu / \\ ? * [ ] :).", False),
            ("2. Sang sheet TongQuan, gõ đúng tên sheet đó vào ô trống đầu tiên của cột 'Mã đơn'. Mọi số liệu tự chạy lên tổng quan.", False),
            ("3. Trong sheet đơn: gõ mã đơn, ngày, nội dung; từng mã hàng một dòng (mã kho, tên/model, ĐVT, số lượng, đơn giá CHƯA VAT).", False),
            ("    Chọn Giao hàng / Xuất hóa đơn ở ô có danh sách, gõ số hóa đơn, hạn thanh toán. VAT mặc định 8% (ô J5).", False),
            ("4. Mỗi lần khách trả tiền: ghi một dòng ở bảng 'Các lần khách thanh toán' (ngày, số tiền, hình thức). Còn nợ tự trừ.", False),
            ("5. Ô trắng là ô nhập. Ô xanh nhạt là công thức, không gõ vào. Thiếu dòng hàng: chèn dòng ở giữa bảng (không chèn dưới dòng Cộng).", False),
            ("Màu cảnh báo: cam = đã giao chưa xuất HĐ · đỏ = quá hạn · vàng = thiếu giá / chưa giao · xanh lá = hoàn tất.", False),
            ("Quá hạn tính theo ngày máy tính so với 'Hạn thanh toán'. Để trống hạn thì không tính quá hạn.", False)]
    for i, (t, b) in enumerate(dong, 1):
        ws.cell(i, 1, t).font = Font(name=F, size=12 if b else 11, bold=b, color=NAVY if b else "000000")


def tao(ra):
    wb = Workbook()
    tq = wb.active
    tq.title = "TongQuan"
    mau = {"o": {"C4": "ĐH-2610-003", "C6": "Đơn hàng Dương Kính", "C8": "VÍ DỤ thật từ phiếu xuất kho: chưa có đơn giá, ngày"},
           "hang": [("KHH", "Máy bơm hỏa tiễn 6SA35/6", "Cái", 2, None), ("KHH", "Máy bơm chìm 200WQ600-12-37", "Cái", 2, None)]}
    d1 = wb.create_sheet("ĐH-2610-003")
    sheet_don(d1, mau)
    m = wb.create_sheet("MAU_DON")
    sheet_don(m)
    m.sheet_properties.tabColor = "A6A6A6"
    d1.sheet_properties.tabColor = BLUE
    tq.sheet_properties.tabColor = NAVY
    huong_dan(wb.create_sheet("HuongDan"))
    tong_quan(tq, ["ĐH-2610-003"])
    wb.calculation.fullCalcOnLoad = True
    ra = pathlib.Path(ra)
    ra.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ra)
    print(ra)


if __name__ == "__main__":
    tao(sys.argv[1] if len(sys.argv) > 1 else ROOT / "cong-no" / "CongNo_ThuLien.xlsx")
