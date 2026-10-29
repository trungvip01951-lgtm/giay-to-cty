#!/usr/bin/env python3
"""Tạo file Excel theo dõi công nợ Trung Kiên: đơn hàng -> nhiều mã hàng, giao hàng, xuất hóa đơn, thu tiền.
Chạy: python3 cong-cu/tao_bang_cong_no.py [ra.xlsx]   (mặc định cong-no/CongNo_TrungKien.xlsx)
CHỈ dùng để tạo file mẫu ban đầu. Khi anh đã nhập dữ liệu thật thì file .xlsx của anh là bản gốc, đừng chạy lại đè lên."""
import datetime as dt
import pathlib
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = pathlib.Path(__file__).resolve().parent.parent
N_DON, N_CT, N_THU, N_KH = 500, 2000, 2000, 200          # số dòng đã dựng sẵn công thức

XANH = PatternFill("solid", fgColor="1F4E78")
XAM = PatternFill("solid", fgColor="EDEDED")           # ô công thức (không nhập)
VANG = PatternFill("solid", fgColor="FFF8DC")          # ô ví dụ
VIEN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=VIEN, right=VIEN, top=VIEN, bottom=VIEN)
TIEN, NGAY = '#,##0;[Red]-#,##0;""', "dd/mm/yyyy"


def dau_bang(ws, tieu_de, rong):
    for i, (t, w) in enumerate(zip(tieu_de, rong), 1):
        c = ws.cell(1, i, t)
        c.font, c.fill = Font(bold=True, color="FFFFFF", name="Arial", size=10), XANH
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{L(len(tieu_de))}{ws.max_row if ws.max_row > 1 else 2}"


def dinh_dang(ws, hang_cuoi, cot_tien=(), cot_ngay=(), cot_cong_thuc=(), cot_so=()):
    for r in range(2, hang_cuoi + 1):
        for c in range(1, ws.max_column + 1):
            x = ws.cell(r, c)
            x.border, x.font = BORDER, Font(name="Arial", size=10)
            if c in cot_tien:
                x.number_format = TIEN
            if c in cot_ngay:
                x.number_format = NGAY
            if c in cot_so:
                x.number_format = "#,##0.##"
            if c in cot_cong_thuc:
                x.fill = XAM


def danh_sach(ws, rng, items=None, nguon=None):
    dv = DataValidation(type="list", formula1=(f'"{",".join(items)}"' if items else nguon), allow_blank=True)
    dv.error, dv.errorTitle, dv.showErrorMessage = "Chọn một giá trị trong danh sách", "Sai giá trị", True
    ws.add_data_validation(dv)
    dv.add(rng)


def tao(ra):
    wb = Workbook()
    th, dh, ct, tt, kh, hd = (wb.active, wb.create_sheet("DonHang"), wb.create_sheet("ChiTiet"),
                              wb.create_sheet("ThuTien"), wb.create_sheet("KhachHang"), wb.create_sheet("HuongDan"))
    th.title = "TongHop"

    # ---------------- KhachHang
    kh.append(["Tên khách hàng", "MST", "Địa chỉ", "Điện thoại / kế toán", "Số đơn", "Tổng phải thu", "Đã thu", "Còn nợ",
               "Còn nợ quá hạn", "Ghi chú"])
    khach = [
        ("CÔNG TY CỔ PHẦN CÔNG NGHỆ FATA", "2601102666", "Lô CN01, KCN Thụy Vân, xã Hy Cương, tỉnh Phú Thọ", "0912 455 310", ""),
        ("CÔNG TY TNHH HOÀNG VIÊN QUẢNG BÁ", "0100143403", "Số 96, đường Tô Ngọc Vân, Phường Tây Hồ, TP Hà Nội", "(84-24)37180888", ""),
        ("CÔNG TY TNHH TIẾN ĐẠI", "0200447853", "Thôn Câu Đông, xã An Quang, TP Hải Phòng", "Ms Luyến 0915.264.659; Ms Hương 0915.144.219", ""),
        ("Công ty TNHH vật tư thiết bị Thu Liên", "", "Hà Nội", "", ""),
    ]
    for i, k in enumerate(khach, 2):
        for c, v in enumerate(k[:4], 1):
            kh.cell(i, c, v)
        kh.cell(i, 10, k[4])
    for r in range(2, N_KH + 1):
        kh[f"E{r}"] = f'=IF(A{r}="","",COUNTIFS(DonHang!$C$2:$C${N_DON},A{r}))'
        kh[f"F{r}"] = f'=IF(A{r}="","",SUMIFS(DonHang!$N$2:$N${N_DON},DonHang!$C$2:$C${N_DON},A{r}))'
        kh[f"G{r}"] = f'=IF(A{r}="","",SUMIFS(DonHang!$O$2:$O${N_DON},DonHang!$C$2:$C${N_DON},A{r}))'
        kh[f"H{r}"] = f'=IF(A{r}="","",F{r}-G{r})'
        kh[f"I{r}"] = f'=IF(A{r}="","",SUMIFS(DonHang!$P$2:$P${N_DON},DonHang!$C$2:$C${N_DON},A{r},DonHang!$R$2:$R${N_DON},">0"))'
    dau_bang(kh, ["Tên khách hàng", "MST", "Địa chỉ", "Điện thoại / kế toán", "Số đơn", "Tổng phải thu", "Đã thu", "Còn nợ",
                  "Còn nợ quá hạn", "Ghi chú"], [38, 14, 44, 30, 8, 16, 16, 16, 16, 24])
    kh.auto_filter.ref = f"A1:J{N_KH}"
    dinh_dang(kh, N_KH, cot_tien=(6, 7, 8, 9), cot_cong_thuc=(5, 6, 7, 8, 9))

    # ---------------- DonHang
    tieu = ["Mã đơn", "Ngày đơn", "Khách hàng", "MST", "Nội dung đơn / Số báo giá, hợp đồng", "Giao hàng", "Ngày giao",
            "Xuất hóa đơn", "Số hóa đơn", "Ngày hóa đơn", "Tiền hàng (chưa VAT)", "VAT %\n(trống = 8%)", "Tiền VAT",
            "Tổng phải thu", "Đã thu", "Còn nợ", "Hạn thanh toán", "Quá hạn (ngày)", "Tình trạng thu tiền", "Cần làm", "Ghi chú"]
    rong = [13, 11, 36, 12, 34, 14, 11, 13, 12, 11, 16, 9, 14, 16, 16, 16, 12, 10, 16, 30, 28]
    mau = [
        ("ĐH-2610-001", dt.date(2026, 10, 5), khach[0][0], "Hợp đồng HĐKT102026-TK-H + Phụ lục 01", "Đã giao", None,
         "Đã xuất", "00002771", dt.date(2026, 10, 17), None, dt.date(2026, 11, 17), "VÍ DỤ thật: hạn 17/11/2026 theo công nợ của FATA"),
        ("ĐH-2610-002", dt.date(2026, 10, 27), khach[1][0], "Báo giá BG-28102026-006", None, None, None, None, None, None, None,
         "VÍ DỤ: chưa cập nhật giao hàng, xuất HĐ, thu tiền"),
        ("ĐH-2609-001", None, khach[2][0], "Hóa đơn 00002469 (hợp đồng + biên bản giao nhận đang làm)", None, None,
         "Đã xuất", "00002469", dt.date(2026, 9, 18), None, None, "VÍ DỤ: chưa cập nhật giao hàng và số đã thu"),
        ("ĐH-2610-003", None, khach[3][0], "Đơn hàng Dương Kính (phiếu xuất kho)", None, None, None, None, None, None, None,
         "VÍ DỤ: chưa có đơn giá"),
    ]
    for i, m in enumerate(mau, 2):
        a, b, c, e, f, g, h, i_, j, l, q, u = m
        for col, v in ((1, a), (2, b), (3, c), (5, e), (6, f), (7, g), (8, h), (9, i_), (10, j), (12, l), (17, q), (21, u)):
            if v is not None:
                dh.cell(i, col, v)
    for r in range(2, N_DON + 1):
        dh[f"D{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},KhachHang!$A$2:$B${N_KH},2,FALSE)&"",""))'
        dh[f"K{r}"] = f'=IF(A{r}="","",SUMIFS(ChiTiet!$I$2:$I${N_CT},ChiTiet!$A$2:$A${N_CT},A{r}))'
        dh[f"M{r}"] = f'=IF(A{r}="","",ROUND(K{r}*IF(L{r}="",0.08,L{r}),0))'
        dh[f"N{r}"] = f'=IF(A{r}="","",K{r}+M{r})'
        dh[f"O{r}"] = f'=IF(A{r}="","",SUMIFS(ThuTien!$D$2:$D${N_THU},ThuTien!$B$2:$B${N_THU},A{r}))'
        dh[f"P{r}"] = f'=IF(A{r}="","",N{r}-O{r})'
        dh[f"R{r}"] = f'=IF(A{r}="","",IF(AND(Q{r}<>"",P{r}>0,TODAY()>Q{r}),TODAY()-Q{r},0))'
        dh[f"S{r}"] = (f'=IF(A{r}="","",IF(N{r}=0,"Chưa có tiền hàng",IF(P{r}<0,"Thu thừa",IF(P{r}=0,"Đã thu đủ",'
                       f'IF(O{r}>0,"Thu một phần","Chưa thu")))))')
        dh[f"T{r}"] = (f'=IF(A{r}="","",IF(AND(F{r}="Đã giao",H{r}<>"Đã xuất"),"Đã giao - CHƯA XUẤT HĐ",'
                       f'IF(AND(H{r}="Đã xuất",F{r}="Chưa giao"),"Đã xuất HĐ - chưa giao",'
                       f'IF(F{r}="","Chưa cập nhật giao hàng",IF(R{r}>0,"Quá hạn "&R{r}&" ngày - đòi tiền",'
                       f'IF(P{r}>0,"Chờ thu tiền",IF(P{r}<0,"Kiểm tra thu thừa",'
                       f'IF(AND(F{r}="Đã giao",H{r}="Đã xuất"),"Hoàn tất","Theo dõi"))))))))')
    dau_bang(dh, tieu, rong)
    dh.auto_filter.ref = f"A1:U{N_DON}"
    dinh_dang(dh, N_DON, cot_tien=(11, 13, 14, 15, 16), cot_ngay=(2, 7, 10, 17),
              cot_cong_thuc=(4, 11, 13, 14, 15, 16, 18, 19, 20))
    for r in range(2, N_DON + 1):
        dh[f"L{r}"].number_format = "0%"
    danh_sach(dh, f"F2:F{N_DON}", ["Chưa giao", "Giao một phần", "Đã giao"])
    danh_sach(dh, f"H2:H{N_DON}", ["Chưa xuất", "Đã xuất"])
    danh_sach(dh, f"C2:C{N_DON}", nguon=f"=KhachHang!$A$2:$A${N_KH}")
    rng = f"A2:U{N_DON}"
    for cong_thuc, mau_nen in ((f'ISNUMBER(SEARCH("CHƯA XUẤT HĐ",$T2))', "F8CBAD"), (f'ISNUMBER(SEARCH("Quá hạn",$T2))', "FF9999"),
                               (f'ISNUMBER(SEARCH("chưa giao",$T2))', "FFE699"), ('$T2="Hoàn tất"', "C6E0B4")):
        dh.conditional_formatting.add(rng, FormulaRule(formula=[cong_thuc], fill=PatternFill("solid", bgColor=mau_nen, fgColor=mau_nen)))

    # ---------------- ChiTiet
    tieu_ct = ["Mã đơn", "Khách hàng", "Mã kho", "Mã hàng / Tên hàng, model", "ĐVT", "Số lượng", "Đơn giá (chưa VAT)", "Ghi chú", "Thành tiền"]
    # cột I là Thành tiền (công thức) để cột nhập liền nhau
    ct_mau = [
        ("ĐH-2610-001", "KHH", "Máy bơm nước ly tâm trục ngang AQUARIS CM65-160/15", "cái", 4, 28200000),
        ("ĐH-2610-001", "KHH", "Máy bơm nước trục rời AQUARIS AQR65-200/15", "cái", 4, 39200000),
        ("ĐH-2610-002", "KHH", "Bơm chìm nước thải PEROTAC 65WC25-15-2.2", "cái", 2, 5800000),
        ("ĐH-2609-001", "KHH", "Máy bơm nước SVm-1500F", "cái", 2, 4500000),
        ("ĐH-2609-001", "KHH", "Ống cốt dù phủ nhựa PVC DK65mm", "mét", 100, 17000),
        ("ĐH-2610-003", "KHH", "Máy bơm hỏa tiễn 6SA35/6", "cái", 2, None),
        ("ĐH-2610-003", "KHH", "Máy bơm chìm 200WQ600-12-37", "cái", 2, None),
    ]
    for i, (ma, kho, ten, dvt, sl, gia) in enumerate(ct_mau, 2):
        ct.cell(i, 1, ma), ct.cell(i, 3, kho), ct.cell(i, 4, ten), ct.cell(i, 5, dvt), ct.cell(i, 6, sl)
        if gia is not None:
            ct.cell(i, 7, gia)
    for r in range(2, N_CT + 1):
        ct[f"B{r}"] = f'=IF(A{r}="","",IFERROR(VLOOKUP(A{r},DonHang!$A$2:$C${N_DON},3,FALSE),"mã đơn chưa có"))'
        ct[f"I{r}"] = f'=IF(OR(F{r}="",G{r}=""),"",ROUND(F{r}*G{r},0))'
    dau_bang(ct, tieu_ct, [13, 34, 8, 48, 8, 10, 16, 24, 16])
    ct.auto_filter.ref = f"A1:I{N_CT}"
    dinh_dang(ct, N_CT, cot_tien=(7, 9), cot_so=(6,), cot_cong_thuc=(2, 9))
    danh_sach(ct, f"A2:A{N_CT}", nguon=f"=DonHang!$A$2:$A${N_DON}")

    # ---------------- ThuTien
    tt.append(["Ngày thu", "Mã đơn", "Khách hàng", "Số tiền thu", "Hình thức", "Số chứng từ / nội dung", "Ghi chú"])
    tt.cell(2, 2, "ĐH-2610-001"), tt.cell(2, 4, 130356000), tt.cell(2, 5, "Chuyển khoản")
    tt.cell(2, 6, "50% hợp đồng gốc (theo công nợ FATA)"), tt.cell(2, 7, "VÍ DỤ thật: ngày thu chưa rõ, anh điền")
    for r in range(2, N_THU + 1):
        tt[f"C{r}"] = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},DonHang!$A$2:$C${N_DON},3,FALSE),"mã đơn chưa có"))'
    dau_bang(tt, ["Ngày thu", "Mã đơn", "Khách hàng", "Số tiền thu", "Hình thức", "Số chứng từ / nội dung", "Ghi chú"],
             [12, 13, 36, 16, 14, 36, 30])
    tt.auto_filter.ref = f"A1:G{N_THU}"
    dinh_dang(tt, N_THU, cot_tien=(4,), cot_ngay=(1,), cot_cong_thuc=(3,))
    danh_sach(tt, f"B2:B{N_THU}", nguon=f"=DonHang!$A$2:$A${N_DON}")
    danh_sach(tt, f"E2:E{N_THU}", ["Chuyển khoản", "Tiền mặt", "Khác"])

    # ---------------- TongHop
    th["A1"], th["A1"].font = "TỔNG HỢP CÔNG NỢ - TRUNG KIÊN", Font(bold=True, size=14, name="Arial")
    th["A2"] = '="Cập nhật theo ngày máy: "&TEXT(TODAY(),"dd/mm/yyyy")'
    chi_so = [
        ("Tổng số đơn hàng", f'=COUNTA(DonHang!A2:A{N_DON})', "0"),
        ("Tổng phải thu (đã gồm VAT)", f"=SUM(DonHang!N2:N{N_DON})", TIEN),
        ("Đã thu", f"=SUM(DonHang!O2:O{N_DON})", TIEN),
        ("CÒN NỢ (khách còn phải trả)", f"=SUM(DonHang!P2:P{N_DON})", TIEN),
        ("Trong đó quá hạn", f'=SUMIFS(DonHang!P2:P{N_DON},DonHang!R2:R{N_DON},">0")', TIEN),
        ("Đơn đã giao nhưng CHƯA xuất HĐ", f'=SUMPRODUCT((DonHang!A2:A{N_DON}<>"")*(DonHang!F2:F{N_DON}="Đã giao")*(DonHang!H2:H{N_DON}<>"Đã xuất"))', "0"),
        ("Đơn đã xuất HĐ nhưng chưa giao", f'=SUMPRODUCT((DonHang!A2:A{N_DON}<>"")*(DonHang!H2:H{N_DON}="Đã xuất")*(DonHang!F2:F{N_DON}="Chưa giao"))', "0"),
        ("Đơn chưa cập nhật giao hàng", f'=SUMPRODUCT((DonHang!A2:A{N_DON}<>"")*(DonHang!F2:F{N_DON}=""))', "0"),
        ("Đơn chưa thu đủ tiền", f'=COUNTIF(DonHang!P2:P{N_DON},">0")', "0"),
    ]
    for i, (ten, ct_, nf) in enumerate(chi_so, 4):
        th.cell(i, 1, ten).font = Font(name="Arial", size=11, bold="CÒN NỢ" in ten)
        c = th.cell(i, 2, ct_)
        c.number_format, c.font, c.alignment = nf, Font(name="Arial", size=12, bold=True), Alignment(horizontal="right")
        th.cell(i, 1).border = th.cell(i, 2).border = BORDER
    th.column_dimensions["A"].width, th.column_dimensions["B"].width = 44, 22
    th["A14"] = "Xem chi tiết: sheet DonHang (cột 'Cần làm' tô màu), KhachHang (công nợ từng khách)."
    th["A14"].font = Font(italic=True, name="Arial", size=10)

    # ---------------- HuongDan
    dong = [
        "CÁCH DÙNG (5 bước)",
        "1. KhachHang: thêm khách mới ở dòng trống đầu tiên (tên, MST, địa chỉ). Tên này sẽ hiện trong danh sách chọn ở DonHang.",
        "2. DonHang: mỗi đơn một dòng, gõ Mã đơn (ví dụ ĐH-2610-001: ĐH - năm tháng - số thứ tự), ngày, chọn khách, ghi số báo giá/hợp đồng.",
        "3. ChiTiet: mỗi mã hàng một dòng. Chọn Mã đơn, gõ mã hàng/model, ĐVT, số lượng, đơn giá CHƯA VAT. Thành tiền tự tính. Một đơn có bao nhiêu mã hàng cũng được.",
        "4. DonHang: chọn 'Giao hàng' (Chưa giao / Giao một phần / Đã giao) và 'Xuất hóa đơn' (Chưa xuất / Đã xuất), gõ số hóa đơn, ngày, hạn thanh toán.",
        "5. ThuTien: mỗi lần khách trả tiền một dòng (ngày, mã đơn, số tiền). Một đơn trả nhiều lần cũng được. Còn nợ tự tính.",
        "",
        "Ô màu xám là công thức, KHÔNG gõ vào. Ô trắng là chỗ nhập. Dòng vàng nhạt/ví dụ: xóa hoặc ghi đè khi dùng thật.",
        "VAT mặc định 8% (ô VAT % để trống). Đơn có VAT khác thì gõ vào, ví dụ 10%.",
        "Cột 'Cần làm' tự báo: Đã giao - CHƯA XUẤT HĐ (đỏ cam), Quá hạn (đỏ), Chưa giao (vàng), Hoàn tất (xanh lá).",
        "Quá hạn tính từ 'Hạn thanh toán' tới hôm nay (theo ngày máy tính). Để trống hạn thì không tính quá hạn.",
        "Dựng sẵn công thức cho 500 đơn, 2.000 dòng hàng, 2.000 lần thu, 200 khách. Cần thêm: kéo công thức của dòng cuối xuống.",
        "Số tiền VAT làm tròn đồng theo từng đơn; nếu hóa đơn lệch vài đồng so với bảng thì ghi chú ở cột Ghi chú.",
    ]
    for i, t in enumerate(dong, 1):
        hd.cell(i, 1, t).font = Font(name="Arial", size=11, bold=(i == 1))
    hd.column_dimensions["A"].width = 150
    ra = pathlib.Path(ra)
    ra.parent.mkdir(parents=True, exist_ok=True)
    wb.calculation.fullCalcOnLoad = True
    wb.save(ra)
    print(ra)


if __name__ == "__main__":
    tao(sys.argv[1] if len(sys.argv) > 1 else ROOT / "cong-no" / "CongNo_TrungKien.xlsx")
