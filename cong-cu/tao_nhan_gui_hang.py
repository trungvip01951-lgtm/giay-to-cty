#!/usr/bin/env python3
"""Nhãn gửi hàng dán lên kiện (mỗi kiện 1 nhãn khổ A5, 2 nhãn/trang A4), PDF.
Dùng JSON bộ 3 phiếu (khach, hang) + khóa "nhan": {"ten_hang", "so_kien", "ghi_chu"}.
Chạy: python3 cong-cu/tao_nhan_gui_hang.py don.json ra.pdf"""
import json, pathlib, sys
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit

ROOT = pathlib.Path(__file__).resolve().parent.parent
TK = json.load(open(ROOT / "du-lieu/cong_ty_trung_kien.json", encoding="utf-8"))["cong_ty"]
FD = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("S", FD + "LiberationSerif-Regular.ttf"))
pdfmetrics.registerFont(TTFont("SB", FD + "LiberationSerif-Bold.ttf"))


def ve(c, x0, y0, w, h, k, nh, i, n):
    pad = 8 * mm
    c.setLineWidth(1.5); c.rect(x0 + 5 * mm, y0 + 5 * mm, w - 10 * mm, h - 10 * mm)
    x, y, W = x0 + pad + 2 * mm, y0 + h - pad - 8 * mm, w - 2 * pad - 4 * mm

    def dong(t, f="S", s=13, gap=1.3):
        nonlocal y
        for l in simpleSplit(t, f, s, W):
            c.setFont(f, s); c.drawString(x, y, l); y -= s * gap

    def ke():
        nonlocal y
        y -= 2
        c.setLineWidth(0.6); c.line(x, y, x + W, y); y -= 20

    dong("NGƯỜI GỬI:", "SB", 14)
    dong(TK["ten"], "SB", 14)
    dong(f"Địa chỉ: {TK['dia_chi']}", s=13)
    dong(f"Điện thoại: {TK['dien_thoai']}", s=13)
    ke()
    dong("NGƯỜI NHẬN:", "SB", 17)
    dong(k["ten"], "SB", 19)
    dong(f"Điện thoại: {k['dien_thoai']}", "SB", 24)
    dong(f"Địa chỉ nhận hàng: {k.get('dia_diem_giao') or k['dia_chi']}", s=16)
    ke()
    dong(f"Hàng: {nh['ten_hang']}", s=15)
    dong(f"Kiện {i}/{n}", "SB", 20)
    if nh.get("ghi_chu"):
        y -= 4
        dong(f"GHI CHÚ: {nh['ghi_chu']}", "SB", 24)


def main(js, ra):
    don = json.load(open(js, encoding="utf-8"))
    k, nh = don["khach"], don["nhan"]
    n = nh["so_kien"]
    c = canvas.Canvas(ra, pagesize=A4)
    W, H = A4
    for i in range(1, n + 1):
        tren = i % 2 == 1
        ve(c, 0, H / 2 if tren else 0, W, H / 2, k, nh, i, n)
        if tren:
            c.setDash(3, 3); c.setLineWidth(0.4); c.line(0, H / 2, W, H / 2); c.setDash()
        if not tren or i == n:
            c.showPage()
    c.save()


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
