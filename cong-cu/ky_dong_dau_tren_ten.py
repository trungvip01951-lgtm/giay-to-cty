#!/usr/bin/env python3
"""Ký và đóng dấu vào PDF đã in sẵn tên người ký (phiếu xuất kho, phiếu bảo hành, biên bản bàn giao... từ phần mềm báo giá).
Mỗi trang: tìm dòng tên đứng riêng (mặc định "Nguyễn Việt Trung", bỏ qua dòng "Người đại diện: ... Chức vụ"), đặt chữ ký + dấu tròn
(ảnh dau-chu-ky/dau_tron_chu_ky_khong_ten.png, không có dấu chức danh vì tên đã in sẵn) vào khoảng trống giữa tiêu đề khối ký và tên.
Chạy: python3 cong-cu/ky_dong_dau_tren_ten.py vao.pdf ra.pdf anh_dau_ky.png [duong_kinh_dau_mm=40] ["Nguyễn Việt Trung"]"""
import io
import re
import subprocess
import sys

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

MM = 72 / 25.4
SEAL_PX = 575                      # đường kính dấu trong ảnh ghép


def tim_cho_ky(pdf, trang, ten):
    html = subprocess.run(["pdftotext", "-bbox", "-f", str(trang), "-l", str(trang), pdf, "-"],
                          capture_output=True, text=True, check=True).stdout
    w = [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in
         re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', html)]
    tu, k = ten.split(), len(ten.split())
    ra = []
    for i in range(len(w) - k + 1):
        if [x[4] for x in w[i:i + k]] != tu:
            continue
        cung_dong = [x for x in w if abs(x[1] - w[i][1]) < 2]
        if any(x[4].startswith(("Chức", "diện:", "Đại")) for x in cung_dong):
            continue                                   # dòng "Người đại diện: ... Chức vụ" ở phần thông tin
        x0, x1, top = w[i][0], w[i + k - 1][2], w[i][1]
        tren = [x for x in w if x[3] < top - 2 and x[3] > top - 90 and x[2] > x0 - 40 and x[0] < x1 + 40]
        day_tren = max((x[3] for x in tren), default=top - 50)
        ra.append(((x0 + x1) / 2, day_tren, top))
    return ra


def main(vao, ra, anh, mm=40.0, ten="Nguyễn Việt Trung"):
    rd = PdfReader(vao)
    im = Image.open(anh)
    sc = mm * MM / SEAL_PX
    gw, gh = im.width * sc, im.height * sc
    wr = PdfWriter()
    for n, pg in enumerate(rd.pages, 1):
        pw, ph = float(pg.mediabox.width), float(pg.mediabox.height)
        cho = tim_cho_ky(vao, n, ten)
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(pw, ph))
        for cx, day_tren, top in cho:
            cy = (day_tren + top) / 2                  # giữa khoảng trống ký (tính từ đỉnh trang)
            x = min(max(cx - gw / 2, 15), pw - gw - 15)
            c.drawImage(ImageReader(im), x, ph - cy - gh / 2, width=gw, height=gh, mask="auto")
            print(f"trang {n}: ký tại x={cx:.0f}pt, y={cy:.0f}pt")
        if not cho:
            print(f"trang {n}: không thấy tên '{ten}', không ký")
        c.save()
        buf.seek(0)
        pg.merge_page(PdfReader(buf).pages[0])
        wr.add_page(pg)
    with open(ra, "wb") as f:
        wr.write(f)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:4], *(float(x) for x in sys.argv[4:5]), *sys.argv[5:6])
