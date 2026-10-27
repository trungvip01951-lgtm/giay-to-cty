#!/usr/bin/env python3
"""Chèn dấu giáp lai (mọi trang) và dấu + chữ ký Giám đốc (khối "ĐẠI DIỆN BÊN BÁN" ở trang cuối) vào PDF hợp đồng.
Chạy: python3 cong-cu/chen_dau_pdf.py vao.pdf ra.pdf anh_ghep.png anh_dau_sach.png [duong_kinh_dau_mm=40]
Ảnh dấu/chữ ký KHÔNG lưu trong kho: đưa đường dẫn khi chạy. Cần pdftotext (poppler), pypdf, reportlab, Pillow."""
import io
import re
import subprocess
import sys

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

MM = 72 / 25.4
SEAL_PX, GHEP_W_PX, GHEP_H_PX = 575, 998, 780      # dấu tròn và ảnh ghép dau_chu_ky_moi_ghep.png


def tim_khoi_ky(pdf, trang):
    """Tìm tiêu đề 'ĐẠI DIỆN BÊN BÁN' ở cột phải: trả về (tâm x, đáy dòng chú thích '(Ký...' ) tính bằng pt từ đỉnh trang."""
    html = subprocess.run(["pdftotext", "-bbox", "-f", str(trang), "-l", str(trang), pdf, "-"],
                          capture_output=True, text=True, check=True).stdout
    w = [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in
         re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', html)]
    cand = [i for i in range(3, len(w)) if w[i][4] == "BÁN" and [x[4] for x in w[i - 3:i]] == ["ĐẠI", "DIỆN", "BÊN"]]
    if not cand:
        raise SystemExit("Không tìm thấy 'ĐẠI DIỆN BÊN BÁN' ở trang cuối")
    i = cand[-1]
    x0, y0 = w[i - 3][0], w[i - 3][1]
    dong = [x for x in w if abs(x[1] - y0) < 2 and x[0] >= x0]
    x1 = max(x[2] for x in dong)
    note = [x for x in w if x[1] > y0 + 5 and x[1] < y0 + 40 and x[0] >= x0 - 10 and x[2] <= x1 + 40]
    y_note = min((x[1] for x in note), default=y0 + 14)
    y_bot = max(x[3] for x in note if abs(x[1] - y_note) < 2) if note else y0 + 14
    return (x0 + x1) / 2, y_bot


def main(vao, ra, anh_ghep, anh_dau, mm=40.0):
    rd = PdfReader(vao)
    n = len(rd.pages)
    pw, ph = float(rd.pages[0].mediabox.width), float(rd.pages[0].mediabox.height)
    seal_pt = mm * MM
    # --- dấu giáp lai: chia dấu tròn thành n lát dọc, mỗi trang một lát sát mép phải, giữa trang
    seal = Image.open(anh_dau).convert("RGBA")
    p = (seal.width - SEAL_PX) // 2
    seal = seal.crop((p, p, p + SEAL_PX, p + SEAL_PX))
    lat = [seal.crop((round(SEAL_PX * k / n), 0, round(SEAL_PX * (k + 1) / n), SEAL_PX)) for k in range(n)]
    # --- dấu + chữ ký ở trang cuối
    cx, y_note = tim_khoi_ky(vao, n)
    gw, gh = GHEP_W_PX / SEAL_PX * seal_pt, GHEP_H_PX / SEAL_PX * seal_pt
    gx, gy_top = cx - gw / 2, y_note + 4
    assert gy_top + gh < ph - 20, "Khối dấu tràn khỏi trang cuối"
    wr = PdfWriter()
    for k, pg in enumerate(rd.pages):
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(pw, ph))
        lw = seal_pt / n
        c.drawImage(ImageReader(lat[k]), pw - lw, ph / 2 - seal_pt / 2, width=lw, height=seal_pt, mask="auto")
        if k == n - 1:
            c.drawImage(ImageReader(anh_ghep), gx, ph - gy_top - gh, width=gw, height=gh, mask="auto")
        c.save()
        buf.seek(0)
        pg.merge_page(PdfReader(buf).pages[0])
        wr.add_page(pg)
    with open(ra, "wb") as f:
        wr.write(f)
    print(f"{ra}: {n} trang, dấu {mm:g} mm, tâm khối ký x={cx:.0f}pt, đỉnh y={gy_top:.0f}pt")


if __name__ == "__main__":
    if len(sys.argv) < 5:
        raise SystemExit(__doc__)
    main(*sys.argv[1:5], *(float(x) for x in sys.argv[5:6]))
