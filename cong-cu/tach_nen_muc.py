#!/usr/bin/env python3
"""Tách dấu đỏ / chữ ký xanh khỏi ảnh chụp trên giấy trắng thành PNG nền trong suốt.
Chạy: python3 cong-cu/tach_nen_muc.py anh.jpg ra.png do|xanh [tam_x,tam_y,ban_kinh]
  do   : mực đỏ (dấu); xanh: mực xanh (chữ ký). Có tam_x,tam_y,ban_kinh thì chỉ lấy trong hình tròn đó
         (dùng khi trong ảnh có dấu khác ở góc). Chụp thẳng, đủ sáng, dấu và chữ ký ký riêng trên giấy trắng."""
import sys

import numpy as np
from PIL import Image, ImageFilter

MAU = {"do": (200, 25, 45), "xanh": (20, 40, 160)}


def tach(anh, ra, mau, vong=None):
    a = np.asarray(Image.open(anh).convert("RGB")).astype(float)
    R, G, B = a[..., 0], a[..., 1], a[..., 2]
    if mau == "do":
        al = np.clip((R - np.maximum(G, B) - 8) / 55, 0, 1) ** 0.85
    else:
        al = np.clip((B - np.maximum(R, G * 0.9) - 6) / 40, 0, 1)
        im = Image.fromarray((al * 255).astype(np.uint8), "L").filter(ImageFilter.MaxFilter(3)).filter(ImageFilter.MinFilter(3))
        al = np.asarray(im).astype(float) / 255
    if vong:
        cx, cy, r = vong
        yy, xx = np.mgrid[0:al.shape[0], 0:al.shape[1]]
        al = al * (np.hypot(xx - cx, yy - cy) < r)
    ys, xs = np.where(al > 0.15)
    p = 12
    y0, y1, x0, x1 = ys.min() - p, ys.max() + p, xs.min() - p, xs.max() + p
    out = np.zeros((y1 - y0, x1 - x0, 4), np.uint8)
    out[..., :3] = MAU[mau]
    out[..., 3] = (al[y0:y1, x0:x1] * 255).astype(np.uint8)
    Image.fromarray(out, "RGBA").save(ra)
    print(ra, out.shape[1], "x", out.shape[0])


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    v = tuple(float(x) for x in sys.argv[4].split(",")) if len(sys.argv) > 4 else None
    tach(sys.argv[1], sys.argv[2], sys.argv[3], v)
