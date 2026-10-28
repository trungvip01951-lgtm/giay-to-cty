#!/usr/bin/env python3
"""Làm bản PDF có dấu trông như mới quét (scan): nền giấy hơi ngả, mờ nhẹ, chữ nhạt đi một chút, nhiễu mịn, lệch trang rất nhẹ.
Chạy: python3 cong-cu/lam_ban_scan.py vao.pdf ra.pdf [dpi=200]      (cần pdftoppm, Pillow, numpy)
Kết quả là PDF dạng ảnh (không chọn/copy được chữ), dùng để gửi khách; giữ riêng bản Word/PDF gốc để sửa."""
import pathlib
import subprocess
import sys
import tempfile

import numpy as np
from PIL import Image, ImageFilter


def scan_trang(im, rng):
    a = np.asarray(im.convert("RGB")).astype(np.float32)
    h, w = a.shape[:2]
    a = np.asarray(Image.fromarray(a.astype(np.uint8)).filter(ImageFilter.GaussianBlur(0.75))).astype(np.float32)
    a = 255 - (255 - a) * 0.88                                  # mực nhạt đi
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    nx, ny = xx / w - 0.5, yy / h - 0.5
    ven = 1 - 0.07 * np.clip((nx ** 2 + ny ** 2) * 2.2, 0, 1)    # tối nhẹ ở mép
    anh = 1 - 0.035 * (nx + 0.5) + 0.02 * (ny + 0.5)              # lệch sáng đèn quét
    giay = np.array([0.975, 0.97, 0.945], np.float32)            # giấy hơi ngả ấm
    a = a * giay * (ven * anh)[..., None]
    a += rng.normal(0, 3.2, (h, w))[..., None]                   # nhiễu hạt
    im2 = Image.fromarray(np.clip(a, 0, 255).astype(np.uint8))
    im2 = im2.rotate(rng.uniform(-0.35, 0.35), resample=Image.BICUBIC, fillcolor=(228, 226, 220))
    return im2


def main(vao, ra, dpi=200):
    rng = np.random.default_rng(2026)
    with tempfile.TemporaryDirectory() as t:
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", vao, f"{t}/p"], check=True)
        trang = [scan_trang(Image.open(p), rng) for p in sorted(pathlib.Path(t).glob("p-*.png"))]
    trang[0].save(ra, "PDF", resolution=dpi, save_all=True, append_images=trang[1:], quality=82)
    print(ra, len(trang), "trang")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 200)
