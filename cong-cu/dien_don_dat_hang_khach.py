#!/usr/bin/env python3
"""Điền thông tin Bên B (Trung Kiên) và hàng hóa vào MẪU ĐƠN ĐẶT HÀNG CỦA KHÁCH (Word), giữ nguyên bố cục của khách.
Dùng cho mẫu có các dòng "BÊN BÁN (Bên B):", "Địa chỉ:", "Mã số thuế:", "Đại diện:", "Tài khoản:" sau phần Bên A,
bảng hàng 6 cột (STT, Tên, ĐVT, SL, Đơn giá, Thành tiền) + 3 dòng tổng, và dòng "Số tiền viết bằng chữ".
Chạy: python3 cong-cu/dien_don_dat_hang_khach.py don.json"""
import copy
import json
import pathlib
import sys

import docx

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402
from tao_denghi_fata import doc_so, so  # noqa: E402

ROOT = m.ROOT
TK, DD, BANK = m.TK, m.DD, m.BANK


def dat_dots(p, gia_tri):
    """Thay run chứa dấu chấm lửng đầu tiên bằng giá trị, xóa các run chấm lửng còn lại."""
    xong = False
    for r in p.runs:
        if "…" in r.text or r.text.strip(".") == "" and r.text.count(".") > 3:
            r.text = "" if xong else gia_tri
            xong = True
    if not xong:
        raise ValueError("không thấy chỗ trống: " + p.text)


def main(f):
    don = json.load(open(f, encoding="utf-8"))
    tien = sum(h["sl"] * h["don_gia"] for h in don["hang"])
    vat = tien * don["vat_phan_tram"] // 100
    tong = tien + vat
    k = don["kiem_tra"]
    assert (tien, vat, tong) == (k["tong_tien_hang"], k["vat"], k["tong"]), (tien, vat, tong)
    d = docx.Document(ROOT / don["mau_khach"])
    ps = d.paragraphs
    i = next(n for n, p in enumerate(ps) if p.text.startswith("BÊN BÁN"))
    dat_dots(ps[i], TK["ten"])
    for p in ps[i + 1:i + 6]:
        t = p.text
        if t.startswith("Địa chỉ"):
            dat_dots(p, TK["dia_chi"])
        elif t.startswith("Mã số thuế"):
            dat_dots(p, TK["mst"])
        elif t.startswith("Đại diện"):
            runs = [r for r in p.runs if "…" in r.text]
            runs[0].text, runs[1].text = DD["ho_ten"], DD["chuc_vu"]
        elif t.startswith("Tài khoản"):
            dat_dots(p, f"{BANK['so_tk']} – Tại {BANK['ten']}")
    # bảng hàng
    tb = next(t for t in d.tables if t.rows[0].cells[0].text.strip() == "STT")
    mau_dong = tb.rows[1]._tr
    for n, h in enumerate(don["hang"]):
        tr = mau_dong if n == 0 else copy.deepcopy(mau_dong)
        if n:
            tb.rows[n]._tr.addnext(tr)
        cells = tb.rows[n + 1].cells
        vals = [str(n + 1), None, h["dvt"], str(h["sl"]), so(h["don_gia"]), so(h["sl"] * h["don_gia"])]
        for c, v in zip(cells, vals):
            p = c.paragraphs[0]
            for r in p.runs[1:]:
                r.text = ""
            if v is not None:
                p.runs[0].text = v
        # ô tên hàng: dòng đầu đậm, các dòng sau thường
        p = cells[1].paragraphs[0]
        p.runs[0].text, p.runs[0].bold = h["ten"][0], True
        for dong in h["ten"][1:]:
            p.runs[-1].add_break()
            r = p.add_run(dong)
            r.bold = False
            r.font.name, r.font.size = p.runs[0].font.name, p.runs[0].font.size
    so_dong = len(don["hang"])
    for off, v in ((1, tien), (2, vat), (3, tong)):
        c = tb.rows[so_dong + off].cells[-1].paragraphs[0]
        c.runs[0].text = so(v)
        for r in c.runs[1:]:
            r.text = ""
    p = next(p for p in d.paragraphs if p.text.startswith("Số tiền viết"))
    dat_dots(p, doc_so(tong))
    if don.get("bo_dong_trong"):
        # bỏ dòng trống đệm và bảng trống cuối mẫu để khối ký còn chỗ đóng dấu (không đổi nội dung của khách)
        for p in list(d.paragraphs):
            if not p.text.strip() and not p._p.xpath(".//w:drawing|.//w:br") and p._p.getparent().tag.endswith("}body"):
                if p._p.getnext() is not None and not p._p.getnext().tag.endswith("}sectPr"):
                    p._p.getparent().remove(p._p)
        for t in list(d.tables):
            if all(not c.text.strip() for r in t.rows for c in r.cells):
                t._tbl.getparent().remove(t._tbl)
    ra = ROOT / don["ra"]
    d.save(ra)
    print(ra, so(tien), so(vat), so(tong), doc_so(tong))


if __name__ == "__main__":
    main(sys.argv[1])
