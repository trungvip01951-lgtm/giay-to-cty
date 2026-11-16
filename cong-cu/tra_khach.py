#!/usr/bin/env python3
"""Tra thông tin khách hàng trong du-lieu/khach_hang.json theo MST, mã hoặc từ khóa tên.
Chạy: python3 cong-cu/tra_khach.py 0111541621   |   python3 cong-cu/tra_khach.py "hà tây"
In khối JSON để chép vào 'khach' của tệp đơn; báo các trường còn thiếu."""
import json, pathlib, sys, unicodedata

ROOT = pathlib.Path(__file__).resolve().parent.parent


def bo_dau(s):
    s = unicodedata.normalize("NFD", s.replace("đ", "d").replace("Đ", "D"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def tim(q):
    db = json.load(open(ROOT / "du-lieu/khach_hang.json", encoding="utf-8"))
    tat_ca = db["khach_hang"] + db.get("nha_cung_cap", [])
    q2 = bo_dau(q).replace(" ", "")
    return [k for k in tat_ca if q.strip() == k.get("mst") or q2 in bo_dau(k.get("ma", "") + k.get("ten", "")).replace(" ", "")]


if __name__ == "__main__":
    kq = tim(" ".join(sys.argv[1:]))
    if not kq:
        sys.exit("Chưa có khách này trong du-lieu/khach_hang.json")
    for k in kq:
        print(json.dumps(k, ensure_ascii=False, indent=2))
        thieu = [f for f in ("ten", "mst", "dia_chi", "dien_thoai", "stk", "ngan_hang", "dai_dien") if k.get(f) == ""]
        if thieu:
            print("Còn thiếu:", ", ".join(thieu))
