#!/usr/bin/env python3
"""Phụ lục hợp đồng nhập khẩu (tiếng Anh) bổ sung cước vận chuyển, đổi điều kiện giá FOB -> CIF.
Trung Kiên là bên mua (THE BUYER). Chạy: python3 cong-cu/tao_phu_luc_nk.py phuluc.json"""
import json, pathlib, sys
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402

ONES = "zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen".split()
TENS = "_ _ twenty thirty forty fifty sixty seventy eighty ninety".split()


def chu(n):
    def ba(x):
        out = []
        if x >= 100:
            out.append(ONES[x // 100] + " hundred"); x %= 100
        if x >= 20:
            out.append(TENS[x // 10] + ("-" + ONES[x % 10] if x % 10 else ""))
        elif x:
            out.append(ONES[x])
        return " ".join(out)
    parts = []
    if n >= 1000:
        parts.append(ba(n // 1000) + " thousand")
    if n % 1000:
        parts.append(ba(n % 1000))
    s = " ".join(parts) + " US dollars only"
    return s[0].upper() + s[1:]


assert chu(75727) == "Seventy-five thousand seven hundred twenty-seven US dollars only"


def usd(x):
    return f"USD {x:,.2f}"


def build(pl, js):
    tong = pl["gia_tri_hang"] + pl["cuoc"]
    assert tong == pl["kiem_tra_tong"], tong
    b, s = pl["ben_mua"], pl["ben_ban"]
    d = m.new_doc()
    d.styles["Normal"].font.size = m.Pt(11)   # tiếng Anh, gọn 1 trang để đủ chỗ dấu bên mua
    m.para(d, "SALES CONTRACT APPENDIX", bold=True, align=m.C, size=15, after=2)
    m.para(d, f"No.: {pl['so_pl']}", bold=True, align=m.C, after=0)
    m.para(d, f"Date: {pl['ngay_pl']} – To Sales Contract No. {pl['so_hd']} dated {pl['ngay_hd']}", italic=True, align=m.C, after=8)
    m.mixed(d, [("THE BUYER: ", True), (b["ten"], True)], after=0)
    m.para(d, f"Address: {b['dia_chi']}", after=0)
    m.para(d, f"Tax code: {b['mst']}", after=0)
    m.para(d, f"Represented by: {b['dai_dien']}", after=6)
    m.mixed(d, [("THE SELLER: ", True), (s["ten"], True)], after=0)
    m.para(d, f"Address: {s['dia_chi']}", after=8)
    m.para(d, f"Based on Sales Contract No. {pl['so_hd']} dated {pl['ngay_hd']} (hereinafter referred to as \"the Contract\"), "
              "the Buyer and the Seller have agreed to amend and supplement the Contract as follows:", after=6)
    m.para(d, "Article 1. Price term", bold=True, after=2)
    m.para(d, f"The price term of the Contract is changed from {pl['dieu_kien_cu']} to {pl['dieu_kien_moi']}.", after=6)
    m.para(d, "Article 2. Freight charge and total contract value", bold=True, after=2)
    m.para(d, f"The ocean freight charge of {usd(pl['cuoc'])} is added to the Contract. "
              "The unit prices, quantities and specifications of the goods remain unchanged. The total contract value is amended as follows:", after=4)
    t = m.table(d, ["No.", "Description", "Amount (USD)"], size=11, rows=
            [["1", "Total value of goods (17 items as per the Contract, unchanged)", f"{pl['gia_tri_hang']:,.2f}"],
             ["2", pl["tuyen"], f"{pl['cuoc']:,.2f}"],
             ["", f"TOTAL CONTRACT VALUE ({pl['dieu_kien_moi'].split(' (')[0].split(',')[0]})", f"{tong:,.2f}"]],
            widths=[1.3, 11.0, 3.7], center_cols=(0, 2))
    for c in t.rows[-1].cells:
        for r in c.paragraphs[0].runs:
            r.bold = True
    m.para(d, f"In words: {chu(tong)}.", bold=True, italic=True, after=6)
    m.para(d, "Article 3. Payment", bold=True, after=2)
    m.para(d, f"The value of goods ({usd(pl['gia_tri_hang'])}) shall be paid according to the payment terms of the Contract. "
              f"The freight charge of {usd(pl['cuoc'])} shall be paid separately by T/T to the Seller's bank account stated in the Contract.", after=6)
    m.para(d, "Article 4. General provisions", bold=True, after=2)
    m.para(d, "All other terms and conditions of the Contract remain unchanged. This Appendix is an integral part of the Contract "
              "and takes effect from the date of signing. This Appendix is made in English in 02 (two) originals, each party keeps 01 (one) original "
              "with the same legal validity.", after=12)
    m.sign_block(d, "THE SELLER", "THE BUYER", left_note="(Signature, full name and stamp)", gap=4)
    p = pathlib.Path(js).parent / (pl["ten_tep"] + ".docx")
    d.save(p)
    return p, tong


if __name__ == "__main__":
    print(build(json.load(open(sys.argv[1], encoding="utf-8")), sys.argv[1]))
