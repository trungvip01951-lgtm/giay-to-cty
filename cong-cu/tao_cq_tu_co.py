#!/usr/bin/env python3
"""Soạn Certificate of Quality (CQ, tiếng Anh) theo danh mục hàng của CO Form E (PDF), để bên xuất/hãng ký, đóng dấu.
Không sửa CO gốc; chỉ đọc danh mục bằng pdftotext. Chạy: python3 cong-cu/tao_cq_tu_co.py co.pdf ra.docx"""
import re, subprocess, sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import tao_mau as m  # noqa: E402

RAC = re.compile(r"\s*E26\w+|\s*Verification:\S+")


def doc_co(pdf):
    L = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True, check=True).stdout.split("\n")
    items, cur = [], None
    for l in L:
        mm = re.match(r'^\s{1,6}(\d{1,2})\s{2,}(?:N/M\s+)?(\S.*?)\s{2,}"PE"\s+(\d+)\s*PIECES?', l)
        if mm:
            cur = {"no": int(mm.group(1)), "desc": mm.group(2).strip(), "qty": int(mm.group(3)), "hs": None}
            items.append(cur)
            continue
        if cur and cur["hs"] is None:
            s = l.strip()
            h = re.search(r"HS CODE:\s*([\d.]+)", s)
            if h:
                cur["hs"] = h.group(1)
            elif s and not s.startswith("(PAGE") and len(l) - len(l.lstrip()) > 15 and "***" not in s and "VIET" not in s and "Nanning" not in s:
                cur["desc"] += " " + s.split("  ")[0]
    for it in items:
        it["desc"] = re.sub(r"\s+", " ", RAC.sub("", it["desc"])).strip()
    assert [i["no"] for i in items] == list(range(1, len(items) + 1)), "Thiếu dòng hàng khi đọc CO"
    return items


def build(co, ra, tt):
    items = doc_co(co)
    d = m.new_doc()
    d.styles["Normal"].font.size = m.Pt(11)
    m.para(d, tt["ben_xuat"], bold=True, align=m.C, size=13, after=0)
    m.para(d, tt["dia_chi_ben_xuat"], align=m.C, after=10)
    m.para(d, "CERTIFICATE OF QUALITY", bold=True, align=m.C, size=16, after=2)
    m.para(d, f"No.: {tt['so']}          Date: {tt['ngay']}", align=m.C, after=10)
    for a, b in (("Buyer / Consignee: ", tt["ben_mua"]), ("Address: ", tt["dia_chi_ben_mua"]),
                 ("Commercial Invoice No.: ", tt["hoa_don"]), ("Certificate of Origin (Form E) No.: ", tt["co"]),
                 ("Means of transport: ", tt["van_chuyen"]), ("Manufacturer / Country of origin: ", tt["xuat_xu"])):
        m.mixed(d, [(a, True), (b, False)], after=0)
    m.para(d, "", after=4)
    m.para(d, "We hereby certify that the goods listed below:", after=2)
    def tach(x):   # "WATER PUMP, MODEL: X, BRAND: Y" -> (Water pump, X, Y)
        mm = re.match(r"(.+?), MODEL: (.+?), BRAND: (.+)$", x)
        if mm:
            return mm.group(1).capitalize(), mm.group(2), mm.group(3)
        a, _, b = x.partition(": ")            # "SPECIALIZED COMPONENT ...: COUPLING DN100"
        return a.capitalize(), b, ""
    t = m.table(d, ["No.", "Description", "Model", "Brand", "HS code", "Q'ty", "Quality"],
                [[str(i["no"]), *tach(i["desc"]), i["hs"], f"{i['qty']} PCS", "New 100%"] for i in items],
                [1.1, 4.4, 3.0, 2.2, 1.6, 1.7, 2.0], size=10, center_cols=(0, 2, 3, 4, 5, 6))
    tong = sum(i["qty"] for i in items)
    c = t.add_row().cells
    a = c[0].merge(c[4]); a.paragraphs[0].alignment = m.R; m.run(a.paragraphs[0], "TOTAL:", bold=True, size=10)
    c[5].paragraphs[0].alignment = m.C; m.run(c[5].paragraphs[0], f"{tong} PCS", bold=True, size=10)
    for x in ["are brand new (100%), unused, manufactured in " + tt["nuoc"] + " in " + tt["nam"] + ";",
              "have been inspected and tested by the manufacturer before shipment;",
              "conform to the manufacturer's technical specifications and quality standards, and are in good working condition."]:
        p = m.para(d, "- " + x, after=1); p.paragraph_format.left_indent = m.Cm(0.6)
    m.para(d, "", after=8)
    s = d.add_table(rows=1, cols=2); s.autofit = False; m.shade_borders(s, borders=False)
    q = s.rows[0].cells[1].paragraphs[0]; q.alignment = m.C; m.run(q, "FOR AND ON BEHALF OF", bold=True)
    q = s.rows[0].cells[1].add_paragraph(); q.alignment = m.C; m.run(q, tt["ben_xuat"], bold=True)
    q = s.rows[0].cells[1].add_paragraph(); q.alignment = m.C; m.run(q, "(Authorized signature and stamp)", italic=True)
    m.fix_layout(s, [6.0, 10.0])
    d.save(ra)
    return len(items), tong


if __name__ == "__main__":
    import json
    print(build(sys.argv[1], sys.argv[2], json.load(open(sys.argv[3], encoding="utf-8"))))
