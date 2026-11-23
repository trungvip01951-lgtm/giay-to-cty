#!/usr/bin/env python3
"""Bản Excel của Certificate of Quality (cùng dữ liệu với tao_cq_tu_co.py). Chạy: python3 cong-cu/tao_cq_excel.py co.pdf ra.xlsx thongtin.json"""
import json, re, sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from tao_cq_tu_co import doc_co  # noqa: E402
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

F = "Times New Roman"


def tach(x):
    mm = re.match(r"(.+?), MODEL: (.+?), BRAND: (.+)$", x)
    if mm:
        return mm.group(1).capitalize(), mm.group(2), mm.group(3)
    a, _, b = x.partition(": ")
    return a.capitalize(), b, ""


def build(co, ra, tt):
    items = doc_co(co)
    wb = Workbook(); ws = wb.active; ws.title = "CQ"
    w = {"A": 6, "B": 30, "C": 17, "D": 12, "E": 10, "F": 10, "G": 11}
    for k, v in w.items():
        ws.column_dimensions[k].width = v
    def dong(r, text, bold=False, size=11, center=False, italic=False):
        ws.merge_cells(f"A{r}:G{r}")
        c = ws[f"A{r}"]; c.value = text
        c.font = Font(name=F, bold=bold, size=size, italic=italic)
        c.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    r = 1
    dong(r, tt["ben_xuat"], True, 13, True); r += 1
    dong(r, tt["dia_chi_ben_xuat"], center=True); r += 2
    dong(r, "CERTIFICATE OF QUALITY", True, 16, True); r += 1
    dong(r, f"No.: {tt['so']}     Date: {tt['ngay']}", center=True); r += 2
    for a, b in (("Buyer / Consignee: ", tt["ben_mua"]), ("Address: ", tt["dia_chi_ben_mua"]),
                 ("Commercial Invoice No.: ", tt["hoa_don"]), ("Certificate of Origin (Form E) No.: ", tt["co"]),
                 ("Means of transport: ", tt["van_chuyen"]), ("Manufacturer / Country of origin: ", tt["xuat_xu"])):
        dong(r, a + b); r += 1
    r += 1
    dong(r, "We hereby certify that the goods listed below:"); r += 1
    th = Side(style="thin"); bd = Border(left=th, right=th, top=th, bottom=th)
    hdr = ["No.", "Description", "Model", "Brand", "HS code", "Q'ty (PCS)", "Quality"]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(r, i, h); c.font = Font(name=F, bold=True); c.border = bd
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r0 = r + 1
    for it in items:
        r += 1
        vals = [it["no"], *tach(it["desc"]), it["hs"], it["qty"], "New 100%"]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v); c.font = Font(name=F); c.border = bd
            c.alignment = Alignment(horizontal="left" if i == 2 else "center", vertical="center", wrap_text=True)
    r += 1
    ws.merge_cells(f"A{r}:E{r}")
    ws[f"A{r}"] = "TOTAL:"; ws[f"A{r}"].font = Font(name=F, bold=True); ws[f"A{r}"].alignment = Alignment(horizontal="right")
    ws[f"F{r}"] = f"=SUM(F{r0}:F{r-1})"; ws[f"F{r}"].font = Font(name=F, bold=True); ws[f"F{r}"].alignment = Alignment(horizontal="center")
    for i in range(1, 8):
        ws.cell(r, i).border = bd
    r += 2
    for x in ["- are brand new (100%), unused, manufactured in " + tt["nuoc"] + " in " + tt["nam"] + ";",
              "- have been inspected and tested by the manufacturer before shipment;",
              "- conform to the manufacturer's technical specifications and quality standards, and are in good working condition."]:
        dong(r, x); ws.row_dimensions[r].height = 30 if len(x) > 90 else 15; r += 1
    r += 2
    for t, b, it in (("FOR AND ON BEHALF OF", True, False), (tt["ben_xuat"], True, False), ("(Authorized signature and stamp)", False, True)):
        ws.merge_cells(f"D{r}:G{r}"); c = ws[f"D{r}"]; c.value = t
        c.font = Font(name=F, bold=b, italic=it); c.alignment = Alignment(horizontal="center", wrap_text=True)
        if t == tt["ben_xuat"]:
            ws.row_dimensions[r].height = 30
        r += 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{r0-1}:{r0-1}"
    wb.calculation.fullCalcOnLoad = True
    wb.save(ra)
    return len(items), sum(i["qty"] for i in items)


if __name__ == "__main__":
    print(build(sys.argv[1], sys.argv[2], json.load(open(sys.argv[3], encoding="utf-8"))))
