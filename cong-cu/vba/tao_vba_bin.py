#!/usr/bin/env python3
"""Đóng gói mã VBA (Module1.bas) thành vbaProject.bin bằng bộ xuất VBA của LibreOffice (chuẩn MS-OVBA).
Chạy: python3 cong-cu/vba/tao_vba_bin.py Module1.bas ra/vbaProject.bin   (cần soffice và python uno)
Mã VBA chỉ dùng ký tự ASCII (chữ Việt viết bằng ChrW) vì dự án VBA dùng bảng mã 1252."""
import pathlib
import subprocess
import sys
import tempfile
import time
import zipfile

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.script import ModuleInfo
from com.sun.star.script.ModuleType import NORMAL


def P(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def main(bas, ra):
    code = pathlib.Path(bas).read_text(encoding="ascii")
    with tempfile.TemporaryDirectory() as t:
        port = 2010
        proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{t}/u",
                                 f"--accept=socket,host=localhost,port={port};urp;"], stderr=subprocess.DEVNULL)
        ctx = None
        for _ in range(80):
            try:
                lc = uno.getComponentContext()
                res = lc.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", lc)
                ctx = res.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (P("Hidden", True),))
        libs = doc.BasicLibraries
        libs.VBACompatibilityMode = True
        libs.ProjectName = "VBAProject"
        if not libs.hasByName("VBAProject"):
            libs.createLibrary("VBAProject")
        lib = libs.getByName("VBAProject")
        mi = ModuleInfo()
        mi.ModuleType = NORMAL
        lib.insertModuleInfo("Module1", mi)
        lib.insertByName("Module1", "Option VBASupport 1\n" + code)
        out = f"{t}/m.xlsm"
        doc.storeToURL("file://" + out, (P("FilterName", "Calc MS Excel 2007 VBA XML"),))
        doc.close(True)
        try:
            desk.terminate()
        except Exception:
            pass
        proc.wait(timeout=60)
        with zipfile.ZipFile(out) as z:
            pathlib.Path(ra).write_bytes(z.read("xl/vbaProject.bin"))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
