#!/usr/bin/env python3
"""Đọc/ghi tệp vbaProject.bin (định dạng CFB + nén MS-OVBA) đủ dùng để sửa bảng mã của dự án VBA.
LibreOffice khi chạy không có locale Windows ghi PROJECTCODEPAGE = 0; Excel cần 1252. sua_code_page() đọc tệp, sửa giá trị
trong luồng VBA/dir rồi ghi lại tệp CFB mới (mọi luồng nhỏ hơn 4096 byte nên nằm trong ministream)."""
import struct

FREE, END, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


def doc_cfb(data):
    """Trả về dict {đường dẫn: bytes} và danh sách storage."""
    ss = 1 << struct.unpack_from("<H", data, 30)[0]
    mss = 1 << struct.unpack_from("<H", data, 32)[0]
    nfat, dir0 = struct.unpack_from("<I", data, 44)[0], struct.unpack_from("<I", data, 48)[0]
    cutoff, mf0 = struct.unpack_from("<I", data, 56)[0], struct.unpack_from("<I", data, 60)[0]
    sec = lambda i: data[512 + i * ss:512 + (i + 1) * ss]  # noqa: E731
    fat = []
    for s in struct.unpack_from("<109I", data, 76)[:nfat]:
        fat += struct.unpack(f"<{ss // 4}I", sec(s))

    def chain(st):
        out = []
        while st < 0xFFFFFFF0:
            out.append(st)
            st = fat[st]
        return out

    dirs = b"".join(sec(i) for i in chain(dir0))
    ents = []
    for i in range(len(dirs) // 128):
        e = dirs[i * 128:(i + 1) * 128]
        nl = struct.unpack_from("<H", e, 64)[0]
        ents.append(dict(name=e[:max(nl - 2, 0)].decode("utf-16le"), typ=e[66], l=struct.unpack_from("<I", e, 68)[0],
                         r=struct.unpack_from("<I", e, 72)[0], c=struct.unpack_from("<I", e, 76)[0],
                         start=struct.unpack_from("<I", e, 116)[0], size=struct.unpack_from("<I", e, 120)[0]))
    ms = b"".join(sec(i) for i in chain(ents[0]["start"])) if ents[0]["start"] < 0xFFFFFFF0 else b""
    mfat = []
    for s in chain(mf0) if mf0 < 0xFFFFFFF0 else []:
        mfat += struct.unpack(f"<{ss // 4}I", sec(s))

    def stream(e):
        if e["size"] < cutoff:
            out, st = b"", e["start"]
            while st < 0xFFFFFFF0:
                out += ms[st * mss:(st + 1) * mss]
                st = mfat[st]
            return out[:e["size"]]
        return b"".join(sec(i) for i in chain(e["start"]))[:e["size"]]

    streams, storages = {}, []

    def walk(i, prefix):
        if i == NOSTREAM or i >= len(ents):
            return
        e = ents[i]
        walk(e["l"], prefix)
        if e["typ"] == 2:
            streams[prefix + e["name"]] = stream(e)
        elif e["typ"] == 1:
            storages.append(prefix + e["name"])
            walk(e["c"], prefix + e["name"] + "/")
        walk(e["r"], prefix)

    walk(ents[0]["c"], "")
    return streams, storages


def giai_nen(b):
    assert b[0] == 1
    out, i = bytearray(), 1
    while i < len(b):
        hdr = struct.unpack_from("<H", b, i)[0]
        size, nen = (hdr & 0x0FFF) + 3, hdr >> 15
        chunk, i, d0 = b[i + 2:i + size], i + size, len(out)
        if not nen:
            out += chunk
            continue
        j = 0
        while j < len(chunk):
            fb = chunk[j]
            j += 1
            for bit in range(8):
                if j >= len(chunk):
                    break
                if not (fb >> bit) & 1:
                    out.append(chunk[j])
                    j += 1
                else:
                    tok = struct.unpack_from("<H", chunk, j)[0]
                    j += 2
                    bc = max((len(out) - d0 - 1).bit_length(), 4)
                    ln, off = (tok & (0xFFFF >> bc)) + 3, (tok >> (16 - bc)) + 1
                    for _ in range(ln):
                        out.append(out[-off])
    return bytes(out)


def nen(data):
    """Nén MS-OVBA chỉ dùng token chữ (hợp lệ theo chuẩn): khúc đủ 4096 byte ghi dạng không nén, khúc cuối ghi dạng nén."""
    out = bytearray(b"\x01")
    for k in range(0, len(data), 4096):
        part = data[k:k + 4096]
        if len(part) == 4096:
            out += struct.pack("<H", 0x3000 | (4096 + 2 - 3)) + part      # khúc không nén: cờ 0, kích thước 4098-3
            continue
        body = bytearray()
        for t in range(0, len(part), 8):
            body.append(0)
            body += part[t:t + 8]
        out += struct.pack("<H", 0xB000 | (len(body) + 2 - 3)) + body
    return bytes(out)


def ghi_cfb(streams, storages):
    """Ghi CFB phiên bản 3, mọi luồng < 4096 byte (ministream)."""
    nodes = [dict(name="Root Entry", typ=5, path="")]
    for s in storages:
        nodes.append(dict(name=s.split("/")[-1], typ=1, path=s))
    for p in streams:
        nodes.append(dict(name=p.split("/")[-1], typ=2, path=p))
    for i, n in enumerate(nodes):
        n.update(id=i, l=NOSTREAM, r=NOSTREAM, c=NOSTREAM, start=END, size=0)
    # ministream
    mini, mfat = bytearray(), []
    for n in nodes:
        if n["typ"] == 2:
            b = streams[n["path"]]
            assert len(b) < 4096
            if not b:
                continue
            k = (len(b) + 63) // 64
            n["start"], n["size"] = len(mfat), len(b)
            mfat += [len(mfat) + t + 1 for t in range(k - 1)] + [END]
            mini += b + b"\0" * (k * 64 - len(b))
    # cây nhị phân con của mỗi storage (sắp theo độ dài tên rồi chữ hoa)
    def cay(ds):
        if not ds:
            return NOSTREAM
        m = len(ds) // 2
        ds[m]["l"], ds[m]["r"] = cay(ds[:m]), cay(ds[m + 1:])
        return ds[m]["id"]
    for parent in [n for n in nodes if n["typ"] in (1, 5)]:
        con = [n for n in nodes if n is not parent and "/".join(n["path"].split("/")[:-1]) == parent["path"]]
        con.sort(key=lambda n: (len(n["name"]), n["name"].upper()))
        parent["c"] = cay(con)
    n_dir = (len(nodes) + 3) // 4
    n_mf = (len(mfat) * 4 + 511) // 512
    n_ms = (len(mini) + 511) // 512
    n_fat = 1
    while n_fat * 128 < n_fat + n_dir + n_mf + n_ms:
        n_fat += 1
    s_dir, s_mf, s_ms = n_fat, n_fat + n_dir, n_fat + n_dir + n_mf
    nodes[0]["start"], nodes[0]["size"] = (s_ms if mini else END), len(mini)
    fat = [FATSECT] * n_fat
    for s0, k in ((s_dir, n_dir), (s_mf, n_mf), (s_ms, n_ms)):
        fat += [s0 + t + 1 for t in range(k - 1)] + ([END] if k else [])
    fat += [FREE] * (n_fat * 128 - len(fat))
    hdr = bytearray(512)
    hdr[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", hdr, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", hdr, 40, 0, n_fat, s_dir, 0, 4096, s_mf if n_mf else END, n_mf, END, 0)
    difat = list(range(n_fat)) + [FREE] * (109 - n_fat)
    struct.pack_into("<109I", hdr, 76, *difat)
    dirb = bytearray()
    for n in nodes:
        e = bytearray(128)
        nm = n["name"].encode("utf-16le") + b"\0\0"
        e[0:len(nm)] = nm
        struct.pack_into("<HBB", e, 64, len(nm), n["typ"], 1)
        struct.pack_into("<III", e, 68, n["l"], n["r"], n["c"])
        struct.pack_into("<II", e, 116, n["start"] if (n["typ"] != 1) else 0, n["size"])
        dirb += e
    while len(dirb) < n_dir * 512:
        e = bytearray(128)
        struct.pack_into("<III", e, 68, NOSTREAM, NOSTREAM, NOSTREAM)
        dirb += e
    mfb = b"".join(struct.pack("<I", x) for x in mfat) + struct.pack("<I", FREE) * (n_mf * 128 - len(mfat))
    out = bytes(hdr) + b"".join(struct.pack("<I", x) for x in fat) + bytes(dirb) + mfb + bytes(mini) + b"\0" * (n_ms * 512 - len(mini))
    return out


def sua_code_page(bin_bytes, code_page=1252):
    streams, storages = doc_cfb(bin_bytes)
    d = bytearray(giai_nen(streams["VBA/dir"]))
    i = 0
    while i < len(d):                                   # đi qua các bản ghi đầu luồng dir tới PROJECTCODEPAGE (0x0003)
        rid, sz = struct.unpack_from("<HI", d, i)
        if rid == 0x0003:
            struct.pack_into("<H", d, i + 6, code_page)
            break
        i += 6 + sz
        if rid == 0x0009:                               # PROJECTVERSION: kích thước khai 4 nhưng thực tế 6 byte dữ liệu
            i += 2
    else:
        raise ValueError("không thấy PROJECTCODEPAGE")
    streams["VBA/dir"] = nen(bytes(d))
    out = ghi_cfb(streams, storages)
    s2, _ = doc_cfb(out)                                # đọc lại để chắc chắn
    assert {k: v for k, v in s2.items() if k != "VBA/dir"} == {k: v for k, v in streams.items() if k != "VBA/dir"}
    assert giai_nen(s2["VBA/dir"]) == bytes(d)
    return out
