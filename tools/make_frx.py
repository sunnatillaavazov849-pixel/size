#!/usr/bin/env python3
"""Size plagini uchun SizeForm.frx faylini yaratadi.

SizeForm boshqaruv elementlarisiz (bo'sh) UserForm: barcha tugma va maydonlar
oyna ochilganda SizePlugin.SizeFormBuild orqali yaratiladi. Shuning uchun .frx
faqat bo'sh forma tavsifini saqlaydi.

.frx tuzilishi (VBA muharriri eksport qiladigan ko'rinish):
  24 baytli sarlavha: 0x0008424C, CFB hajmi, 0, 0, eni, bo'yi (twips)
  so'ng Compound File (MS-CFB): \\x01CompObj, f (MS-OFORMS FormControl), o

Ishlatish:  python3 tools/make_frx.py src/SizeForm.frx
"""
import struct
import sys

# --- MS-OFORMS FormControl ("f" oqimi), elementlarsiz ---
# Haqiqiy Office faylidagi bo'sh bo'lmagan formadan olingan sarlavha:
# versiya 0.4, PropMask 0x0C100C08 (NextAvailableID, DisplayedSize, LogicalSize,
# Font, ShapeCookie, DrawBuffer), shrift Tahoma (StdFont).
F_PREFIX = bytes.fromhex(
    "00042400080c100c03000000ffff000003000000007d0000"   # versiya, cbForm, PropMask, DataBlock
    "db2c0000ae3000000000000000000000"                   # DisplayedSize, LogicalSize (HIMETRIC)
    "0352e30b918fce119de300aa004bb851"                   # StdFont GUID
    "01" "0000" "00" "9001" "44420100" "06"              # StdFont: versiya, charset, flags, weight 400, 8.25 pt, nom uzunligi
    + "Tahoma".encode("ascii").hex()
)

CLIENT_W_TW = 6800      # 340 pt
CLIENT_H_TW = 9200      # 460 pt
CLIENT_LEFT_TW = 120
CLIENT_TOP_TW = 465

# Root \x01CompObj: "Microsoft Forms 2.0 Form", Forms.Form.1, CLSID {C62A69F0-16DC-11CE-9E98-00AA00574A4F}
COMPOBJ = bytes.fromhex(
    "0100FEFF030A0000FFFFFFFFF0692AC6DC16CE119E9800AA00574A4F"
    "190000004D6963726F736F667420466F726D7320322E3020466F726D"
    "0010000000456D626564646564204F626A656374000D000000466F72"
    "6D732E466F726D2E3100F439B271000000000000000000000000"
)


def twips_to_himetric(tw):
    return int(round(tw / 1440.0 * 2540.0))


def form_stream():
    f = bytearray(F_PREFIX)
    # DisplayedSize = mijoz maydoni o'lchami
    struct.pack_into("<LL", f, 24, twips_to_himetric(CLIENT_W_TW), twips_to_himetric(CLIENT_H_TW))
    # FormSiteData: CountOfSiteClassInfo = 0, CountOfSites = 0, CountOfBytes = 0
    f += struct.pack("<HLL", 0, 0, 0)
    return bytes(f)


# --- minimal MS-CFB (v3, 512 baytli sektorlar) ---
FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def cfb_name_key(name):
    return (len(name), name.upper())


def dir_entry(name, typ, left, right, child, start, size, clsid=b"\0" * 16):
    n = name.encode("utf-16-le") + b"\0\0" if name else b""
    e = bytearray(128)
    e[0:len(n)] = n
    struct.pack_into("<H", e, 64, len(n))
    e[66] = typ          # 1 = storage, 2 = stream, 5 = root
    e[67] = 1            # qora (black)
    struct.pack_into("<LLL", e, 68, left, right, child)
    e[80:96] = clsid
    struct.pack_into("<L", e, 116, start)
    struct.pack_into("<L", e, 120, size)
    return bytes(e)


def build_cfb(streams):
    """streams: [(nom, bayt)], hammasi 4096 baytdan kichik (mini oqim)."""
    # mini oqim
    mini = bytearray()
    minifat = []
    starts = []
    for _, data in streams:
        if not data:
            starts.append(ENDOFCHAIN)
            continue
        first = len(mini) // 64
        n = (len(data) + 63) // 64
        starts.append(first)
        for i in range(n):
            minifat.append(first + i + 1 if i < n - 1 else ENDOFCHAIN)
        mini += data + b"\0" * (n * 64 - len(data))

    # sektorlar: 0 = FAT, 1 = katalog, 2 = MiniFAT, 3.. = mini oqim
    n_mini_sect = (len(mini) + 511) // 512
    fat = [FATSECT, ENDOFCHAIN, ENDOFCHAIN]
    for i in range(n_mini_sect):
        fat.append(3 + i + 1 if i < n_mini_sect - 1 else ENDOFCHAIN)
    assert len(fat) <= 128

    # katalog: ildiz + oqimlar; oddiy muvozanatlashgan ikkilik daraxt
    order = sorted(range(len(streams)), key=lambda i: cfb_name_key(streams[i][0]))
    ids = {i: k + 1 for k, i in enumerate(order)}      # katalogdagi raqam
    left = {}; right = {}

    def build(lo, hi):
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        node = ids[order[mid]]
        left[node] = build(lo, mid - 1)
        right[node] = build(mid + 1, hi)
        return node

    root_child = build(0, len(order) - 1)
    entries = [dir_entry("Root Entry", 5, NOSTREAM, NOSTREAM, root_child,
                         3 if mini else ENDOFCHAIN, len(mini))]
    for k, i in enumerate(order):
        node = k + 1
        name, data = streams[i]
        entries.append(dir_entry(name, 2, left[node], right[node], NOSTREAM, starts[i], len(data)))
    while len(entries) < 4:
        entries.append(dir_entry("", 0, NOSTREAM, NOSTREAM, NOSTREAM, 0, 0))
    assert len(entries) == 4

    hdr = bytearray(512)
    hdr[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", hdr, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<L", hdr, 40, 0)                 # katalog sektorlari soni (v3 da 0)
    struct.pack_into("<L", hdr, 44, 1)                 # FAT sektorlari soni
    struct.pack_into("<L", hdr, 48, 1)                 # birinchi katalog sektori
    struct.pack_into("<L", hdr, 52, 0)
    struct.pack_into("<L", hdr, 56, 4096)              # mini oqim chegarasi
    struct.pack_into("<L", hdr, 60, 2 if minifat else ENDOFCHAIN)
    struct.pack_into("<L", hdr, 64, 1 if minifat else 0)
    struct.pack_into("<L", hdr, 68, ENDOFCHAIN)        # DIFAT yo'q
    struct.pack_into("<L", hdr, 72, 0)
    difat = [0] + [FREESECT] * 108
    struct.pack_into("<109L", hdr, 76, *difat)

    def sector(words):
        words = list(words) + [FREESECT] * (128 - len(words))
        return struct.pack("<128L", *words)

    out = bytearray(hdr)
    out += sector(fat)
    out += b"".join(entries)
    out += sector(minifat)
    out += mini + b"\0" * (n_mini_sect * 512 - len(mini))
    return bytes(out)


def build_frx():
    cfb = build_cfb([("\x01CompObj", COMPOBJ), ("f", form_stream()), ("o", b"")])
    cx = CLIENT_W_TW + 2 * CLIENT_LEFT_TW
    cy = CLIENT_H_TW + CLIENT_TOP_TW + CLIENT_LEFT_TW
    return struct.pack("<LLLLLL", 0x0008424C, len(cfb), 0, 0, cx, cy) + cfb


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "SizeForm.frx"
    data = build_frx()
    with open(path, "wb") as fh:
        fh.write(data)
    print(path, len(data), "bytes")
