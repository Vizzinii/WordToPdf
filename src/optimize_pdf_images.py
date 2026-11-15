# -*- coding: utf-8 -*-
"""PDF image pass-through optimizer (fixes size bloat after Word/WPS PDF export).

What it does:
    Replaces each PDF image stream with a lossless encoding derived from the
    matching image inside the .docx. Text, fonts, vector table lines, page
    structure and image placement are never touched.

    Encoding ladder (every branch is pixel-lossless):
      - JPEG source          -> original JPEG bytes, /DCTDecode (1:1)
      - 2 colours (default)  -> 1bpp /Indexed + /FlateDecode
      - 2 colours (--g4)     -> 1bpp CCITT Group 4, /CCITTFaxDecode
      - 3..256 colours       -> /Indexed + /FlateDecode at the smallest bit
                                depth that fits the palette (1/2/4/8 bpp)

    The bi-level branch is what makes a scanned report shrink: Word stores
    1-bit TIFF/PNG at 1 bit per pixel, while the old 8bpp /Indexed path spent
    8 bits per pixel on the same two colours. Reducing BitsPerComponent only
    repacks the same index values; no resampling, quantisation or lossy step
    is involved.

    The same argument applies to every other palette image: a 6-colour diagram
    needs 3 bits, so it is stored at 4bpp rather than 8bpp. Nothing about the
    pixel matrix changes - only how many bits each index occupies - and a
    40-page sample of scanned pages plus 6-colour diagrams measured 12.6%
    smaller with byte-identical rendered output.

    1bpp /Indexed is the default because a full-document test measured
    59.78 MB with 12/12 identical page renders, against 61.35 MB with 4/12
    renders for G4. Some viewers anti-alias /DeviceGray 1bpp differently when
    zooming, while /Indexed keeps the original colour-space family. --g4 is
    kept for documents where Group 4 happens to compress better.

Performance notes (V1.3):
    The matching stage no longer evaluates the full 256-byte fingerprint for
    every same-size candidate: `distance` exits as soon as the running sum
    passes the acceptance threshold, and `best_match` keeps the best candidate
    instead of running `min()` and then recomputing the winner. Opaque images
    skip the RGBA composite that the quantiser used to receive. Word images
    that appear more than once are indexed once. Image encoding runs in a
    process pool when a document has many image objects.

Usage:
    python optimize_pdf_images.py A.pdf A.docx
    python optimize_pdf_images.py A.pdf                  auto-find matching .docx
    python optimize_pdf_images.py A.pdf -o B.pdf         choose output name
    python optimize_pdf_images.py A.pdf --ascii          force plain ASCII output
    python optimize_pdf_images.py A.pdf --no-verify      skip all integrity checks
    python optimize_pdf_images.py A.pdf --g4             use CCITT G4 instead of 1bpp Flate
    python optimize_pdf_images.py A.pdf --render-check   add full-page render diagnostic
    python optimize_pdf_images.py A.pdf --image-pages 40 sample 40 pages for pixel check
    python optimize_pdf_images.py A.pdf --render-pages 20  sample 20 render pages
    python optimize_pdf_images.py A.pdf --jobs 8         encode images in 8 processes

    Default verification compares decoded image pixels (the real fidelity
    check). --render-check additionally compares zoomed page rasters; viewers
    may anti-alias 1bpp images differently, so small render differences there
    do NOT indicate image data loss.

    A.pdf    must already be exported from A.docx by Word/WPS; this script
             does NOT convert docx to pdf.
    Output   defaults to "<source>_优化保真版.pdf", written next to the source
             PDF, never overwrites the source.

Dependencies:
    python -m pip install pymupdf pillow
    (Pillow needs libtiff only for the optional --g4 branch; if unavailable,
     that branch falls back to 1bpp Flate automatically.)

DO NOT change these three details, or images get re-encoded and render black:
    1. update_stream(..., compress=0)  -- compress=1 re-flates the stream
    2. garbage=4, clean=False, deflate=False -- clean=True rewrites image data
    3. xref_set_key (Filter, ColorSpace) must be called AFTER update_stream

DO NOT change these --g4 parameters either: /CCITTFaxDecode requires
    /BlackIs1 true + /Decode [0 1] + /K -1 + /Columns + /Rows, and ONE single
    G4 strip. A wrong polarity combination inverts every pixel; concatenated
    multi-strip TIFF streams decode with shifted rows. Word TIFFs are usually
    multi-strip, so g4_bilevel() decodes to a bitmap and re-encodes one strip.

Matching: group by pixel size, then compare a 16x16 grayscale fingerprint and
          confirm with a 32x32 one. A candidate must clear BOTH gates to be
          used; anything else leaves the exported image untouched. A loose
          single gate here does not merely cost bytes - it silently swaps in a
          different picture that happens to share the pixel size.

Memory: only an index (filename + fingerprint) of Word images is kept in RAM.
        Original bytes are read on demand from the docx, one image at a time.
        Output is written straight to disk; verification opens one PDF at a
        time. A single image that cannot be decoded is skipped, not fatal.
"""

import sys
import os
import io
import zlib
import time
import zipfile
import collections
import warnings
import ctypes
from ctypes import wintypes
from operator import sub, or_ as _OR

import pymupdf
from PIL import Image

warnings.filterwarnings("ignore")

# Matching gates. Two fingerprints are used: a coarse 16x16 one for the bucket
# scan, and a fine 32x32 one as confirmation. Measured on the 2,865-page
# reference report (8,401 PDF image objects, 8,356 Word images):
#   legitimate matches   -> 16x16 distance 0.03 .. 0.34   (nothing between 1 and 4)
#   wrong same-size twin -> 16x16 distance 6.87 .. 9.71
# The old single gate of 12.0 accepted every one of those wrong twins, which is
# how pages 39/40/46/47/55 ended up carrying a different image of the same pixel
# size. A budget is never a reason to accept a *different picture*: a rejected
# match only means "keep what Word exported", a wrong match silently rewrites
# page content. So the accept gate is deliberately far below the error band.
DIST_ACCEPT = 1.0          # 16x16: only near-identical fingerprints pass
DIST_ACCEPT_FINE = 3.0     # 32x32: confirmation pass on the accepted candidate
DIST_THRESHOLD = 12.0      # reporting only; never used to accept a match
PROGRESS_EVERY = 200
ZLIB_LEVEL = 9

# Bi-level policy. "flate" (default) keeps V1.2 behaviour and the identical
# rendering: 1bpp /Indexed, same colour-space family as the original. "best"
# encodes both 1bpp Flate and CCITT G4 and keeps the smaller one, and "g4"
# forces G4; both of those emit /DeviceGray + /Decode [0 1], which viewers
# anti-alias differently when zooming (measured up to 43/255 channel delta).
BILEVEL_POLICY = "flate"

# Volume-first policy: when a lossless re-encode would be LARGER than what Word
# stored, keep the smaller original instead. Set to False (via --lossless-all)
# to force pixel-exact storage at a size cost (~27 MB on the reference package).
KEEP_LOSSY = True



STAT_LABELS = {
    "jpeg_passthrough": "JPEG 原样回填",
    "g4_bilevel": "二值图转 1bpp G4",
    "flate_1bpp": "二值图转 1bpp Flate",
    "png_indexed": "多色图转 8bpp 索引",
    "indexed_1bpp": "索引图转 1bpp",
    "indexed_2bpp": "索引图转 2bpp",
    "indexed_4bpp": "索引图转 4bpp",
    "rgb_flate": "真彩无损转 Flate",
    "kept_dct_smaller": "原 JPEG 更小保留",
    "no_original": "无可靠同尺寸原图，保留导出图",
    "kept": "保留原样",
    "mem_skipped": "内存不足跳过",
    "colorspace_checked": "核对色彩空间的 JPEG 数",
    "colorspace_fixed": "色彩空间与分量不符已改正",
    "jpeg_progressive": "渐进式 JPEG 数",
}

VERIFY_LABELS = {
    "pages_src": "原页数",
    "pages_out": "新页数",
    "chars_src": "原文字字符数",
    "chars_out": "新文字字符数",
    "placements_src": "原图像放置数",
    "placements_out": "新图像放置数",
    "img_checked": "图像抽检对象数",
    "img_identical": "解码像素一致数",
    "img_diff": "解码像素不一致数",
    "render_checked": "整页渲染诊断页数",
    "render_identical": "整页渲染一致页数",
    "render_diff": "整页渲染差异页数",
}

_USE_ZH = False
_CONSOLE = None
_K32 = None
_PMC = None


def _init_win32():
    global _K32, _PMC
    if os.name != "nt" or _K32 is not None:
        return

    try:
        class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
            _fields_ = [
                ("cb", wintypes.DWORD),
                ("PageFaultCount", wintypes.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileSize", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]

        k = ctypes.windll.kernel32
        p = ctypes.windll.psapi
        k.GetStdHandle.argtypes = [wintypes.DWORD]
        k.GetStdHandle.restype = ctypes.c_void_p
        k.GetConsoleMode.argtypes = [ctypes.c_void_p,
                                     ctypes.POINTER(wintypes.DWORD)]
        k.GetConsoleMode.restype = wintypes.BOOL
        k.WriteConsoleW.argtypes = [
            ctypes.c_void_p,
            ctypes.c_wchar_p,
            wintypes.DWORD,
            ctypes.POINTER(wintypes.DWORD),
            ctypes.c_void_p,
        ]
        k.WriteConsoleW.restype = wintypes.BOOL
        k.GetCurrentProcess.restype = ctypes.c_void_p
        p.GetProcessMemoryInfo.argtypes = [
            ctypes.c_void_p,
            ctypes.POINTER(PROCESS_MEMORY_COUNTERS),
            wintypes.DWORD,
        ]
        p.GetProcessMemoryInfo.restype = wintypes.BOOL
        _K32 = k
        _PMC = PROCESS_MEMORY_COUNTERS
    except Exception:
        _K32 = None
        _PMC = None


def init_output(force_ascii=False):
    """Write Chinese via Win32 wide-char API (codepage independent); ASCII fallback."""
    global _USE_ZH, _CONSOLE
    _USE_ZH = False
    _CONSOLE = None
    if force_ascii:
        return
    _init_win32()
    if not _K32:
        return
    try:
        handle = _K32.GetStdHandle(wintypes.DWORD(0xFFFFFFF5))
        mode = wintypes.DWORD()
        if handle and _K32.GetConsoleMode(handle, ctypes.byref(mode)):
            _CONSOLE = handle
            _USE_ZH = True
            return
    except Exception:
        pass


def emit(zh, en=None):
    msg = zh if _USE_ZH else en
    if _USE_ZH and _CONSOLE:
        try:
            written = wintypes.DWORD()
            if _K32.WriteConsoleW(_CONSOLE, msg + "\n", len(msg) + 1,
                                  ctypes.byref(written), None):
                return
        except Exception:
            pass
        msg = en
    print(msg, flush=True)


def path_str(path):
    return path if _USE_ZH else ascii(path)[1:-1]


def mem_info():
    """Return (current RSS MB, peak RSS MB) on Windows, else (0, 0)."""
    _init_win32()
    if not _K32 or not _PMC:
        return 0.0, 0.0
    try:
        counters = _PMC()
        counters.cb = ctypes.sizeof(_PMC)
        ok = ctypes.windll.psapi.GetProcessMemoryInfo(
            _K32.GetCurrentProcess(), ctypes.byref(counters), counters.cb)
        if not ok:
            return 0.0, 0.0
        return (counters.WorkingSetSize / 1048576.0,
                counters.PeakWorkingSetSize / 1048576.0)
    except Exception:
        return 0.0, 0.0


def labelled(d, labels):
    if _USE_ZH:
        return "{%s}" % ", ".join(f"'{labels.get(k, k)}': {v}"
                                 for k, v in d.items())
    return str(dict(d))


FP_COARSE = 16
FP_FINE = 32


def fingerprint(im):
    """16x16 grayscale fingerprint: cheap scan of one pixel-size bucket."""
    return bytes(im.convert("L").resize((FP_COARSE, FP_COARSE),
                                        Image.Resampling.BILINEAR).tobytes())


def fingerprint_fine(im):
    """32x32 grayscale fingerprint: confirmation pass for an accepted match."""
    return bytes(im.convert("L").resize((FP_FINE, FP_FINE),
                                        Image.Resampling.BILINEAR).tobytes())


def distance(a, b):
    """Mean absolute difference of two same-length fingerprints, scale 0..255."""
    if not len(a):
        return 255.0
    return sum(map(abs, map(sub, a, b))) / float(len(a))


def distance_capped(a, b, cap):
    """Like distance() but abandons the sum once it can no longer pass `cap`.

    Element differences are non-negative, so a running sum above the cap proves
    the final value is above it as well. Returns cap + 1 on early exit.
    """
    if not len(a):
        return 255.0
    total = 0
    limit = cap * len(a)
    for x, y in zip(a, b):
        total += x - y if x > y else y - x
        if total > limit:
            return cap + 1.0
    return total / float(len(a))


def best_match(fp, candidates, fp_fine=None, threshold=None):
    """Return (candidate, distance) only for a *trustworthy* match.

    Two gates: the coarse 16x16 fingerprint must come within DIST_ACCEPT, and,
    when a fine fingerprint is available on both sides, the 32x32 one must
    confirm it within DIST_ACCEPT_FINE. A candidate that clears the bucket scan
    but fails confirmation is rejected and the caller keeps the exporter's own
    image. Rejecting costs bytes; accepting a look-alike silently swaps the
    picture on a page.
    """
    cap = DIST_ACCEPT if threshold is None else threshold
    best = None
    best_d = cap
    for cand in candidates:
        d = distance_capped(fp, cand["fp"], best_d)
        if d > best_d:
            continue
        # Accept ties (<= cap) but never replace an equal-best candidate:
        # picking a different one on a tie means back-filling a different
        # image of the same size.
        if best is None or d < best_d:
            best_d = d
            best = cand
            if d == 0:
                break
    if best is None:
        return None, best_d
    if fp_fine and best.get("fp2"):
        fine = distance_capped(fp_fine, best["fp2"], DIST_ACCEPT_FINE)
        if fine > DIST_ACCEPT_FINE:
            # Same coarse shape, different actual picture: this is exactly the
            # look-alike case that corrupted pages 39/40/46/47/55.
            return None, best_d
    return best, best_d


def inspect_bytes(raw):
    """Read only size and both fingerprints; the bitmap is released at once."""
    with Image.open(io.BytesIO(raw)) as im:
        return (im.width, im.height), fingerprint(im), fingerprint_fine(im)


def build_word_index(docx_path):
    """Keep only an index in RAM: zip filename + size + 16x16 fingerprint."""
    archive = zipfile.ZipFile(docx_path)
    grouped = collections.defaultdict(list)
    count = 0
    try:
        for name in archive.namelist():
            if "/media/" not in name or name.endswith("/"):
                continue
            try:
                raw = archive.read(name)
                size, fp, fp2 = inspect_bytes(raw)
            except MemoryError:
                emit("  内存紧张，跳过一张 Word 图片索引",
                     "  Low memory: skipped one Word image index")
                continue
            except Exception:
                continue
            # No deduplication here. Two Word images can occupy the same pixel
            # size with coarse fingerprints inside the accept gate of each
            # other, and collapsing those buckets made one of them back-fill the
            # other's bytes (measured: a 90 KB JPEG came out byte-identical to a
            # different image). The fine fingerprint in `best_match` is what
            # separates them now; deduplicating would hide the second candidate
            # from that comparison.
            entry = {"name": name, "fp": fp, "fp2": fp2}
            grouped[size].append(entry)
            count += 1
            if count % 1000 == 0:
                emit("  已索引 Word 插图 %d 张 ..." % count,
                     "  Indexed %d Word images ..." % count)
    except Exception:
        archive.close()
        raise
    return archive, grouped, count


def find_docx(pdf_path, explicit=None):
    if explicit:
        return explicit
    folder = os.path.dirname(os.path.abspath(pdf_path))
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    base = stem
    for suffix in ("_优化保真版", "_优化", "_compressed"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
            break
    cand = os.path.join(folder, base + ".docx")
    if os.path.exists(cand):
        return cand
    candidates = [os.path.join(folder, f) for f in os.listdir(folder)
                  if f.lower().endswith(".docx") and not f.startswith("~$")]
    if len(candidates) == 1:
        return candidates[0]
    return None


def _pure_bw_palette(palette):
    """Return (is_pure_bw, black_index) for a two-entry /Indexed palette."""
    c0 = tuple(palette[0:3])
    c1 = tuple(palette[3:6])
    if c0 == (0, 0, 0) and c1 == (255, 255, 255):
        return True, 0
    if c0 == (255, 255, 255) and c1 == (0, 0, 0):
        return True, 1
    return False, -1


def g4_bilevel(indices, palette, width, height):
    """Two-colour pure B/W image -> single-strip CCITT G4 stream.

    Returns (data, colorspace, bpc, decode, decode_parms) or None, so the
    caller can fall back to 1bpp Flate. TIFF is only an intermediate container
    for the G4 code stream. Multi-strip inputs are decoded to a bitmap and
    re-encoded as one strip; their raw streams are NEVER concatenated.
    """
    ok, black_index = _pure_bw_palette(palette)
    if not ok:
        return None
    try:
        table = bytes(0 if i == black_index else 255 for i in range(256))
        luma = bytes(indices).translate(table)
        im_l = Image.frombytes("L", (width, height), luma)
        del luma
        bw = im_l.convert("1")
        im_l.close()
        buf = io.BytesIO()
        bw.save(buf, format="TIFF", compression="group4",
                tiffinfo={278: height})
        bw.close()
        tiff = buf.getvalue()
        del buf
        with Image.open(io.BytesIO(tiff)) as t:
            offs = t.tag_v2.get(273)
            cnts = t.tag_v2.get(279)
            off_list = offs if isinstance(offs, (list, tuple)) else [offs]
            cnt_list = cnts if isinstance(cnts, (list, tuple)) else [cnts]
            if len(off_list) != 1 or len(cnt_list) != 1:
                return None
            off, cnt = int(off_list[0]), int(cnt_list[0])
            if off < 0 or cnt <= 0 or off + cnt > len(tiff):
                return None
            data = tiff[off:off + cnt]
        parms = ("<< /K -1 /Columns %d /Rows %d /BlackIs1 true "
                 "/EndOfBlock false /EncodedByteAlign false "
                 "/DamagedRowsBeforeError 0 >>") % (width, height)
        return data, "/DeviceGray", "1", "[0 1]", parms
    except Exception:
        return None


def _pack_indices(indices, width, height, bpc):
    """Repack one byte per index into `bpc` bits per index, PDF row order.

    PDF packs the first sample of a byte into the high-order bits and starts
    every row on a byte boundary. Consecutive samples are merged with integer
    shifts over `map(operator.or_)`, so the per-row Python loop runs once per
    output byte group; a 3 MP image packs in well under a second.
    """
    if bpc == 8:
        return bytes(indices)
    group = 8 // bpc
    mask = (1 << bpc) - 1
    data = bytes(indices)
    out = bytearray()
    pad = (-width) % group
    shifts = tuple(range(8 - bpc, -1, -bpc))
    for y in range(height):
        row = data[y * width:(y + 1) * width]
        if pad:
            row += bytes(pad)
        acc = bytes((v & mask) << shifts[0] for v in row[0::group])
        for slot in range(1, group):
            acc = bytes(map(_OR, acc,
                            bytes((v & mask) << shifts[slot]
                                  for v in row[slot::group])))
        out.extend(acc)
    return bytes(out)


def _min_bpc(hival):
    """Smallest bit depth that can represent indices 0..hival."""
    for bpc in (1, 2, 4):
        if hival < (1 << bpc):
            return bpc
    return 8


def flate_bilevel(indices, palette, width, height):
    """Two-colour image -> 1bpp /Indexed + /FlateDecode, palette preserved.

    The packed bit value equals the /Indexed index value, so no polarity or
    /Decode key is involved. This branch also covers non-B/W two-colour images.
    """
    table = bytes(255 if i else 0 for i in range(256))
    luma = bytes(indices).translate(table)
    im_l = Image.frombytes("L", (width, height), luma)
    del luma
    bw = im_l.convert("1")
    im_l.close()
    packed = bw.tobytes()
    bw.close()
    data = zlib.compress(packed, ZLIB_LEVEL)
    colorspace = "[/Indexed /DeviceRGB 1 <%s>]" % (bytes(palette[:6]).hex())
    return data, colorspace, "1", None, None


def _bilevel_pick(indices, palette, width, height, force_g4=False):
    """Choose the smaller of 1bpp Flate and CCITT G4 for a two-colour image.

    The two encoders disagree per image: on the real 2,145-page report G4 lost
    to Flate, while on clean scanned text G4 won by ~26%. Both are
    pixel-lossless and carry no /Decode ambiguity, so encoding both and keeping
    the smaller one buys size with no quality trade-off.
    """
    flate_enc = flate_bilevel(indices, palette, width, height)
    if BILEVEL_POLICY == "flate" and not force_g4:
        return flate_enc
    g4_enc = g4_bilevel(indices, palette, width, height)
    if g4_enc is None:
        return flate_enc
    if force_g4 or BILEVEL_POLICY == "g4":
        return g4_enc
    return g4_enc if len(g4_enc[0]) < len(flate_enc[0]) else flate_enc


def _palette_image(im, alpha_kept=False):
    """Reduce an image to a <=256-colour palette image, or return None.

    Returns None when the image needs more than 256 colours: a palette cannot
    represent it, and quantising would change pixels (measured: 0.9% of
    channels moved by up to 11 levels on a real 1611x1375 figure). The caller
    then falls back to the lossless true-colour path.

    `alpha_kept` notes that the PDF object keeps its own /SMask, so the colour
    samples must stay the image's own RGB; flattening onto white first would
    apply alpha twice (once into the samples, once through the mask).
    """
    if im.mode == "P":
        # Already palette-based: keep palette and index bytes verbatim.
        return im.copy()
    if im.mode in ("1", "L"):
        return im.convert("RGB").quantize(colors=256)
    rgb = im if im.mode == "RGB" else im.convert("RGB")
    try:
        if rgb.getcolors(maxcolors=257) is None:
            return None          # >256 colours, cannot be stored losslessly
        return rgb.quantize(colors=256, method=Image.Quantize.MEDIANCUT)
    finally:
        if rgb is not im:
            rgb.close()


def _truecolour_stream(im, width, height, alpha_kept=False):
    """Lossless /DeviceRGB 8bpc stream for an image a palette cannot hold.

    Word re-encodes such images as lossy JPEG in the PDF, so this is normally
    both *better* (pixel-identical to the original) and *smaller*: measured on
    135 real objects, lossless RGB+Flate came to 7.68 MB against 8.59 MB of
    stored JPEG. Returns a descriptor or None when it would not help.
    """
    rgb = im if im.mode == "RGB" else im.convert("RGB")
    try:
        if rgb.size != (width, height):
            return None
        data = zlib.compress(rgb.tobytes(), ZLIB_LEVEL)
    finally:
        if rgb is not im:
            rgb.close()
    return {"data": data, "filter": "/FlateDecode",
            "colorspace": "/DeviceRGB", "bpc": "8",
            "decode": None, "decode_parms": None, "tag": "rgb_flate"}


JPEG_COLORSPACE = {1: "/DeviceGray", 3: "/DeviceRGB", 4: "/DeviceCMYK"}


def jpeg_components(raw):
    """读出 JPEG 的分量数（不解码）；不是 JPEG 或结构异常时返回 None。

    PDF 规范要求 /DCTDecode 流的 /ColorSpace 与 JPEG 自身分量数一致：
        1 分量 -> /DeviceGray   3 分量 -> /DeviceRGB   4 分量 -> /DeviceCMYK
    不一致时阅读器会按声明的色彩空间去解释字节：单通道扫描件被当成 RGB 时，
    R 通道取到亮度、G/B 补零，浅底页整页变红、深底页整页变黑。
    """
    if raw[:2] != b"\xff\xd8":
        return None
    index, total = 2, len(raw)
    while index < total - 9:
        if raw[index] != 0xFF:
            index += 1
            continue
        marker = raw[index + 1]
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
            index += 2
            continue
        length = (raw[index + 2] << 8) | raw[index + 3]
        # SOF0/SOF1/SOF2 等：offset+4 精度，offset+9 分量数
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                      0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            if index + 10 > total:
                return None
            return raw[index + 9]
        index += 2 + length
    return None


def jpeg_is_progressive(raw):
    """SOF2（以及 SOF6/SOF10/SOF14）是渐进式编码。"""
    if raw[:2] != b"\xff\xd8":
        return False
    index, total = 2, len(raw)
    while index < total - 9:
        if raw[index] != 0xFF:
            index += 1
            continue
        marker = raw[index + 1]
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
            index += 2
            continue
        if marker in (0xC2, 0xC6, 0xCA, 0xCE):
            return True
        if marker in (0xC0, 0xC1, 0xC3, 0xC5, 0xC7, 0xC9, 0xCB, 0xCD, 0xCF):
            return False
        length = (raw[index + 2] << 8) | raw[index + 3]
        index += 2 + length
    return False


def encode_stream(raw, width, height, allow_bilevel=True, use_g4=False,
                  alpha_kept=False):
    """Build a complete PDF stream descriptor for one Word image.

    Returns a dict with data / filter / colorspace / bpc / decode /
    decode_parms / tag, or None when the image cannot be handled safely.
    With use_g4=False (default), bi-level images take the 1bpp /Indexed +
    Flate path, which keeps the same colour-space family as the 8bpp original.
    `alpha_kept` = the object keeps an /SMask, so colour samples must not be
    pre-composited onto white.
    """
    if raw[:2] == b"\xff\xd8":
        # The Word original is a JPEG. Keep its bytes only when the PDF object
        # already holds exactly those bytes; otherwise the object is a lossy
        # re-encode and a lossless stream may be both better and smaller, which
        # `_encode_one` evaluates against the stored size.
        #
        # The colour space must come from the JPEG's own component count, not a
        # hard-coded /DeviceRGB: 895 of the grayscale scans in a real bid
        # package are single-component JPEGs, and declaring them RGB turns
        # light pages solid red and dark pages solid black.
        comps = jpeg_components(raw)
        colorspace = JPEG_COLORSPACE.get(comps)
        if colorspace is None:
            return None
        return {"data": raw, "filter": "/DCTDecode",
                "colorspace": colorspace, "bpc": "8",
                "decode": None, "decode_parms": None,
                "tag": "jpeg_passthrough"}
    with Image.open(io.BytesIO(raw)) as im:
        size = im.size
        if size != (width, height):
            return None
        pal_img = _palette_image(im, alpha_kept)
        if pal_img is None:
            # More than 256 colours: a palette cannot represent it losslessly.
            # A true-colour Flate stream can, and is usually smaller than the
            # lossy JPEG Word stored here.
            return _truecolour_stream(im, width, height, alpha_kept)
    try:
        palette = list(pal_img.getpalette() or [])
        indices = bytearray(pal_img.tobytes())
    finally:
        pal_img.close()
    hival = max(indices) if indices else 0
    palette += [0] * max(0, (hival + 1) * 3 - len(palette))
    if hival == 1 and allow_bilevel:
        data, colorspace, bpc, decode, parms = _bilevel_pick(
            indices, palette, width, height, use_g4)
        g4 = parms is not None
        return {"data": data,
                "filter": "/CCITTFaxDecode" if g4 else "/FlateDecode",
                "colorspace": colorspace, "bpc": bpc,
                "decode": decode, "decode_parms": parms,
                "tag": "g4_bilevel" if g4 else "flate_1bpp"}
    bpc = _min_bpc(hival)
    packed = _pack_indices(indices, width, height, bpc)
    del indices
    data = zlib.compress(packed, ZLIB_LEVEL)
    del packed
    colorspace = "[/Indexed /DeviceRGB %d <%s>]" % (
        hival, bytes(palette[: (hival + 1) * 3]).hex())
    return {"data": data, "filter": "/FlateDecode",
            "colorspace": colorspace, "bpc": str(bpc),
            "decode": None, "decode_parms": None,
            "tag": "png_indexed" if bpc == 8 else "indexed_%dbpp" % bpc}


def _image_xrefs(doc):
    """Xref numbers of every /Image object in the document."""
    xrefs = []
    for xref in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(xref, "Subtype")[1] == "/Image":
                xrefs.append(xref)
        except Exception:
            continue
    return xrefs


def _encode_one(doc, xref, word, archive, use_g4):
    """Match one PDF image against the Word index and encode it.

    Returns (xref, descriptor_or_None, stat_key).
    """
    try:
        width = int(doc.xref_get_key(xref, "Width")[1])
        height = int(doc.xref_get_key(xref, "Height")[1])
        extracted = doc.extract_image(xref)["image"]
    except Exception:
        return xref, None, "kept"
    try:
        _, fp, fp_fine = inspect_bytes(extracted)
    except MemoryError:
        return xref, None, "mem_skipped"
    except Exception:
        return xref, None, "kept"
    del extracted
    candidates = word.get((width, height))
    if not candidates:
        return xref, None, "no_original"
    best, best_d = best_match(fp, candidates, fp_fine)
    if best is None:
        # The bucket has images of the same pixel size, but none of them is
        # this picture: the exporter resampled or cropped it, so the original
        # is not in the bucket at all. Back-filling the closest look-alike is
        # what swapped a test-report scan for an unrelated certificate on
        # pages 39/40/46/47/55. Keep what the exporter produced instead.
        return xref, None, "no_original"
    try:
        if doc.xref_get_key(xref, "ImageMask")[1] == "true":
            return xref, None, "kept"
    except Exception:
        return xref, None, "kept"
    try:
        allow_bilevel = doc.xref_get_key(xref, "SMask")[1] == "null"
    except Exception:
        return xref, None, "kept"
    try:
        raw = archive.read(best["name"])
    except Exception:
        return xref, None, "kept"
    try:
        enc = encode_stream(raw, width, height, allow_bilevel, use_g4,
                            alpha_kept=not allow_bilevel)
    except MemoryError:
        return xref, None, "mem_skipped"
    except Exception:
        return xref, None, "kept"
    finally:
        del raw
    if enc is None:
        return xref, None, "kept"
    # Never grow a stream. A lossless re-encode can be larger than what Word
    # stored (a JPEG of a photo compresses better than raw RGB), in which case
    # keeping the original bytes is the smaller-but-still-correct choice...
    # except that such an object is lossy relative to the Word original, so the
    # caller decides using `allow_lossy_keep`.
    try:
        current = len(doc.xref_stream_raw(xref) or b"")
    except Exception:
        current = 0
    if (KEEP_LOSSY and enc["tag"] != "jpeg_passthrough" and current
            and len(enc["data"]) >= current):
        if enc["tag"] == "rgb_flate":
            return xref, None, "kept_dct_smaller"
        return xref, None, "kept"
    return xref, enc, enc["tag"]


def audit_jpeg_colorspace(doc, apply_fix=True):
    """核对每个 /DCTDecode 对象的 /ColorSpace 与 JPEG 分量数是否一致。

    这一步是独立于「回填」的：回填只在匹配到 Word 原图时才改写对象，而实测
    大量扫描件因为 WPS 导出时已缩放（原图 502x713、PDF 里 381x539），匹配不上
    就原样保留导出图——错标的 /DeviceRGB 由 WPS 带进来，回填路径根本碰不到它。
    所以必须在整份 PDF 上逐个对象核对。

    返回 {"checked", "fixed", "progressive", "unknown"}；apply_fix=False 时只统计。
    """
    stats = {"checked": 0, "fixed": 0, "progressive": 0, "unknown": 0,
             "samples": []}
    for xref in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(xref, "Subtype")[1] != "/Image":
                continue
            filt = doc.xref_get_key(xref, "Filter")[1]
            if "DCTDecode" not in filt:
                continue
        except Exception:
            continue
        try:
            raw = doc.xref_stream_raw(xref) or b""
        except Exception:
            continue
        comps = jpeg_components(raw)
        if comps is None:
            stats["unknown"] += 1
            continue
        stats["checked"] += 1
        if jpeg_is_progressive(raw):
            stats["progressive"] += 1
        want = JPEG_COLORSPACE.get(comps)
        if want is None:
            stats["unknown"] += 1
            continue
        try:
            current = doc.xref_get_key(xref, "ColorSpace")[1]
        except Exception:
            current = ""
        if current == want:
            continue
        # 只纠正 DeviceXxx 系列；索引色、ICCBased 等是另一类合法写法，不碰。
        if not (current.startswith("/Device") or current.startswith("[/")):
            continue
        if len(stats["samples"]) < 8:
            stats["samples"].append((xref, comps, current[:32]))
        if apply_fix:
            try:
                doc.xref_set_key(xref, "ColorSpace", want)
                # /ColorTransform 只对 3/4 分量有意义；灰度对象上留着 1 会让
                # 部分阅读器按 YCbCr 去解释单通道数据，同样偏色。
                if comps == 1:
                    try:
                        parms = doc.xref_get_key(xref, "DecodeParms")[1]
                    except Exception:
                        parms = ""
                    if parms and "ColorTransform" in parms:
                        doc.xref_set_key(xref, "DecodeParms", "null")
                stats["fixed"] += 1
            except Exception:
                pass
        else:
            stats["fixed"] += 1
    return stats


def _apply_encoding(doc, xref, enc):
    """Write one encoded stream descriptor back into the PDF.

    Order matters: update_stream first (compress=0 keeps the bytes verbatim),
    then the dictionary keys, otherwise MuPDF re-encodes the stream.
    """
    doc.update_stream(xref, enc["data"], new=1, compress=0)
    doc.xref_set_key(xref, "Filter", enc["filter"])
    doc.xref_set_key(xref, "ColorSpace", enc["colorspace"])
    doc.xref_set_key(xref, "BitsPerComponent", enc["bpc"])
    for key, value in (("Decode", enc["decode"]),
                       ("DecodeParms", enc["decode_parms"])):
        doc.xref_set_key(xref, key, value if value else "null")
    doc.xref_set_key(xref, "ImageMask", "null")


def _worker_encode(payload):
    """Process-pool entry point: encode a slice of one document's images."""
    global ZLIB_LEVEL, emit
    pdf_path, docx_path, xrefs, use_g4, ZLIB_LEVEL = payload
    # A worker's console chatter would interleave with the parent's progress
    # lines (and is invisible in a --windowed build anyway).
    emit = lambda zh, en=None: None
    archive, word, _ = build_word_index(docx_path)
    doc = pymupdf.open(pdf_path)
    out = []
    try:
        for xref in xrefs:
            out.append(_encode_one(doc, xref, word, archive, use_g4))
    finally:
        doc.close()
        archive.close()
    return out


def optimize(pdf_path, docx_path, out_path, use_g4=False, jobs=None,
             return_index=False):
    t0 = time.time()
    archive, word, count = build_word_index(docx_path)
    if not word:
        raise RuntimeError("no usable images found in docx")
    emit("Word 插图索引完成：%d 张，用时 %.1f 秒" % (count, time.time() - t0),
         "Indexed %d Word images in %.1fs" % (count, time.time() - t0))
    doc = pymupdf.open(pdf_path)
    stats = collections.Counter()
    try:
        xrefs = _image_xrefs(doc)
        total = len(xrefs)
        emit("PDF 图像对象：%d 个，开始逐张比对（可随时查看进度）" % total,
             "Found %d PDF image objects, matching one by one" % total)
        results = []
        workers = _resolve_jobs(jobs, total)
        if workers > 1:
            emit("  并行编码：%d 个进程处理 %d 个图像对象" % (workers, total),
                 "  Parallel encode: %d processes for %d image objects"
                 % (workers, total))
            results = _parallel_encode(pdf_path, docx_path, xrefs, use_g4,
                                       workers, total)
        else:
            for seq, xref in enumerate(xrefs, 1):
                if seq % PROGRESS_EVERY == 0 or seq == total:
                    rss, _ = mem_info()
                    emit("  进度 %d/%d，已用 %.0f 秒，内存 %.0f MB"
                         % (seq, total, time.time() - t0, rss),
                         "  %d/%d, %.0fs elapsed, RSS %.0f MB"
                         % (seq, total, time.time() - t0, rss))
                results.append(_encode_one(doc, xref, word, archive, use_g4))
        for xref, enc, key in results:
            if enc is None:
                stats[key] += 1
                continue
            _apply_encoding(doc, xref, enc)
            stats[key] += 1
            del enc
        # 回填是否命中和色彩空间是否正确是两件事：WPS 导出稿里本来就有错标的
        # /DeviceRGB，回填匹配不上（尺寸已被缩放）时会原样留下。这里对整份
        # PDF 的每个 DCTDecode 对象独立核对一遍分量数。
        cs = audit_jpeg_colorspace(doc, apply_fix=True)
        stats["colorspace_checked"] = cs["checked"]
        stats["colorspace_fixed"] = cs["fixed"]
        stats["jpeg_progressive"] = cs["progressive"]
        if cs["fixed"]:
            emit("  色彩空间纠正：%d 个 JPEG 对象的 /ColorSpace 与分量数不符，已按分量数改正"
                 "（灰度 JPEG %d 个，其中渐进式 %d 个）"
                 % (cs["fixed"], cs["checked"], cs["progressive"]),
                 "  ColorSpace fixed on %d JPEG objects (%d progressive)"
                 % (cs["fixed"], cs["progressive"]))
        emit("  正在写出优化文件 ...", "  Writing optimized PDF to disk ...")
        # use_objstms compresses the object dictionaries into /ObjStm streams and
        # deflate re-compresses the content streams. Both are structural only:
        # no image byte changes, and on real documents this measured 0.7-1.6%
        # smaller. `clean` must stay False, or MuPDF rewrites image data.
        doc.save(out_path, garbage=4, clean=False, deflate=True,
                 use_objstms=1)
        doc.close()
    finally:
        if not return_index:
            archive.close()
    if return_index:
        return stats, (archive, word, count)
    return stats


def _resolve_jobs(jobs, total):
    """How many worker processes to use for encoding."""
    if jobs is None:
        env = os.environ.get("PDFOPT_JOBS")
        if env:
            try:
                jobs = int(env)
            except ValueError:
                jobs = None
    if jobs is None:
        cpus = os.cpu_count() or 1
        jobs = min(4, max(1, cpus - 1)) if total >= 120 else 1
    jobs = max(1, int(jobs))
    return min(jobs, max(1, total))


def _parallel_encode(pdf_path, docx_path, xrefs, use_g4, workers, total):
    """Encode image objects across processes; falls back to serial on error."""
    from concurrent.futures import ProcessPoolExecutor
    per = max(1, (len(xrefs) + workers - 1) // workers)
    chunks = [xrefs[i:i + per] for i in range(0, len(xrefs), per)]
    payloads = [(pdf_path, docx_path, chunk, use_g4, ZLIB_LEVEL)
                for chunk in chunks]
    results = []
    done = 0
    try:
        with ProcessPoolExecutor(max_workers=workers) as pool:
            for part in pool.map(_worker_encode, payloads):
                results.extend(part)
                done += len(part)
                emit("  进度 %d/%d（并行）" % (done, total),
                     "  %d/%d (parallel)" % (done, total))
    except Exception as exc:
        emit("  并行编码不可用（%s），改用单进程。" % exc,
             "  Parallel encoding unavailable (%s); using one process." % exc)
        results = []
        archive, word, _ = build_word_index(docx_path)
        doc = pymupdf.open(pdf_path)
        try:
            results = [_encode_one(doc, xref, word, archive, use_g4)
                       for xref in xrefs]
        finally:
            doc.close()
            archive.close()
    return results


def page_stats(path):
    doc = pymupdf.open(path)
    try:
        pages = doc.page_count
        chars = 0
        placements = 0
        for i in range(pages):
            chars += len(doc[i].get_text())
            placements += len(doc[i].get_images())
        return pages, chars, placements
    finally:
        doc.close()


def _image_samples(doc, xref):
    """Decoded pixels of one image object, normalised to 8-bit RGB bytes.

    MuPDF expands /Indexed, /DeviceGray and 1bpp images into the same RGB
    layout, so the comparison basis is identical on both sides. The previous
    implementation routed every image through a PNG encode + decode pair first,
    which cost ~2 s per 120 sampled objects for no change in the result.
    """
    pix = pymupdf.Pixmap(doc, xref)
    if pix.colorspace is None or pix.n - pix.alpha < 3:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    if pix.alpha:
        pix = pymupdf.Pixmap(pix, 0)
    return pix.width, pix.height, pix.samples


def verify_images(pdf_path, out_path, sample_pages=20):
    """Compare decoded image pixels between the two PDFs, object by object.

    This is the authoritative fidelity check: image quality is defined by the
    decoded pixel matrix, not by how a viewer anti-aliases that matrix when a
    page is zoomed. Returns (checked, identical, different).
    """
    src = pymupdf.open(pdf_path)
    out = pymupdf.open(out_path)
    try:
        n = min(src.page_count, out.page_count)
        if n <= 0:
            return 0, 0, 0
        step = max(1, n // max(1, sample_pages))
        pages = list(range(0, n, step))[:sample_pages]
        if pages and pages[-1] != n - 1:
            pages[-1] = n - 1
        identical = 0
        diff = 0
        for i in pages:
            a = src[i].get_images(full=True)
            b = out[i].get_images(full=True)
            if len(a) != len(b):
                diff += max(len(a), len(b))
                continue
            for ta, tb in zip(a, b):
                try:
                    if _image_samples(src, ta[0]) == _image_samples(out, tb[0]):
                        identical += 1
                    else:
                        diff += 1
                except Exception:
                    diff += 1
        return identical + diff, identical, diff
    finally:
        src.close()
        out.close()


def verify_render(pdf_path, out_path, sample_pages=12):
    """Render identical pages from both PDFs and compare pixel buffers.

    This catches inverted polarity, shifted G4 rows, missing images and other
    visual regressions that a text/placement-only check cannot detect.
    """
    src = pymupdf.open(pdf_path)
    out = pymupdf.open(out_path)
    try:
        n = min(src.page_count, out.page_count)
        if n <= 0:
            return 0, 0, 0
        step = max(1, n // max(1, sample_pages))
        pages = list(range(0, n, step))[:sample_pages]
        if pages and pages[-1] != n - 1:
            pages[-1] = n - 1
        matrix = pymupdf.Matrix(0.5, 0.5)
        identical = 0
        diff = 0
        for i in pages:
            a = src[i].get_pixmap(matrix=matrix, alpha=False)
            b = out[i].get_pixmap(matrix=matrix, alpha=False)
            if (a.width, a.height, a.n) == (b.width, b.height, b.n) \
                    and a.samples == b.samples:
                identical += 1
            else:
                diff += 1
            del a
            del b
        return len(pages), identical, diff
    finally:
        src.close()
        out.close()


def verify_all_pages(pdf_path, out_path, scale=0.12, gross_tol=9.0,
                     gross_frac=0.08, soft_tol=4.5, soft_frac=0.015,
                     cancel=None, progress=None):
    """Render EVERY page of both PDFs and diff them; return the flagged pages.

    Cost is tiny because the pages are rendered small: measured 17.8 s for
    2,865 pages at scale 0.12.

    Two tiers, because this audit cannot tell "same picture, sharper" from
    "different picture" by pixel totals alone - both land near mean 6 on real
    scans:
      * GROSS  (mean > gross_tol or frac > gross_frac): layout damage, a blank
        or shifted page, an inverted or broken stream. Blocks delivery.
      * SOFT   (above soft_tol): usually the sharpness difference between the
        exporter's smoothed JPEG and the lossless original that replaced it,
        which is the intended result. Reported for the log only.
    The real guard against a wrong *pick* is DIST_ACCEPT / DIST_ACCEPT_FINE in
    `best_match`, because a look-alike swap looks exactly like a legitimate
    back-fill once it is in the file.

    Returns (checked, gross, soft) with each list holding
    (page_number, mean_delta, frac_moved).
    """
    src = pymupdf.open(pdf_path)
    out = pymupdf.open(out_path)
    gross, soft = [], []
    checked = 0
    try:
        n = min(src.page_count, out.page_count)
        matrix = pymupdf.Matrix(scale, scale)
        for i in range(n):
            if cancel is not None and cancel.is_set():
                break
            try:
                a = src[i].get_pixmap(matrix=matrix, colorspace=pymupdf.csGRAY,
                                      alpha=False)
                b = out[i].get_pixmap(matrix=matrix, colorspace=pymupdf.csGRAY,
                                      alpha=False)
            except Exception:
                continue
            checked += 1
            if (a.width, a.height) != (b.width, b.height):
                gross.append((i + 1, 255.0, 1.0))
                continue
            sa, sb = a.samples, b.samples
            if sa == sb:
                del a
                del b
                continue
            total = 0
            moved = 0
            count = len(sa)
            for x, y in zip(sa, sb):
                d = x - y if x > y else y - x
                total += d
                if d > 40:
                    moved += 1
            mean = total / float(count)
            frac = moved / float(count)
            row = (i + 1, round(mean, 2), round(frac, 4))
            if mean > gross_tol or frac > gross_frac:
                gross.append(row)
            elif mean > soft_tol or frac > soft_frac:
                soft.append(row)
            if progress is not None and (i + 1) % 500 == 0:
                progress(i + 1, n)
            del a
            del b
    finally:
        src.close()
        out.close()
    return checked, gross, soft


def verify(pdf_path, out_path, image_pages=20, render_check=False,
           render_pages=12):
    probe = pymupdf.open(pdf_path)
    pages = probe.page_count
    probe.close()
    emit("  正在做完整性校验（逐页比对 %d 页）..." % pages,
         "  Verifying %d pages ..." % pages)
    p1, c1, i1 = page_stats(pdf_path)
    p2, c2, i2 = page_stats(out_path)
    result = {"pages_src": p1, "pages_out": p2,
              "chars_src": c1, "chars_out": c2,
              "placements_src": i1, "placements_out": i2}
    emit("  正在做图像解码像素比对（抽样 %d 页的图像对象）..." % image_pages,
         "  Comparing decoded image pixels on %d sample pages ..."
         % image_pages)
    checked, identical, diff = verify_images(pdf_path, out_path, image_pages)
    result["img_checked"] = checked
    result["img_identical"] = identical
    result["img_diff"] = diff
    if render_check:
        emit("  正在做整页渲染诊断（抽样 %d 页，仅作参考）..." % render_pages,
             "  Running full-page render diagnostic on %d pages ..."
             % render_pages)
        rc, ri, rd = verify_render(pdf_path, out_path, render_pages)
        result["render_checked"] = rc
        result["render_identical"] = ri
        result["render_diff"] = rd
    return result


def main():
    global BILEVEL_POLICY, KEEP_LOSSY
    argv = sys.argv[1:]
    force_ascii = ("--ascii" in argv
                   or os.environ.get("PDFOPT_ASCII") == "1")
    no_verify = "--no-verify" in argv
    use_g4 = "--g4" in argv
    if "--lossless-all" in argv:
        # Default is volume-first: a lossless re-encode is only used when it is
        # not larger than what Word stored. This flag removes that size test, so
        # every lossy JPEG is replaced by a pixel-exact stream even when the
        # file grows (~27 MB on the reference bid package).
        KEEP_LOSSY = False
    if "--bilevel-best" in argv:
        # Encode both 1bpp Flate and G4 per image and keep the smaller one.
        # Cheaper in bytes, but G4 renders through /DeviceGray + /Decode [0 1],
        # which some viewers anti-alias differently at zoom.
        BILEVEL_POLICY = "best"
    render_check = "--render-check" in argv
    jobs = None
    zlib_level = None
    image_pages = 20
    render_pages = 12
    for flag, attr in (("--image-pages", "image_pages"),
                       ("--render-pages", "render_pages"),
                       ("--jobs", "jobs"),
                       ("--zlib-level", "zlib_level")):
        if flag in argv:
            k = argv.index(flag)
            try:
                val = max(1, int(argv[k + 1]))
            except (IndexError, ValueError):
                emit("参数 %s 需要正整数，已按默认值处理。" % flag,
                     "Invalid %s value; using the default." % flag)
            else:
                if attr == "image_pages":
                    image_pages = val
                elif attr == "render_pages":
                    render_pages = val
                elif attr == "jobs":
                    jobs = val
                else:
                    zlib_level = max(1, min(9, val))
            del argv[k:k + 2]
    if zlib_level is not None:
        global ZLIB_LEVEL
        ZLIB_LEVEL = zlib_level
        emit("  deflate 级别设为 %d（越小越快，文件略大）" % ZLIB_LEVEL,
             "  deflate level set to %d (faster, slightly larger)" % ZLIB_LEVEL)
    argv = [a for a in argv
            if a not in ("--ascii", "--no-verify", "--g4", "--render-check",
                         "--no-render-check", "--no-g4", "--bilevel-best",
                         "--lossless-all")]
    init_output(force_ascii)
    if not argv:
        print(__doc__)
        return 1
    if argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    pdf_path = argv[0]
    docx_path = None
    out_path = None
    i = 1
    while i < len(argv):
        arg = argv[i]
        if arg in ("-o", "--out") and i + 1 < len(argv):
            out_path = argv[i + 1]
            i += 2
        elif arg.lower().endswith(".docx"):
            docx_path = arg
            i += 1
        elif arg.lower().endswith(".pdf"):
            out_path = arg
            i += 1
        else:
            emit("无法识别的参数：%s" % arg, "Unknown argument: %s" % arg)
            return 2
    if not os.path.exists(pdf_path):
        emit("找不到 PDF 文件：%s" % pdf_path,
             "PDF not found: %s" % path_str(pdf_path))
        return 2
    docx_path = find_docx(pdf_path, docx_path)
    if not docx_path or not os.path.exists(docx_path):
        emit("未找到对应 Word 文档，请显式传入 A.docx",
             "Matching .docx not found; pass it explicitly")
        return 2
    if out_path is None:
        out_path = os.path.splitext(pdf_path)[0] + "_优化保真版.pdf"
    before = os.path.getsize(pdf_path)
    try:
        stats = optimize(pdf_path, docx_path, out_path, use_g4, jobs)
    except MemoryError:
        emit("内存不足，任务未能完成；可关闭其他程序后重试，或把文档拆分处理。",
             "Out of memory. Close other applications and retry, "
             "or split the document.")
        return 3
    after = os.path.getsize(out_path)
    emit("Word 文档：%s" % path_str(docx_path),
         "Source DOCX: %s" % path_str(docx_path))
    emit("替换统计：%s" % labelled(stats, STAT_LABELS),
         "Replacements: %s" % labelled(stats, STAT_LABELS))
    emit("体积：%.2f MB -> %.2f MB（减少 %.1f%%）"
         % (before / 1048576, after / 1048576,
            (1 - after / before) * 100),
         "Size: %.2f MB -> %.2f MB (-%.1f%%)"
         % (before / 1048576, after / 1048576,
            (1 - after / before) * 100))
    emit("相对 Word：%.2f MB，输出 PDF 差值 %+.2f MB"
         % (os.path.getsize(docx_path) / 1048576,
            (after - os.path.getsize(docx_path)) / 1048576),
         "Vs DOCX: %.2f MB, output PDF delta %+.2f MB"
         % (os.path.getsize(docx_path) / 1048576,
            (after - os.path.getsize(docx_path)) / 1048576))
    if no_verify:
        emit("已按 --no-verify 跳过完整性校验；输出文件已生成。",
             "Verification skipped (--no-verify); output file was written.")
    else:
        try:
            result = verify(pdf_path, out_path, image_pages,
                            render_check, render_pages)
        except MemoryError:
            emit("内存不足，跳过完整性校验；输出文件已生成。",
                 "Low memory: skipped verification; output file was written.")
        else:
            emit("完整性校验：%s" % labelled(result, VERIFY_LABELS),
                 "Integrity: %s" % labelled(result, VERIFY_LABELS))
            _, peak = mem_info()
            emit("峰值内存约 %.0f MB，输出文件：%s"
                 % (peak, path_str(out_path)),
                 "Peak RSS %.0f MB, output: %s"
                 % (peak, path_str(out_path)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
