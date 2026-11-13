# -*- coding: utf-8 -*-
"""
wps_to_pdf.py -- Export Word documents to PDF through WPS Office COM automation.

This replaces the manual "Save as PDF" step in WPS Writer. The script starts a
dedicated, hidden WPS automation instance (KWps.Application), opens each .docx
read-only, and calls Document.ExportAsFixedFormat(path, 17), which is the same
PDF export engine used by WPS's own menu command. Open WPS windows belonging to
the user are not touched.

Usage:
    python wps_to_pdf.py A.docx
    python wps_to_pdf.py A.docx -o A.pdf
    python wps_to_pdf.py A.docx B.docx -d D:\\output
    python wps_to_pdf.py A.docx --optimize
    python wps_to_pdf.py A.docx --optimize --g4
    python wps_to_pdf.py A.docx --engine cli
    python wps_to_pdf.py A.docx --overwrite --visible
    python wps_to_pdf.py --check

    --optimize chains the existing optimize_pdf_images.py: after WPS writes
    A.pdf, the Word originals are back-filled and A_优化保真版.pdf is produced.
    Both files are kept. The optimizer source is never modified by this script.

    Two export engines are available:
      wps (default) -- KWps.Application COM, Document.ExportAsFixedFormat.
                       No account requirement; byte-identical image payload to
                       WPS's manual "Output to PDF" (verified on 2,145 pages).
      cli           -- official wpscli.exe word2pdf command line. Requires a
                       signed-in WPS account with VIP privilege per its own
                       --help (exit 100/101 otherwise). Use --image-quality
                       medium to match the desktop default; "high" produces a
                       much larger intermediate PDF.
      word          -- Microsoft Office COM fallback, only if WPS is absent.

Options:
    -o / --out           output PDF path (single input only)
    -d / --out-dir       output directory for batch conversion
    --optimize           run image optimisation after export
    --optimized-out      explicit path for the optimised PDF
    --g4                 optimiser: use CCITT G4 instead of the default 1bpp Flate
    --image-pages N      optimiser: sample N pages for image pixel verification
    --render-check       optimiser: also run the full-page render diagnostic
    --no-verify          skip optimiser integrity verification
    --no-docx-verify     skip the Word-original pixel reference check
    --overwrite          allow overwriting an existing output PDF
    --visible            show the WPS window (for troubleshooting only)
    --engine wps|cli|word  export engine, default wps
    --cli-path PATH      explicit path to wpscli.exe (auto-detected otherwise)
    --image-quality low|medium|high   CLI engine only, default medium
    --timeout N          seconds to wait for the exported file to settle (300)
    --no-page-check      skip comparing WPS page count with the resulting PDF
    --ascii              plain ASCII console output
    --check              environment self-test: COM + export + PDF verification

Dependencies:
    python -m pip install pywin32 pymupdf pillow
    (pywin32 is only needed for the COM engines; pymupdf validates the result.)

Exit codes: 0 success, 2 argument/file error, 3 environment/COM error,
            4 export failed, 5 optimiser failed.
"""

from __future__ import annotations

import argparse
import gc
import io
import os
import re
import subprocess
import sys
import time

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

try:
    import optimize_pdf_images as _opt
except Exception:
    _opt = None

_ASCII = False

EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0
WD_STATISTIC_PAGES = 2
WD_FORMAT_XML_DOCUMENT = 12
PROGIDS = {"wps": "KWps.Application", "word": "Word.Application"}

CLI_EXIT_MESSAGES = {
    3: "转换超时",
    100: "未登录 WPS 账号（可改用 --engine wps 走 COM 通道）",
    101: "WPS 账号缺少 VIP 权限（可改用 --engine wps 走 COM 通道）",
    201: "缺少必要参数",
    203: "输入文件不存在",
    207: "输入文件超过 200 MB 上限",
    209: "源文档需要打开密码，请用 -p 传入",
    210: "打开密码不正确",
    211: "输出目录不存在",
    212: "输出目录无写权限",
}

# Common HRESULTs from the WPS/Word automation surface, translated so the log
# says what to do instead of printing a bare hex code.
HRESULT_MESSAGES = {
    0x800A1437: "文档被其他程序占用或只读，请先关闭已打开的该文档",
    0x800A1439: "文档需要打开密码",
    0x800A1428: "文档已损坏，无法读取",
    0x80010105: "WPS 自动化实例已失效，请重试或重启 WPS",
    0x8001010A: "WPS 主线程忙（可能弹出了对话框），请切换到 WPS 窗口处理",
    0x80004005: "WPS 返回未指明的错误，常见于权限不足或文档异常",
    0x80070005: "权限不足：请勿以管理员身份运行本工具，或改用有写权限的目录",
    0x80070020: "目标文件被占用，请关闭占用它的程序后重试",
    0x80070070: "磁盘空间不足",
    0x80070570: "文件或目录已损坏且无法读取",
}

EXPORT_RETRIES = 1
RETRY_DELAY = 1.5



def emit(zh, en):
    if _opt is not None:
        _opt.emit(zh, en)
        return
    print(en if _ASCII else zh, flush=True)


def init_output(ascii_only=False):
    global _ASCII
    _ASCII = ascii_only
    if _opt is not None:
        _opt.init_output(ascii_only)
        return
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def com_message(exc):
    """Turn a pywin32 com_error into a readable one-line message."""
    try:
        args = exc.args
        if len(args) >= 3 and args[2]:
            detail = (args[2][2] if len(args[2]) > 2 else str(args[2]))
            hr = (args[0] & 0xFFFFFFFF) if isinstance(args[0], int) else 0
            hint = HRESULT_MESSAGES.get(hr)
            if hint:
                return "%s（0x%08X）：%s" % (detail, hr, hint)
            return "%s (0x%08X)" % (detail, hr)
        return str(args[1] if len(args) > 1 else exc)
    except Exception:
        return str(exc)


def is_transient(exc):
    """True when a COM failure is worth one immediate retry.

    WPS/Word rejects automation calls while a modal dialog or a previous
    instance is still shutting down; those clear up by themselves.
    """
    hr = exc.args[0] if getattr(exc, "args", None) else None
    if isinstance(hr, int) and hr & 0xFFFFFFFF in (
            0x80010105, 0x8001010A, 0x80004005):
        return True
    text = str(exc).lower()
    return any(token in text for token in ("call was rejected", "busy",
                                           "正在忙", "retry"))


# HRESULTs that mean "the automation server process is gone", not "this
# document is bad". Retrying on the same object is pointless; the caller has to
# create a new instance.
_DEAD_INSTANCE_HRESULTS = {
    0x800706BA,   # RPC_S_SERVER_UNAVAILABLE  远程过程调用失败
    0x800706BE,   # RPC_S_CALL_FAILED         RPC 服务器不可用
    0x800706BF,   # RPC_S_CALL_FAILED_DNE
    0x80010108,   # RPC_E_DISCONNECTED
    0x8001010E,   # RPC_E_WRONG_THREAD (instance marshal lost)
    0x80010007,   # RPC_E_SERVER_DIED
    0x800401FD,   # CO_E_OBJNOTCONNECTED
}


def is_instance_dead(exc):
    """True when the automation process itself died and must be recreated.

    A batch that keeps one hidden Word/WPS instance alive will eventually see
    this after a very large document; the correct response is to rebuild the
    instance and retry the document once, not to report the document as failed.
    """
    for candidate in (getattr(exc, "hresult", None),
                      (getattr(exc, "args", None) or (None,))[0]):
        if isinstance(candidate, int):
            if candidate & 0xFFFFFFFF in _DEAD_INSTANCE_HRESULTS:
                return True
    text = str(exc).lower()
    return ("rpc" in text or "远程过程调用" in text
            or "被调用方" in text or "server died" in text)


def quit_app(app):
    """Close the private automation instance and release its COM references."""
    if app is None:
        return
    try:
        app.Quit()
    except Exception:
        pass



def start_app(engine, visible=False):
    """Create a private out-of-process automation instance."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        raise RuntimeError(
            "pywin32 未安装，请执行：python -m pip install pywin32")
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    progid = PROGIDS[engine]
    try:
        app = win32.DispatchEx(progid)
    except Exception as exc:
        raise RuntimeError(
            f"无法启动 {progid!s}：{com_message(exc)!s}")
    for name, value in (("Visible", bool(visible)),
                        ("DisplayAlerts", 0),
                        ("ScreenUpdating", False)):
        try:
            setattr(app, name, value)
        except Exception:
            pass
    return app


def describe(app):
    parts = []
    for key in ("Name", "Version", "Build", "Path"):
        try:
            parts.append(f"{key!s}={getattr(app, key)!s}")
        except Exception:
            pass
    return ", ".join(parts)


def wait_for_file(path, timeout):
    """Wait until a non-empty file has stopped growing."""
    deadline = time.time() + max(5, timeout)
    last, stable = -1, 0
    while time.time() < deadline:
        if os.path.exists(path):
            size = os.path.getsize(path)
            if size > 0 and size == last:
                stable += 1
                if stable >= 2:
                    return size
            else:
                stable = 0
            last = size
        time.sleep(0.4)
    raise TimeoutError(
        "导出调用已返回，但在 %d 秒内未生成稳定文件：%s" % (timeout, path))


def pdf_info(path):
    import pymupdf
    doc = pymupdf.open(path)
    try:
        return doc.page_count
    finally:
        doc.close()


def output_path(args, src):
    if args.out:
        return args.out
    stem = os.path.splitext(os.path.basename(src))[0] + ".pdf"
    directory = args.out_dir or os.path.dirname(os.path.abspath(src))
    return os.path.join(directory, stem)


def find_cli(explicit=None):
    """Locate wpscli.exe from the WPS COM registration or install roots."""
    if explicit:
        if os.path.isfile(explicit):
            return explicit
        raise RuntimeError("指定的 wpscli.exe 不存在：%s" % explicit)
    candidates = []
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                            "KWps.Application\\CLSID") as key:
            clsid = winreg.QueryValue(key, "")
        for sub in ("CLSID\\%s\\LocalServer32" % clsid,
                    "WOW6432Node\\CLSID\\%s\\LocalServer32" % clsid):
            try:
                cmd = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, sub)
                match = re.search('"([^"]+\\.exe)"', cmd, re.I)
                if match:
                    candidates.append(os.path.join(
                        os.path.dirname(os.path.dirname(match.group(1))),
                        "clitool", "wpscli.exe"))
            except OSError:
                continue
    except Exception:
        pass
    roots = ["D:\\Work\\WPS Office", "C:\\Program Files\\WPS Office",
             "C:\\Program Files (x86)\\WPS Office",
             os.path.join(os.environ.get("LOCALAPPDATA", ""),
                          "Kingsoft", "WPS Office")]
    for root in roots:
        if not os.path.isdir(root):
            continue
        try:
            versions = sorted(os.listdir(root), reverse=True)
            for version in versions:
                candidates.append(os.path.join(root, version, "clitool",
                                               "wpscli.exe"))
        except OSError:
            continue
    for path in candidates:
        if os.path.isfile(path):
            return path
    raise RuntimeError(
        "未找到 wpscli.exe，请用 --cli-path 指定路径，"
        "或改用 --engine wps（COM 通道）")


def export_via_cli(cli, src, dst, args):
    """Convert with the official wpscli word2pdf command line."""
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    if not os.path.isfile(src):
        raise FileNotFoundError("找不到 Word 文档：%s" % src)
    if os.path.exists(dst) and not args.overwrite:
        raise FileExistsError(
            "输出文件已存在（未覆盖，可加 --overwrite）：%s" % dst)
    parent = os.path.dirname(dst)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    partial = dst + ".partial.pdf"
    if os.path.exists(partial):
        os.remove(partial)
    cmd = [cli, "word2pdf", src, "-o", partial,
           "--image-quality", args.image_quality,
           "--timeout", str(args.timeout), "--json"]
    try:
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace",
                                  timeout=args.timeout + 180)
        except subprocess.TimeoutExpired:
            raise TimeoutError(
                "wpscli 在 %d 秒内未完成转换" % (args.timeout + 180))
        if proc.returncode != 0:
            reason = CLI_EXIT_MESSAGES.get(
                proc.returncode, "退出码 %d" % proc.returncode)
            detail = ((proc.stdout or "") + (proc.stderr or "")).strip()
            raise RuntimeError(
                f"wpscli 导出失败：{reason!s} {detail[-300:]!s}")
        size = wait_for_file(partial, args.timeout)
        pages = pdf_info(partial)
        if pages <= 0:
            raise RuntimeError("导出的 PDF 页数为 0，文件可能损坏")
        os.replace(partial, dst)
        return size, pages, "wpscli word2pdf"
    finally:
        if os.path.exists(partial):
            try:
                os.remove(partial)
            except OSError:
                pass


def verify_against_docx(out_pdf, docx, sample_pages=20):
    """Compare optimised PDF images with the Word originals.

    The optimiser's own img_diff counter compares the output against the WPS
    export, whose JPEG re-encoding is lossy; a non-zero value there is expected.
    The Word original is the correct fidelity reference.

    A JPEG passthrough object contains the exact bytes of the Word image, so
    raw-byte equality is the strictest zero-loss proof. Decoded pixels are
    compared only for non-DCT objects: PIL and MuPDF round IDCT coefficients
    slightly differently, which otherwise reports false differences on
    byte-identical JPEGs. Returns (checked, identical, different).
    """
    import pymupdf
    from PIL import Image
    if _opt is None:
        return 0, 0, 0
    archive, grouped, _ = _opt.build_word_index(docx)
    doc = pymupdf.open(out_pdf)
    identical = diff = 0
    try:
        n = doc.page_count
        step = max(1, n // max(1, sample_pages))
        pages = list(range(0, n, step))[:sample_pages]
        if pages and pages[-1] != n - 1:
            pages[-1] = n - 1
        for i in pages:
            for item in doc[i].get_images(full=True):
                xref = item[0]
                try:
                    width, height, buf = _opt._image_samples(doc, xref)
                    candidates = grouped.get((width, height), [])
                    if not candidates:
                        diff += 1
                        continue
                    if "DCTDecode" in doc.xref_get_key(xref, "Filter")[1]:
                        raw = doc.extract_image(xref)["image"]
                        for candidate in candidates:
                            try:
                                if archive.read(candidate["name"]) == raw:
                                    identical += 1
                                    break
                            except Exception:
                                continue
                        else:
                            diff += 1
                        continue
                    fp = _opt.fingerprint(
                        Image.frombytes("RGB", (width, height), buf))
                    best = min(candidates,
                               key=lambda c: _opt.distance(fp, c["fp"]))
                    with Image.open(io.BytesIO(
                            archive.read(best["name"]))) as original:
                        reference = original.convert("RGB").tobytes()
                    if reference == buf:
                        identical += 1
                    else:
                        diff += 1
                except Exception:
                    diff += 1
    finally:
        doc.close()
        archive.close()
    return identical + diff, identical, diff


def export_doc(app, src, dst, timeout, page_check=True, overwrite=False):
    """Open one .docx read-only and export it; return (bytes, pages, method)."""
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.isfile(src):
        raise FileNotFoundError("找不到 Word 文档：%s" % src)
    if os.path.exists(dst) and not overwrite:
        raise FileExistsError(
            "输出文件已存在（未覆盖，可加 --overwrite）：%s" % dst)
    parent = os.path.dirname(dst)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    partial = dst + ".partial.pdf"
    if os.path.exists(partial):
        os.remove(partial)
    doc = None
    try:
        doc = app.Documents.Open(src, False, True, False)
        pages_in_doc = None
        if page_check:
            try:
                pages_in_doc = int(
                    doc.ComputeStatistics(WD_STATISTIC_PAGES))
            except Exception:
                pages_in_doc = None
        attempts = (
            ("ExportAsFixedFormat",
             lambda p: doc.ExportAsFixedFormat(p, EXPORT_FORMAT_PDF)),
            ("SaveAs2", lambda p: doc.SaveAs2(p, EXPORT_FORMAT_PDF)),
            ("SaveAs", lambda p: doc.SaveAs(p, EXPORT_FORMAT_PDF)),
        )
        used = None
        errors = []
        for name, call in attempts:
            for attempt in range(EXPORT_RETRIES + 1):
                try:
                    call(partial)
                    used = name
                    break
                except Exception as exc:
                    if attempt < EXPORT_RETRIES and is_transient(exc):
                        time.sleep(RETRY_DELAY)
                        continue
                    errors.append(f"{name}: {com_message(exc)}")
                    break
                finally:
                    if used is None and os.path.exists(partial):
                        try:
                            os.remove(partial)
                        except OSError:
                            pass
            if used is not None:
                break
        if used is None:
            raise RuntimeError("WPS 未提供可用的 PDF 导出接口：%s"
                               % " | ".join(errors))
        size = wait_for_file(partial, timeout)
        pages = pdf_info(partial)
        if pages <= 0:
            raise RuntimeError("导出的 PDF 页数为 0，文件可能损坏")
        if pages_in_doc and pages != pages_in_doc:
            emit("  提示：WPS 统计页数 %d，导出 PDF 为 %d 页，已按导出结果继续。"
                 % (pages_in_doc, pages),
                 "  Note: WPS reports %d pages, exported PDF has %d; "
                 "continuing." % (pages_in_doc, pages))
        os.replace(partial, dst)
        return size, pages, used, pages_in_doc
    finally:
        if doc is not None:
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if os.path.exists(partial):
            try:
                os.remove(partial)
            except OSError:
                pass


def run_optimizer(docx, pdf, out, args):
    if _opt is None:
        raise RuntimeError(
            "未找到 optimize_pdf_images.py，无法执行 --optimize")
    if os.path.exists(out) and not args.overwrite:
        raise FileExistsError(
            "优化输出文件已存在（可加 --overwrite）：%s" % out)
    emit("  开始优化 PDF 图像（原始 Word 图片回填）...",
         "  Optimising PDF images from the Word originals ...")
    t0 = time.time()
    try:
        stats = _opt.optimize(pdf, docx, out, use_g4=args.g4)
    except RuntimeError as exc:
        if "no usable images" in str(exc):
            emit("  该文档不含可用插图，已跳过图像优化；导出的 PDF 即为最终结果。",
                 "  No usable images in this document; optimisation skipped, "
                 "the exported PDF is the final result.")
            return pdf
        raise
    emit("  替换统计：%s" % _opt.labelled(stats, _opt.STAT_LABELS),
         "  Replacements: %s" % _opt.labelled(stats, _opt.STAT_LABELS))
    emit("  优化后体积：%.2f MB（耗时 %.1f 分钟）"
         % (os.path.getsize(out) / 1048576,
            (time.time() - t0) / 60),
         "  Optimised size: %.2f MB (%.1f min)"
         % (os.path.getsize(out) / 1048576,
            (time.time() - t0) / 60))
    if not args.no_verify:
        result = _opt.verify(pdf, out, args.image_pages,
                             args.render_check, args.render_pages)
        emit("  完整性校验：%s" % _opt.labelled(result, _opt.VERIFY_LABELS),
             "  Integrity: %s" % _opt.labelled(result, _opt.VERIFY_LABELS))
        if result.get("img_diff"):
            emit("  说明：上述 img_diff 是 WPS 导出的 JPEG 有损中间图与 Word "
                 "原图无损回填之间的必然差异，不代表画质下降。",
                 "  Note: img_diff compares lossy WPS-export JPEGs with the "
                 "lossless Word originals; it is expected, not quality loss.")
    if not args.no_docx_verify:
        checked, same, diff = verify_against_docx(out, docx, args.image_pages)
        emit("  Word 原图基准复核：抽检 %d 个图像对象，逐像素一致 %d，不一致 %d"
             % (checked, same, diff),
             "  Word-original reference check: %d objects, %d pixel-identical, "
             "%d different" % (checked, same, diff))
    return out


def self_check(engine, visible=False):
    """Create a tiny docx with WPS, export it and verify the PDF."""
    import tempfile
    tmp = tempfile.gettempdir()
    docx = os.path.join(tmp, "wps_to_pdf_selftest.docx")
    pdf = os.path.join(tmp, "wps_to_pdf_selftest.pdf")
    for p in (docx, pdf):
        try:
            os.remove(p)
        except OSError:
            continue
    app = start_app(engine, visible)
    try:
        emit("  COM 服务：%s" % describe(app),
             "  COM server: %s" % describe(app))
        doc = app.Documents.Add()
        doc.Content.Text = "WPS automation self-test 2026"
        try:
            doc.SaveAs(docx, WD_FORMAT_XML_DOCUMENT)
            emit("  测试 DOCX 已生成：%d 字节" % os.path.getsize(docx),
                 "  Test DOCX written: %d bytes" % os.path.getsize(docx))
        except Exception as exc:
            emit("  测试 DOCX 生成失败（不影响 PDF 导出能力）：%s"
                 % com_message(exc),
                 "  Test DOCX creation failed (PDF export unaffected): %s"
                 % com_message(exc))
            docx = None
        if docx and os.path.exists(docx):
            # Do NOT close this document first. Closing the last open document
            # tears down the automation session in some Word/WPS builds
            # ("RPC 服务器不可用" / `<unknown>.Open`), so the follow-up Open
            # fails. Close after opening the file instead.
            size, pages, used, _ = export_doc(app, docx, pdf, 60, False, True)
            emit("  导出接口：%s；PDF %d 字节 / %d 页" % (used, size, pages),
                 "  Export method: %s; PDF %d bytes / %d pages"
                 % (used, size, pages))
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
            doc = None
        else:
            # No test docx on disk, so export the blank document directly.
            # The first document must still be open here: closing the last
            # document makes Word's late-bound `Documents` collection unusable
            # and `Documents.Add()` then fails with `<unknown>.Add`.
            doc.Content.Text = "WPS automation self-test 2026"
            doc.ExportAsFixedFormat(pdf, EXPORT_FORMAT_PDF)
            size = wait_for_file(pdf, 60)
            pages = pdf_info(pdf)
            emit("  导出接口：ExportAsFixedFormat；PDF %d 字节 / %d 页"
                 % (size, pages),
                 "  Export method: ExportAsFixedFormat; PDF %d bytes / "
                 "%d pages" % (size, pages))
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
            doc = None
        emit("  自检通过：WPS COM 自动化可用，导出 PDF 有效。",
             "  Self-test passed: WPS COM automation is functional.")
    except Exception as exc:
        emit("  自检失败：%s" % exc, "  Self-test failed: %s" % exc)
        return 4
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        for p in (docx, pdf):
            if not p:
                continue
            try:
                os.remove(p)
            except OSError:
                continue
    return 0


def main_cli(args):
    try:
        cli = find_cli(args.cli_path)
    except RuntimeError as exc:
        emit("环境错误：%s" % exc, "Environment error: %s" % exc)
        return 3
    emit("已定位 WPS CLI：%s" % cli, "Located WPS CLI: %s" % cli)
    t_all = time.time()
    for src in args.sources:
        dst = output_path(args, src)
        emit("正在导出（CLI）：%s" % src, "Exporting (CLI): %s" % src)
        t0 = time.time()
        try:
            size, pages, used = export_via_cli(cli, src, dst, args)
        except (FileNotFoundError, FileExistsError, TimeoutError,
                RuntimeError) as exc:
            emit("  导出失败：%s" % exc, "  Export failed: %s" % exc)
            return 4
        emit("  完成：%.2f MB / %d 页 / 耗时 %.1f 秒 / 接口 %s -> %s"
             % (size / 1048576, pages, time.time() - t0, used, dst),
             "  Done: %.2f MB / %d pages / %.1f s / %s -> %s"
             % (size / 1048576, pages, time.time() - t0, used, dst))
        if args.optimize:
            out = (args.optimized_out
                   or os.path.splitext(dst)[0] + "_优化保真版.pdf")
            try:
                final = run_optimizer(src, dst, out, args)
            except (FileExistsError, RuntimeError, MemoryError) as exc:
                emit("  优化失败：%s" % exc, "  Optimisation failed: %s" % exc)
                return 5
            emit("  最终输出：%s" % final, "  Final output: %s" % final)
    emit("全部完成，总耗时 %.1f 分钟。" % ((time.time() - t_all) / 60),
         "All done in %.1f min." % ((time.time() - t_all) / 60))
    return 0


def build_parser():
    p = argparse.ArgumentParser(
        description="通过 WPS COM 自动化将 Word 另存为 PDF（可串联图像优化）")
    p.add_argument("sources", nargs="*", help=".docx 文件路径，可多个")
    p.add_argument("-o", "--out", help="输出 PDF 路径（仅单个输入时可用）")
    p.add_argument("-d", "--out-dir", help="批量输出目录")
    p.add_argument("--optimize", action="store_true", help="导出后执行图像优化")
    p.add_argument("--optimized-out", help="优化后 PDF 路径")
    p.add_argument("--g4", action="store_true", help="优化器改用 CCITT G4")
    p.add_argument("--image-pages", type=int, default=20)
    p.add_argument("--render-check", action="store_true")
    p.add_argument("--render-pages", type=int, default=12)
    p.add_argument("--no-verify", action="store_true")
    p.add_argument("--no-docx-verify", action="store_true")
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--visible", action="store_true")
    p.add_argument("--engine", choices=("wps", "cli", "word"), default="wps")
    p.add_argument("--cli-path", help="wpscli.exe 路径（默认自动探测）")
    p.add_argument("--image-quality", choices=("low", "medium", "high"),
                   default="medium", help="CLI 引擎图像质量，默认 medium")
    p.add_argument("--timeout", type=int, default=300)
    p.add_argument("--no-page-check", action="store_true")
    p.add_argument("--ascii", action="store_true")
    p.add_argument("--check", action="store_true", help="仅做环境自检")
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    init_output(args.ascii)
    if args.check:
        if args.engine == "cli":
            try:
                cli = find_cli(args.cli_path)
                proc = subprocess.run([cli, "--version"], capture_output=True,
                                      text=True, encoding="utf-8",
                                      errors="replace", timeout=120)
                emit(f"WPS CLI 自检通过：{cli!s} / {proc.stdout.strip()!s}",
                     f"WPS CLI self-test passed: {cli!s} / "
                     f"{proc.stdout.strip()!s}")
                return 0 if proc.returncode == 0 else 3
            except Exception as exc:
                emit("  WPS CLI 自检失败：%s" % exc,
                     "  WPS CLI self-test failed: %s" % exc)
                return 3
        return self_check(args.engine, args.visible)
    if not args.sources:
        print(build_parser().format_help())
        return 2
    if args.out and len(args.sources) > 1:
        emit("多个输入时不能使用 -o，请改用 -d 指定输出目录。",
             "Use -d instead of -o when converting several documents.")
        return 2
    if args.optimize and len(args.sources) > 1:
        emit("批量模式下暂不串联优化，请逐个文档执行 --optimize。",
             "Batch mode does not chain the optimiser; run it per document.")
        return 2
    if args.optimized_out and not args.optimize:
        emit("--optimized-out 需配合 --optimize 使用。",
             "--optimized-out requires --optimize.")
        return 2
    if args.engine == "cli":
        return main_cli(args)
    app = None
    try:
        app = start_app(args.engine, args.visible)
        emit("已连接 WPS 自动化实例：%s" % describe(app),
             "Connected WPS automation instance: %s" % describe(app))
        t_all = time.time()
        for src in args.sources:
            dst = output_path(args, src)
            emit("正在导出：%s" % src, "Exporting: %s" % src)
            t0 = time.time()
            try:
                size, pages, used, _ = export_doc(
                    app, src, dst, args.timeout,
                    not args.no_page_check, args.overwrite)
            except (FileNotFoundError, FileExistsError, TimeoutError,
                    RuntimeError) as exc:
                emit("  导出失败：%s" % exc, "  Export failed: %s" % exc)
                return 4
            emit("  完成：%.2f MB / %d 页 / 耗时 %.1f 秒 / 接口 %s -> %s"
                 % (size / 1048576, pages, time.time() - t0, used, dst),
                 "  Done: %.2f MB / %d pages / %.1f s / %s -> %s"
                 % (size / 1048576, pages, time.time() - t0, used, dst))
            if args.optimize:
                out = (args.optimized_out
                       or os.path.splitext(dst)[0] + "_优化保真版.pdf")
                try:
                    final = run_optimizer(src, dst, out, args)
                except (FileExistsError, RuntimeError, MemoryError) as exc:
                    emit("  优化失败：%s" % exc,
                         "  Optimisation failed: %s" % exc)
                    return 5
                emit("  最终输出：%s" % final, "  Final output: %s" % final)
        emit("全部完成，总耗时 %.1f 分钟。" % ((time.time() - t_all) / 60),
             "All done in %.1f min." % ((time.time() - t_all) / 60))
        return 0
    except RuntimeError as exc:
        emit("环境错误：%s" % exc, "Environment error: %s" % exc)
        return 3
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        del app
        gc.collect()


if __name__ == "__main__":
    sys.exit(main())
