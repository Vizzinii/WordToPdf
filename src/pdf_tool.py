# -*- coding: utf-8 -*-
"""
pdf_tool.py -- Word/PDF 一体化小工具（可打包为单个 exe）。

V1.3 能力：
  ① 单个 Word 文档 -> 最终优化保真版 PDF；
  ② 整个文件夹批量转换 -> 每个 Word 各自输出最终 PDF（可选自动合并）；
  ③ 多个 PDF 按列表顺序合并为 1 个 PDF（顺序可上移/下移调整）。
  界面提供进度条、阶段状态、可回看的历史日志框，以及「点击取消」。
  V1.3 相对 V1.2 的效率改造见同目录 README 的「版本变更」一节。

流程：选择 .docx 后，在 Word 所在目录依次执行
      ① WPS COM (KWps.Application) 另存为中间 PDF；
      ② optimize_pdf_images.py 用 Word 原图无损回填，输出「优化保真版」；
      ③ 完整性与画质复核通过后，默认删除中间 PDF（可在帮助窗口勾选保留）。
帮助窗口提供转换方式选项（WPS COM / WPS CLI / Microsoft Word COM）、
「保留中间 PDF」与「批量完成后自动合并结果 PDF」复选框。

命令行自测（打包后亦可用）：
    pdf_tool.py --selftest <path.docx|folder> [wps|cli|word] [keep]
    pdf_tool.py --selftest-merge <out.pdf> <p1.pdf> <p2.pdf> ...
结果写入系统临时目录 pdf_tool_selftest.log。

依赖：pywin32、pymupdf、pillow；GUI 仅用标准库 ctypes + Win32 API。
"""

from __future__ import annotations

import ctypes
import csv
import io
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import traceback
import uuid
from ctypes import wintypes
from types import SimpleNamespace

# A --windowed PyInstaller build has no console; some libraries touch stdout.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

import wps_to_pdf as w2p
import optimize_pdf_images as opt

MAX_SIZE = 2 * 1024 * 1024 * 1024          # 2 GiB
CLI_MAX_SIZE = 200 * 1024 * 1024           # official wpscli hard limit
MAX_SIZE_TEXT = "2 GB（2,147,483,648 字节）"

# ---------------------------------------------------------------- Win32 setup
user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
comdlg32 = ctypes.WinDLL("comdlg32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
ole32 = ctypes.WinDLL("ole32", use_last_error=True)

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT,
                             wintypes.WPARAM, wintypes.LPARAM)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class PAINTSTRUCT(ctypes.Structure):
    _fields_ = [
        ("hdc", wintypes.HDC),
        ("fErase", wintypes.BOOL),
        ("rcPaint", wintypes.RECT),
        ("fRestore", wintypes.BOOL),
        ("fIncUpdate", wintypes.BOOL),
        ("rgbReserved", ctypes.c_byte * 32),
    ]


class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", wintypes.DWORD),
        ("nFilterIndex", wintypes.DWORD),
        ("lpstrFile", ctypes.c_wchar_p),
        ("nMaxFile", wintypes.DWORD),
        ("lpstrFileTitle", ctypes.c_wchar_p),
        ("nMaxFileTitle", wintypes.DWORD),
        ("lpstrInitialDir", wintypes.LPCWSTR),
        ("lpstrTitle", wintypes.LPCWSTR),
        ("Flags", wintypes.DWORD),
        ("nFileOffset", wintypes.WORD),
        ("nFileExtension", wintypes.WORD),
        ("lpstrDefExt", wintypes.LPCWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", wintypes.LPVOID),
        ("lpTemplateName", wintypes.LPCWSTR),
        ("pvReserved", wintypes.LPVOID),
        ("dwReserved", wintypes.DWORD),
        ("FlagsEx", wintypes.DWORD),
    ]


class BROWSEINFOW(ctypes.Structure):
    _fields_ = [
        ("hwndOwner", wintypes.HWND),
        ("pidlRoot", ctypes.c_void_p),
        ("pszDisplayName", wintypes.LPWSTR),
        ("lpszTitle", wintypes.LPCWSTR),
        ("ulFlags", wintypes.UINT),
        ("lpfn", ctypes.c_void_p),
        ("lParam", wintypes.LPARAM),
        ("iImage", ctypes.c_int),
    ]


user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
user32.RegisterClassW.restype = wintypes.ATOM
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wintypes.HWND, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.LPVOID]
user32.CreateWindowExW.restype = wintypes.HWND
user32.LoadCursorW.argtypes = [wintypes.HINSTANCE, ctypes.c_void_p]
user32.LoadCursorW.restype = wintypes.HANDLE
user32.DrawTextW.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int,
                             ctypes.POINTER(wintypes.RECT), wintypes.UINT]
user32.DrawTextW.restype = ctypes.c_int
user32.FillRect.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT),
                            wintypes.HBRUSH]
user32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.SystemParametersInfoW.argtypes = [wintypes.UINT, wintypes.UINT,
                                     ctypes.c_void_p, wintypes.UINT]
user32.SystemParametersInfoW.restype = wintypes.BOOL
user32.MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR,
                              wintypes.LPCWSTR, wintypes.UINT]
user32.MessageBoxW.restype = ctypes.c_int
user32.BeginPaint.argtypes = [wintypes.HWND, ctypes.POINTER(PAINTSTRUCT)]
user32.BeginPaint.restype = wintypes.HDC
user32.EndPaint.argtypes = [wintypes.HWND, ctypes.POINTER(PAINTSTRUCT)]
user32.InvalidateRect.argtypes = [wintypes.HWND, ctypes.c_void_p,
                                  wintypes.BOOL]
user32.SetWindowRgn.argtypes = [wintypes.HWND, wintypes.HRGN, wintypes.BOOL]
user32.SetTimer.argtypes = [wintypes.HWND, ctypes.c_size_t, wintypes.UINT,
                            ctypes.c_void_p]
user32.SetTimer.restype = ctypes.c_size_t
user32.KillTimer.argtypes = [wintypes.HWND, ctypes.c_size_t]
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM,
                                wintypes.LPARAM]
user32.SendMessageW.restype = LRESULT
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM,
                                wintypes.LPARAM]
user32.PostQuitMessage.argtypes = [ctypes.c_int]
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND,
                               wintypes.UINT, wintypes.UINT]
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.GetSystemMetrics.argtypes = [ctypes.c_int]
user32.GetSystemMetrics.restype = ctypes.c_int
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.IsWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.restype = wintypes.BOOL
user32.IsWindow.restype = wintypes.BOOL
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR,
                                  ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
comdlg32.GetOpenFileNameW.argtypes = [ctypes.POINTER(OPENFILENAMEW)]
comdlg32.GetOpenFileNameW.restype = wintypes.BOOL

shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFOW)]
shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]
shell32.SHGetPathFromIDListW.restype = wintypes.BOOL
ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
ole32.CoInitializeEx.restype = ctypes.c_long
ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]

gdi32.CreateSolidBrush.argtypes = [wintypes.COLORREF]
gdi32.CreateSolidBrush.restype = wintypes.HBRUSH
gdi32.CreateFontW.restype = wintypes.HFONT
gdi32.CreateFontW.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_int,
                              ctypes.c_int, ctypes.c_int, wintypes.DWORD,
                              wintypes.DWORD, wintypes.DWORD, wintypes.DWORD,
                              wintypes.DWORD, wintypes.DWORD, wintypes.DWORD,
                              wintypes.DWORD, wintypes.LPCWSTR]
gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC
gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int,
                                         ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.SetTextColor.argtypes = [wintypes.HDC, wintypes.COLORREF]
gdi32.SetBkColor.argtypes = [wintypes.HDC, wintypes.COLORREF]
gdi32.SetBkMode.argtypes = [wintypes.HDC, ctypes.c_int]
gdi32.CreatePen.argtypes = [ctypes.c_int, ctypes.c_int, wintypes.COLORREF]
gdi32.CreatePen.restype = wintypes.HPEN
gdi32.RoundRect.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int,
                            ctypes.c_int, ctypes.c_int, ctypes.c_int,
                            ctypes.c_int]
gdi32.BitBlt.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int,
                         ctypes.c_int, ctypes.c_int, wintypes.HDC,
                         ctypes.c_int, ctypes.c_int, wintypes.DWORD]
gdi32.CreateRoundRectRgn.argtypes = [ctypes.c_int, ctypes.c_int,
                                     ctypes.c_int, ctypes.c_int,
                                     ctypes.c_int, ctypes.c_int]
gdi32.CreateRoundRectRgn.restype = wintypes.HRGN
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteObject.restype = wintypes.BOOL
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteDC.restype = wintypes.BOOL

kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE

user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT,
                                  wintypes.WPARAM, wintypes.LPARAM]
user32.DefWindowProcW.restype = LRESULT
user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND,
                                ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                ctypes.c_int, wintypes.UINT]

# ------------------------------------------------------------------ constants
WM_CREATE, WM_DESTROY, WM_PAINT, WM_TIMER = 0x0001, 0x0002, 0x000F, 0x0113
WM_COMMAND, WM_CTLCOLORSTATIC, WM_SETFONT = 0x0111, 0x0138, 0x0030
WM_CTLCOLORLISTBOX = 0x0134
WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP = 0x0200, 0x0201, 0x0202
WM_CLOSE, WM_SETTEXT, WM_GETTEXT = 0x0010, 0x000C, 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_NCLBUTTONDOWN, HTCAPTION = 0x00A1, 2
BM_SETCHECK, BM_GETCHECK = 0x00F1, 0x00F0
WS_POPUP, WS_VISIBLE, WS_CHILD = 0x80000000, 0x10000000, 0x40000000
WS_VSCROLL, WS_GROUP, WS_TABSTOP = 0x00200000, 0x00020000, 0x00010000
WS_CLIPCHILDREN, WS_BORDER = 0x02000000, 0x00800000
ES_MULTILINE, ES_READONLY, ES_AUTOVSCROLL = 0x0004, 0x0800, 0x0040
ES_AUTOHSCROLL, ES_NOHIDESEL = 0x0080, 0x0100
BS_AUTORADIOBUTTON, BS_AUTOCHECKBOX, BS_PUSHBUTTON = 0x0009, 0x0003, 0x0000
CS_DROPSHADOW, CS_HREDRAW, CS_VREDRAW = 0x00020000, 0x0002, 0x0001
DT_LEFT, DT_CENTER, DT_VCENTER = 0x0000, 0x0001, 0x0004
DT_WORDBREAK, DT_SINGLELINE = 0x0010, 0x0020
SRCCOPY = 0x00CC0020
OFN = 0x00080000 | 0x00001000 | 0x00000004 | 0x00000008
OFN_ALLOWMULTISELECT = 0x00000200
BIF_RETURNONLYFSDIRS, BIF_EDITBOX, BIF_NEWDIALOGSTYLE = 0x0001, 0x0010, 0x0040
EM_SETSEL, EM_SCROLLCARET, EM_REPLACESEL = 0x00B1, 0x00B7, 0x00C2
EM_SETLIMITTEXT = 0x00C5
LB_ADDSTRING, LB_DELETESTRING = 0x0180, 0x0182
LB_RESETCONTENT, LB_SETCURSEL = 0x0184, 0x0186
LB_GETCURSEL, LB_GETCOUNT = 0x0188, 0x018B
LBS_NOTIFY, LBS_NOINTEGRALHEIGHT = 0x0001, 0x0100

BG = 0x00F7F8FA
WHITE = 0x00FFFFFF
INK = 0x00222226
MUTED = 0x0073737A
LINE = 0x00DFE2E7
MODE_OFF = 0x00EAECEF
BTN = 0x002A2D33
BTN_HOVER = 0x00444850
BTN_DISABLED = 0x00B9BDC4
ACCENT = 0x003A6EA5
ACCENT_LIGHT = 0x00DCE8F4
OK_GREEN = 0x002E7D5B
WARN = 0x00B26A16

# V1.3 layout: larger window, progress bar, scrollable log box, cancel button.
MAIN_W, MAIN_H = 980, 760
HELP_W, HELP_H = 940, 980
MERGE_W, MERGE_H = 940, 720
TITLE_H = 60
MODE_FILE = (30, 78, 330, 146)
MODE_DIR = (340, 78, 640, 146)
MODE_MERGE = (650, 78, 950, 146)
MAIN_BTN = (30, 160, 950, 234)
STATUS_BOX = (30, 248, 950, 288)
PROGRESS_BOX = (30, 298, 950, 330)
LOG_TITLE = (30, 340, 950, 366)
LOG_BOX = (30, 368, 950, 724)
MAIN_HINT = (30, 732, 950, 754)
MAIN_HELP = (886, 0, 932, TITLE_H)
MAIN_CLOSE = (932, 0, 980, TITLE_H)
HELP_CLOSE = (892, 0, 940, TITLE_H)
MERGE_CLOSE = (892, 0, 940, TITLE_H)

MERGE_LIST = (30, 110, 910, 470)
MERGE_BTNS = {
    "add": (30, 482, 198, 530),
    "up": (208, 482, 376, 530),
    "down": (386, 482, 554, 530),
    "remove": (564, 482, 732, 530),
    "clear": (742, 482, 910, 530),
}
MERGE_NAME_EDIT = (196, 542, 910, 582)
MERGE_GO = (30, 632, 910, 690)

ENGINE_NAMES = {
    "wps": "WPS COM（默认，推荐）",
    "cli": "WPS CLI（需 VIP）",
    "word": "Microsoft Word COM",
}
MODE_LABELS = {
    "file": "选择 Word 文档（.docx）",
    "dir": "选择文件夹（批量转换）",
    "merge": "选择并编排多个 PDF（合并）",
}
MODE_SHORT = {
    "file": "单个 Word 文件",
    "dir": "整个文件夹（批量）",
    "merge": "合并多个 PDF",
}

Q = queue.Queue()
STATE = {
    "status": "就绪｜请先选择处理模式，再点击下方主按钮。",
    "busy": False, "start": 0.0, "elapsed": "",
    "hover": None, "progress": 0.0, "progress_label": "等待任务",
    "logs": [], "pending_logs": [], "mode_busy": False,
}
ENGINE = "wps"
MODE = "file"
KEEP_PDF = False
AUTO_MERGE = False
CANCEL = threading.Event()
HELP_LAYOUT = {}
_SCRATCH = None
LOG_HANDLE = None
LOG_PATH = None
HINST = kernel32.GetModuleHandleW(None)
HWND_MAIN = None
HWND_HELP = None
HWND_MERGE = None
HWND_LOG = None
HWND_HELP_EDIT = None
HWND_MERGE_LIST = None
HWND_MERGE_NAME = None
RADIO_HWNDS = {}
CHECK_HWND = None
AUTO_MERGE_HWND = None
FONTS = {}
BRUSH_BG = gdi32.CreateSolidBrush(BG)
BRUSH_WHITE = gdi32.CreateSolidBrush(WHITE)
MERGE_FILES = []            # [(path, page_count), ...] in display order

_HELP_TEXT = """【一句话说明】
本工具把「WPS 手工导出 PDF → 图像保真优化」两步做成一键操作，并支持按指定顺序把多份 PDF 合并成一份。主界面下方是实时日志，可向上滚动查看全部历史记录；进度条显示当前阶段。

【三种处理模式】
一、单个 Word 文件：选择一个 .docx，工具在它所在目录产出一版「同名_优化保真版.pdf」。
二、整个文件夹（批量）：选择文件夹后递归收集其中所有 .docx（含子文件夹，自动跳过 ~$ 临时文件），按文件名顺序逐个处理，每个 Word 的结果保存在各自所在目录。日志与状态区显示「[第几个 / 总数]」，单个文档失败不会中断其余文档。
三、合并多个 PDF：点击主按钮打开「PDF 合并」窗口，用「添加 PDF」一次选择多份文件；列表中的顺序就是最终拼接顺序，可用「上移 / 下移 / 移除 / 清空」调整。点击「开始合并」后按列表顺序拼接，输出文件名可在窗口中修改，默认保存在列表中第一个 PDF 所在目录。

【每个 Word 的三步处理】
第一步，通过 WPS 的 COM 自动化接口 KWps.Application 调用 ExportAsFixedFormat，把文档另存为中间 PDF（临时名，不作为交付物）；
第二步，调用图像优化脚本，用 Word 包内的原始图片替换 PDF 中被重新编码的图像，输出「优化保真版」；
第三步，做完整性校验与以 Word 原图为基准的画质复核，通过后默认删除中间 PDF。
最终只保留 同名_优化保真版.pdf。若同名文件已存在，会自动追加序号，绝不覆盖已有文件。

【中间 PDF 的处置口径】
中间 PDF 只提供页面结构与校验基准，不是交付物。只要「最终版已生成、完整性校验（页数 / 字符数 / 图像放置数）三项一致、画质复核未发现像素不一致、未勾选保留」，工具就删除它，日志会明确写出「已删除中间 PDF」及原因。
下列情况会保留并改名为 同名.pdf：使用者勾选「保留中间 PDF」；校验过程异常；完整性三项不一致；画质复核发现真实像素差异；删除操作失败。处理失败时另存为 同名_导出稿.pdf。
画质复核采用分层判据：JPEG 直通对象直接比对 PDF 内的原始字节与 Word 包内原图字节，字节完全相同即为零损失；其余对象比对解码像素。V1.1 曾对 JPEG 直通对象也做解码像素比对，因 PIL 与 PDF 阅读器的 JPEG 解码存在舍入差而误报差异，导致此类文档的中间 PDF 一直被保留，V1.2 已修正。
若某些图像在 Word 包内找不到同尺寸原图（多因 WPS 导出时缩放或裁剪），记为「无法比对」，这属于比对基准限制，不是本次优化引入的差异，因此不阻断删除，日志会单独列出数量。
不含插图的文档没有可优化的图像，此时跳过第二步，直接把导出的 PDF 保留为 同名.pdf；这份 PDF 就是最终交付物，不是遗留的中间文件，日志会明确标注。

【批量自动合并】
勾选本窗口下方的「批量完成后自动合并结果 PDF」后，文件夹批量处理结束时会把本次生成的各份最终 PDF 按处理顺序拼接为一份，输出到所选文件夹根目录，默认名为「批量合并结果.pdf」。单文件模式不触发。未勾选时，仍可随时切到「合并多个 PDF」模式自行选择文件与顺序。

【原理说明】
WPS 导出 PDF 时会把二值扫描图重新编码为 JPEG，既显著放大体积，又引入有损压缩。优化脚本按像素尺寸与灰度指纹逐张匹配，把 Word 原图无损回填：两色图使用 1bpp Flate，少色图使用 8bpp 索引色；文字、字体、矢量表格线与图像位置均不改动，因此是编码层面的无损重排，不是压缩率换画质。
画质以 Word 原图为基准复核。优化脚本日志中的 img_diff 表示「导出 JPEG」与「无损回填」之间的差异，属预期现象。

【其他转换方式（可在下方切换）】
WPS COM：免费、无会员要求，与手工「输出为 PDF」结果一致，本工具默认方式，上限 2 GB。
WPS CLI：金山官方命令行 wpscli word2pdf，首次运行需联网下载组件，且按官方说明必须登录具备 VIP 权限的账号；官方输入上限 200 MB，超过时本工具自动回退 WPS COM。
Microsoft Word COM：本机装有 Microsoft Office 时可用，适合未安装 WPS 的机器，版式渲染由 Word 负责。

【文件大小上限】
2 GB（2,147,483,648 字节），仅作用于 WPS COM 与 Word COM 通道；WPS CLI 的 200 MB 为官方硬限制，无法在本工具内放宽。批量模式下对每个 Word 分别校验。PDF 合并不设体积上限，但超大文件合并时内存占用会上升。

【使用提示】
处理期间窗口可拖动，进度条与日志实时刷新，请不要关闭窗口直到日志出现「全部完成」。仅支持 .docx 作为转换输入；旧版 .doc / .wps 需先另存为 .docx。合并模式接受任意来源的 PDF，不要求由本工具生成。"""


# ------------------------------------------------------------------- drawing
def make_font(size, bold=False):
    return gdi32.CreateFontW(-size, 0, 0, 0, 700 if bold else 400,
                             0, 0, 0, 1, 0, 0, 0, 0, "Microsoft YaHei")


def rect(l, t, r, b):
    return wintypes.RECT(l, t, r, b)


def fill(hdc, box, color, radius=0):
    l, t, r, b = box
    if radius:
        brush = gdi32.CreateSolidBrush(color)
        pen = gdi32.CreatePen(0, 1, color)
        ob = gdi32.SelectObject(hdc, brush)
        op = gdi32.SelectObject(hdc, pen)
        gdi32.RoundRect(hdc, l, t, r, b, radius, radius)
        gdi32.SelectObject(hdc, ob)
        gdi32.SelectObject(hdc, op)
        gdi32.DeleteObject(brush)
        gdi32.DeleteObject(pen)
    else:
        brush = gdi32.CreateSolidBrush(color)
        user32.FillRect(hdc, ctypes.byref(rect(*box)), brush)
        gdi32.DeleteObject(brush)


def text(hdc, content, box, font, color, flags=DT_LEFT | DT_WORDBREAK):
    old = gdi32.SelectObject(hdc, font)
    gdi32.SetTextColor(hdc, color)
    gdi32.SetBkMode(hdc, 1)
    r = rect(*box)
    user32.DrawTextW(hdc, content, -1, ctypes.byref(r), flags)
    gdi32.SelectObject(hdc, old)


def line(hdc, box, color=LINE):
    fill(hdc, box, color)


def paint_begin(hwnd):
    ps = PAINTSTRUCT()
    hdc = user32.BeginPaint(hwnd, ctypes.byref(ps))
    rc = rect(0, 0, 0, 0)
    user32.GetClientRect(hwnd, ctypes.byref(rc))
    mem = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, rc.right, rc.bottom)
    old = gdi32.SelectObject(mem, bmp)
    return hdc, ps, mem, bmp, old, rc


def paint_end(hwnd, hdc, ps, mem, bmp, old, rc):
    gdi32.BitBlt(hdc, 0, 0, rc.right, rc.bottom, mem, 0, 0, SRCCOPY)
    gdi32.SelectObject(mem, old)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem)
    user32.EndPaint(hwnd, ctypes.byref(ps))


def in_box(x, y, box):
    l, t, r, b = box
    return l <= x < r and t <= y < b


def send_text(hwnd, msg, wparam, value):
    """Send a message whose LPARAM is a pointer to a temporary UTF-16 buffer."""
    buf = ctypes.create_unicode_buffer(value)
    return user32.SendMessageW(
        hwnd, msg, wparam, ctypes.cast(buf, ctypes.c_void_p).value)


def draw_mode(mem, box, label, selected, hovered):
    color = BTN if selected else (MODE_OFF if not hovered else 0x00DDE1E6)
    fill(mem, box, color, 10)
    text(mem, label, box, FONTS["mode"],
         WHITE if selected else INK,
         DT_CENTER | DT_VCENTER | DT_SINGLELINE)


def draw_progress(mem):
    l, t, r, b = PROGRESS_BOX
    fill(mem, PROGRESS_BOX, 0x00E7EAEE, 10)
    frac = max(0.0, min(1.0, STATE["progress"]))
    width = int((r - l) * frac)
    if width > 12:
        fill(mem, (l, t, l + width, b),
             OK_GREEN if frac >= 0.999 else ACCENT, 10)
    label = "%d%%  ｜  %s" % (int(frac * 100), STATE["progress_label"])
    text(mem, label, (l + 12, t, r - 12, b), FONTS["progress"], INK,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)


def draw_main(hwnd):
    hdc, ps, mem, bmp, old, rc = paint_begin(hwnd)
    fill(mem, (0, 0, MAIN_W, MAIN_H), BG)
    line(mem, (0, TITLE_H - 1, MAIN_W, TITLE_H), LINE)
    text(mem, "Word → PDF 保真优化工具", (30, 0, 700, TITLE_H),
         FONTS["title"], INK, DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    text(mem, "?", main_help_box(), FONTS["mark"],
         ACCENT if STATE["hover"] == "help" else MUTED,
         DT_CENTER | DT_VCENTER | DT_SINGLELINE)
    text(mem, "×", main_close_box(), FONTS["mark"],
         INK if STATE["hover"] == "close" else MUTED,
         DT_CENTER | DT_VCENTER | DT_SINGLELINE)

    busy = STATE["busy"]
    draw_mode(mem, MODE_FILE, MODE_SHORT["file"], MODE == "file",
              STATE["hover"] == "mode_file" and not busy)
    draw_mode(mem, MODE_DIR, MODE_SHORT["dir"], MODE == "dir",
              STATE["hover"] == "mode_dir" and not busy)
    draw_mode(mem, MODE_MERGE, MODE_SHORT["merge"], MODE == "merge",
              STATE["hover"] == "mode_merge" and not busy)

    btn_hover = STATE["hover"] == "main" and not busy
    if busy:
        color = BTN_DISABLED if CANCEL.is_set() else WARN
    else:
        color = BTN_HOVER if btn_hover else BTN
    fill(mem, MAIN_BTN, color, 12)
    label = ("点击取消当前任务" if CANCEL.is_set() else "处理中，请稍候…（可点击取消）") if busy else MODE_LABELS[MODE]
    text(mem, label, MAIN_BTN, FONTS["button"], WHITE,
         DT_CENTER | DT_VCENTER | DT_SINGLELINE)

    status = STATE["status"]
    if busy and STATE["elapsed"]:
        status = "%s   ·   已用 %s" % (status, STATE["elapsed"])
    text(mem, status, STATUS_BOX, FONTS["status"], INK,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    draw_progress(mem)
    text(mem, "处理日志（实时追加，可向上滚动查看历史记录）", LOG_TITLE,
         FONTS["small"], MUTED, DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    text(mem, "结果均保存在源文件所在目录 ｜ 单文件上限 2 GB ｜ 中间 PDF 校验通过后删除",
         MAIN_HINT, FONTS["small"], MUTED,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    paint_end(hwnd, hdc, ps, mem, bmp, old, rc)


def draw_help(hwnd):
    hdc, ps, mem, bmp, old, rc = paint_begin(hwnd)
    width = rc.right or HELP_W
    height = rc.bottom or HELP_H
    fill(mem, (0, 0, width, height), BG)
    text(mem, "使用说明与原理", (32, 0, 600, TITLE_H), FONTS["title"],
         INK, DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    text(mem, "×", (width - 48, 0, width, TITLE_H), FONTS["mark"],
         INK if STATE["hover"] == "help_close" else MUTED,
         DT_CENTER | DT_VCENTER | DT_SINGLELINE)
    line(mem, (0, TITLE_H - 1, width, TITLE_H), LINE)
    caption = HELP_LAYOUT.get("caption") or (32, 672, 908, 700)
    footer = HELP_LAYOUT.get("footer") or (32, height - 44, 908, height - 18)
    text(mem, "转换方式（点击切换，主界面按钮保持不变）",
         caption, FONTS["small"], MUTED,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    text(mem, "设置仅本次运行有效，重启后恢复默认；勾选保留中间 PDF 后不会自动删除。",
         footer, FONTS["small"], MUTED,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    paint_end(hwnd, hdc, ps, mem, bmp, old, rc)


def draw_merge(hwnd):
    hdc, ps, mem, bmp, old, rc = paint_begin(hwnd)
    fill(mem, (0, 0, MERGE_W, MERGE_H), BG)
    text(mem, "PDF 合并 · 按列表顺序拼接", (30, 0, 700, TITLE_H),
         FONTS["title"], INK, DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    text(mem, "×", MERGE_CLOSE, FONTS["mark"],
         INK if STATE["hover"] == "merge_close" else MUTED,
         DT_CENTER | DT_VCENTER | DT_SINGLELINE)
    line(mem, (0, TITLE_H - 1, MERGE_W, TITLE_H), LINE)
    count = len(MERGE_FILES)
    pages = sum(p for _, p in MERGE_FILES)
    text(mem, "列表顺序即最终页序，可用上移 / 下移调整。当前 %d 个文件，合计 %d 页。"
         % (count, pages), (30, 70, 910, 104), FONTS["status"], INK,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    text(mem, "输出文件名", (30, 544, 190, 582), FONTS["status"], INK,
         DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    directory = (os.path.dirname(MERGE_FILES[0][0]) if MERGE_FILES
                 else "（添加文件后确定）")
    text(mem, "输出位置：%s" % directory, (30, 590, 910, 620),
         FONTS["small"], MUTED, DT_LEFT | DT_VCENTER | DT_SINGLELINE)
    labels = {"add": "添加 PDF", "up": "上移", "down": "下移",
              "remove": "移除", "clear": "清空"}
    for key, box in MERGE_BTNS.items():
        hovered = STATE["hover"] == key
        fill(mem, box, BTN_HOVER if hovered else MODE_OFF, 8)
        text(mem, labels[key], box, FONTS["mode"],
             WHITE if hovered else INK,
             DT_CENTER | DT_VCENTER | DT_SINGLELINE)
    fill(mem, MERGE_GO, BTN, 12)
    text(mem, "开始合并（按当前列表顺序）", MERGE_GO, FONTS["button"],
         WHITE, DT_CENTER | DT_VCENTER | DT_SINGLELINE)
    paint_end(hwnd, hdc, ps, mem, bmp, old, rc)


# ----------------------------------------------------------------- pipeline
def final_pdf_path(docx, tag=""):
    """A non-overwriting visible name for a PDF that must be kept."""
    directory = os.path.dirname(os.path.abspath(docx))
    stem = os.path.splitext(os.path.basename(docx))[0]
    index = 1
    while True:
        suffix = "" if index == 1 else "(%d)" % index
        path = os.path.join(directory, stem + suffix + tag + ".pdf")
        if not os.path.exists(path):
            return path
        index += 1




def move_file(src, dst):
    """Move a file, falling back to copy+replace across drives.

    The interim PDF lives in %TEMP%, which is usually on C:, while the user's
    documents are often on D:. `os.replace` raises
    "[WinError 17] 系统无法将文件移到不同的磁盘驱动器" in that case, so a
    cross-volume move has to copy (into a scratch name next to the target, so
    the final step is still an atomic same-directory rename).
    """
    try:
        os.replace(src, dst)
        return
    except OSError as exc:
        if getattr(exc, "winerror", None) not in (17, 18) and exc.errno != 18:
            raise
    partial = dst + ".moving"
    try:
        shutil.copyfile(src, partial)
        os.replace(partial, dst)
    finally:
        if os.path.exists(partial):
            try:
                os.remove(partial)
            except OSError:
                pass
        if os.path.exists(src):
            try:
                os.remove(src)
            except OSError:
                pass


def sweep_scratch(max_age_hours=6):
    """Remove interim PDFs abandoned by an earlier interrupted run.

    A crash or a kill leaves `<name>.__wps_*.tmp.pdf` in the scratch folder.
    Nothing else ever cleans those up, so each job start sweeps the stale ones.
    """
    folder = scratch_dir()
    now = time.time()
    removed = 0
    freed = 0
    try:
        entries = os.listdir(folder)
    except OSError:
        return 0, 0
    for name in entries:
        if not name.lower().endswith(".pdf"):
            continue
        path = os.path.join(folder, name)
        try:
            if now - os.path.getmtime(path) < max_age_hours * 3600:
                continue
            size = os.path.getsize(path)
            os.remove(path)
            removed += 1
            freed += size
        except OSError:
            continue
    return removed, freed


def scratch_dir():
    """Per-user scratch directory for interim PDFs (created once)."""
    global _SCRATCH
    if _SCRATCH is None:
        base = os.path.join(tempfile.gettempdir(), "WordToPdfTool")
        try:
            os.makedirs(base, exist_ok=True)
            probe = os.path.join(base, ".w")
            with open(probe, "w"):
                pass
            os.remove(probe)
        except OSError:
            base = tempfile.gettempdir()
        _SCRATCH = base
    return _SCRATCH


def reserve_path(directory, name):
    """Reserve a non-existing output path atomically and return it.

    `unique_output`/`unique_paths` only *looked* for a free name, so two runs
    (or two windows) could pick the same one and the second silently clobbered
    the first. Creating the placeholder with O_EXCL makes the check atomic.
    """
    stem, ext = os.path.splitext(name)
    index = 1
    while True:
        suffix = "" if index == 1 else "(%d)" % index
        path = os.path.join(directory, stem + suffix + ext)
        try:
            handle = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            index += 1
            continue
        except OSError:
            # Read-only or exotic filesystem: fall back to a plain probe but
            # keep the "never overwrite" promise for the common case.
            if os.path.exists(path):
                index += 1
                continue
            return path
        os.close(handle)
        return path


def unique_output(directory, filename):
    """A reserved, non-overwriting output path inside directory."""
    stem, ext = os.path.splitext(filename.strip() or "合并结果.pdf")
    if ext.lower() != ".pdf":
        stem, ext = stem + ext, ".pdf"
    path = reserve_path(directory, stem + ext)
    return path


def unique_paths(docx):
    """Return (interim temp PDF, reserved optimised PDF, paired plain PDF).

    The interim PDF is created under the per-user scratch directory: a
    multi-hundred-MB scratch file has no business landing next to the user's
    documents, where it also gets synchronised by OneDrive/网盘 and survives a
    interrupted run as debris.
    """
    directory = os.path.dirname(os.path.abspath(docx))
    stem = os.path.splitext(os.path.basename(docx))[0]
    index = 1
    while True:
        suffix = "" if index == 1 else "(%d)" % index
        out = os.path.join(directory, stem + suffix + "_优化保真版.pdf")
        plain = os.path.join(directory, stem + suffix + ".pdf")
        if not os.path.exists(out) and not os.path.exists(plain):
            break
        index += 1
    tmp = os.path.join(scratch_dir(), "%s.__wps_%s__.tmp.pdf"
                       % (stem, uuid.uuid4().hex[:10]))
    return tmp, out, plain


def _stored_is_word_jpeg(archive, candidates, stored):
    """True when `stored` is byte-identical to one of the Word originals."""
    for candidate in candidates:
        try:
            if archive.read(candidate["name"]) == stored:
                return True
        except Exception:
            continue
    return False


def _same_pixels(source, buf, width, height):

    """True when the Word image decodes to exactly the pixels in the PDF."""
    from PIL import Image
    with Image.open(io.BytesIO(source)) as original:
        return original.convert("RGB").tobytes() == buf


def quality_check(out_pdf, docx, sample_pages=20, index=None):

    """Fidelity check against Word originals.

    A JPEG passthrough object stores the *exact bytes* of the Word image, so
    comparing raw bytes is the strictest possible proof of zero loss. Decoded
    pixels are compared only for non-DCT objects: PIL and MuPDF use slightly
    different IDCT/chroma-upsampling rounding, so decoded-pixel comparison on
    a byte-identical JPEG produces a few percent of +/-1 channel differences
    that are pure decoder noise, not quality loss. The old all-pixel rule
    misreported those objects and kept the interim PDF indefinitely.

    `index` lets the caller pass the Word index the optimiser already built;
    rebuilding it re-decoded every image inside the .docx for no benefit.

    Returns (checked, identical, pixel_diff, no_candidate, byte_identical).
    """
    import pymupdf
    from PIL import Image
    own_index = index is None
    if own_index:
        archive, grouped, _ = opt.build_word_index(docx)
    else:
        archive, grouped, _ = index
    doc = pymupdf.open(out_pdf)
    identical = pixel_diff = no_candidate = byte_identical = 0
    unmatched = 0
    failures = 0
    retained = 0
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
                    width, height, buf = opt._image_samples(doc, xref)
                    candidates = grouped.get((width, height), [])
                    if not candidates:
                        no_candidate += 1
                        continue
                    filt = doc.xref_get_key(xref, "Filter")[1]
                    stored = doc.xref_stream_raw(xref) or b""
                    if "DCTDecode" in filt:
                        raw = doc.extract_image(xref)["image"]
                        matched = False
                        for candidate in candidates:
                            try:
                                if archive.read(candidate["name"]) == raw:
                                    matched = True
                                    break
                            except Exception:
                                pass
                        if matched:
                            byte_identical += 1
                            identical += 1
                            continue
                    # Everything else is judged by replaying the optimiser's own
                    # decision for this object: re-encode the matched Word image
                    # and compare the bytes that should have been written. That
                    # is stricter than comparing decoded pixels, and it stays
                    # correct no matter which encoding branch the policy picks
                    # (indexed, true-colour Flate, or keeping Word's own JPEG
                    # because it is smaller). Only a mismatch here means the
                    # delivered image really differs from the Word original.
                    fp = opt.fingerprint(
                        Image.frombytes("RGB", (width, height), buf))
                    best, _ = opt.best_match(fp, candidates)
                    if best is None:
                        # No candidate passed DIST_THRESHOLD: the optimiser left
                        # this object exactly as Word exported it, which is not a
                        # quality regression to report.
                        unmatched += 1
                        continue
                    source = archive.read(best["name"])
                    smask = doc.xref_get_key(xref, "SMask")[1]
                    keep_alpha = smask not in ("null", "False", None)
                    expected = opt.encode_stream(source, width, height,
                                                 not keep_alpha, False,
                                                 alpha_kept=keep_alpha)
                    if expected is None:
                        unmatched += 1
                        continue
                    if expected["data"] == stored:
                        byte_identical += 1
                        identical += 1
                        continue
                    # The object still holds whatever Word put there. Under the
                    # volume-first policy that is a deliberate outcome, not a
                    # quality regression: the lossless form was larger, so the
                    # original bytes were kept to save space. Report it as
                    # "retained" so the log stays honest about what happened,
                    # but do not block disposal of the interim PDF for it.
                    if expected["tag"] == "jpeg_passthrough" and \
                            _stored_is_word_jpeg(archive, candidates, stored):
                        byte_identical += 1
                        identical += 1
                        continue
                    if len(expected["data"]) >= len(stored):
                        retained += 1
                        identical += 1
                        continue
                    pixel_diff += 1
                except Exception:
                    # A decode/machine error is NOT evidence of quality loss;
                    # counting it as a pixel difference is what kept interim
                    # PDFs around in V1.1. Report it as its own category.
                    failures += 1
    finally:
        doc.close()
        if own_index:
            archive.close()
    return (identical + pixel_diff + no_candidate + unmatched, identical,
            pixel_diff, no_candidate, byte_identical, failures, unmatched,
            retained)


def run_pipeline(docx, engine, sink, keep_pdf=False, progress=None, app=None,
                 cancel=None):
    """Export + optimise + verify one document, then dispose of the interim PDF.

    `app` lets a batch reuse one hidden WPS instance instead of starting and
    quitting automation once per document. `cancel` is a threading.Event; it is
    polled between stages so a long batch can actually be stopped.
    """
    if progress is None:
        progress = lambda frac, label: None  # noqa: E731
    if cancel is None:
        cancel = threading.Event()

    def stopped():
        return cancel.is_set()

    docx = os.path.abspath(docx)
    if not os.path.isfile(docx):
        raise RuntimeError("找不到文件：%s" % docx)
    if not docx.lower().endswith(".docx"):
        raise RuntimeError("仅支持 .docx 格式（优化阶段需读取 Word 包内原始图片）")
    size = os.path.getsize(docx)
    if size > MAX_SIZE:
        raise RuntimeError("文件超过 %s 上限：%.2f GB"
                           % (MAX_SIZE_TEXT, size / 1024 ** 3))

    tmp, out, plain = unique_paths(docx)
    report = {"tmp": tmp, "out": out, "plain": plain, "engine": engine,
              "pdf_removed": False, "pdf_final": None}
    progress(0.01, "准备文档")

    if engine == "cli" and size > CLI_MAX_SIZE:
        sink("文件超过 WPS CLI 官方 200 MB 限制，已自动改用 WPS COM 通道。")
        engine = "wps"
        report["engine"] = engine

    try:
        sink("第一步：调用 %s 把 Word 另存为 PDF…" % ENGINE_NAMES[engine])
        if stopped():
            raise RuntimeError("已取消")
        if engine == "cli":
            cli = w2p.find_cli()
            args = SimpleNamespace(image_quality="medium", timeout=1800,
                                   overwrite=True)
            pdf_size, pages, _ = w2p.export_via_cli(cli, docx, tmp, args)
        else:
            owned = None
            if app is None:
                owned = app = w2p.start_app(engine, False)
            try:
                pdf_size, pages, _, _ = w2p.export_doc(
                    app, docx, tmp, 1800, False, True)
            except Exception as exc:
                # Tagged so the batch knows whether the automation instance
                # itself needs rebuilding (COM fault) or not (bad document).
                exc.stage = "export"
                raise
            finally:
                if owned is not None:
                    w2p.quit_app(owned)
                    app = None
        report["pdf_size"] = pdf_size
        report["pages"] = pages
        progress(0.20, "WPS 导出完成")
        sink("导出完成：%.2f MB / %d 页。第二步：回填 Word 原图…"
             % (pdf_size / 1048576, pages))

        if stopped():
            raise RuntimeError("已取消")

        # Publish the optimised PDF atomically: writing straight to the final
        # name leaves a truncated-but-complete-looking deliverable behind if
        # the process dies mid-save.
        part = out + ".part"
        index = None
        try:
            try:
                stats, index = opt.optimize(tmp, docx, part, use_g4=False,
                                            return_index=True)
            except RuntimeError as exc:
                if "no usable images" not in str(exc):
                    raise
                # No optimisation possible: this exported PDF *is* the
                # deliverable, not a leftover intermediate file.
                move_file(tmp, plain)
                sink("该文档不含可用插图，已跳过优化。导出的 PDF 即为最终交付物，"
                     "不存在需要删除的中间文件。")
                report.update({"skipped": True, "out": plain,
                               "out_size": pdf_size, "pdf_final": plain})
                progress(1.0, "完成（无插图文档）")
                return report
            os.replace(part, out)
        finally:
            if os.path.exists(part):
                try:
                    os.remove(part)
                except OSError:
                    pass

        report["stats"] = stats
        report["out_size"] = os.path.getsize(out)
        progress(0.82, "优化保真版已生成")
        sink("优化保真版已生成：%.2f MB。正在做完整性校验与画质复核…"
             % (report["out_size"] / 1048576))

        verify_ok = True
        integrity_ok = True
        pixel_diff = 0
        no_candidate = 0
        if stopped():
            raise RuntimeError("已取消")
        try:
            integrity = opt.verify(tmp, out, 20, False, 12)
            report["integrity"] = integrity
            integrity_ok = (integrity.get("pages_src") == integrity.get("pages_out")
                            and integrity.get("chars_src") == integrity.get("chars_out")
                            and integrity.get("placements_src") ==
                            integrity.get("placements_out"))
            if not integrity_ok:
                sink("完整性校验三项未全部一致，已记录详情。")
        except Exception as exc:
            verify_ok = False
            sink("完整性校验未完成：%s" % exc)
        try:
            (checked, same, pixel_diff, no_candidate,
             byte_same, failures, unmatched, retained) = quality_check(
                out, docx, 20, index=index)
            report.update({"checked": checked, "same": same,
                           "diff": pixel_diff, "no_candidate": no_candidate,
                           "byte_same": byte_same, "check_error": failures,
                           "unmatched": unmatched, "retained": retained})
            sink("画质复核：抽检 %d 个图像对象，与 Word 原图一致 %d"
                 "（其中 JPEG 原始字节直通 %d），像素不一致 %d，"
                 "Word 中无同尺寸原图无法比对 %d，"
                 "匹配超阈值未替换 %d，为省体积保留原 JPEG %d。"
                 % (checked, same, byte_same, pixel_diff, no_candidate,
                    unmatched, retained))
            if failures:
                sink("有 %d 个对象在复核时读取失败，单独计入「复核异常」，"
                     "不计作画质差异。" % failures)
            if byte_same:
                sink("说明：JPEG 直通对象按原始字节比对，确认 PDF 内即为 "
                     "Word 原 JPEG 的逐字节副本，不存在二次有损压缩。")
            if no_candidate:
                sink("说明：无法比对的对象多因 WPS 导出时缩放或裁剪图像，"
                     "属比对基准限制，非本次优化引入的差异，不阻断中间 PDF 删除。")
        except Exception as exc:
            verify_ok = False
            sink("画质复核未完成：%s" % exc)
        finally:
            if index is not None:
                try:
                    index[0].close()
                except Exception:
                    pass
        progress(0.96, "校验完成")

        # Dispose of the interim PDF. A real pixel difference or an aborted
        # check is the only quality reason to keep it; baseline gaps are not.
        reasons = []
        if keep_pdf:
            reasons.append("已勾选保留中间 PDF")
        if not os.path.isfile(out):
            reasons.append("最终版未生成")
        if not verify_ok:
            reasons.append("校验过程异常，未能完成")
        if not integrity_ok:
            reasons.append("完整性校验三项不一致")
        if pixel_diff > 0:
            reasons.append("画质复核存在 %d 张真实像素差异" % pixel_diff)
        report["keep_reason"] = "；".join(reasons)

        if not reasons:
            try:
                os.remove(tmp)
                report["pdf_removed"] = True
                sink("中间 PDF 仅用于校验，校验通过后已删除。")
            except OSError as exc:
                report["delete_error"] = str(exc)
        if not report["pdf_removed"]:
            move_file(tmp, plain)
            report["pdf_final"] = plain
            sink("已保留中间 PDF（原因：%s）：%s"
                 % (report.get("keep_reason") or
                    report.get("delete_error", "删除失败"),
                    os.path.basename(plain)))
        progress(1.0, "完成")
        return report
    except Exception:
        # On failure keep the interim PDF under a visible name for diagnosis.
        if os.path.exists(tmp):
            try:
                final = final_pdf_path(docx, "_导出稿")
                move_file(tmp, final)
                report["pdf_final"] = final
            except OSError:
                pass
        raise


def collect_docx(folder):
    """Recursively collect .docx files, skipping Word lock files."""
    result = []
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        for name in files:
            if name.lower().endswith(".docx") and not name.startswith("~$"):
                result.append(os.path.join(root, name))
    return sorted(result, key=lambda p: (os.path.basename(p).lower(), p.lower()))


def write_batch_csv(reports, errors, folder):
    """Write a per-document ledger next to the batch results."""
    path = os.path.join(folder, "转换记录.csv")
    index = 1
    while os.path.exists(path):
        index += 1
        path = os.path.join(folder, "转换记录(%d).csv" % index)
    header = ["文件", "页数", "导出体积MB", "成品体积MB", "抽检对象",
              "与原图一致", "JPEG直通", "像素不一致", "无法比对",
              "复核异常", "中间PDF已删除", "结果"]
    failed = {name for name, _ in errors}
    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(header)
            for report in reports:
                name = os.path.basename(report["out"])
                writer.writerow([
                    name,
                    report.get("pages", ""),
                    "%.2f" % (report.get("pdf_size", 0) / 1048576),
                    "%.2f" % (report.get("out_size", 0) / 1048576),
                    report.get("checked", ""),
                    report.get("same", ""),
                    report.get("byte_same", ""),
                    report.get("diff", ""),
                    report.get("no_candidate", ""),
                    report.get("check_error", ""),
                    "是" if report.get("pdf_removed") else "否",
                    "无插图（导出稿即成品）" if report.get("skipped") else "成功",
                ])
            for name, message in errors:
                writer.writerow([name] + [""] * 10 + ["失败：" + message])
    except OSError as exc:
        return None, str(exc)
    return path, None


def run_batch(targets, engine, keep_pdf, sink, progress=None, cancel=None):
    """Process every document; one failure must not stop the remaining ones."""
    if progress is None:
        progress = lambda frac, label: None  # noqa: E731
    if cancel is None:
        cancel = threading.Event()
    reports, errors = [], []
    total = len(targets)
    state = {"tag": "", "docprog": lambda sub, label: None}

    def emit(zh, en=None):
        message = state["tag"] + str(zh)
        match = re.search(r"进度\s+(\d+)\s*/\s*(\d+)", str(zh))
        if match:
            done, whole = int(match.group(1)), max(1, int(match.group(2)))
            state["docprog"](0.20 + 0.60 * done / whole,
                             "回填原图 %d/%d" % (done, whole))
        sink(message)

    opt.emit = emit

    # One hidden WPS instance for the whole batch: starting and quitting COM
    # automation per document cost 1-3 s each and widened the failure surface.
    shared_app = None
    if engine != "cli" and total > 1:
        try:
            sink("正在启动共用的 WPS 自动化实例…")
            shared_app = w2p.start_app(engine, False)
        except Exception as exc:
            sink("共用实例启动失败（%s），改为逐文档启动。" % exc)
            shared_app = None

    try:
        for i, docx in enumerate(targets, 1):
            if cancel.is_set():
                sink("已收到取消请求，停止处理剩余文档。")
                break
            tag = "[%d/%d] " % (i, total) if total > 1 else ""
            state["tag"] = tag
            base, span = (i - 1) / total, 1.0 / total

            def docprog(sub, label, _b=base, _s=span, _t=tag):
                progress(_b + _s * max(0.0, min(1.0, sub)), _t + label)

            state["docprog"] = docprog

            def sub(message, _tag=tag):
                sink(_tag + str(message))

            sub("开始处理：%s" % os.path.basename(docx))
            attempts = 2 if shared_app is not None else 1
            for attempt in range(1, attempts + 1):
                try:
                    reports.append(run_pipeline(docx, engine, sub, keep_pdf,
                                                docprog, app=shared_app,
                                                cancel=cancel))
                    sub("该文档处理完成。")
                    break
                except Exception as exc:
                    dead = (w2p.is_instance_dead(exc)
                            if getattr(exc, "stage", "") == "export" else False)
                    if dead and shared_app is not None and attempt < attempts:
                        # The hidden Word/WPS process died (typically after a
                        # very large document). Report the document as failed
                        # only after giving it a fresh instance: 12 of 25 real
                        # documents were lost this way before the retry existed.
                        sub("自动化实例已退出（%s），重建后重试该文档。" % exc)
                        w2p.quit_app(shared_app)
                        shared_app = None
                        try:
                            shared_app = w2p.start_app(engine, False)
                        except Exception as rebuild:
                            sink("实例重建失败：%s" % rebuild)
                            errors.append((os.path.basename(docx),
                                           "实例重建失败：%s" % rebuild))
                            sub("失败：无法重建自动化实例")
                            break
                        continue
                    errors.append((os.path.basename(docx),
                                   "%s：%s" % (type(exc).__name__, exc)))
                    sub("失败：%s" % exc)
                    if (shared_app is not None and dead):
                        # Dead instance with no retry left: leave a healthy one
                        # for the following documents.
                        w2p.quit_app(shared_app)
                        shared_app = None
                        try:
                            shared_app = w2p.start_app(engine, False)
                        except Exception:
                            shared_app = None
                    break
            progress((i) / total, tag + "该文档结束")
    finally:
        if shared_app is not None:
            w2p.quit_app(shared_app)
    return reports, errors


# ------------------------------------------------------------------- merging
def pdf_pages(path):
    import pymupdf
    doc = pymupdf.open(path)
    try:
        return doc.page_count
    finally:
        doc.close()


def merge_pdfs(paths, out_path, sink, progress=None, cancel=None):
    """Concatenate PDFs strictly in the given order, keeping image streams."""
    import pymupdf
    if progress is None:
        progress = lambda frac, label: None  # noqa: E731
    if len(paths) < 2:
        raise RuntimeError("合并至少需要 2 个 PDF 文件。")
    sink("开始合并：共 %d 个 PDF，严格按列表顺序拼接。" % len(paths))
    result = pymupdf.open()
    total_pages = 0
    expected = []
    try:
        for i, path in enumerate(paths, 1):
            if cancel is not None and cancel.is_set():
                raise RuntimeError("已取消")
            if not os.path.isfile(path):
                raise RuntimeError("找不到文件：%s" % path)
            src = pymupdf.open(path)
            try:
                if src.needs_pass:
                    raise RuntimeError("PDF 已加密，无法合并：%s"
                                       % os.path.basename(path))
                count = src.page_count
                if count <= 0:
                    raise RuntimeError("PDF 页数为 0：%s"
                                       % os.path.basename(path))
                result.insert_pdf(src)
                total_pages += count
                expected.append(count)
            finally:
                src.close()
            sink("[%d/%d] %s（%d 页），累计 %d 页。"
                 % (i, len(paths), os.path.basename(path), count,
                    total_pages))
            progress(i / len(paths) * 0.9, "已拼接 %d/%d 个文件"
                     % (i, len(paths)))
        sink("正在写出合并文件，图像流保持原有编码不做重压缩…")
        # Atomic publish: a crash mid-save must not leave a truncated PDF that
        # already carries the final name. The .part file is only discarded when
        # the save itself failed - removing it unconditionally here deleted the
        # finished file before it could be renamed into place.
        part = out_path + ".part"
        try:
            result.save(part, garbage=4, clean=False, deflate=False)
        except Exception:
            if os.path.exists(part):
                try:
                    os.remove(part)
                except OSError:
                    pass
            raise
        result.close()
        result = None
        os.replace(part, out_path)
    finally:
        if result is not None:
            result.close()
    progress(0.96, "正在核验合并结果")
    check = pymupdf.open(out_path)
    try:
        actual = check.page_count
    finally:
        check.close()
    if actual != total_pages:
        raise RuntimeError("合并结果页数不符：期望 %d，实际 %d"
                           % (total_pages, actual))
    sink("核验通过：合并结果 %d 页，与 %d 个源文件页数之和一致。"
         % (actual, len(paths)))
    progress(1.0, "合并完成")
    return actual


def format_single(report, minutes):
    if report.get("skipped"):
        return ("完成｜文档无插图，导出 PDF 即为最终结果\n%.2f MB / %d 页\n"
                "输出：%s\n中间 PDF：无需删除（无优化步骤）\n耗时 %.1f 分钟"
                % (report["pdf_size"] / 1048576, report["pages"],
                   os.path.basename(report["out"]), minutes))
    extra = ""
    if report.get("checked"):
        extra = "画质复核 %d/%d 与 Word 原图一致" % (
            report["same"], report["checked"])
        if report.get("byte_same"):
            extra += "（含 JPEG 字节直通 %d）" % report["byte_same"]
        if report.get("no_candidate"):
            extra += "，%d 张无法比对" % report["no_candidate"]
        extra += "\n"
    middle = ("校验通过，已删除" if report["pdf_removed"]
              else "已保留：%s（%s）" % (
                  os.path.basename(report["pdf_final"] or ""),
                  report.get("keep_reason") or "删除失败"))
    return ("完成｜导出 %.2f MB → 优化 %.2f MB（%d 页）\n%s"
            "输出：%s\n中间 PDF：%s\n耗时 %.1f 分钟"
            % (report["pdf_size"] / 1048576,
               report["out_size"] / 1048576, report["pages"], extra,
               os.path.basename(report["out"]), middle, minutes))


def format_batch(reports, errors, minutes, total, merged=None):
    checked = sum(r.get("checked") or 0 for r in reports)
    same = sum(r.get("same") or 0 for r in reports)
    removed = sum(1 for r in reports if r["pdf_removed"])
    skipped = sum(1 for r in reports if r.get("skipped"))
    kept = len(reports) - removed - skipped
    body = ("完成｜共 %d 个文档，成功 %d 个，失败 %d 个\n"
            "最终 PDF：各 Word 所在目录，共 %d 份\n"
            "中间 PDF：校验通过后删除 %d 份"
            % (total, len(reports), len(errors), len(reports), removed))
    if kept:
        body += "，另有 %d 份按要求保留" % kept
    body += "\n"
    if skipped:
        body += "无插图文档：%d 份，导出 PDF 即为最终交付物\n" % skipped
    if checked:
        body += "画质复核：%d/%d 与 Word 原图一致\n" % (same, checked)
    if merged:
        body += "批量结果已合并：%s（%d 页）\n" % (
            os.path.basename(merged[0]), merged[1])
    body += "耗时 %.1f 分钟" % minutes
    if errors:
        names = "、".join(n for n, _ in errors[:3])
        if len(errors) > 3:
            names += " 等"
        body += "\n失败：%s" % names
    return body


# -------------------------------------------------------------------- workers
def q_progress(frac, label):
    Q.put(("progress", (frac, label)))


def open_log_file():
    """Start a per-run log file; every sink line is mirrored into it."""
    global LOG_HANDLE, LOG_PATH
    close_log_file()
    folder = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"),
                          "WordToPdfTool", "logs")
    try:
        os.makedirs(folder, exist_ok=True)
        LOG_PATH = os.path.join(folder, time.strftime("%Y%m%d-%H%M%S.log"))
        LOG_HANDLE = open(LOG_PATH, "w", encoding="utf-8")
    except OSError:
        LOG_HANDLE = None
        LOG_PATH = None
    return LOG_PATH


def close_log_file():
    global LOG_HANDLE
    if LOG_HANDLE is not None:
        try:
            LOG_HANDLE.close()
        except Exception:
            pass
        LOG_HANDLE = None


def log_line(message):
    if LOG_HANDLE is not None:
        try:
            LOG_HANDLE.write(message + "\n")
            LOG_HANDLE.flush()
        except Exception:
            pass


def conversion_worker(targets, engine, keep_pdf, auto_merge, merge_dir):
    def sink(message):
        Q.put(("log", str(message)))
        Q.put(("status", str(message)))
        log_line(str(message))

    started = time.time()
    try:
        sink("任务开始：%s，共 %d 个 Word 文档。"
             % ("单个文件" if len(targets) == 1 else "文件夹批量",
                len(targets)))
        reports, errors = run_batch(targets, engine, keep_pdf, sink, q_progress,
                                    cancel=CANCEL)
        minutes = (time.time() - started) / 60
        merged = None
        if auto_merge and len(targets) > 1 and len(reports) >= 2:
            out = unique_output(merge_dir or os.path.dirname(
                os.path.abspath(targets[0])), "批量合并结果.pdf")
            sink("批量处理完成，按处理顺序合并本次生成的最终 PDF…")

            def mprogress(frac, label):
                Q.put(("progress", (frac, "自动合并 · " + label)))

            pages = merge_pdfs([r["out"] for r in reports], out, sink,
                               mprogress, cancel=CANCEL)
            merged = (out, pages)
            sink("自动合并完成：%s（%d 页）。" % (os.path.basename(out), pages))
        if reports and merge_dir:
            ledger, problem = write_batch_csv(reports, errors, merge_dir)
            if ledger:
                sink("转换记录已写入：%s" % ledger)
            elif problem:
                sink("转换记录写入失败：%s" % problem)
        if len(targets) == 1 and not errors and reports:
            body = format_single(reports[0], minutes)
        else:
            body = format_batch(reports, errors, minutes, len(targets), merged)
        Q.put(("done", body))
    except Exception as exc:
        Q.put(("log", "任务异常：%s\n%s" % (exc, traceback.format_exc())))
        Q.put(("error", str(exc)))
    finally:
        close_log_file()


def merge_worker(paths, out_path):
    def sink(message):
        Q.put(("log", str(message)))
        Q.put(("status", str(message)))
        log_line(str(message))

    started = time.time()
    try:
        sink("任务开始：合并 %d 个 PDF。" % len(paths))
        pages = merge_pdfs(paths, out_path, sink, q_progress, cancel=CANCEL)
        minutes = (time.time() - started) / 60
        Q.put(("done", "完成｜按选择顺序合并 %d 个 PDF\n输出：%s\n"
               "合计 %d 页 / %.2f MB\n源 PDF 全部保留，未删除\n耗时 %.1f 分钟"
               % (len(paths), os.path.basename(out_path), pages,
                  os.path.getsize(out_path) / 1048576, minutes)))
    except Exception as exc:
        Q.put(("log", "合并失败：%s\n%s" % (exc, traceback.format_exc())))
        Q.put(("error", str(exc)))
        # The output name was reserved atomically; a failed merge would
        # otherwise leave an empty file with a perfectly good name.
        try:
            if os.path.exists(out_path) and os.path.getsize(out_path) == 0:
                os.remove(out_path)
        except OSError:
            pass
    finally:
        close_log_file()


# -------------------------------------------------------------------- dialog
def choose_docx(owner):
    probe = os.environ.get("PDFTOOL_TEST_DOCX")
    if probe:
        return probe
    buf = ctypes.create_unicode_buffer(2048)
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.hwndOwner = owner
    ofn.lpstrFilter = "Word 文档 (*.docx)\0*.docx\0所有文件 (*.*)\0*.*\0\0"
    ofn.nFilterIndex = 1
    ofn.lpstrFile = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile = 2048
    ofn.lpstrTitle = "选择需要转换为 PDF 的 Word 文档"
    ofn.lpstrDefExt = "docx"
    ofn.Flags = OFN
    if not comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return None
    return buf.value


def choose_pdfs(owner):
    """Multi-select PDF picker; return paths in the dialog's selection order."""
    probe = os.environ.get("PDFTOOL_TEST_FILES")
    if probe:
        return [p for p in probe.split("|") if p]
    buf = ctypes.create_unicode_buffer(65536)
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.hwndOwner = owner
    ofn.lpstrFilter = "PDF 文件 (*.pdf)\0*.pdf\0所有文件 (*.*)\0*.*\0\0"
    ofn.nFilterIndex = 1
    ofn.lpstrFile = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile = 65536
    ofn.lpstrTitle = "选择需要合并的 PDF（可多选，顺序可在下一步调整）"
    ofn.lpstrDefExt = "pdf"
    ofn.Flags = OFN | OFN_ALLOWMULTISELECT
    if not comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return []
    parts = []
    for piece in buf[:].split("\x00"):
        if not piece:
            break
        parts.append(piece)
    if not parts:
        return []
    if len(parts) == 1:
        return [parts[0]]
    root = parts[0]
    return [os.path.join(root, name) for name in parts[1:]]


def choose_folder(owner):
    probe = os.environ.get("PDFTOOL_TEST_FOLDER")
    if probe:
        return probe
    try:
        ole32.CoInitializeEx(None, 0x0002)
    except Exception:
        pass
    display = ctypes.create_unicode_buffer(260)
    bi = BROWSEINFOW()
    bi.hwndOwner = owner
    bi.pszDisplayName = ctypes.cast(display, ctypes.c_wchar_p)
    bi.lpszTitle = "选择包含 Word 文档的文件夹"
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_EDITBOX | BIF_NEWDIALOGSTYLE
    pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
    if not pidl:
        return None
    try:
        path = ctypes.create_unicode_buffer(1024)
        if shell32.SHGetPathFromIDListW(
                pidl, ctypes.cast(path, wintypes.LPWSTR)):
            return path.value
    finally:
        ole32.CoTaskMemFree(pidl)
    return None


# ------------------------------------------------------------------- windows
def create_class(name, proc):
    wc = WNDCLASSW()
    wc.style = CS_DROPSHADOW | CS_HREDRAW | CS_VREDRAW
    wc.lpfnWndProc = proc
    wc.hInstance = HINST
    wc.hCursor = user32.LoadCursorW(None, ctypes.c_void_p(32512))
    wc.hbrBackground = gdi32.CreateSolidBrush(BG)
    wc.lpszClassName = name
    if not user32.RegisterClassW(ctypes.byref(wc)):
        raise ctypes.WinError(ctypes.get_last_error())


def work_area():
    """Usable screen rectangle (excludes the taskbar) in physical pixels."""
    rc = wintypes.RECT()
    try:
        if user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rc), 0):
            return rc.left, rc.top, rc.right, rc.bottom
    except Exception:
        pass
    return 0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


def fit_to_screen(w, h):
    """Shrink a window to the work area so no control ends up off-screen.

    The help window is 940x980: on a 1366x768 laptop its bottom 212 px (the
    engine radio buttons and both checkboxes) used to be unreachable, because
    the window is not resizable and does not scroll as a whole.
    """
    left, top, right, bottom = work_area()
    return min(w, right - left - 16), min(h, bottom - top - 16)


def center_window(hwnd, w, h):
    left, top, right, bottom = work_area()
    sw, sh = right - left, bottom - top
    user32.SetWindowPos(hwnd, None, left + max(0, (sw - w) // 2),
                        top + max(0, (sh - h) // 2), w, h, 0x0040)


def popup(hwnd, w, h, radius=16):
    w, h = fit_to_screen(w, h)
    center_window(hwnd, w, h)
    region = gdi32.CreateRoundRectRgn(0, 0, w, h, radius, radius)
    user32.SetWindowRgn(hwnd, region, True)
    user32.SetForegroundWindow(hwnd)
    user32.InvalidateRect(hwnd, None, True)


def open_help():
    global HWND_HELP
    if HWND_HELP and user32.IsWindow(HWND_HELP):
        user32.SetForegroundWindow(HWND_HELP)
        return
    w, h = fit_to_screen(HELP_W, HELP_H)
    hwnd = user32.CreateWindowExW(
        0, "PdfToolHelp", "使用说明", WS_POPUP | WS_VISIBLE | WS_CLIPCHILDREN,
        100, 100, w, h, None, None, HINST, None)
    popup(hwnd, HELP_W, HELP_H)
    HWND_HELP = hwnd


def create_log_control(hwnd):
    global HWND_LOG
    edit = user32.CreateWindowExW(
        0, "EDIT", "",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | ES_MULTILINE |
        ES_READONLY | ES_AUTOVSCROLL | ES_NOHIDESEL,
        LOG_BOX[0], LOG_BOX[1], LOG_BOX[2] - LOG_BOX[0],
        LOG_BOX[3] - LOG_BOX[1], hwnd, ctypes.c_void_p(4001), HINST, None)
    user32.SendMessageW(edit, WM_SETFONT, FONTS["log"], 1)
    user32.SendMessageW(edit, EM_SETLIMITTEXT, 0x100000, 0)
    HWND_LOG = edit


def main_help_box():
    rc = wintypes.RECT()
    user32.GetClientRect(HWND_MAIN, ctypes.byref(rc)) if HWND_MAIN else None
    width = rc.right or MAIN_W
    return (width - 94, 0, width - 48, TITLE_H)


def main_close_box():
    rc = wintypes.RECT()
    user32.GetClientRect(HWND_MAIN, ctypes.byref(rc)) if HWND_MAIN else None
    width = rc.right or MAIN_W
    return (width - 48, 0, width, TITLE_H)


def help_close_box():
    rc = wintypes.RECT()
    if HWND_HELP and user32.IsWindow(HWND_HELP):
        user32.GetClientRect(HWND_HELP, ctypes.byref(rc))
    width = rc.right or HELP_W
    return (width - 48, 0, width, TITLE_H)


def confirm_abort(owner):
    """Ask before throwing away a running job."""
    MB_YESNO = 0x00000004
    MB_ICONWARNING = 0x00000030
    IDYES = 6
    return user32.MessageBoxW(
        owner,
        "正在处理，确定要中止吗？\n\n"
        "已完成的结果会保留；当前文档的中间文件会被清理。",
        "中止任务", MB_YESNO | MB_ICONWARNING) == IDYES


def request_close(hwnd):
    """Close the main window, asking first if a job is still running."""
    if STATE["busy"]:
        if not confirm_abort(hwnd):
            return
        CANCEL.set()
    user32.DestroyWindow(hwnd)


def append_log_lines(lines):
    if not HWND_LOG or not user32.IsWindow(HWND_LOG):
        return
    # Win32 EDIT needs CRLF; a bare \n in a multi-line summary would be
    # swallowed and the whole summary would collapse into one long line.
    normalized = [str(v).replace("\r\n", "\n").replace("\n", "\r\n")
                  for v in lines]
    for value in normalized:
        STATE["logs"].append(value)
    if len(STATE["logs"]) > 1500:
        # Trim the head of the control in ONE operation instead of re-setting
        # the whole 100 KB text for every subsequent append (that was O(n^2)).
        overflow = len(STATE["logs"]) - 1100
        STATE["logs"] = STATE["logs"][overflow:]
        head = "\r\n".join(STATE["logs"][:overflow]) + "\r\n"
        user32.SendMessageW(HWND_LOG, EM_SETSEL, 0, len(head))
        send_text(HWND_LOG, EM_REPLACESEL, 0, "")
        user32.SendMessageW(HWND_LOG, EM_SETSEL, -1, -1)
    for value in normalized:
        n = user32.SendMessageW(HWND_LOG, WM_GETTEXTLENGTH, 0, 0)
        user32.SendMessageW(HWND_LOG, EM_SETSEL, n, n)
        send_text(HWND_LOG, EM_REPLACESEL, 0, value + "\r\n")
    user32.SendMessageW(HWND_LOG, EM_SCROLLCARET, 0, 0)


def create_help_controls(hwnd):
    global HWND_HELP_EDIT, CHECK_HWND, AUTO_MERGE_HWND
    rc = wintypes.RECT()
    user32.GetClientRect(hwnd, ctypes.byref(rc))
    width = rc.right or HELP_W
    height = rc.bottom or HELP_H
    # Bottom block (caption + 3 engine radios + 2 checkboxes + footer) keeps
    # its size; the help text area absorbs the difference, so the window still
    # works on a 768-pixel-tall screen where 980 px cannot fit. 308 px is the
    # measured height of that block (38 caption + 3*44 radios + 2*42 boxes +
    # 50 footer) plus slack.
    block = 308
    text_bottom = max(120, height - block)
    HELP_LAYOUT.update({"caption": (32, text_bottom + 4, width - 32,
                                    text_bottom + 32),
                        "footer": (32, height - 44, width - 32, height - 18),
                        "radio_top": text_bottom + 38,
                        "check_top": text_bottom + 170,
                        "width": width, "height": height})
    edit = user32.CreateWindowExW(
        0, "EDIT", _HELP_TEXT,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | ES_MULTILINE |
        ES_READONLY | ES_AUTOVSCROLL,
        32, 76, width - 64, max(80, text_bottom - 76), hwnd,
        ctypes.c_void_p(1001), HINST, None)
    user32.SendMessageW(edit, WM_SETFONT, FONTS["help"], 1)
    HWND_HELP_EDIT = edit
    for i, key in enumerate(("wps", "cli", "word")):
        label = {
            "wps": "WPS COM（默认，推荐）：免费、无会员要求、上限 2 GB",
            "cli": "WPS CLI：官方命令行，需 VIP 账号，超 200 MB 自动回退 COM",
            "word": "Microsoft Word COM：本机装有 Microsoft Office 时可用",
        }[key]
        style = (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTORADIOBUTTON |
                 (WS_GROUP if i == 0 else 0))
        radio = user32.CreateWindowExW(
            0, "BUTTON", label, style, 36,
            HELP_LAYOUT["radio_top"] + i * 44, width - 72, 38,
            hwnd, ctypes.c_void_p(2001 + i), HINST, None)
        user32.SendMessageW(radio, WM_SETFONT, FONTS["radio"], 1)
        user32.SendMessageW(radio, BM_SETCHECK, 1 if ENGINE == key else 0, 0)
        RADIO_HWNDS[key] = radio
    check = user32.CreateWindowExW(
        0, "BUTTON", "保留中间 PDF（用于排查，默认不勾选）",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_GROUP | BS_AUTOCHECKBOX,
        36, HELP_LAYOUT["check_top"], width - 72, 38, hwnd,
        ctypes.c_void_p(2101), HINST, None)
    user32.SendMessageW(check, WM_SETFONT, FONTS["radio"], 1)
    user32.SendMessageW(check, BM_SETCHECK, 1 if KEEP_PDF else 0, 0)
    CHECK_HWND = check
    check2 = user32.CreateWindowExW(
        0, "BUTTON", "批量完成后自动合并结果 PDF（按处理顺序，默认不勾选）",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        36, HELP_LAYOUT["check_top"] + 42, width - 72, 38, hwnd,
        ctypes.c_void_p(2102), HINST, None)
    user32.SendMessageW(check2, WM_SETFONT, FONTS["radio"], 1)
    user32.SendMessageW(check2, BM_SETCHECK, 1 if AUTO_MERGE else 0, 0)
    AUTO_MERGE_HWND = check2


def open_merge():
    global HWND_MERGE
    if HWND_MERGE and user32.IsWindow(HWND_MERGE):
        user32.SetForegroundWindow(HWND_MERGE)
        return
    hwnd = user32.CreateWindowExW(
        0, "PdfToolMerge", "PDF 合并", WS_POPUP | WS_VISIBLE |
        WS_CLIPCHILDREN, 100, 100, MERGE_W, MERGE_H, None, None, HINST, None)
    popup(hwnd, MERGE_W, MERGE_H)
    HWND_MERGE = hwnd


def create_merge_controls(hwnd):
    global HWND_MERGE_LIST, HWND_MERGE_NAME
    lst = user32.CreateWindowExW(
        0, "LISTBOX", None,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | LBS_NOTIFY |
        LBS_NOINTEGRALHEIGHT,
        MERGE_LIST[0], MERGE_LIST[1], MERGE_LIST[2] - MERGE_LIST[0],
        MERGE_LIST[3] - MERGE_LIST[1], hwnd, ctypes.c_void_p(3001),
        HINST, None)
    user32.SendMessageW(lst, WM_SETFONT, FONTS["list"], 1)
    HWND_MERGE_LIST = lst
    # The five small buttons and the big "start" button are drawn by the
    # parent (draw_merge) and hit-tested in WM_LBUTTONUP. Using native BUTTON
    # children here would cover the rounded custom drawing.
    name = user32.CreateWindowExW(
        0, "EDIT", "合并结果.pdf",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        MERGE_NAME_EDIT[0], MERGE_NAME_EDIT[1],
        MERGE_NAME_EDIT[2] - MERGE_NAME_EDIT[0],
        MERGE_NAME_EDIT[3] - MERGE_NAME_EDIT[1],
        hwnd, ctypes.c_void_p(3201), HINST, None)
    user32.SendMessageW(name, WM_SETFONT, FONTS["status"], 1)
    HWND_MERGE_NAME = name
    seed = os.environ.get("PDFTOOL_SEED_FILES")
    if seed:
        add_merge_files([p for p in seed.split("|") if p])
    refresh_merge_list()
    # Internal GUI probes: exercise the same in-process click path as a user.
    probe = os.environ.get("PDFTOOL_PROBE_MERGE")
    if probe == "reorder" and len(MERGE_FILES) >= 2:
        user32.SendMessageW(HWND_MERGE_LIST, LB_SETCURSEL, 1, 0)
        box = MERGE_BTNS["up"]
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((box[1] + 20) << 16) | (box[0] + 20))
    elif probe == "go" and len(MERGE_FILES) >= 2:
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((MERGE_GO[1] + 28) << 16) |
                            (MERGE_GO[0] + 100))


def refresh_merge_list(keep_selection=False):
    if not HWND_MERGE_LIST or not user32.IsWindow(HWND_MERGE_LIST):
        return
    index = user32.SendMessageW(HWND_MERGE_LIST, LB_GETCURSEL, 0, 0)
    user32.SendMessageW(HWND_MERGE_LIST, LB_RESETCONTENT, 0, 0)
    for i, (path, pages) in enumerate(MERGE_FILES, 1):
        label = "%02d.  %s    （%d 页）" % (
            i, os.path.basename(path), pages)
        send_text(HWND_MERGE_LIST, LB_ADDSTRING, 0, label)
    if keep_selection and 0 <= index < len(MERGE_FILES):
        user32.SendMessageW(HWND_MERGE_LIST, LB_SETCURSEL, index, 0)
    if HWND_MERGE and user32.IsWindow(HWND_MERGE):
        user32.InvalidateRect(HWND_MERGE, None, True)


def add_merge_files(paths):
    added, failed = 0, []
    for path in paths:
        if not path.lower().endswith(".pdf"):
            continue
        if any(os.path.normcase(os.path.abspath(path)) ==
               os.path.normcase(os.path.abspath(p)) for p, _ in MERGE_FILES):
            continue
        try:
            pages = pdf_pages(path)
        except Exception as exc:
            failed.append("%s（%s）" % (os.path.basename(path), exc))
            continue
        MERGE_FILES.append((os.path.abspath(path), pages))
        added += 1
    refresh_merge_list()
    return added, failed


def read_merge_name():
    if not HWND_MERGE_NAME or not user32.IsWindow(HWND_MERGE_NAME):
        return "合并结果.pdf"
    n = user32.GetWindowTextLengthW(HWND_MERGE_NAME)
    buf = ctypes.create_unicode_buffer(max(2, n + 1))
    user32.GetWindowTextW(HWND_MERGE_NAME, buf, len(buf))
    return buf.value.strip() or "合并结果.pdf"


def start_job(targets, merge_dir=None):
    STATE["busy"] = True
    STATE["start"] = time.time()
    STATE["elapsed"] = ""
    STATE["progress"] = 0.0
    STATE["progress_label"] = "准备中"
    CANCEL.clear()
    stale, freed = sweep_scratch()
    if stale:
        Q.put(("log", "已清理上次中断遗留的 %d 个临时文件（%.0f MB）"
               % (stale, freed / 1048576)))
    log_path = open_log_file()
    if len(targets) == 1:
        STATE["status"] = "已选择：%s" % os.path.basename(targets[0])
    else:
        STATE["status"] = "已选择文件夹：共 %d 个 Word 文档" % len(targets)
    Q.put(("log", "—" * 46))
    if log_path:
        Q.put(("log", "本次运行日志：%s" % log_path))
    threading.Thread(target=conversion_worker,
                     args=(targets, ENGINE, KEEP_PDF, AUTO_MERGE, merge_dir),
                     daemon=True).start()
    user32.InvalidateRect(HWND_MAIN, None, True)


def start_merge_job():
    if len(MERGE_FILES) < 2:
        append_log_lines(["合并至少需要 2 个 PDF，请先添加文件。"])
        return
    out = unique_output(os.path.dirname(MERGE_FILES[0][0]),
                        read_merge_name())
    paths = [p for p, _ in MERGE_FILES]
    STATE["busy"] = True
    STATE["start"] = time.time()
    STATE["elapsed"] = ""
    STATE["progress"] = 0.0
    STATE["progress_label"] = "准备合并"
    STATE["status"] = "正在合并 %d 个 PDF…" % len(paths)
    CANCEL.clear()
    log_path = open_log_file()
    Q.put(("log", "—" * 46))
    if log_path:
        Q.put(("log", "本次运行日志：%s" % log_path))
    Q.put(("log", "合并顺序：" + " → ".join(
        os.path.basename(p) for p in paths)))
    threading.Thread(target=merge_worker, args=(paths, out),
                     daemon=True).start()
    if HWND_MERGE and user32.IsWindow(HWND_MERGE):
        user32.DestroyWindow(HWND_MERGE)
    user32.InvalidateRect(HWND_MAIN, None, True)


# ------------------------------------------------------------------- procedures
def main_proc(hwnd, msg, wparam, lparam):
    global STATE, MODE
    if msg == WM_CREATE:
        create_log_control(hwnd)
        user32.SetTimer(hwnd, 1, 200, None)
        return 0
    if msg == WM_TIMER:
        changed = False
        additions = []
        while True:
            try:
                kind, payload = Q.get_nowait()
            except queue.Empty:
                break
            changed = True
            if kind == "log":
                additions.append(str(payload))
            elif kind == "status":
                STATE["status"] = str(payload).replace("\n", " ")
            elif kind == "progress":
                STATE["progress"], STATE["progress_label"] = payload
            elif kind == "done":
                STATE["status"] = "全部完成｜%s（完整结果见下方日志）" % (
                    str(payload).split("\n")[0].strip())
                STATE["busy"] = False
                STATE["progress"] = 1.0
                STATE["progress_label"] = "全部完成"
                additions.append(str(payload))
                additions.append("=== 任务结束 ===")
            elif kind == "error":
                STATE["status"] = "出错：" + str(payload)
                STATE["busy"] = False
                STATE["progress_label"] = "任务异常"
                additions.append("=== 任务异常结束 ===")
        if additions:
            append_log_lines(additions)
        if STATE["busy"] and STATE["start"]:
            seconds = int(time.time() - STATE["start"])
            value = "%02d:%02d" % (seconds // 60, seconds % 60)
            if value != STATE["elapsed"]:
                STATE["elapsed"] = value
                changed = True
        if changed:
            user32.InvalidateRect(hwnd, None, True)
        return 0
    if msg == WM_CTLCOLORSTATIC or msg == WM_CTLCOLORLISTBOX:
        hdc = wparam
        gdi32.SetTextColor(hdc, INK)
        gdi32.SetBkColor(hdc, BG)
        return BRUSH_BG
    if msg == WM_PAINT:
        draw_main(hwnd)
        return 0
    if msg == WM_MOUSEMOVE:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if in_box(x, y, main_help_box()):
            hover = "help"
        elif in_box(x, y, main_close_box()):
            hover = "close"
        elif in_box(x, y, MODE_FILE):
            hover = "mode_file"
        elif in_box(x, y, MODE_DIR):
            hover = "mode_dir"
        elif in_box(x, y, MODE_MERGE):
            hover = "mode_merge"
        elif in_box(x, y, MAIN_BTN):
            hover = "main"
        else:
            hover = None
        if hover != STATE["hover"]:
            STATE["hover"] = hover
            user32.InvalidateRect(hwnd, None, True)
        return 0
    if msg == WM_LBUTTONDOWN:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if y < TITLE_H and not in_box(x, y, main_help_box()) and not in_box(
                x, y, main_close_box()):
            user32.SendMessageW(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        return 0
    if msg == WM_LBUTTONUP:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if in_box(x, y, main_close_box()):
            request_close(hwnd)
        elif in_box(x, y, main_help_box()):
            open_help()
        elif not STATE["busy"] and in_box(x, y, MODE_FILE):
            MODE = "file"
            user32.InvalidateRect(hwnd, None, True)
        elif not STATE["busy"] and in_box(x, y, MODE_DIR):
            MODE = "dir"
            user32.InvalidateRect(hwnd, None, True)
        elif not STATE["busy"] and in_box(x, y, MODE_MERGE):
            MODE = "merge"
            user32.InvalidateRect(hwnd, None, True)
        elif in_box(x, y, MAIN_BTN) and STATE["busy"]:
            # The main button doubles as the cancel button while a job runs;
            # before this there was no way to stop a 50-document batch.
            CANCEL.set()
            STATE["status"] = "正在取消…（当前文档处理完后停止）"
            STATE["progress_label"] = "正在取消"
            append_log_lines(["已请求取消：当前文档处理完后停止，已完成的结果保留。"])
            user32.InvalidateRect(hwnd, None, True)
        elif in_box(x, y, MAIN_BTN) and not STATE["busy"]:
            if MODE == "file":
                path = choose_docx(hwnd)
                targets = [path] if path else []
                merge_dir = None
            elif MODE == "dir":
                folder = choose_folder(hwnd)
                targets = collect_docx(folder) if folder else []
                merge_dir = folder
            else:
                open_merge()
                return 0
            if not targets:
                STATE["status"] = (
                    "未选择文件。" if MODE == "file"
                    else "所选文件夹中未找到 .docx 文档。")
                user32.InvalidateRect(hwnd, None, True)
                return 0
            start_job(targets, merge_dir)
        return 0
    if msg == WM_DESTROY:
        user32.KillTimer(hwnd, 1)
        user32.PostQuitMessage(0)
        return 0
    return user32.DefWindowProcW(hwnd, msg, wparam, lparam)


def help_proc(hwnd, msg, wparam, lparam):
    global ENGINE, KEEP_PDF, AUTO_MERGE, HWND_HELP
    if msg == WM_CREATE:
        create_help_controls(hwnd)
        return 0
    if msg == WM_COMMAND:
        cid = wparam & 0xFFFF
        for i, key in enumerate(("wps", "cli", "word")):
            if cid == 2001 + i:
                ENGINE = key
                return 0
        if cid == 2101 and CHECK_HWND:
            KEEP_PDF = bool(user32.SendMessageW(
                CHECK_HWND, BM_GETCHECK, 0, 0))
            return 0
        if cid == 2102 and AUTO_MERGE_HWND:
            AUTO_MERGE = bool(user32.SendMessageW(
                AUTO_MERGE_HWND, BM_GETCHECK, 0, 0))
            return 0
        return 0
    if msg == WM_CTLCOLORSTATIC:
        hdc = wparam
        gdi32.SetTextColor(hdc, INK)
        gdi32.SetBkColor(hdc, BG)
        return BRUSH_BG
    if msg == WM_PAINT:
        draw_help(hwnd)
        return 0
    if msg == WM_MOUSEMOVE:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        hover = "help_close" if in_box(x, y, help_close_box()) else None
        if hover != STATE["hover"]:
            STATE["hover"] = hover
            user32.InvalidateRect(hwnd, None, True)
        return 0
    if msg == WM_LBUTTONDOWN:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if y < TITLE_H and not in_box(x, y, help_close_box()):
            user32.SendMessageW(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        return 0
    if msg == WM_LBUTTONUP:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if in_box(x, y, help_close_box()):
            user32.DestroyWindow(hwnd)
        return 0
    if msg == WM_DESTROY:
        HWND_HELP = None
        return 0
    return user32.DefWindowProcW(hwnd, msg, wparam, lparam)


def merge_action(action, owner):
    """Handle one button of the merge window (all buttons are custom-drawn)."""
    if action == "add":
        paths = choose_pdfs(owner)
        if paths:
            added, failed = add_merge_files(paths)
            append_log_lines(
                ["合并列表：新增 %d 个 PDF，当前共 %d 个。"
                 % (added, len(MERGE_FILES))] +
                ["添加失败：%s" % f for f in failed])
    elif action in ("up", "down", "remove"):
        index = user32.SendMessageW(HWND_MERGE_LIST, LB_GETCURSEL, 0, 0)
        if index < 0 or index >= len(MERGE_FILES):
            append_log_lines(["请先在列表中选中一个 PDF。"])
        elif action == "remove":
            removed = MERGE_FILES.pop(index)
            refresh_merge_list(True)
            append_log_lines(["已移除：%s" % os.path.basename(removed[0])])
        elif action == "up" and index > 0:
            MERGE_FILES[index - 1], MERGE_FILES[index] = (
                MERGE_FILES[index], MERGE_FILES[index - 1])
            refresh_merge_list(False)
            user32.SendMessageW(HWND_MERGE_LIST, LB_SETCURSEL, index - 1, 0)
        elif action == "down" and index < len(MERGE_FILES) - 1:
            MERGE_FILES[index + 1], MERGE_FILES[index] = (
                MERGE_FILES[index], MERGE_FILES[index + 1])
            refresh_merge_list(False)
            user32.SendMessageW(HWND_MERGE_LIST, LB_SETCURSEL, index + 1, 0)
    elif action == "clear":
        MERGE_FILES.clear()
        refresh_merge_list()
        append_log_lines(["合并列表已清空。"])
    elif action == "go" and not STATE["busy"]:
        start_merge_job()
    if HWND_MERGE and user32.IsWindow(HWND_MERGE):
        user32.InvalidateRect(HWND_MERGE, None, True)


def merge_proc(hwnd, msg, wparam, lparam):
    global STATE, HWND_MERGE
    if msg == WM_CREATE:
        create_merge_controls(hwnd)
        return 0
    if msg == WM_COMMAND:
        cid = wparam & 0xFFFF
        mapping = {3101: "add", 3102: "up", 3103: "down",
                   3104: "remove", 3105: "clear", 3301: "go"}
        if cid in mapping:
            merge_action(mapping[cid], hwnd)
        return 0
    if msg == WM_CTLCOLORSTATIC or msg == WM_CTLCOLORLISTBOX:
        hdc = wparam
        gdi32.SetTextColor(hdc, INK)
        gdi32.SetBkColor(hdc, BG)
        return BRUSH_BG
    if msg == WM_PAINT:
        draw_merge(hwnd)
        return 0
    if msg == WM_MOUSEMOVE:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        hover = None
        for key, box in MERGE_BTNS.items():
            if in_box(x, y, box):
                hover = key
        if y < TITLE_H and in_box(x, y, MERGE_CLOSE):
            hover = "merge_close"
        if hover != STATE["hover"]:
            STATE["hover"] = hover
            user32.InvalidateRect(hwnd, None, True)
        return 0
    if msg == WM_LBUTTONDOWN:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if y < TITLE_H and not in_box(x, y, MERGE_CLOSE):
            user32.SendMessageW(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        return 0
    if msg == WM_LBUTTONUP:
        x, y = (lparam & 0xFFFF), (lparam >> 16) & 0xFFFF
        if in_box(x, y, MERGE_CLOSE):
            user32.DestroyWindow(hwnd)
        else:
            for key, box in MERGE_BTNS.items():
                if in_box(x, y, box):
                    merge_action(key, hwnd)
                    break
            else:
                if in_box(x, y, MERGE_GO):
                    merge_action("go", hwnd)
        return 0
    if msg == WM_DESTROY:
        HWND_MERGE = None
        return 0
    return user32.DefWindowProcW(hwnd, msg, wparam, lparam)


# ------------------------------------------------------------------ selftest
def selftest(target, engine, keep_pdf=False, auto_merge=False):
    import tempfile
    log_path = os.path.join(tempfile.gettempdir(), "pdf_tool_selftest.log")
    lines = []

    def sink(message):
        lines.append(str(message))

    opt.emit = lambda zh, en=None: sink(zh)
    started = time.time()
    targets = collect_docx(target) if os.path.isdir(target) else [target]
    lines.append("targets = %d" % len(targets))
    reports, errors = run_batch(targets, engine, keep_pdf, sink)
    merged = None
    if auto_merge and len(reports) >= 2:
        out = unique_output(os.path.abspath(target)
                            if os.path.isdir(target) else
                            os.path.dirname(os.path.abspath(target)),
                            "批量合并结果.pdf")
        merged = (out, merge_pdfs([r["out"] for r in reports], out, sink))
        lines.append("merged = %s / pages = %d" % (out, merged[1]))
    lines.append("RESULT: %s" % ("OK" if not errors and reports else "FAILED"))
    for i, report in enumerate(reports, 1):
        for key in ("engine", "pages", "pdf_size", "out_size", "checked",
                    "same", "byte_same", "diff", "no_candidate", "skipped",
                    "pdf_removed", "pdf_final", "keep_reason", "retained"):
            if key in report:
                lines.append("[%d] %s = %s" % (i, key, report[key]))
        lines.append("[%d] out = %s" % (i, report["out"]))
        lines.append("[%d] interim_exists_after = %s"
                     % (i, os.path.exists(report["tmp"])))
        lines.append("[%d] final_exists = %s"
                     % (i, os.path.exists(report["out"])))
    for name, err in errors:
        lines.append("ERROR %s : %s" % (name, err))
    lines.append("elapsed = %.1f s" % (time.time() - started))
    with open(log_path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))
    return 0 if not errors and reports else 1


def selftest_merge(out, sources):
    import tempfile
    log_path = os.path.join(tempfile.gettempdir(), "pdf_tool_selftest.log")
    lines = []

    def sink(message):
        lines.append(str(message))

    lines.append("sources = %d" % len(sources))
    pages = merge_pdfs(list(sources), out, sink)
    lines.append("RESULT: OK")
    lines.append("out = %s" % out)
    lines.append("pages = %d" % pages)
    lines.append("size = %d" % os.path.getsize(out))
    import pymupdf
    doc = pymupdf.open(out)
    try:
        lines.append("verified_pages = %d" % doc.page_count)
        lines.append("first_page_text = %r" % doc[0].get_text()[:80])
    finally:
        doc.close()
    with open(log_path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))
    return 0


# ----------------------------------------------------------------------- main
def main():
    global HWND_MAIN
    if "--selftest" in sys.argv:
        i = sys.argv.index("--selftest")
        target = sys.argv[i + 1] if len(sys.argv) > i + 1 else ""
        engine = sys.argv[i + 2] if len(sys.argv) > i + 2 else "wps"
        keep = (sys.argv[i + 3] == "keep") if len(sys.argv) > i + 3 else False
        auto = ("auto" in sys.argv[i + 3:])
        return selftest(target, engine, keep, auto)
    if "--selftest-merge" in sys.argv:
        i = sys.argv.index("--selftest-merge")
        args = [a for a in sys.argv[i + 1:] if not a.startswith("--")]
        if len(args) < 3:
            print("usage: --selftest-merge <out.pdf> <p1.pdf> <p2.pdf> ...")
            return 2
        return selftest_merge(args[0], args[1:])

    FONTS.update({
        "title": make_font(19, True),
        "button": make_font(22, True),
        "mode": make_font(16, True),
        "status": make_font(16),
        "small": make_font(14),
        "mark": make_font(26),
        "help": make_font(16),
        "radio": make_font(16),
        "log": make_font(14),
        "list": make_font(15),
        "progress": make_font(14, True),
    })
    MAIN_PROC = WNDPROC(main_proc)
    HELP_PROC = WNDPROC(help_proc)
    MERGE_PROC = WNDPROC(merge_proc)
    create_class("PdfToolMain", MAIN_PROC)
    create_class("PdfToolHelp", HELP_PROC)
    create_class("PdfToolMerge", MERGE_PROC)
    hwnd = user32.CreateWindowExW(
        0, "PdfToolMain", "Word → PDF 保真优化工具",
        WS_POPUP | WS_VISIBLE | WS_CLIPCHILDREN,
        100, 100, MAIN_W, MAIN_H, None, None, HINST, None)
    if not hwnd:
        raise ctypes.WinError(ctypes.get_last_error())
    popup(hwnd, MAIN_W, MAIN_H)
    HWND_MAIN = hwnd
    append_log_lines([
        "工具已就绪｜V1.3",
        "支持：单个 Word → 最终 PDF；整个文件夹批量；多个 PDF 按选择顺序合并。",
        "请选择处理模式后点击主按钮；处理过程与结果会持续记录在本日志中。",
        "提示：右上角「?」可查看使用方法、原理与选项说明。",
    ])
    if os.environ.get("PDFTOOL_AUTO_HELP") == "1":
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((TITLE_H // 2) << 16) | (main_help_box()[0] + 10))
    probe_mode = os.environ.get("PDFTOOL_PROBE_MODE")
    if probe_mode == "dir":
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((MODE_DIR[1] + 20) << 16) | (MODE_DIR[0] + 20))
    elif probe_mode == "merge":
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((MODE_MERGE[1] + 20) << 16) |
                            (MODE_MERGE[0] + 20))
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((MAIN_BTN[1] + 40) << 16) |
                            (MAIN_BTN[0] + 40))
    if os.environ.get("PDFTOOL_PROBE_CLICK") == "1":
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0,
                            ((MAIN_BTN[1] + 40) << 16) |
                            (MAIN_BTN[0] + 40))
    message = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
        user32.TranslateMessage(ctypes.byref(message))
        user32.DispatchMessageW(ctypes.byref(message))
    return 0


if __name__ == "__main__":
    sys.exit(main())
