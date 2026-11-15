# -*- coding: utf-8 -*-
"""PDF 跳转全量校验与补链层。

用途
----
Word 文档里的目录、交叉引用大多是 ``HYPERLINK \\l "书签"`` 域（或原生
``<w:hyperlink w:anchor>``）。WPS/Word 导出 PDF 时会把它们转成 /Link 注解，
但实测存在漏转：域结构异常、书签被删过、标题改写、模板换代，都会让 PDF 里那
段文字看着像链接却点不动。本模块在「图片已经回填完成」的交付 PDF 上做一次
**全量**（不抽样）校验，并把缺的那几条按 Word 原稿的书签目标补回去。

它只增删 /Link 注解与相关字典键，不动页面内容流、不动图像流、不动版式，因此
可以和图像回填阶段任意先后执行；放在合成总 PDF 之前，每份分册各自补齐，合成
后各分册内部的跳转仍然成立。

与「按文字猜」的旧做法的区别
----------------------------
1. 源位置按**文档顺序对齐**：Word 里第 k 个显示文字相同的域，对应 PDF 里第 k
   处出现，而不是「第一个命中就贴」。7 条「（点此返回）」不会再互相覆盖。
2. 目标页解析优先用 **docx 书签所在标题**去匹配 PDF 大纲，其次才用显示文字，
   再次才用编号前缀；都对不上就**放弃并报告**，绝不瞎猜一个页。
3. 校验既有注解本身是否可用（目标页越界、目标缺失、热区退化），不只是数个数。
4. 补完注解后按 xref 差集取回刚建的那条，写 /P、/F、/Border，不再「取 Annots
   里最后一个」——页面原本有别的注解时会改错对象。

命令行
------
    python link_fix.py --docx A.docx --pdf B.pdf              # 只校验，就地补链
    python link_fix.py --docx A.docx --pdf B.pdf --dry-run    # 只出报告不改文件
    python link_fix.py --pdf B.pdf                            # 无 docx：只审既有注解
    python link_fix.py --docx A.docx --pdf B.pdf --json r.json
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile

import pymupdf
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# 显示文字太短（"1"、"—"）时，包含式匹配会把整页文字都当成同一个链接。
MIN_TEXT = 2
# 包含式匹配允许的长度差上限：短的要占到长的这个比例以上才算同一条。
COVER_RATIO = 0.6
# 书签所在段落文字截断长度，只用于去大纲里比对标题。
HEADING_CAP = 120
# 热区高度相对字号的比例。实测 WPS 为 1.493，取 1.5 与之对齐。
HOT_ZONE_FACTOR = 1.5

_SPACE = re.compile(r"[\s\u3000\u00a0\u200b]+")
_NUM = re.compile(r"^\d+(?:\.\d+)*")


def norm(text):
    """去掉所有空白（含全角空格、NBSP、零宽），用于跨格式比对。"""
    return _SPACE.sub("", text or "")


def _para_text(para):
    return "".join(t.text or "" for t in para.iter(W + "t"))


# --------------------------------------------------------------- docx 侧读取
def _read_rels(archive):
    """rId -> 外部 URL（只关心超链接关系）。"""
    try:
        raw = archive.read("word/_rels/document.xml.rels")
    except KeyError:
        return {}
    root = etree.fromstring(raw)
    out = {}
    for rel in root:
        if not rel.get("Target"):
            continue
        if rel.get("TargetMode") == "External":
            out[rel.get("Id")] = rel.get("Target")
    return out


def _fields_in_para(para, rels):
    """按文档顺序取出一个段落里的全部超链接。"""
    out = []
    kids = list(para)
    state = {"begin": None, "sep": None, "instr": []}
    for i, kid in enumerate(kids):
        local = etree.QName(kid).localname
        if local == "hyperlink":
            anchor = kid.get(W + "anchor")
            rid = kid.get(R + "id")
            display = "".join(t.text or "" for t in kid.iter(W + "t"))
            if anchor:
                out.append({"kind": "goto", "target": anchor,
                            "display": display, "native": True})
            elif rid and rels.get(rid):
                out.append({"kind": "uri", "target": rels[rid],
                            "display": display, "native": True})
            continue
        # 复杂域：begin -> instrText* -> separate -> 结果 run* -> end
        for fc in kid.iter(W + "fldChar"):
            ftype = fc.get(W + "fldCharType")
            if ftype == "begin" and state["begin"] is None:
                state = {"begin": i, "sep": None, "instr": []}
            elif ftype == "separate" and state["begin"] is not None:
                state["sep"] = i
            elif ftype == "end" and state["begin"] is not None:
                if state["sep"] is not None:
                    shown = "".join(
                        "".join(t.text or "" for t in kids[j].iter(W + "t"))
                        for j in range(state["sep"] + 1, i))
                    instr = "".join(state["instr"])
                    link = _parse_instr(instr, rels)
                    if link is not None:
                        link["display"] = shown
                        link["native"] = False
                        out.append(link)
                state = {"begin": None, "sep": None, "instr": []}
        for it in kid.iter(W + "instrText"):
            if state["begin"] is not None:
                state["instr"].append(it.text or "")
    return out


def _parse_instr(instr, rels):
    if "HYPERLINK" not in instr.upper():
        return None
    rest = instr[instr.upper().index("HYPERLINK") + 9:]
    anchor = re.search(r'\\l\s+"([^"]*)"', rest)
    if anchor:
        return {"kind": "goto", "target": anchor.group(1)}
    # 外链：形如 HYPERLINK [\o "提示"] "https://..."
    quoted = re.findall(r'"([^"]*)"', rest)
    for cand in quoted:
        if "://" in cand or cand.lower().startswith("mailto:"):
            return {"kind": "uri", "target": cand}
    if quoted:
        return {"kind": "uri", "target": quoted[-1]}
    return None


def read_docx_links(docx_path):
    """返回 (fields, bookmarks, stats)。

    fields    : 文档顺序的 [{kind, target, display, native, order}]
    bookmarks : {书签名: 所在段落文字}，用于把书签翻译成标题再去对大纲
    """
    archive = zipfile.ZipFile(docx_path)
    try:
        rels = _read_rels(archive)
        root = etree.fromstring(archive.read("word/document.xml"))
    finally:
        archive.close()

    fields = []
    for para in root.iter(W + "p"):
        for link in _fields_in_para(para, rels):
            link["display"] = link.get("display") or ""
            if link["kind"] == "goto" and not link["target"]:
                continue
            fields.append(link)

    bookmarks = {}
    bookmark_para = {}
    for bs in root.iter(W + "bookmarkStart"):
        name = bs.get(W + "name") or ""
        if not name or name.startswith("_GoBack"):
            continue
        parent = bs.getparent()
        # 书签通常与标题同段；个别书签单独成段，向上一层找一次。
        if etree.QName(parent).localname != "p":
            parent = parent.getparent() if parent is not None else None
        text = _para_text(parent)[:HEADING_CAP] if parent is not None else ""
        bookmarks[name] = text
        bookmark_para[name] = parent

    for i, link in enumerate(fields):
        link["order"] = i
        if link["kind"] == "goto":
            link["target_heading"] = bookmarks.get(link["target"], "")
            link["dangling"] = link["target"] not in bookmarks
        else:
            link["target_heading"] = ""
            link["dangling"] = False

    stats = {"docx_fields": len(fields),
             "goto": sum(1 for f in fields if f["kind"] == "goto"),
             "uri": sum(1 for f in fields if f["kind"] == "uri"),
             "bookmarks": len(bookmarks),
             "dangling": sum(1 for f in fields if f["dangling"])}
    return fields, bookmarks, stats


# ------------------------------------------------------------------ PDF 侧读取
def _link_text(page, rect):
    try:
        return norm(page.get_textbox(pymupdf.Rect(rect)))
    except Exception:
        return ""


def collect_links(doc):
    """全量收集 /Link 注解，按 (页, 纵坐标, 横坐标) 排序。"""
    out = []
    for pno in range(doc.page_count):
        page = doc[pno]
        try:
            links = page.get_links()
        except Exception:
            continue
        for link in links:
            rect = link.get("from")
            if rect is None:
                continue
            item = {"page": pno,
                    "rect": pymupdf.Rect(rect),
                    "text": _link_text(page, rect),
                    "xref": link.get("xref"),
                    "kind": link.get("kind"),
                    "dest": link.get("page"),
                    "uri": link.get("uri"),
                    "named": link.get("nameddest")}
            out.append(item)
    out.sort(key=lambda a: (a["page"], round(a["rect"].y0, 1),
                            round(a["rect"].x0, 1)))
    return out


def named_destinations(doc):
    """命名目标：名字 -> 0-based 页。

    WPS 会把大纲目标注册成 ``wl0``、``wl1``… 这类命名目标，Word 的书签名
    （``hd36C1``）并不会出现在 PDF 里，所以这张表只用来核对既有链接的目标是否
    有效，不能用来把书签翻译成页——那一步只能靠大纲标题。
    """
    result = {}
    try:
        names = doc.resolve_names()
    except Exception:
        return result
    for name, value in names.items():
        page = None
        if isinstance(value, dict):
            page = value.get("page")
        elif isinstance(value, (list, tuple)) and value:
            page = value[0]
        if isinstance(page, int) and page >= 0:
            result[name] = page
    return result


def collect_outline(doc):
    """大纲：(标题, 0-based 起始页)，按文档顺序。"""
    out = []
    try:
        toc = doc.get_toc(simple=True)
    except Exception:
        return out
    for entry in toc:
        if len(entry) < 3:
            continue
        title, page = entry[1], entry[2]
        if page and page > 0:
            out.append((title, page - 1))
    return out


# -------------------------------------------------------------------- 匹配逻辑
def text_matches(a, b):
    """两段规范化后的文字是否可视为同一条链接的文字。"""
    if not a or not b:
        return False
    if a == b:
        return True
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    if len(short) < MIN_TEXT:
        return False
    if short not in long_:
        return False
    return len(short) / float(len(long_)) >= COVER_RATIO


def align_fields(fields, annots):
    """把 docx 域与已有注解配对。

    做法是「以文字为键，同文按序取用」，不是全局保序扫描。原因：目录表里的
    链接按纵向坐标排出来后，几何顺序和 Word 文档顺序并不一致（一个单元格折行
    就会把后面的条目顶到前面），全局保序会把紧邻的两条判成缺失。这里：

      1. 先按规范化文字**精确**配对，同一文字出现多次时按先后次序一一取用，
         `（点此返回）` 那 7 条各自对上自己的那处，不会互相吞掉；
      2. 再用包含式配对剩余项（长文字被拆到两行/两页时用得上），按吻合度从高
         到低取未被占用的注解；
      3. 仍配不上的才算缺失。

    返回 (matched, missing)，matched 元素为 (field, annot, 依据)。
    """
    used = [False] * len(annots)
    matched, pending = [], []

    by_text = {}
    for idx, a in enumerate(annots):
        key = a.get("text") or ""
        if key:
            by_text.setdefault(key, []).append(idx)
    by_uri = {}
    for idx, a in enumerate(annots):
        if a.get("uri"):
            by_uri.setdefault(a["uri"], []).append(idx)

    for field in fields:
        if field["kind"] == "uri":
            pool = by_uri.get(field["target"]) or []
            hit = next((j for j in pool if not used[j]), None)
            if hit is None:
                pending.append(field)
            else:
                used[hit] = True
                matched.append((field, annots[hit], "外链"))
            continue
        pool = by_text.get(norm(field["display"])) or []
        hit = next((j for j in pool if not used[j]), None)
        if hit is None:
            pending.append(field)
        else:
            used[hit] = True
            matched.append((field, annots[hit], "精确"))

    still = []
    for field in pending:
        if field["kind"] == "uri":
            still.append(field)
            continue
        want = norm(field["display"])
        if len(want) < MIN_TEXT:
            still.append(field)
            continue
        best = None
        for j, a in enumerate(annots):
            if used[j] or a.get("kind") not in (None, pymupdf.LINK_GOTO):
                continue
            text = a.get("text", "")
            if not text_matches(want, text):
                continue
            short = min(len(want), len(text))
            long_ = max(len(want), len(text), 1)
            score = short / float(long_)
            cand = (-score, j)
            if best is None or cand < best[0]:
                best = (cand, j)
        if best is None:
            still.append(field)
        else:
            j = best[1]
            used[j] = True
            matched.append((field, annots[j], "近似"))
    return matched, still


def audit_existing(annots, page_count):
    """既有注解自身的健康度：目标页越界、目标缺失、热区退化。"""
    bad = []
    for a in annots:
        rect = a["rect"]
        if rect.is_empty or rect.width < 1 or rect.height < 1:
            bad.append({"page": a["page"] + 1, "why": "热区退化", "text": a["text"]})
            continue
        if a.get("kind") == pymupdf.LINK_GOTO:
            dest = a.get("dest")
            if dest is None:
                bad.append({"page": a["page"] + 1, "why": "目标缺失",
                            "text": a["text"]})
            elif dest < 0 or dest >= page_count:
                bad.append({"page": a["page"] + 1,
                            "why": "目标页越界(%s)" % dest, "text": a["text"]})
    return bad


# ---------------------------------------------------------------- 文字定位索引
def build_line_index(doc):
    """全量行索引：[(页, block号, line号, 规范化行文字)]。

    只存定位用的三元组和文字，不存字符框——1371 页逐字符存下来会吃掉几百 MB。
    真正要用某一行时再对那一页取一次 rawdict，取字符级坐标。
    """
    index = []
    for pno in range(doc.page_count):
        page = doc[pno]
        try:
            data = page.get_text("dict")
        except Exception:
            continue
        for bi, block in enumerate(data.get("blocks", [])):
            for li, line in enumerate(block.get("lines", [])):
                text = norm("".join(s.get("text", "")
                                    for s in line.get("spans", [])))
                if text:
                    index.append((pno, bi, li, text))
    return index


def _chars_of_line(page, block_i, line_i, cache):
    """取一行的字符、坐标、行框与最大字号；cache 以页号为键复用 rawdict。"""
    pno = page.number
    if pno not in cache:
        try:
            cache[pno] = page.get_text("rawdict")["blocks"]
        except Exception:
            cache[pno] = []
    blocks = cache[pno]
    try:
        line = blocks[block_i]["lines"][line_i]
    except (IndexError, KeyError, TypeError):
        return [], "", None, 0.0
    chars, boxes, size = [], [], 0.0
    for span in line.get("spans", []):
        size = max(size, float(span.get("size") or 0))
        for ch in span.get("chars", []):
            c = ch.get("c", "")
            if not c or _SPACE.match(c):
                continue
            chars.append(c)
            boxes.append(pymupdf.Rect(ch["bbox"]))
    return boxes, "".join(chars), pymupdf.Rect(line["bbox"]), size


def _union(boxes):
    rect = pymupdf.Rect(boxes[0])
    for box in boxes[1:]:
        rect.include_rect(box)
    return rect


def _hot_zone(want_box, center_y, font_size):
    """热区：横向取字符实际范围，纵向按字号放大到行距高度。

    实测 WPS 生成的 /Link 热区高度是字号的 1.5 倍（10.45 pt 字 → 15.60 pt 高，
    中位比 1.493），纵向与行中心对称（偏差中位 0.05 pt）。只按字形包围盒贴会
    上下各少 2–4 pt，点击容错变差，也与 WPS 原生注解对不齐。
    """
    rect = pymupdf.Rect(want_box)
    height = max(font_size * HOT_ZONE_FACTOR, rect.height)
    if center_y is not None:
        rect.y0 = center_y - height / 2.0
        rect.y1 = center_y + height / 2.0
    return rect


def _rect_in_line(page, block_i, line_i, want, cache):
    boxes, text, line_box, size = _chars_of_line(page, block_i, line_i, cache)
    if not text:
        return None
    start = text.find(want)
    if start < 0:
        return None
    run = _union(boxes[start:start + len(want)])
    center = (line_box.y0 + line_box.y1) / 2.0 if line_box else \
        (run.y0 + run.y1) / 2.0
    return _hot_zone(run, center, size)


def _rect_across_lines(page, block_i, first_i, line_count, want, cache):
    """跨行兜底：把同一 block 里连续几行拼起来再找，热区取并集。"""
    boxes, text = [], ""
    first_box = last_box = None
    size = 0.0
    for offset in range(line_count):
        line_boxes, line_text, line_box, line_size = _chars_of_line(
            page, block_i, first_i + offset, cache)
        if not line_text:
            break
        if first_box is None:
            first_box = line_box
        last_box = line_box
        size = max(size, line_size)
        text += line_text
        boxes.extend(line_boxes)
    start = text.find(want)
    if start < 0:
        return None
    run = _union(boxes[start:start + len(want)])
    center = run.y0 + run.height / 2.0
    if first_box is not None and last_box is not None:
        center = (first_box.y0 + last_box.y1) / 2.0
    return _hot_zone(run, center, size)


def find_occurrences(doc, display, index, max_join=3):
    """显示文字在 PDF 里的所有出现位置（按文档顺序，字符级精确热区）。

    先在行索引里做**纯子串**预筛（不做长度比限制：链接文字常常只是整句里的一
    小段，例如「为方便各位专家评审，如需返回目录可（点此返回）。」里只标
    「（点此返回）」），命中后再提取该行的字符框算精确热区；单行找不到时按同
    一 block 内连续 2–3 行拼接再试，覆盖折行的情况。
    """
    want = norm(display)
    if len(want) < MIN_TEXT:
        return []
    by_page = {}
    for pno, bi, li, text in index:
        if want in text:
            by_page.setdefault(pno, []).append((bi, li))
    if not by_page:
        return []

    cache = {}
    found = []
    for pno in sorted(by_page):
        page = doc[pno]
        hits = []
        for bi, li in by_page[pno]:
            rect = _rect_in_line(page, bi, li, want, cache)
            if rect is not None:
                hits.append((bi, li, rect))
        if not hits and max_join > 1:
            for bi, li in by_page[pno]:
                for count in range(2, max_join + 1):
                    rect = _rect_across_lines(page, bi, li, count, want, cache)
                    if rect is not None:
                        hits.append((bi, li, rect))
                        break
        # 同一行命中多次时只保留一个（子串预筛已保证整行唯一）
        seen = set()
        for bi, li, rect in hits:
            key = (bi, li)
            if key in seen:
                continue
            seen.add(key)
            found.append((pno, rect))
    cache.clear()
    return found


def resolve_target(field, outline, named, page_count, hint=None):
    """该跳到第几页，按证据强弱依次尝试。

    1. 同名命名目标（少见于 WPS 导出，有则最权威）；
    2. **同文字的其他链接已经指向的页**——同一段显示文字在别处已经有可点的注解
       且目标一致，这比任何文字猜测都可靠。`（点此返回）`这类回目录链接全靠它：
       它们的书签名是 ml36R1…R7，书签所在段落是目录行，用标题去匹配会跳正文；
    3. 书签所在标题去匹配大纲；
    4. 显示文字去匹配大纲；
    5. 编号前缀（14.1.2 → 大纲里以 14.1.2 开头的那条）；
    6. 标题/文字的近似包含匹配。

    都对不上返回 (None, "无匹配")，由调用方报告，绝不瞎猜。
    """
    bookmark = field.get("target") or ""
    if bookmark and bookmark in named:
        page = named[bookmark]
        if page is not None and 0 <= page < page_count:
            return page, "命名目标"

    if hint is not None and 0 <= hint < page_count:
        return hint, "同类链接"

    heading = norm(field.get("target_heading") or "")
    display = norm(field.get("display") or "")
    for candidate, how in ((heading, "书签标题"), (display, "显示文字")):
        if not candidate:
            continue
        for title, page in outline:
            if norm(title) == candidate:
                return page, how

    number = _NUM.match(display)
    if number:
        prefix = number.group(0)
        best = None
        for title, page in outline:
            ntitle = norm(title)
            if ntitle.startswith(prefix):
                # 优先最具体的编号：14.1.2 优于 14
                if best is None or len(prefix) > best[0]:
                    best = (len(prefix), page)
        if best:
            return best[1], "编号前缀"

    for candidate, how in ((heading, "书签标题近似"), (display, "显示文字近似")):
        if not candidate:
            continue
        best = None
        for title, page in outline:
            ntitle = norm(title)
            if not ntitle:
                continue
            short, long_ = sorted((candidate, ntitle), key=len)
            if short in long_ and len(short) / float(len(long_)) >= 0.75:
                score = len(short)
                if best is None or score > best[0]:
                    best = (score, page, title)
        if best:
            return best[1], how
    return None, "无匹配"


# ---------------------------------------------------------------------- 主流程
def _annot_xrefs(doc, page):
    try:
        raw = doc.xref_get_key(page.xref, "Annots")[1] or ""
    except Exception:
        return set()
    return set(int(x) for x in re.findall(r"(\d+) 0 R", raw))


def _decorate(doc, page, before, rect):
    """给刚插入的那条注解补 /P、/F、/Border，返回其 xref 或 None。"""
    after = _annot_xrefs(doc, page)
    new = sorted(after - before)
    if not new:
        return None
    xref = new[-1]
    try:
        doc.xref_set_key(xref, "P", "%d 0 R" % page.xref)
        doc.xref_set_key(xref, "F", "4")          # Print
        doc.xref_set_key(xref, "Border", "[0 0 0]")
    except Exception:
        return xref
    return xref


def analyze(pdf_path, docx_path=None, fields=None, index=None):
    """只读分析：返回报告字典，不修改任何文件。"""
    doc = pymupdf.open(pdf_path)
    try:
        page_count = doc.page_count
        annots = collect_links(doc)
        outline = collect_outline(doc)
        named = named_destinations(doc)
        report = {
            "pdf": os.path.basename(pdf_path),
            "pages": page_count,
            "pdf_links": len(annots),
            "goto_links": sum(1 for a in annots
                              if a.get("kind") == pymupdf.LINK_GOTO),
            "uri_links": sum(1 for a in annots if a.get("uri")),
            "outline": len(outline),
            "named_dests": len(named),
            "bad_existing": audit_existing(annots, page_count),
        }
        if fields is None:
            report.update({"docx": None, "docx_fields": None,
                           "matched": None, "missing": [], "repairable": [],
                           "unresolved": [], "dangling": 0})
            return report

        matched, missing = align_fields(fields, annots)
        report["docx"] = os.path.basename(docx_path) if docx_path else None
        report["docx_fields"] = len(fields)
        report["matched"] = len(matched)
        report["matched_loose"] = sum(1 for _, _, how in matched
                                      if how == "近似")
        report["missing"] = [{"display": f["display"],
                              "target": f["target"],
                              "kind": f["kind"],
                              "dangling": f.get("dangling", False)}
                             for f in missing]
        report["dangling"] = sum(1 for f in missing if f.get("dangling"))

        if not missing:
            report["repairable"] = []
            report["unresolved"] = []
            return report

        if index is None:
            index = build_line_index(doc)
        report["repairable"], report["unresolved"] = _plan(
            doc, fields, missing, outline, named, page_count, index, annots,
            matched)
        return report
    finally:
        doc.close()


def _is_covered(pno, rect, annots):
    """这处文字是不是已经有链接了（热区重叠即算，避免贴第二条重复链接）。"""
    for a in annots:
        if a["page"] != pno:
            continue
        try:
            if a["rect"].intersects(rect):
                return True
        except Exception:
            continue
    return False


def _dest_hints(fields, matched, page_count):
    """同文字链接已指向的页：{规范化显示文字: 目标页}。

    取众数而不是「全部一致才采信」：7 条「（点此返回）」里有一条因为排版原因
    原来就指向另一页（dest=1），严格一致会把整组都判成无提示，退回按书签标题
    去猜——而那正好会跳到正文页而不是目录页。只要众数够明显（≥2 票且过半）就
    采信；票数打平则放弃，交给后面的文字匹配。
    """
    votes = {}
    for field, annot, _how in matched:
        if field["kind"] != "goto":
            continue
        dest = annot.get("dest")
        if not isinstance(dest, int) or dest < 0 or dest >= page_count:
            continue
        votes.setdefault(norm(field["display"]), []).append(dest)
    hints = {}
    for key, dests in votes.items():
        counts = {}
        for dest in dests:
            counts[dest] = counts.get(dest, 0) + 1
        dest, count = max(counts.items(), key=lambda kv: kv[1])
        if count == len(dests) or (count >= 2 and count / float(len(dests)) >= 0.5):
            hints[key] = dest
    return hints


def _plan(doc, fields, missing, outline, named, page_count, index, annots,
          matched):
    """为每条缺失域算好热区与目标页；算不出的进 unresolved。"""
    repairable, unresolved = [], []
    hints = _dest_hints(fields, matched, page_count)
    # 同一显示文字在**全部**域里的序号。只在缺失项里数会把序号整体前移：
    # 正文 p1 与目录 p3 都有「14.1.2…」，目录那条丢了却按序号 0 贴回正文。
    ordinal, counter = {}, {}
    for field in fields:
        if field["kind"] != "goto":
            continue
        key = norm(field["display"])
        if len(key) < MIN_TEXT:
            continue
        ordinal[field["order"]] = counter.get(key, 0)
        counter[key] = counter.get(key, 0) + 1

    for field in missing:
        want = norm(field["display"])
        if field["kind"] == "uri":
            unresolved.append({"display": field["display"],
                               "why": "外链无法在本地重建目标",
                               "target": field["target"]})
            continue
        if len(want) < MIN_TEXT:
            unresolved.append({"display": field["display"],
                               "why": "显示文字过短，无法安全定位",
                               "target": field["target"]})
            continue
        occurrences = find_occurrences(doc, field["display"], index)
        if not occurrences:
            unresolved.append({"display": field["display"],
                               "why": "PDF 文字层找不到该文字",
                               "target": field["target"]})
            continue
        covered = [_is_covered(pno, rect, annots) for pno, rect in occurrences]
        want_index = ordinal.get(field["order"], 0)
        pick = None
        if want_index < len(occurrences) and not covered[want_index]:
            pick = want_index
        else:
            for i, taken in enumerate(covered):
                if not taken:
                    pick = i
                    break
        if pick is None:
            pick = want_index if want_index < len(occurrences) else len(occurrences) - 1
        pno, rect = occurrences[pick]
        dest, how = resolve_target(field, outline, named, page_count,
                                   hints.get(want))
        if dest is None:
            unresolved.append({"display": field["display"],
                               "why": "找不到目标页（书签=%s）" % field["target"],
                               "target": field["target"]})
            continue
        repairable.append({"display": field["display"],
                           "target": field["target"],
                           "page": pno,
                           "rect": tuple(round(v, 2) for v in rect),
                           "dest": dest,
                           "how": how,
                           "why": "未找到对应注解"})
    return repairable, unresolved


def repair(pdf_path, docx_path, dry_run=False, report_json=None):
    """全量校验 + 就地补链。返回报告字典。

    补链写入方式：复制成 .part，对副本 saveIncr，再原子替换原文件。失败时原
    文件保持可用，.part 被清理。
    """
    fields, _bookmarks, docx_stats = read_docx_links(docx_path)
    report = analyze(pdf_path, docx_path, fields)
    report["docx_stats"] = docx_stats
    report["inserted"] = 0

    if dry_run or not report["repairable"]:
        if report_json:
            with open(report_json, "w", encoding="utf-8") as handle:
                json.dump(report, handle, ensure_ascii=False, indent=2)
        return report

    part = pdf_path + ".links.part"
    shutil.copy(pdf_path, part)
    inserted = 0
    try:
        doc = pymupdf.open(part)
        try:
            for item in report["repairable"]:
                page = doc[item["page"]]
                before = _annot_xrefs(doc, page)
                try:
                    page.insert_link({"kind": pymupdf.LINK_GOTO,
                                      "from": pymupdf.Rect(item["rect"]),
                                      "page": item["dest"]})
                except Exception as exc:
                    item["why"] = "插入失败：%s" % exc
                    continue
                xref = _decorate(doc, page, before, pymupdf.Rect(item["rect"]))
                item["xref"] = xref
                inserted += 1
            if inserted:
                doc.saveIncr()
        finally:
            doc.close()
        if inserted:
            os.replace(part, pdf_path)
    except Exception:
        if os.path.exists(part):
            try:
                os.remove(part)
            except OSError:
                pass
        raise
    if os.path.exists(part):
        try:
            os.remove(part)
        except OSError:
            pass

    report["inserted"] = inserted
    # 补完再审一次，证明补的东西真的生效
    after = analyze(pdf_path, docx_path, fields)
    report["after"] = {"pdf_links": after["pdf_links"],
                       "matched": after["matched"],
                       "missing": len(after["missing"]),
                       "bad_existing": len(after["bad_existing"])}
    if report_json:
        with open(report_json, "w", encoding="utf-8") as handle:
            json.dump(report, handle, ensure_ascii=False, indent=2)
    return report


def audit_only(pdf_path):
    """没有 docx 时的退化模式：只审既有注解是否可用。"""
    doc = pymupdf.open(pdf_path)
    try:
        annots = collect_links(doc)
        return {"pdf": os.path.basename(pdf_path),
                "pages": doc.page_count,
                "pdf_links": len(annots),
                "bad_existing": audit_existing(annots, doc.page_count)}
    finally:
        doc.close()


def merge_summary(reports, enabled=True):
    """合并前的补链总账；GUI 与自测两条路径都调用同一份，避免口径分叉。"""
    if not enabled or not reports:
        return None
    total = 0
    failed = 0
    unresolved = 0
    for report in reports:
        link = report.get("link") or {}
        total += link.get("inserted") or 0
        unresolved += link.get("unresolved") or 0
        if report.get("link_error"):
            failed += 1
    line = "跳转校验汇总：%d 份分册新增链接 %d 条" % (len(reports), total)
    if unresolved:
        line += "，%d 条未能自动补回" % unresolved
    if failed:
        line += "，%d 份补链失败" % failed
    return line


def format_report(report):
    """给 GUI 日志用的一行摘要 + 明细行。"""
    lines = []
    if report.get("docx_fields") is None:
        lines.append("跳转校验：%d 页，/Link 注解 %d 个，异常 %d 个"
                     % (report["pages"], report["pdf_links"],
                        len(report.get("bad_existing") or [])))
    else:
        lines.append("跳转校验：Word 域 %d 条，PDF 注解 %d 个，已对上 %d 条，"
                     "缺 %d 条" % (report["docx_fields"], report["pdf_links"],
                                   report.get("matched") or 0,
                                   len(report.get("missing") or [])))
        if report.get("dangling"):
            lines.append("  其中 %d 条域指向 Word 里不存在的书签（原稿本身是悬空引用）"
                         % report["dangling"])
        for item in report.get("repairable") or []:
            lines.append("  补链：第 %d 页「%s」→ 第 %d 页（%s）"
                         % (item["page"] + 1, item["display"][:28],
                            item["dest"] + 1, item["how"]))
        for item in report.get("unresolved") or []:
            lines.append("  未补：「%s」%s" % (item["display"][:28], item["why"]))
        for item in report.get("bad_existing") or []:
            lines.append("  既有注解异常：第 %d 页「%s」%s"
                         % (item["page"], item["text"][:20], item["why"]))
        if report.get("inserted") is not None:
            after = report.get("after") or {}
            lines.append("补链结果：新增 %d 条，复核后注解 %s 个，仍缺 %s 条"
                         % (report["inserted"], after.get("pdf_links"),
                            after.get("missing")))
    return lines


def main():
    parser = argparse.ArgumentParser(
        description="PDF 跳转全量校验与补链（只动 /Link 注解，不动内容与版式）")
    parser.add_argument("--docx", help="Word 原稿，提供权威的链接清单")
    parser.add_argument("--pdf", required=True, help="待校验的 PDF")
    parser.add_argument("--dry-run", action="store_true",
                        help="只出报告，不修改 PDF")
    parser.add_argument("--json", help="把完整报告写成 JSON")
    args = parser.parse_args()

    if not args.docx:
        report = audit_only(args.pdf)
        for line in format_report(report):
            print(line)
        if args.json:
            with open(args.json, "w", encoding="utf-8") as handle:
                json.dump(report, handle, ensure_ascii=False, indent=2)
        return 0

    report = repair(args.pdf, args.docx, args.dry_run, args.json)
    for line in format_report(report):
        print(line)
    return 0 if not (report.get("unresolved") or
                     report.get("bad_existing")) else 1


if __name__ == "__main__":
    sys.exit(main())
