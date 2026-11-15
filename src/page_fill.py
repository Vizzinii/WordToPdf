# -*- coding: utf-8 -*-
"""投标文件页码回填层：把「对应页码 / 佐证材料对应页码」这类空缺按总 PDF 的真实页码补齐。

为什么需要「导出跑两遍」
------------------------
页码值取决于该材料在总 PDF 里的位置，而总 PDF 由各分册拼接而成，于是：

    页码值  ← 分册在总 PDF 里的偏移 + 册内页
    总 PDF  ← 各分册先导出成 PDF
    分册 PDF ← 页码已填好的 docx

不需要真的先合出一份几百 MB 的总 PDF。总页码可以算出：

    总页码 = 该分册之前所有分册页数之和 + 该材料在本分册内的页序号

所以流程是「探测轮导出 → 求解页码 → 回填 docx → 定稿轮导出 → 合并」。
回填会改变表格内容，极端情况下可能引起重排、页数变化，进而平移后续分册的
页码；因此定稿轮之后要做一次收敛检测，最多迭代 MAX_ROUNDS 轮。

页码怎么定位（不依赖 docx 的自动编号）
--------------------------------------
分册 PDF 的正文里，每份材料都以「编号 + 标题」开头（`8.1.1 检测报告（报告编号：
…）`、`9.2.1.1 …`），这些标题在文字层里可以精确读到。于是对每个需要填页码的
表格：

  1. 读出该表 N 个数据行；
  2. 在分册 PDF 里找出**编号后缀连续为 1..N、页序单调不减**的那一组标题；
  3. 第 i 行就取第 i 个标题所在页。

这一步是自校验的：只有恰好 N 条、后缀严格连续、页序不回退才采用，否则不填并
报「结构不匹配」。它同时避开了目录页里的条目——目录页上的编号是跳跃的（28、
67、96…），不会构成 1..N 的连续序列。

行内有唯一标识（报告编号 / 证书编号 / 专利号 / 合同编号）时，优先用标识直接
检索，作为比结构匹配更强的证据；两者冲突时以标识为准并记录。

只改目标单元格里的空文本节点，不动段落属性（字体、字号、颜色、下划线）与
表格结构，因此版式与样式保持不变。
"""
import copy
import os
import re
import shutil
import zipfile

import pymupdf
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 表头里出现这些词就认为该列需要填页码。
PAGE_HEADER_HINTS = ("对应页码", "页码", "对应投标文件页码", "佐证材料对应页码")
# 表头里出现这些词，该行是表头/合计/占位行，不填。
SKIP_ROW_HINTS = ("序号", "合计", "……", "...")
# 行内可用于直接检索的标识列（表头包含即算）。
ID_HEADER_HINTS = ("报告编号", "证书编号", "专利号", "合同编号", "编号")
# 同一批标题的最长编号层级，防止把 9.2.1.1 误拆成 9.2.1 + 1.1。
MAX_NUM_DEPTH = 6
MAX_ROUNDS = 3

NUM_HEAD = re.compile(r"^\s*(\d+(?:\.\d+){1,5})\s*[、.．\s]?\s*(.*)$")
NUM_TAIL = re.compile(r"(\d+(?:\.\d+)+)\s*$")
SPACE = re.compile(r"[\s\u3000\u00a0\u200b]+")


def norm(text):
    return SPACE.sub("", text or "")


def cell_text(tc):
    return "".join(t.text or "" for t in tc.iter(W + "t")).strip()


# ------------------------------------------------------------------ docx 读取
def _para_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()


def read_tables(docx_path):
    """读出所有含页码列的表格。

    返回 [{index, header, page_cols, id_cols, rows, context}]
      rows    : [{row_no, cells, page_cells:[{col, text}], ident:{col:text}}]
      context : 表格之前最近的若干段文字（用于在 PDF 里定位表格所属小节）
    """
    with zipfile.ZipFile(docx_path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    body = root.find(W + "body")
    if body is None:
        return []

    tables = []
    kids = list(body)
    ti = 0
    for pos, kid in enumerate(kids):
        if etree.QName(kid).localname != "tbl":
            continue
        index = ti
        ti += 1
        trs = kid.findall(W + "tr")
        if len(trs) < 2:
            continue
        header = [cell_text(tc) for tc in trs[0].findall(W + "tc")]
        page_cols = [i for i, h in enumerate(header)
                     if any(hint in h for hint in PAGE_HEADER_HINTS)]
        if not page_cols:
            continue
        id_cols = [i for i, h in enumerate(header)
                   if any(hint in h for hint in ID_HEADER_HINTS)
                   and i not in page_cols]
        context = []
        for back in range(pos - 1, max(-1, pos - 6), -1):
            text = _para_text(kids[back]) if etree.QName(
                kids[back]).localname == "p" else ""
            if text:
                context.append(text)
            if len(context) >= 3:
                break
        rows = []
        for rno, tr in enumerate(trs[1:], start=2):
            tcs = tr.findall(W + "tc")
            cells = [cell_text(tc) for tc in tcs]
            if not any(c for c in cells):
                continue
            joined = "".join(cells)
            if any(hint in joined for hint in SKIP_ROW_HINTS) and \
                    len(norm(joined)) < 20:
                continue
            page_cells = []
            for ci in page_cols:
                if ci >= len(tcs):
                    continue
                page_cells.append({"col": ci, "text": cells[ci]})
            ident = {}
            for ci in id_cols:
                if ci < len(cells) and cells[ci] and cells[ci] != "/":
                    ident[ci] = cells[ci]
            rows.append({"row_no": rno, "cells": cells,
                         "page_cells": page_cells, "ident": ident})
        if rows:
            tables.append({"index": index, "header": header,
                           "page_cols": page_cols, "id_cols": id_cols,
                           "rows": rows, "context": context})
    return tables


# ---------------------------------------------------------------- PDF 标题索引
def extract_headings(doc):
    """提取正文里「编号 + 标题」的位置，**一页可以有多条**。

    实测 p556 同时排着 `9.2.1.1 …` 与 `9.2.1.2 …`（两个专利证书标题在同一页），
    因此不能「每页只取第一条」。同时必须跳过目录页：目录页上 1..N 的编号是齐全
    且连续的，会和正文构成两份一模一样的候选。

    目录页的判据必须用「二级目录/三级目录」这种表头词，不能用「目录」两字：材料
    页脚常写「如需返回目录可（点此返回）」，用宽判据会把大部分正文页误判成目录页
    （实测 26 页的分册只剩 2 条标题可提）。
    """
    out = []
    toc_markers = ("二级目录", "三级目录")
    toc_pages = set()
    for pno in range(doc.page_count):
        try:
            head = doc[pno].get_text()[:800]
        except Exception:
            continue
        if any(marker in head for marker in toc_markers):
            toc_pages.add(pno)
    # 目录可能连续多页：从第一个目录页起，把紧随其后、且没有正文标题的页也算进去
    if toc_pages:
        last = max(toc_pages)
        for pno in range(last + 1, min(last + 8, doc.page_count)):
            try:
                text = doc[pno].get_text()
            except Exception:
                break
            first = [l.strip() for l in text.splitlines() if l.strip()][:1]
            if first and NUM_HEAD.match(first[0]):
                toc_pages.add(pno)
            else:
                break

    for pno in range(doc.page_count):
        if pno in toc_pages:
            continue
        try:
            lines = []
            for block in doc[pno].get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    text = "".join(s.get("text", "")
                                   for s in line.get("spans", [])).strip()
                    if text:
                        lines.append(text)
        except Exception:
            continue
        # 标题都排在页首；取开头连续的编号行
        for text in lines[:14]:
            m = NUM_HEAD.match(text)
            if not m:
                if len(out) and out[-1]["page"] == pno:
                    break
                continue
            number = m.group(1).rstrip(".")
            parts = number.split(".")
            if len(parts) < 2 or not m.group(2).strip():
                continue
            out.append({"page": pno, "number": number, "parts": parts,
                        "text": text[:160]})
    return out


def _sequences_for_prefix(headings, prefix, depth):
    """某个父编号下、指定层级的标题，按尾号索引：{尾号: 标题}。"""
    want = prefix + "."
    table = {}
    pages = {}
    for h in headings:
        if len(h["parts"]) != depth or not h["number"].startswith(want):
            continue
        try:
            tail = int(h["parts"][-1])
        except ValueError:
            continue
        table.setdefault(tail, h)
        pages.setdefault(tail, h["page"])
    return table, pages


def _runs_with_suffix_1_to_n(headings, n):
    """找「尾号 1..n」的编号组，返回候选列表。

    表格行号与标题尾号是同一条编号的两面：第 i 行对应尾号 i。真实文档会缺号
    （实测 9.2.2 尾号 1..982 但有 113 个缺号），缺号的页由调用方用相邻页插值补，
    因此这里只要求「尾号 1 存在、最大尾号 >= n」，不要求连续。
    """
    prefixes = {}
    for h in headings:
        depth = len(h["parts"])
        if depth < 2:
            continue
        prefixes.setdefault(depth, set()).add(".".join(h["parts"][:-1]))

    candidates = []
    for depth, group in prefixes.items():
        for prefix in group:
            table, pages = _sequences_for_prefix(headings, prefix, depth)
            if not table or 1 not in table:
                continue
            if max(table) < n:
                continue
            ordered = [pages[k] for k in sorted(pages)]
            if any(ordered[i] > ordered[i + 1] for i in range(len(ordered) - 1)):
                continue
            candidates.append({"prefix": prefix, "depth": depth,
                               "table": table, "pages": pages,
                               "start_page": pages[1],
                               "span": ordered[-1] - ordered[0],
                               "count": len(table)})
    candidates.sort(key=lambda c: (c["start_page"], -c["depth"]))
    return candidates


def _pages_for_rows(candidate, n):
    """按尾号取页；缺号用相邻已编号页插值，返回 (pages, 缺失行号列表)。"""
    table = candidate["table"]
    pages = candidate["pages"]
    known = sorted(pages)
    out = []
    gaps = []
    for row in range(1, n + 1):
        if row in pages:
            out.append(pages[row])
            continue
        gaps.append(row)
        before = [k for k in known if k < row]
        after = [k for k in known if k > row]
        if before and after:
            out.append(pages[before[-1]])
        elif before:
            out.append(pages[before[-1]])
        elif after:
            out.append(pages[after[0]])
        else:
            return None, []
    return out, gaps


def _heading_for_text(headings, text, need=10):
    """在标题里找包含这段文字的那条，返回标题项。"""
    want = norm(text)
    if len(want) < need:
        return None
    hits = [h for h in headings if want in norm(h["text"])]
    if len(hits) == 1:
        return hits[0]
    return hits[0] if hits else None


def _parent_prefix(headings, context, context_pages):
    """推断表格所属的父编号：先按前文匹配标题，再退回上下文页之前最近的标题。"""
    for text in context:
        h = _heading_for_text(headings, text, need=8)
        if h is not None:
            return h["number"]
    if context_pages:
        floor = min(context_pages)
        before = [h for h in headings if h["page"] <= floor]
        if before:
            before.sort(key=lambda h: (h["page"], len(h["parts"])))
            return before[-1]["number"]
    return None


def locate_table_pages(headings, rows, context_pages, context=None):
    """把一张表的 N 行映射到 N 个册内页。返回 (pages, how) 或 (None, 原因)。"""
    n = len(rows)
    if n == 0:
        return None, "无数据行"
    candidates = _runs_with_suffix_1_to_n(headings, n)
    if not candidates:
        return None, "PDF 里找不到 %d 条连续编号的标题" % n

    parent = _parent_prefix(headings, context or [], context_pages)
    chosen = None
    if parent:
        # 子编号 = 父编号 + "." + 序号，层级比父编号多一层
        p_depth = len(parent.split("."))
        exact = [c for c in candidates
                 if c["prefix"] == parent and c["depth"] == p_depth + 1]
        if exact:
            chosen = exact[0]
        else:
            deeper = [c for c in candidates
                      if c["prefix"].startswith(parent + ".")
                      and c["depth"] == p_depth + 2]
            if deeper:
                deeper.sort(key=lambda c: (c["start_page"], -c["depth"]))
                chosen = deeper[0]
    if chosen is None:
        floor = min(context_pages) if context_pages else 0
        after = [c for c in candidates if c["start_page"] >= floor]
        pool = after or candidates
        # 起点最靠近表格前文优先；同起点时取层级更深的（子树优先）
        pool.sort(key=lambda c: (c["start_page"], -c["depth"]))
        chosen = pool[0]
    pages, gaps = _pages_for_rows(chosen, n)
    if pages is None:
        return None, "候选 %s 无法覆盖 %d 行" % (chosen["prefix"], n)
    how = "编号 %s（%d 条命中" % (chosen["prefix"], chosen["count"])
    if gaps:
        how += "，%d 行按相邻页推定" % len(gaps)
    how += "）"
    return pages, how


def find_context_pages(doc, context):
    """在 PDF 里找表格前文所在页，作为结构匹配的下界。"""
    pages = []
    for text in context:
        want = norm(text)[:24]
        if len(want) < 8:
            continue
        for pno in range(min(doc.page_count, 400)):
            if want and want in norm(doc[pno].get_text()):
                pages.append(pno)
                break
    return pages


def match_row_to_heading(headings, row, table_rows=None, row_index=None):
    """用行内最有辨识度的文字去标题里找它所属的那一节，返回标题项或 None。

    只用标题文本做比对，不扫全页文字：表格本身那一页也印着报告编号、证书编号，
    扫全文会把「证据页」定位到「表格页」上（这正是第一版复现率只有 0.2% 的原因）。

    命中多条时判为歧义、返回 None，交给结构匹配兜底，绝不随手取第一条——实测
    「合同金额：…」这类反复出现的文字会同时命中几十条标题，随手取第一条会把
    几十行都填成同一页。
    """
    page_cols = {pc["col"] for pc in row["page_cells"]}
    tokens = []
    for col, value in row["ident"].items():
        token = norm(value)
        if len(token) >= 6:
            tokens.append((1000 + len(token), token))
    for ci, text in enumerate(row["cells"]):
        if ci in page_cols:
            continue
        token = norm(text)
        # 太短的单元格（「发明」「通过」「正偏离」）不具辨识度
        if len(token) >= 10:
            tokens.append((len(token), token))
    tokens.sort(reverse=True)
    for _score, token in tokens:
        hits = [h for h in headings if token in norm(h["text"])]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            # 同一编号出现多次（多页重复印同一标题）时，若全在同一页组内，
            # 取第一处；否则放弃。
            numbers = {h["number"] for h in hits}
            if len(numbers) == 1:
                return hits[0]
    return None


# -------------------------------------------------------------------- 页码求解
CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7,
          "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12, "十三": 13,
          "十四": 14, "十五": 15, "十六": 16, "十七": 17, "十八": 18,
          "十九": 19, "二十": 20, "二十一": 21, "二十二": 22, "二十三": 23,
          "二十四": 24, "二十五": 25}
CHAPTER_CN = re.compile(r"^\s*([一二三四五六七八九十]+)\s*[、.．\s]")
# 引用形式不在行首，前面还有「投标文件章节：」，因此单独一条宽匹配。
CHAPTER_REF = re.compile(r"([一二三四五六七八九十]{1,3})\s*[、.．]\s*\S")


def volume_order(folder):
    """按文件名前的数字前缀排序；无前缀的排最后并按名称。"""
    items = []
    for name in os.listdir(folder):
        if name.startswith("~$"):
            continue
        low = name.lower()
        if low.endswith(".docx"):
            kind = "docx"
        elif low.endswith(".xlsx"):
            kind = "xlsx"
        else:
            continue
        m = re.match(r"^(\d+)", name)
        items.append((int(m.group(1)) if m else 10 ** 6, name, kind))
    items.sort()
    return items


def _vol_title(doc):
    """分册正文第一行，通常就是「章号 + 一级标题」。"""
    for pno in range(min(doc.page_count, 4)):
        try:
            for line in doc[pno].get_text().splitlines():
                line = line.strip()
                if len(line) >= 4 and CHAPTER_CN.match(line):
                    return line[:100]
        except Exception:
            continue
    return ""


def build_plan(folder, pdf_of, progress=None):
    """求解全部分册的全部空缺页码。

    `pdf_of(name)` 给出一个分册文件名对应的 PDF 路径。
    返回 {"volumes": [...], "cells": [...], "unresolved": [...]}

    两遍扫描：第一遍只取页数与章号，第二遍才解析每一格。跨分册引用（偏差表里
    写着「投标文件章节：三、技术参数响应表…」）和按章号定位（封面的评分项）都
    需要先知道全局布局。
    """
    volumes = []
    offset = 0
    for order, (_, name, kind) in enumerate(volume_order(folder), start=1):
        pdf_path = pdf_of(name)
        if not pdf_path or not os.path.exists(pdf_path):
            volumes.append({"order": order, "file": name, "kind": kind,
                            "pages": 0, "offset": offset, "missing": True,
                            "title": "", "chapter": 0})
            continue
        doc = pymupdf.open(pdf_path)
        try:
            pages = doc.page_count
            title = _vol_title(doc) if kind == "docx" else ""
        finally:
            doc.close()
        m = CHAPTER_CN.match(title)
        volumes.append({
            "order": order, "file": name, "kind": kind, "pages": pages,
            "offset": offset, "title": title,
            "chapter": CN_NUM.get(m.group(1), 0) if m else 0})
        offset += pages

    by_chapter = {}
    for vol in volumes:
        if vol.get("chapter"):
            by_chapter.setdefault(vol["chapter"], vol)

    cells = []
    unresolved = []
    for vol in volumes:
        name = vol["file"]
        if vol.get("missing"):
            unresolved.append({"file": name, "why": "该分册还没有 PDF"})
            continue
        if vol["kind"] != "docx":
            continue
        pdf_path = pdf_of(name)
        doc = pymupdf.open(pdf_path)
        try:
            headings = extract_headings(doc)
            for table in read_tables(os.path.join(folder, name)):
                ctx_pages = find_context_pages(doc, table["context"])
                row_pages, row_sources, how = resolve_rows(
                    headings, table["rows"], ctx_pages, table["context"])
                for ri, row in enumerate(table["rows"]):
                    if not row["page_cells"]:
                        continue
                    chapter_ref = None
                    for ci, text in enumerate(row["cells"]):
                        chapter_ref = parse_chapter_ref(text)
                        if chapter_ref:
                            break
                    for pc in row["page_cells"]:
                        if pc["text"].strip() and re.search(r"\d", pc["text"]):
                            continue
                        page_in_vol = row_pages[ri]
                        source = row_sources[ri]
                        base = vol["offset"]
                        if page_in_vol is None and chapter_ref and \
                                chapter_ref in by_chapter:
                            # 跨分册：指向别的分册，取其起始页
                            target = by_chapter[chapter_ref]
                            page_in_vol = 0
                            base = target["offset"]
                            source = "跨分册 第%s章 %s" % (
                                chapter_ref, target["file"][:12])
                        if page_in_vol is None:
                            unresolved.append({
                                "file": name, "table": "T%d" % table["index"],
                                "row": row["row_no"], "col": pc["col"],
                                "cells": [c[:24] for c in row["cells"]],
                                "why": how or "无可信定位"})
                            continue
                        total = base + page_in_vol + 1
                        cells.append({
                            "file": name, "table": "T%d" % table["index"],
                            "row": row["row_no"], "col": pc["col"],
                            "page_in_volume": page_in_vol + 1,
                            "total_page": total, "source": source,
                            "label": (row["cells"][0] if row["cells"] else "")})
        finally:
            doc.close()
    return {"volumes": volumes, "cells": cells,
            "unresolved": unresolved, "total_pages": offset}


# -------------------------------------------------------------------- 写入 docx
def chapter_prefix(volumes):
    """每个分册的一级章号：{文件: 章号}。

    分册的正文都以「章号 + 一级标题」开头（`六、投标产品的相关生产/入网许可证…`），
    章号用中文数字。这是跨分册引用唯一稳定的锚点：24 号的偏差表里写着「投标文件
    章节：三、技术参数响应表…」，指的是第 25 号那一册的章号，不是它自己的。
    """
    out = {}
    for vol in volumes:
        m = CHAPTER_CN.match(vol.get("title") or "")
        if m:
            out[vol["file"]] = CN_NUM.get(m.group(1), 0)
    return out


def parse_chapter_ref(text):
    """从「投标文件章节：三、技术参数响应表…」里取出章号（中文数字 -> int）。"""
    if not text:
        return None
    for match in CHAPTER_REF.finditer(text):
        value = CN_NUM.get(match.group(1))
        if value:
            return value
    return None


def format_value(total_page, template="第{n}页"):
    if template == "plain":
        return str(total_page)
    if "{n}" in template:
        return template.format(n=total_page)
    return "%s%d%s" % (template, total_page, "") if template else str(total_page)


def _clear_cell_runs(tc):
    """清掉单元格里所有段落的所有 run，保留第一个段落的 pPr。"""
    paras = tc.findall(W + "p")
    if not paras:
        p = etree.SubElement(tc, W + "p")
        paras = [p]
    for extra in paras[1:]:
        tc.remove(extra)
    para = paras[0]
    for r in para.findall(W + "r"):
        para.remove(r)
    for h in para.findall(W + "hyperlink"):
        para.remove(h)
    return para


def _make_run(para, text, rpr_template, color="000000", underline=False):
    """写一个 run：复制原 rPr 作为字体基准，再按要求固定颜色与下划线。"""
    run = etree.SubElement(para, W + "r")
    if rpr_template is not None:
        rpr = copy.deepcopy(rpr_template)
        for tag in ("color", "u"):
            for node in rpr.findall(W + tag):
                rpr.remove(node)
        run.append(rpr)
    else:
        rpr = etree.SubElement(run, W + "rPr")
    if color:
        node = etree.SubElement(rpr, W + "color")
        node.set(W + "val", color)
    if not underline:
        node = etree.SubElement(rpr, W + "u")
        node.set(W + "val", "none")
    t = etree.SubElement(run, W + "t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def write_filled_docx(src, dst, cells, template="第{n}页",
                      color="000000", underline=False):
    """把求解结果写进 docx 副本；只改目标单元格的空文本。返回实际写入数。"""
    with zipfile.ZipFile(src) as archive:
        raw = archive.read("word/document.xml")
        infos = archive.infolist()
        blobs = {i.filename: archive.read(i.filename) for i in infos}
    root = etree.fromstring(raw)
    body = root.find(W + "body")

    by_table = {}
    for cell in cells:
        by_table.setdefault(cell["table"], []).append(cell)

    written = 0
    ti = 0
    for kid in body:
        if etree.QName(kid).localname != "tbl":
            continue
        key = "T%d" % ti
        ti += 1
        todo = by_table.get(key)
        if not todo:
            continue
        trs = kid.findall(W + "tr")
        for cell in todo:
            rno = cell["row"] - 1
            if rno < 0 or rno >= len(trs):
                continue
            tcs = trs[rno].findall(W + "tc")
            if cell["col"] >= len(tcs):
                continue
            tc = tcs[cell["col"]]
            rpr = None
            for r in tc.iter(W + "r"):
                found = r.find(W + "rPr")
                if found is not None:
                    rpr = found
                    break
            para = _clear_cell_runs(tc)
            _make_run(para, format_value(cell["total_page"], template), rpr,
                      color=color, underline=underline)
            cell["value"] = format_value(cell["total_page"], template)
            written += 1

    blobs["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = dst + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            ni = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            ni.compress_type = zipfile.ZIP_DEFLATED
            ni.external_attr = info.external_attr
            ni.internal_attr = info.internal_attr
            ni.create_system = info.create_system
            ni.extract_version = 20
            ni.flag_bits = 0
            zout.writestr(ni, blobs[info.filename])
    os.replace(tmp, dst)
    return written


def resolve_rows(headings, rows, context_pages, context):
    """给一张表的每一行定册内页；返回 (pages, sources, 说明)。

    pages[i] 为 None 表示这一行没有可信定位。三级证据：
      1. 行内文字唯一命中某条标题 → 用该标题所在页；
      2. 结构匹配：表格列 N 行对应编号 1..N 的标题（缺号按相邻页插值）；
      3. 行内带明确编号（8.1.14 这类）而表格本身结构不成立时，直接用编号查。
    全部失败就留空并写进未解析报告，绝不填一个「看起来差不多」的页。
    """
    n = len(rows)
    pages = [None] * n
    sources = [""] * n

    struct_pages, how = locate_table_pages(headings, rows, context_pages,
                                           context)
    for i, row in enumerate(rows):
        hit = match_row_to_heading(headings, row)
        if hit is not None:
            pages[i] = hit["page"]
            sources[i] = "标题命中 %s" % hit["number"]

    # 结构匹配只用于「有行仍未定」且表格确实呈 1..N 结构的情况。
    # 否则会把整表强行对齐到不相关的章节标题（实测 32 号 T2 的 29 行被对齐到
    # 10.1/10.2 的章节标题，29 行全填成同一页）。
    if struct_pages is not None and any(p is None for p in pages):
        by_number = {h["number"]: h["page"] for h in headings}
        # 结构可信的前提：至少一半行能通过「行内编号」或「唯一标题命中」独立
        # 验证，或表格首列就是从 1 开始的序号。
        numbered_rows = sum(1 for row in rows
                            if str(row["cells"][0]).strip().isdigit()) \
            if rows and rows[0]["cells"] else 0
        structural = numbered_rows >= max(2, len(rows) * 0.9)
        confirmed = sum(1 for p in pages if p is not None)
        if structural or confirmed >= len(rows) * 0.5:
            for i, row in enumerate(rows):
                if pages[i] is not None:
                    continue
                explicit = _explicit_number(row, by_number)
                if explicit is not None:
                    pages[i] = explicit
                    sources[i] = "行内编号"
                    continue
                pages[i] = struct_pages[i]
                sources[i] = how

    return pages, sources, how


def _explicit_number(row, by_number):
    """行里是否写着本册某个小节号（8.1.14 / 9.2.2.12）。"""
    page_cols = {pc["col"] for pc in row["page_cells"]}
    for ci, text in enumerate(row["cells"]):
        if ci in page_cols:
            continue
        for token in re.findall(r"\d+(?:\.\d+){1,5}", text or ""):
            if token in by_number:
                return by_number[token]
    return None


def write_report(plan, path):
    lines = ["文件\t表格\t行\t列\t填入页码\t册内页\t依据"]
    for cell in plan["cells"]:
        lines.append("%s\t%s\t第%d行\t第%d列\t%s\t第%d页\t%s" % (
            cell["file"], cell["table"], cell["row"], cell["col"] + 1,
            cell.get("value") or format_value(cell["total_page"]),
            cell["page_in_volume"], cell["source"]))
    if plan["unresolved"]:
        lines.append("")
        lines.append("=== 未解析（未写入，需人工确认）===")
        for item in plan["unresolved"]:
            lines.append("%s\t%s\t第%s行\t第%s列\t%s\t%s" % (
                item.get("file"), item.get("table", ""), item.get("row", ""),
                (item.get("col") or 0) + 1, item.get("why"),
                " / ".join(item.get("cells") or [])[:80]))
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))
    return path


# ---------------------------------------------------------------- 一轮完整回填
def fill_round(src_folder, out_folder, pdf_of, template="第{n}页",
               color="000000", underline=False, report_path=None,
               progress=None):
    """求解 + 写出：把 `src_folder` 里所有 docx 复制到 `out_folder` 并填入页码。

    xlsx 原样复制（页码列只出现在 docx 的表格里）。返回 (plan, 写入统计)。
    """
    plan = build_plan(src_folder, pdf_of, progress)
    by_file = {}
    for cell in plan["cells"]:
        by_file.setdefault(cell["file"], []).append(cell)

    os.makedirs(out_folder, exist_ok=True)
    written = {}
    for _order, (_, name, kind) in enumerate(volume_order(src_folder), start=1):
        src = os.path.join(src_folder, name)
        dst = os.path.join(out_folder, name)
        if kind == "xlsx":
            shutil.copy2(src, dst)
            written[name] = 0
            continue
        cells = by_file.get(name) or []
        if not cells:
            shutil.copy2(src, dst)
            written[name] = 0
            continue
        written[name] = write_filled_docx(
            src, dst, cells, template=template, color=color,
            underline=underline)
    if report_path:
        write_report(plan, report_path)
    return plan, written
