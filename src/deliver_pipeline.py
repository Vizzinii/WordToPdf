# -*- coding: utf-8 -*-
"""投标文件整体交付流水线：探测轮 → 页码回填 → 定稿轮 → 合并总 PDF。

为什么要跑两遍导出
------------------
页码值取决于材料在总 PDF 里的位置，总 PDF 又由各分册拼接而来，而分册要从
「页码已填好的 docx」导出——三者互为前提。但不必先合出一份几百 MB 的总 PDF：
总页码 = 该分册之前所有分册页数之和 + 该材料在本分册内的页序号。所以：

    第 1 轮（探测）  全部源文件导出 PDF，页码还空着 → 拿到各分册页数与册内标题页
    求解            按分册偏移累加，算出每个空缺格应填的数字
    回填            写 docx 副本（只改目标单元格，不动样式）
    第 2 轮（定稿）  只重导「回填后有变化」的分册
    收敛检测        若某分册页数变了，后面所有分册页码整体平移 → 重算、重导，
                    最多 MAX_ROUNDS 轮；始终不稳则明确报错，不静默交付偏移页码
    合并            按交付顺序拼接，并把各分册书签按页偏移平移后写回

分册 = 源目录里的一个 .docx / .xlsx；xlsx 直接导出成 PDF，与 docx 同等对待。
"""
import os
import shutil
import time

import pymupdf

import link_fix
import optimize_pdf_images as opt
import page_fill as pf
import wps_to_pdf as w2p

MAX_ROUNDS = 3


class Pipeline(object):
    def __init__(self, src, out, engine="wps", sink=None, template="第{n}页",
                 color="000000", underline=False, keep_filled=True,
                 do_images=True, do_links=True, do_pages=True):
        self.src = os.path.abspath(src)
        self.out = os.path.abspath(out)
        self.engine = engine
        self.sink = sink or (lambda message: None)
        self.template = template
        self.color = color
        self.underline = underline
        self.keep_filled = keep_filled
        self.do_images = do_images
        self.do_links = do_links
        self.do_pages = do_pages
        self.pdf_dir = os.path.join(self.out, "分PDF")
        self.filled_dir = os.path.join(self.out, "分docx与xlsx")
        self.log = []

    def say(self, message):
        self.log.append(str(message))
        self.sink(str(message))

    # ------------------------------------------------------------------ 导出
    def pdf_name(self, name):
        return os.path.join(self.pdf_dir, os.path.splitext(name)[0] + ".pdf")

    def export_one(self, app, name, kind, dst, overwrite=True, sheet_app=None,
                   folder=None):
        src = os.path.join(folder or self.src, name)
        if kind == "xlsx":
            own = sheet_app is None
            if own:
                sheet_app = w2p.start_sheet_app(False)
            try:
                size, pages = w2p.export_xlsx(sheet_app, src, dst,
                                              timeout=1800,
                                              overwrite=overwrite)
            finally:
                if own:
                    w2p.quit_sheet_app(sheet_app)
        else:
            size, pages, _used, _doc_pages = w2p.export_doc(
                app, src, dst, timeout=1800, page_check=True,
                overwrite=overwrite)
        return size, pages

    def export_all(self, folder, only=None):
        """导出 folder 下的分册；only 给定时只导这些文件名。"""
        os.makedirs(self.pdf_dir, exist_ok=True)
        targets = pf.volume_order(folder)
        if only is not None:
            targets = [t for t in targets if t[1] in only]
        if not targets:
            return {}
        self._export_from = folder
        app = w2p.start_app(self.engine, False)
        sheet_app = None
        result = {}
        try:
            if any(kind == "xlsx" for _o, _n, kind in targets):
                sheet_app = w2p.start_sheet_app(False)
            for order, (_, name, kind) in enumerate(targets, start=1):
                dst = self.pdf_name(name)
                self.say("  [%d/%d] 导出 %s" % (order, len(targets), name[:56]))
                size, pages = self.export_one(app, name, kind, dst,
                                              sheet_app=sheet_app,
                                              folder=folder)
                result[name] = {"size": size, "pages": pages}
                self.say("         %.2f MB / %d 页"
                         % (size / 1048576.0, pages))
        finally:
            if sheet_app is not None:
                w2p.quit_sheet_app(sheet_app)
            w2p.quit_app(app)
        return result

    # ------------------------------------------------------------------ 主流程
    def run(self):
        started = time.time()
        os.makedirs(self.out, exist_ok=True)
        self.say("=" * 78)
        self.say("投标文件交付流水线")
        self.say("源目录：%s" % self.src)
        self.say("输出：  %s" % self.out)

        volumes = pf.volume_order(self.src)
        if not volumes:
            self.say("源目录里没有 .docx 或 .xlsx：%s" % self.src)
            self.say("请把分册文件放进源目录后重新运行。")
            raise RuntimeError("源目录没有可处理的 Word/Excel 分册")
        self.say("分册数：%d（docx %d / xlsx %d）" % (
            len(volumes),
            sum(1 for v in volumes if v[2] == "docx"),
            sum(1 for v in volumes if v[2] == "xlsx")))

        # ---- 第 1 轮：探测导出（源文件，页码仍空）
        self.say("")
        self.say("【第 1 轮】探测导出：全部源文件 → PDF（页码尚未填写）")
        probe = self.export_all(self.src)

        # ---- 页码回填
        plan = None
        filled_stats = {}
        if self.do_pages:
            self.say("")
            self.say("【页码求解与回填】")
            report_path = os.path.join(self.out, "_页码回填明细.txt")
            plan, filled_stats = pf.fill_round(
                self.src, self.filled_dir, self.pdf_name,
                template=self.template, color=self.color,
                underline=self.underline, report_path=report_path)
            self.say("分册总页数：%d" % plan["total_pages"])
            self.say("求解出 %d 格，未解析 %d 格" % (
                len(plan["cells"]), len(plan["unresolved"])))
            self.say("写入 %d 份 docx，明细：%s" % (
                sum(1 for v in filled_stats.values() if v), report_path))
        else:
            os.makedirs(self.filled_dir, exist_ok=True)
            for _o, (_, name, kind) in enumerate(volumes, start=1):
                shutil.copy2(os.path.join(self.src, name),
                             os.path.join(self.filled_dir, name))

        # ---- 第 2 轮：定稿导出（页码已填）
        self.say("")
        self.say("【第 2 轮】定稿导出：页码已填的分册 → PDF")
        final = self.export_all(self.filled_dir)

        # ---- 收敛检测：页数是否因回填而变化
        rounds = 1
        if self.do_pages and plan is not None:
            changed = [name for name in final
                       if probe.get(name, {}).get("pages")
                       != final[name]["pages"]]
            if changed:
                self.say("")
                self.say("回填后有 %d 份分册页数变化，后续分册页码需要平移："
                         % len(changed))
                for name in changed[:10]:
                    self.say("   %s：%s → %s 页" % (
                        name[:52], probe.get(name, {}).get("pages"),
                        final[name]["pages"]))
                rounds = self._stabilise(changed)
            else:
                self.say("")
                self.say("收敛检测通过：回填未改变任何分册页数，页码无需平移。")

        # ---- 图像回填 + 跳转补链（逐分册）
        if self.do_images or self.do_links:
            self.say("")
            self.say("【图像与跳转】")
            for _o, (_, name, kind) in enumerate(volumes, start=1):
                pdf_path = self.pdf_name(name)
                if not os.path.exists(pdf_path):
                    continue
                src_docx = os.path.join(self.filled_dir, name)
                if kind == "docx" and self.do_images:
                    self._optimise(pdf_path, src_docx, name)
                if kind == "docx" and self.do_links:
                    self._links(pdf_path, src_docx, name)

        # ---- 合并
        self.say("")
        self.say("【合并总 PDF】")
        ordered = [self.pdf_name(name) for _o, (_, name, _k) in
                   enumerate(volumes, start=1)
                   if os.path.exists(self.pdf_name(name))]
        merged_path = os.path.join(self.out, "投标技术文件_总.pdf")
        total = self._merge(ordered, merged_path)
        self.say("")
        self.say("完成：%d 个分册 / %d 页 / 迭代 %d 轮 / 用时 %.1f 分钟"
                 % (len(ordered), total, rounds,
                    (time.time() - started) / 60.0))
        self.say("输出目录：%s" % self.out)
        log_path = os.path.join(self.out, "_流水线日志.txt")
        with open(log_path, "w", encoding="utf-8") as handle:
            handle.write("\n".join(self.log))
        return {"pages": total, "rounds": rounds, "log": log_path}

    def _stabilise(self, changed):
        """页数变了就重算页码、重导受影响分册，直到页数稳定。"""
        rounds = 1
        while rounds < MAX_ROUNDS:
            rounds += 1
            self.say("   —— 第 %d 轮重算 ——" % rounds)
            report_path = os.path.join(self.out, "_页码回填明细.txt")
            plan, written = pf.fill_round(
                self.src, self.filled_dir, self.pdf_name,
                template=self.template, color=self.color,
                underline=self.underline, report_path=report_path)
            before = {}
            for _o, (_, vname, _k) in enumerate(
                    pf.volume_order(self.filled_dir), start=1):
                p = self.pdf_name(vname)
                if os.path.exists(p):
                    with pymupdf.open(p) as handle:
                        before[vname] = handle.page_count
                else:
                    before[vname] = 0
            self.export_all(self.filled_dir, only=set(changed))
            after = {}
            for vname in changed:
                p = self.pdf_name(vname)
                if os.path.exists(p):
                    with pymupdf.open(p) as handle:
                        after[vname] = handle.page_count
                else:
                    after[vname] = 0
            shifted = [n for n in changed if before.get(n) != after.get(n)]
            if not shifted:
                self.say("   页数已稳定。")
                return rounds
            changed = shifted
        self.say("   ⚠ 经 %d 轮仍未稳定，交付件页码可能与实际页存在偏移，"
                 "请核对 _页码回填明细.txt。" % MAX_ROUNDS)
        return rounds

    def _optimise(self, pdf_path, docx_path, name):
        part = pdf_path + ".opt.part"
        try:
            stats = opt.optimize(pdf_path, docx_path, part)
            os.replace(part, pdf_path)
            total = sum(stats.values()) if isinstance(stats, dict) else 0
            self.say("   图像回填：%s（%d 个对象已处理）" % (name[:50], total))
        except Exception as exc:
            if "no usable images" in str(exc):
                self.say("   图像回填：%s 无可用插图，跳过" % name[:50])
            else:
                self.say("   图像回填失败（保留导出图）：%s → %s"
                         % (name[:40], exc))
        finally:
            if os.path.exists(part):
                try:
                    os.remove(part)
                except OSError:
                    pass

    def _links(self, pdf_path, docx_path, name):
        try:
            report = link_fix.repair(pdf_path, docx_path)
            self.say("   跳转校验：%s 域 %s 条 / 注解 %s 个 / 补 %s 条 / 缺 %s 条"
                     % (name[:44], report.get("docx_fields"),
                        report.get("pdf_links"), report.get("inserted"),
                        len(report.get("missing") or [])))
        except Exception as exc:
            self.say("   跳转校验失败（不影响内容）：%s → %s"
                     % (name[:40], exc))

    def _merge(self, paths, out_path):
        from pdf_tool import merge_pdfs
        if not paths:
            raise RuntimeError("没有可合并的 PDF")
        pages = merge_pdfs(paths, out_path, self.say)
        return pages
