# -*- coding: utf-8 -*-
"""deliver_pipeline 的命令行入口。

    python run_deliver.py --src <源目录> --out <输出目录>
    python run_deliver.py --src ... --out ... --plan-only     # 只出页码计划
    python run_deliver.py --src ... --out ... --no-images --no-links
"""
import argparse
import os
import sys

import page_fill as pf


def pdf_of_from(pdf_dir):
    def resolve(name):
        path = os.path.join(pdf_dir, os.path.splitext(name)[0] + ".pdf")
        return path if os.path.exists(path) else None
    return resolve


def plan_only(src, pdf_dir, report_path):
    plan = pf.build_plan(src, pdf_of_from(pdf_dir))
    print("分册 %d 份，总页数 %d" % (len(plan["volumes"]), plan["total_pages"]))
    for vol in plan["volumes"]:
        print("  %2d %-5s 章=%-3s 页=%-5d 起始=%-6d %s" % (
            vol["order"], vol["kind"], vol.get("chapter"), vol["pages"],
            vol["offset"] + 1, vol["file"][:46]))
    print("\n可填 %d 格 / 未解析 %d 格"
          % (len(plan["cells"]), len(plan["unresolved"])))
    by_file = {}
    for cell in plan["cells"]:
        by_file[cell["file"]] = by_file.get(cell["file"], 0) + 1
    for name, count in sorted(by_file.items()):
        print("    %-58s %4d 格" % (name[:56], count))
    if plan["unresolved"]:
        print("\n未解析示例：")
        for item in plan["unresolved"][:12]:
            print("    %s %s 第%s行  %s  %s" % (
                item.get("file", "")[:34], item.get("table"),
                item.get("row"), item.get("why"),
                " / ".join(item.get("cells") or [])[:46]))
    pf.write_report(plan, report_path)
    print("\n明细已写入：%s" % report_path)
    return plan


def main():
    parser = argparse.ArgumentParser(description="投标文件整体交付流水线")
    parser.add_argument("--src", required=True, help="源目录（docx/xlsx 所在）")
    parser.add_argument("--out", required=True, help="输出目录")
    parser.add_argument("--engine", default="wps", choices=("wps", "word", "cli"))
    parser.add_argument("--plan-only", action="store_true",
                        help="只根据已有 PDF 出页码计划，不导出、不写文件")
    parser.add_argument("--pdf-dir", help="--plan-only 时使用的 PDF 目录")
    parser.add_argument("--no-images", action="store_true")
    parser.add_argument("--no-links", action="store_true")
    parser.add_argument("--no-pages", action="store_true")
    parser.add_argument("--plain", action="store_true",
                        help="页码写成纯数字，不加“第…页”")
    args = parser.parse_args()

    if args.plan_only:
        pdf_dir = args.pdf_dir or os.path.join(args.out, "分PDF")
        report = os.path.join(args.out, "_页码回填明细.txt")
        os.makedirs(args.out, exist_ok=True)
        plan_only(args.src, pdf_dir, report)
        return 0

    from deliver_pipeline import Pipeline
    pipeline = Pipeline(
        args.src, args.out, engine=args.engine,
        sink=lambda m: print(m, flush=True),
        template="plain" if args.plain else "第{n}页",
        do_images=not args.no_images,
        do_links=not args.no_links,
        do_pages=not args.no_pages)
    pipeline.run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
