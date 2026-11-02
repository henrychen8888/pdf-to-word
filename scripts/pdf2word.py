#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pdf2word.py — PDF 转 Word（双引擎 + 覆盖率择优 + 质量自检）

引擎:
  pdf2docx  —— PyMuPDF 版面分析，表格结构识别强，速度快（主通道）
  layout    —— 自研版面引擎，文字零丢失优先，表格识别较弱（兜底通道）

择优策略 (--engine auto，默认):
  先跑 pdf2docx，用「字符多重集覆盖率」核验；
  覆盖率 < 0.95 时自动改用 layout 重转，取覆盖率更高的产物。
  --engine best: 两个引擎都跑，直接取优（更慢，质量最高）。

用法:
  python pdf2word.py <pdf|目录> [...] [-o 输出目录] [--engine auto|pdf2docx|layout|best]
                     [--pages 1-10] [--report 报告.json]

退出码: 0 全部成功; 1 有文件失败; 2 环境不可用
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
import warnings
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

warnings.filterwarnings("ignore", category=FutureWarning)
for _n in ("", "pdf2docx", "fitz", "pymupdf"):
    logging.getLogger(_n).setLevel(logging.ERROR)

# 质量阈值
COV_PASS = 0.98        # 优
COV_OK = 0.95          # 可接受；低于此值触发兜底引擎
SCAN_CHARS_PER_PAGE = 30


# ---------------------------------------------------------------- 基础度量
def norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def pdf_facts(path: str, pages: tuple[int, int] | None = None) -> dict:
    """原始 PDF 事实：总页数、实际参与转换的页数、正文净字符、唯一图片数、是否加密。

    pages 为 0 基闭区间 (start, end)。指定页范围时，字符与图片只统计该范围，
    否则「只转前 5 页」会被拿去和全文比对，覆盖率被误判为个位数。
    """
    import pymupdf
    doc = pymupdf.open(path)
    try:
        if doc.is_encrypted and not doc.authenticate(""):
            return {"pages": doc.page_count, "conv_pages": 0,
                    "chars": 0, "images": 0, "encrypted": True}
        if pages:
            start = max(0, pages[0])
            end = min(pages[1] + 1, doc.page_count)
        else:
            start, end = 0, doc.page_count
        parts, xrefs = [], set()
        for i in range(start, end):
            page = doc[i]
            parts.append(page.get_text())
            for info in page.get_images(full=True):
                xrefs.add(info[0])
        txt = norm("".join(parts))
        return {"pages": doc.page_count, "conv_pages": end - start,
                "chars": len(txt), "text": txt,
                "images": len(xrefs), "encrypted": False}
    finally:
        doc.close()


def docx_facts(path: str) -> dict:
    """产物 docx 事实。

    表格文本必须按 XML 结构 (w:tr / w:tc) 遍历读取。
    不可用 id(cell._tc) 去重：lxml 元素代理对象回收后 id 会复用，
    不同单元格会被误判为同一个，字符数被严重低估（实测 60 格只读到 11 格）。
    """
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(path)
    paras = [p.text for p in doc.paragraphs]
    cells = []
    for tb in doc.tables:
        for tr in tb._tbl.tr_lst:
            for tc in tr.tc_lst:
                cells.append("".join(t.text or "" for t in tc.iter(qn("w:t"))))
    txt = norm("".join(paras) + "".join(cells))
    # 真表格：至少 2 行 2 列。pdf2docx 会把代码块/URL 退化成大量 1×1 表格，
    # 这类「表格」不是结构信息，反而降低可编辑性，不计入。
    real_tables = sum(1 for tb in doc.tables
                      if len(tb.rows) > 1 and len(tb.columns) > 1)
    imgs = 0
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
            imgs = len(re.findall(r"<a:blip\b", xml))
    except Exception:
        pass
    return {"chars": len(txt), "text": txt, "tables": len(doc.tables),
            "real_tables": real_tables,
            "images": imgs, "paras": len([p for p in paras if p.strip()])}


def coverage(src_text: str, dst_text: str) -> dict:
    """字符多重集覆盖率：不依赖段落顺序，表格重排不会造成假阴性。"""
    cs, cd = Counter(src_text), Counter(dst_text)
    if not cs:
        return {"coverage": 0.0, "missing": 0, "extra": sum(cd.values())}
    common = sum((cs & cd).values())
    return {"coverage": round(common / sum(cs.values()), 4),
            "missing": sum((cs - cd).values()),
            "extra": sum((cd - cs).values())}


# ---------------------------------------------------------------- 引擎
def has_pdf2docx() -> bool:
    try:
        import pdf2docx  # noqa: F401
        return True
    except Exception:
        return False


def has_layout() -> bool:
    try:
        import pymupdf  # noqa: F401
        from docx import Document  # noqa: F401
        return True
    except Exception:
        return False


def run_pdf2docx(pdf: str, docx: str, pages) -> None:
    from pdf2docx import Converter
    prev = logging.root.manager.disable
    logging.disable(logging.WARNING)
    cv = None
    try:
        cv = Converter(pdf)
        if pages:
            cv.convert(docx, start=pages[0], end=pages[1])
        else:
            cv.convert(docx, start=0, end=None)
    finally:
        if cv is not None:
            try: cv.close()
            except Exception: pass
        logging.disable(prev)


def run_layout(pdf: str, docx: str, pages) -> None:
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    import layout_engine
    layout_engine.convert(pdf, docx, pages)


ENGINES = {
    "pdf2docx": (run_pdf2docx, has_pdf2docx),
    "layout": (run_layout, has_layout),
}


# ---------------------------------------------------------------- 单文件
def collect_inputs(targets) -> list[str]:
    out, seen = [], set()
    for t in targets:
        t = os.path.abspath(t)
        if os.path.isdir(t):
            for root, _d, files in os.walk(t):
                for fn in sorted(files):
                    if fn.lower().endswith(".pdf"):
                        p = os.path.join(root, fn)
                        if p not in seen:
                            seen.add(p); out.append(p)
        elif os.path.isfile(t):
            if t not in seen:
                seen.add(t); out.append(t)
        else:
            print(f"[跳过] 不存在: {t}", file=sys.stderr)
    return out


def out_path_for(pdf: str, outdir: str | None) -> str:
    d = outdir or os.path.dirname(pdf)
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, os.path.splitext(os.path.basename(pdf))[0] + ".docx")


def try_engine(eng: str, pdf: str, docx: str, pages) -> dict:
    fn = ENGINES[eng][0]
    rec = {"engine": eng, "ok": False, "seconds": 0, "coverage": 0.0, "error": None}
    t0 = time.time()
    try:
        if os.path.exists(docx):
            os.remove(docx)
        fn(pdf, docx, pages)
        rec["seconds"] = round(time.time() - t0, 2)
        if not os.path.exists(docx) or os.path.getsize(docx) < 1024:
            raise RuntimeError("产物为空或过小")
        rec["ok"] = True
    except Exception as e:
        rec["seconds"] = round(time.time() - t0, 2)
        rec["error"] = f"{type(e).__name__}: {e}"
    return rec


def _is_better(a: dict, b: dict | None) -> bool:
    """判断引擎结果 a 是否优于 b。

    主指标是文字覆盖率，但覆盖率差距在 2 个百分点以内时，改比真表格结构
    （至少 2×2）：保住表格结构比多挤出 1% 的字更有实用价值。
    """
    if b is None:
        return True
    if a["coverage"] - b["coverage"] >= 0.02:
        return True
    if b["coverage"] - a["coverage"] >= 0.02:
        return False
    return a.get("real_tables", 0) > b.get("real_tables", 0)


def convert_one(pdf: str, outdir: str | None, strategy: str, pages, quiet=False) -> dict:
    out_docx = out_path_for(pdf, outdir)
    rec = {"pdf": pdf, "docx": out_docx, "status": "fail", "engine": None,
           "coverage": 0.0, "quality": None, "warnings": [], "error": None,
           "engine_runs": []}

    try:
        src = pdf_facts(pdf, pages)
    except Exception as e:
        rec["error"] = f"无法读取 PDF（可能损坏或加密）: {e}"
        return rec
    if src.get("encrypted"):
        rec["error"] = "PDF 已加密，请先解密或「打印为 PDF」后重试"
        return rec
    rec.update(pages=src["pages"], conv_pages=src["conv_pages"],
               pdf_chars=src["chars"], pdf_images=src["images"])

    # 引擎执行顺序
    if strategy == "best":
        order = [e for e in ("pdf2docx", "layout") if ENGINES[e][1]()]
    elif strategy == "auto":
        order = [e for e in ("pdf2docx", "layout") if ENGINES[e][1]()]
    else:
        order = [strategy] if ENGINES[strategy][1]() else []
    if not order:
        rec["error"] = "无可用引擎：请 pip install pdf2docx（或确认 PyMuPDF / python-docx 可用）"
        return rec

    best = None
    for i, eng in enumerate(order):
        run = try_engine(eng, pdf, out_docx, pages)
        if run["ok"]:
            dst = docx_facts(out_docx)
            cov = coverage(src["text"], dst["text"])
            run.update(coverage=cov["coverage"], missing=cov["missing"],
                       extra=cov["extra"], tables=dst["tables"],
                       real_tables=dst["real_tables"],
                       docx_chars=dst["chars"], docx_images=dst["images"],
                       docx_paras=dst["paras"])
            if _is_better(run, best):
                best = run
                rec.update(engine=eng, coverage=run["coverage"],
                           docx_chars=dst["chars"], docx_tables=dst["tables"],
                           docx_real_tables=dst["real_tables"],
                           docx_images=dst["images"], docx_paras=dst["paras"],
                           seconds=run["seconds"])
        rec["engine_runs"].append({k: run[k] for k in
                                   ("engine", "ok", "seconds", "coverage", "error")
                                   if k in run})
        if not quiet:
            tag = "OK" if run["ok"] else "FAIL"
            extra = f"覆盖率 {run.get('coverage', 0):.1%}" if run["ok"] else run["error"]
            print(f"    [{eng}] {tag} {run['seconds']}s {extra}")

        # auto：首个引擎达标即止；best：全部跑完
        if strategy == "auto" and run["ok"] and run["coverage"] >= COV_OK:
            break
        if strategy == "auto" and not run["ok"] and i == len(order) - 1:
            break

    if best is None:
        rec["error"] = "所有引擎均失败: " + "; ".join(
            f"{r['engine']}={r['error']}" for r in rec["engine_runs"])
        return rec

    # 若最优引擎不是最后写入磁盘的那个，重新生成
    if rec["engine"] != rec["engine_runs"][-1]["engine"]:
        rerun = try_engine(rec["engine"], pdf, out_docx, pages)
        if not rerun["ok"]:
            rec["warnings"].append(f"重生成最优产物失败: {rerun['error']}")

    rec["status"] = "ok"

    # ---- 质量判定 ----
    cov = rec["coverage"]
    per_page = src["chars"] / max(src["conv_pages"], 1)
    scanned = src["chars"] < SCAN_CHARS_PER_PAGE * max(src["conv_pages"], 1)
    if scanned:
        rec["quality"] = "ocr_needed"
        rec["warnings"].append(
            f"疑似扫描件/图片型 PDF（每页约 {per_page:.0f} 字符），"
            "文字层不可靠，需先 OCR 再转换")
    elif cov >= COV_PASS:
        rec["quality"] = "pass"
    elif cov >= COV_OK:
        rec["quality"] = "warn"
        rec["warnings"].append(f"文字覆盖率 {cov:.1%}，可能有少量内容丢失，建议抽查")
    else:
        rec["quality"] = "fail"
        rec["warnings"].append(f"文字覆盖率仅 {cov:.1%}，产物可能严重缺内容")

    if src["images"] > 0 and rec.get("docx_images", 0) == 0:
        rec["warnings"].append(f"原 PDF 含 {src['images']} 个图片对象，产物中未检出图片")
    if rec.get("docx_chars", 0) > src["chars"] * 1.02 and src["chars"] > 0:
        rec["warnings"].append(
            f"产物字符数明显多于原文（{rec['docx_chars']} vs {src['chars']}），可能存在重复内容")

    return rec


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description="PDF 转 Word（双引擎 + 质量自检）")
    ap.add_argument("inputs", nargs="+", help="PDF 文件或目录")
    ap.add_argument("-o", "--outdir", default=None, help="输出目录（默认与源文件同目录）")
    ap.add_argument("--engine", default="auto",
                    choices=["auto", "pdf2docx", "layout", "best"])
    ap.add_argument("--pages", default=None, help="页范围，如 1-10（1 起算，含两端）")
    ap.add_argument("--report", default=None, help="JSON 报告输出路径")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    pages = None
    if args.pages:
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", args.pages.strip())
        if not m:
            print(f"ERROR: --pages 格式应为 1-10，收到 {args.pages}", file=sys.stderr)
            return 2
        pages = (max(int(m.group(1)) - 1, 0), int(m.group(2)) - 1)

    files = collect_inputs(args.inputs)
    if not files:
        print("ERROR: 未找到 PDF 文件", file=sys.stderr)
        return 2

    avail = [e for e in ENGINES if ENGINES[e][1]()]
    if not avail:
        print("ERROR: 无可用引擎。请执行: pip install pdf2docx PyMuPDF python-docx",
              file=sys.stderr)
        return 2

    print(f"引擎: {args.engine} (可用: {', '.join(avail)}) | 文件: {len(files)}")
    results = []
    for i, pdf in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {os.path.basename(pdf)}")
        rec = convert_one(pdf, args.outdir, args.engine, pages, args.quiet)
        results.append(rec)
        if rec["status"] == "ok":
            mark = {"pass": "PASS", "warn": "WARN",
                    "fail": "FAIL", "ocr_needed": "OCR"}.get(rec["quality"], "?")
            cp, tp = rec.get("conv_pages", rec["pages"]), rec["pages"]
            pg = f"{cp}/{tp}页" if cp != tp else f"{tp}页"
            print(f"    => {mark} [{rec['engine']}] {rec['seconds']}s "
                  f"覆盖率 {rec['coverage']:.1%} | {pg} "
                  f"表{rec.get('docx_real_tables',0)} 图{rec.get('docx_images',0)}")
            print(f"       {rec['docx']}")
            for w in rec["warnings"]:
                print(f"       ! {w}")
        else:
            print(f"    => FAIL {rec['error']}")

    ok = [r for r in results if r["status"] == "ok"]
    bad = [r for r in results if r["status"] != "ok"]
    print(f"\n完成: {len(ok)} 成功 / {len(bad)} 失败")
    for r in bad:
        print(f"  失败 {os.path.basename(r['pdf'])}: {r['error']}")

    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"报告: {args.report}")

    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
