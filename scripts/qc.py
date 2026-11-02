#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
qc.py — PDF→DOCX 转换质量核验（不依赖文字顺序）

指标说明：
  coverage   = 多重集覆盖率 = 产物中出现、且数量不超过原文的字符数 / 原文字符数
               （不依赖段落顺序，表格重排不会造成假阴性）
  extra      = 产物中原文没有的字符数（幻觉 / 异常插入）
  tables/images/paras  结构计数

用法:
  python qc.py <原PDF> <产物docx> [--json]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def pdf_text(path: str) -> tuple[str, int, int]:
    import pymupdf
    doc = pymupdf.open(path)
    try:
        parts = []
        xrefs = set()
        for p in doc:
            parts.append(p.get_text())
            for info in p.get_images(full=True):
                xrefs.add(info[0])
        return norm("".join(parts)), doc.page_count, len(xrefs)
    finally:
        doc.close()


def docx_text(path: str) -> tuple[str, int, int]:
    """提取产物全文（段落 + 表格）。

    表格按 XML 结构 (w:tr / w:tc) 遍历取文本。不可用 id(cell._tc) 去重：
    lxml 元素代理对象被回收后 id 会复用，会把不同单元格误判为同一个，
    造成字符数被严重低估（实测 60 格只统计到 11 格）。
    """
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for tb in doc.tables:
        for tr in tb._tbl.tr_lst:
            for tc in tr.tc_lst:
                parts.append("".join(t.text or "" for t in tc.iter(qn("w:t"))))
    imgs = 0
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
            imgs = len(re.findall(r"<a:blip\b", xml))
    except Exception:
        pass
    return norm("".join(parts)), len(doc.tables), imgs


def check(pdf: str, docx: str) -> dict:
    src, pages, src_imgs = pdf_text(pdf)
    dst, tables, dst_imgs = docx_text(docx)
    cs, cd = Counter(src), Counter(dst)
    common = sum((cs & cd).values())
    missing = cs - cd
    extra = cd - cs
    coverage = common / len(src) if src else 0.0
    return {
        "pdf": pdf, "docx": docx,
        "pages": pages,
        "pdf_chars": len(src), "docx_chars": len(dst),
        "coverage": round(coverage, 4),
        "missing_chars": sum(missing.values()),
        "extra_chars": sum(extra.values()),
        "missing_top": missing.most_common(8),
        "extra_top": extra.most_common(8),
        "pdf_images": src_imgs, "docx_images": dst_imgs,
        "tables": tables,
        "verdict": ("PASS" if coverage >= 0.95 else
                    "WARN" if coverage >= 0.80 else "FAIL"),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("docx")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    r = check(a.pdf, a.docx)
    if a.json:
        print(json.dumps(r, ensure_ascii=False, indent=2))
    else:
        print(f"覆盖率 {r['coverage']:.1%}  [{r['verdict']}]")
        print(f"  原文 {r['pdf_chars']} 字符 / 产物 {r['docx_chars']} 字符 "
              f"(缺 {r['missing_chars']} / 多 {r['extra_chars']})")
        print(f"  页数 {r['pages']} | 表格 {r['tables']} | 图片 {r['pdf_images']}->{r['docx_images']}")
        if r["missing_top"]:
            print("  缺失最多字符:", r["missing_top"][:6])
        if r["extra_chars"]:
            print("  多余字符:", r["extra_top"][:6])
    return 0 if r["verdict"] != "FAIL" else 1


if __name__ == "__main__":
    sys.exit(main())
