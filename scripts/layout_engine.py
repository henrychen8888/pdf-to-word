# -*- coding: utf-8 -*-
"""
layout_engine.py — 基于 PyMuPDF 版面自研的 PDF→DOCX 通道

设计目标：**文字零丢失**优先于像素级版式还原。
适用于 PPT 导出型 PDF（整页背景图 + 叠加文字）、多重文本框、特殊中文字体
等 pdf2docx 容易丢字的场景。

还原内容：
  - 文字（保留字号 / 颜色 / 加粗 / 斜体，段落按行合并）
  - 表格（PyMuPDF find_tables 结构识别）
  - 图片（原图提取，按版面宽度缩放）
  - 页面尺寸与源 PDF 一致
"""
from __future__ import annotations

import io
import re

import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# 与表格区域重叠超过该比例的文字块视为表格内容，不再重复输出
TABLE_OVERLAP = 0.45
# 覆盖页面面积超过该比例的图视为背景图
BG_RATIO = 0.55


def _overlap(a, b) -> float:
    """a、b 两个 (x0,y0,x1,y1) 的交集面积占 a 面积的比例。"""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    area = max((ax1 - ax0) * (ay1 - ay0), 1e-6)
    return inter / area


def _rgb(color_int: int) -> RGBColor:
    v = int(color_int or 0)
    return RGBColor((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF)


def _set_cjk_font(run, font_name: str):
    """同时设置西文与东亚字体，避免中文回退成宋体。"""
    name = font_name or "Microsoft YaHei"
    # 去掉 Bold/Italic 等 PS 后缀
    for suf in ("-Bold", "-Italic", "-BoldItalic", "Bold", "Italic", ",Bold", ",Italic"):
        name = name.replace(suf, "")
    name = name.strip() or "Microsoft YaHei"
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)


_ASCII_END = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789,.;:!?)]}")
_ASCII_START = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789([{")


def _join_lines(lines: list[str]) -> str:
    """把同一文本块的多行合并成段落：中文直接相连，西文之间补空格。"""
    out = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        if not out:
            out = ln
            continue
        prev, nxt = out[-1], ln[0]
        if prev in _ASCII_END and nxt in _ASCII_START:
            out += " " + ln
        elif prev == "-":                     # 连字符断词
            out += ln
        else:
            out += ln
    return out


def _chart_name(el):
    from lxml import etree
    return etree.QName(el).localname


def _make_anchor(inline_el, behind: bool, off_x_emu: int, off_y_emu: int):
    """把 <wp:inline> 改造成 <wp:anchor>，使其脱离流式排版（可衬于文字下方）。

    Word 对 anchor 子元素顺序有强校验，必须是：
    simplePos, positionH, positionV, extent, effectExtent, wrapNone, docPr,
    cNvGraphicFramePr, graphic
    """
    import copy
    from lxml import etree
    from docx.oxml import OxmlElement

    anchor = OxmlElement("wp:anchor")
    for k, v in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0"),
                 ("simplePos", "0"), ("relativeHeight", "1"),
                 ("behindDoc", "1" if behind else "0"), ("locked", "0"),
                 ("layoutInCell", "1"), ("allowOverlap", "1")):
        anchor.set(k, v)

    sp = OxmlElement("wp:simplePos")
    sp.set("x", "0"); sp.set("y", "0")
    anchor.append(sp)

    for tag, val in (("wp:positionH", off_x_emu), ("wp:positionV", off_y_emu)):
        pos = OxmlElement(tag)
        pos.set("relativeFrom", "page")
        off = OxmlElement("wp:posOffset")
        off.text = str(int(val))
        pos.append(off)
        anchor.append(pos)

    kids = {_chart_name(c): c for c in inline_el}
    for name in ("extent", "effectExtent"):
        if name in kids:
            anchor.append(copy.deepcopy(kids[name]))
    anchor.append(OxmlElement("wp:wrapNone"))
    for name in ("docPr", "cNvGraphicFramePr", "graphic"):
        if name in kids:
            anchor.append(copy.deepcopy(kids[name]))
    return anchor


class _PageWriter:
    def __init__(self, doc, page):
        self.doc = doc
        self.page = page
        self.page_rect = page.rect
        self.content_w = self.page_rect.width - 72  # 左右各 36pt 页边距
        self._pic_id = 1000

    # ---------- 文字 ----------
    def add_text_block(self, block):
        lines, style = [], None
        for line in block.get("lines", []):
            texts = []
            for span in line.get("spans", []):
                texts.append(span.get("text", ""))
                if style is None and span.get("text", "").strip():
                    style = span
            if texts:
                lines.append("".join(texts))
        text = _join_lines(lines)
        if not text.strip():
            return
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        # 首行缩进按左侧边距比例
        try:
            indent = max(block["bbox"][0] - self.page_rect.x0 - 36, 0)
            if indent > 6:
                p.paragraph_format.left_indent = Pt(min(indent, 200))
        except Exception:
            pass

        # 逐 span 写入以保留同一块内的多样式
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                t = span.get("text", "")
                if not t:
                    continue
                run = p.add_run(t)
                sz = span.get("size") or 10.5
                run.font.size = Pt(round(sz, 1))
                flags = span.get("flags", 0)
                run.font.bold = bool(flags & 16)
                run.font.italic = bool(flags & 2)
                run.font.color.rgb = _rgb(span.get("color", 0))
                _set_cjk_font(run, span.get("font", ""))

    # ---------- 表格 ----------
    def add_table(self, tab):
        data = tab.extract()
        if not data:
            return
        rows = len(data)
        cols = max(len(r) for r in data) if data else 0
        if rows == 0 or cols == 0:
            return
        t = self.doc.add_table(rows=rows, cols=cols)
        try:
            t.style = "Table Grid"
        except Exception:
            pass
        for ri, row in enumerate(data):
            for ci in range(cols):
                cell = t.cell(ri, ci)
                val = row[ci] if ci < len(row) else ""
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run((val or "").replace("\n", " "))
                run.font.size = Pt(9)
                _set_cjk_font(run, "")
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ---------- 图片 ----------
    def add_image(self, xref, bbox, is_bg: bool):
        """全部图片按原始坐标浮动定位。

        背景图衬于文字下方 (behindDoc=1)，前景图浮于文字上方。
        若改用内嵌 (inline) 插入，每张图都会撑开流式版面，
        实测 30 页文档会膨胀到 68~114 页。
        """
        try:
            info = self.doc_part_image(xref)
            if not info:
                return
            img_bytes, _ext = info
            w_pt = min(bbox[2] - bbox[0], self.page_rect.width)
            h_pt = bbox[3] - bbox[1]
            if w_pt < 8 or h_pt < 8:
                return
            stream = io.BytesIO(img_bytes)

            p = self.doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = Pt(1)
            run = p.add_run()
            run.font.size = Pt(1)
            run.add_picture(stream, width=Pt(w_pt), height=Pt(h_pt))
            inline = run._element.find(".//" + qn("wp:inline"))
            if inline is None:
                return
            anchor = _make_anchor(
                inline,
                behind=bool(is_bg),
                off_x_emu=int((bbox[0] - self.page_rect.x0) * 12700),
                off_y_emu=int((bbox[1] - self.page_rect.y0) * 12700),
            )
            drawing = run._element.find(".//" + qn("w:drawing"))
            if drawing is not None:
                drawing.replace(inline, anchor)
        except Exception:
            pass

    def doc_part_image(self, xref):
        try:
            img = self.doc_src.extract_image(xref)
            return img["image"], img.get("ext", "png")
        except Exception:
            return None


def convert(pdf: str, docx: str, pages: tuple[int, int] | None = None,
            progress=None) -> None:
    src = pymupdf.open(pdf)
    try:
        if src.is_encrypted and not src.authenticate(""):
            raise RuntimeError("PDF 已加密")
        start = pages[0] if pages else 0
        end = min(pages[1] + 1, src.page_count) if pages else src.page_count
        start = max(0, start)

        out = Document()
        first = src[start].rect
        sec = out.sections[0]
        sec.page_width = Pt(first.width)
        sec.page_height = Pt(first.height)
        sec.left_margin = sec.right_margin = Pt(36)
        sec.top_margin = sec.bottom_margin = Pt(36)

        for idx, pno in enumerate(range(start, end)):
            page = src[pno]
            if idx > 0:
                out.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

            writer = _PageWriter(out, page)
            writer.doc_src = src

            # 表格区域
            tables = []
            try:
                for tab in page.find_tables().tables:
                    tables.append(tab)
            except Exception:
                pass

            d = page.get_text("dict")
            page_area = (page.rect.x1 - page.rect.x0) * (page.rect.y1 - page.rect.y0)

            def _block_chars(b):
                if b.get("type") != 0:
                    return 0
                txt = "".join(s.get("text", "") for l in b.get("lines", [])
                              for s in l.get("spans", []))
                return len(re.sub(r"\s+", "", txt))

            # 表格可靠性评估：抽出来的字 >= 区域内文字 90% 才认定可用，
            # 否则放弃表格结构、保留文字块，避免「跳过文字 + 表格没抽全」双丢失。
            usable_tables = []
            for tab in tables:
                try:
                    data = tab.extract() or []
                except Exception:
                    continue
                tab_chars = sum(len(re.sub(r"\s+", "", c or "")) for row in data for c in row)
                region_chars = sum(_block_chars(b) for b in d.get("blocks", [])
                                   if b.get("type") == 0
                                   and _overlap(tuple(b["bbox"]), tuple(tab.bbox)) > TABLE_OVERLAP)
                if region_chars == 0 or tab_chars >= region_chars * 0.9:
                    usable_tables.append(tab)
            tboxes = [tuple(t.bbox) for t in usable_tables]
            tables = usable_tables

            # ---- 按阅读顺序组装 ----
            # PyMuPDF 的 blocks 顺序即正确阅读顺序（分栏时按栏优先），
            # 因此保持原序，仅把表格/图片按纵向位置插入到相应块之前。
            text_blocks = []      # 顺序输出序列
            img_blocks = []       # (y, xref, bbox, is_bg)
            for b in d.get("blocks", []):
                bbox = tuple(b["bbox"])
                if b.get("type") == 1:
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    is_bg = area / max(page_area, 1) >= BG_RATIO
                    xref = b.get("xref") or _find_xref(page, bbox)
                    if xref:
                        img_blocks.append((bbox[1], xref, bbox, is_bg))
                else:
                    if any(_overlap(bbox, tb) > TABLE_OVERLAP for tb in tboxes):
                        continue          # 表格区域内的文字已由表格承载
                    text_blocks.append(b)

            # 每个表格/非背景图找插入位置：输出序列中首个 y0 大于其 y0 的位置
            def _insert_pos(y0):
                for i, tb_ in enumerate(text_blocks):
                    if tb_["bbox"][1] > y0:
                        return i
                return len(text_blocks)

            marks = []   # (pos, sub, kind, payload)
            for tab in tables:
                marks.append((_insert_pos(tab.bbox[1]), 0, "table", tab))
            fg_imgs = [t for t in img_blocks if not t[3]]
            bg_imgs = [t for t in img_blocks if t[3]]
            for y0, xref, bbox, is_bg in fg_imgs:
                marks.append((_insert_pos(y0), 1, "img", (xref, bbox, is_bg)))
            marks.sort(key=lambda m: (m[0], m[1]))

            # 背景图统一先输出（视觉上位于文字之后 / 底层）
            for _y, xref, bbox, is_bg in bg_imgs:
                writer.add_image(xref, bbox, True)

            mi = 0
            for i, b in enumerate(text_blocks):
                while mi < len(marks) and marks[mi][0] <= i:
                    _, _, kind, payload = marks[mi]
                    if kind == "table":
                        writer.add_table(payload)
                    else:
                        xref, bbox, is_bg = payload
                        writer.add_image(xref, bbox, is_bg)
                    mi += 1
                writer.add_text_block(b)
            while mi < len(marks):
                _, _, kind, payload = marks[mi]
                if kind == "table":
                    writer.add_table(payload)
                else:
                    xref, bbox, is_bg = payload
                    writer.add_image(xref, bbox, is_bg)
                mi += 1

            if progress:
                progress(idx + 1, end - start)

        out.save(docx)
    finally:
        src.close()


def _find_xref(page, bbox):
    """文本块无 xref 时，按 bbox 反查图片 xref。"""
    try:
        for info in page.get_image_info(xrefs=True):
            ib = tuple(info["bbox"])
            if _overlap(bbox, ib) > 0.5:
                return info.get("xref")
    except Exception:
        pass
    return None
