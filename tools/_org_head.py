# -*- coding: utf-8 -*-
"""Лекция «Организация труда при выполнении электромонтажных работ» — учебный вариант."""

import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1eo-25\mdk-04-01\3-organizaciya-truda"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Организация труда при выполнении электромонтажных работ.docx")
FONT = "Times New Roman"
ACCENT = "3F7D5F"
NOTE_FILL = "EEF4EF"


def _rfonts(rpr):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)


def style_font(style, size=None, bold=None, color=None):
    rpr = style.element.get_or_add_rPr()
    _rfonts(rpr)
    style.font.name = FONT
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
        col = rpr.find(qn("w:color"))
        if col is not None:
            for a in ("w:themeColor", "w:themeShade", "w:themeTint"):
                if col.get(qn(a)) is not None:
                    del col.attrib[qn(a)]


def rich(p, text, size=None, bold=None, italic=None):
    for tok in re.split(r"(\*\*[^*]+\*\*|~[^~]+~|\^[^\^]+\^|\*[^*]+\*)", text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            r = p.add_run(tok[2:-2]); _rfonts(r._element.get_or_add_rPr()); r.bold = True
        elif tok.startswith("~") and tok.endswith("~") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.font.subscript = True
        elif tok.startswith("^") and tok.endswith("^") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.font.superscript = True
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.italic = True
        else:
            r = p.add_run(tok); _rfonts(r._element.get_or_add_rPr())
        r.font.name = FONT
        if size is not None:
            r.font.size = Pt(size)
        if bold is not None:
            r.bold = bold
        if italic is not None:
            r.italic = italic


def para(doc, text, indent=0.8, after=2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.05
    rich(p, text)


def heading(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    rich(p, text)


def _note_pPr(p):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), NOTE_FILL)
    pPr.append(shd)
    bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), ACCENT)
    bdr.append(left)
    pPr.append(bdr)


def note(doc, title, items):
    first = doc.add_paragraph()
    first.paragraph_format.left_indent = Cm(0.5)
    first.paragraph_format.right_indent = Cm(0.2)
    first.paragraph_format.space_before = Pt(4)
    first.paragraph_format.space_after = Pt(1)
    first.paragraph_format.line_spacing = 1.05
    first.paragraph_format.keep_with_next = True
    first.paragraph_format.keep_together = True
    rich(first, "**" + title + "**", size=10)
    _note_pPr(first)

    for i, item in enumerate(items):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.85)
        p.paragraph_format.right_indent = Cm(0.2)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        p.paragraph_format.space_after = Pt(4 if i == len(items) - 1 else 1)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.keep_together = True
        if i < len(items) - 1:
            p.paragraph_format.keep_with_next = True
        rich(p, "— " + item, size=10)
        _note_pPr(p)


def mark_border(p):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "8")
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), "8FBFA3")
    bdr.append(left)
    pPr.insert_element_before(
        bdr,
        "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
        "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
        "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
        "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
        "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
        "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
    )


def figure(doc, name, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    sh = p.add_run().add_picture(os.path.join(IMG, name), width=Cm(width_cm))
    sp = sh._inline.graphic.graphicData.pic.spPr
    ln = OxmlElement("a:ln"); ln.set("w", "0"); ln.append(OxmlElement("a:noFill")); sp.append(ln)
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_before = Pt(1)
    c.paragraph_format.space_after = Pt(4)
    rich(c, caption, size=9, italic=True)


def shade(cell, color="E8F0EC"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def table(doc, caption, header, rows, widths):
    c = doc.add_paragraph()
    c.paragraph_format.space_before = Pt(4)
    c.paragraph_format.space_after = Pt(2)
    c.paragraph_format.keep_with_next = True
    rich(c, caption, size=9, italic=True)

    t = doc.add_table(rows=1, cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for i, h in enumerate(header):
        cell = t.rows[0].cells[i]
        cell.width = widths[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        rich(p, h, size=9, bold=True)
        shade(cell)

    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].width = widths[i]
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY
            rich(p, value, size=9)

    after = doc.add_paragraph()
    after.paragraph_format.space_after = Pt(0)
    after.paragraph_format.line_spacing = 1.0
    rich(after, "", size=6)


