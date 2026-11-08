# -*- coding: utf-8 -*-
"""Замена графика время-токовых характеристик B, C, D в готовой работе."""

import math
import os

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Cm
from PIL import Image, ImageDraw, ImageFont

DOCX = (
    r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\circuit-breaker-selection\docs"
    r"\Практическая работа — Принцип действия и установка автоматического выключателя.docx"
)
PNG = (
    r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\circuit-breaker-selection\images"
    r"\time-current-curve.png"
)
ARIAL = r"C:\Windows\Fonts\arial.ttf"
ARIAL_BOLD = r"C:\Windows\Fonts\arialbd.ttf"
INK = (20, 35, 27)
GREY = (90, 107, 97)


def fnt(size, bold=False):
    return ImageFont.truetype(ARIAL_BOLD if bold else ARIAL, size)


def ctext(d, xy, text, f, fill=INK):
    box = d.textbbox((0, 0), text, font=f)
    d.text((xy[0] - (box[2] - box[0]) / 2, xy[1] - (box[3] - box[1]) / 2 - box[1]), text, font=f, fill=fill)


def draw(path):
    W, H = 1600, 900
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    f18, f20, f22 = fnt(18), fnt(20), fnt(22)

    left, right, top, bottom = 170, 1480, 70, 740

    def px(k):
        return left + (math.log10(k) / math.log10(20)) * (right - left)

    def py(t):
        return bottom - ((math.log10(t) + 2) / 6) * (bottom - top)

    for k in (1, 2, 3, 5, 10, 20):
        d.line((px(k), bottom, px(k), bottom + 10), fill=INK, width=2)
        ctext(d, (px(k), bottom + 34), str(k), f18)
    for t, label in ((0.01, "0,01"), (0.1, "0,1"), (1, "1"), (100, "100"), (10000, "10 000")):
        d.line((left - 10, py(t), left, py(t)), fill=INK, width=2)
        d.text((20, py(t) - 12), label, font=f18, fill=INK)

    for k in (2, 3, 5, 10):
        d.line((px(k), top, px(k), bottom), fill=(232, 237, 233), width=2)
    for t in (0.1, 1, 100):
        d.line((left, py(t), right, py(t)), fill=(232, 237, 233), width=2)

    # тепловая зона
    pts = []
    k = 1.05
    while k <= 20:
        t = 60 / (k * k - 1)
        if 0.01 <= t <= 10000:
            pts.append((px(k), py(t)))
        k += 0.05
    d.line(pts, fill=(47, 107, 79), width=6, joint="curve")
    d.text((px(1.4), py(300)), "тепловая зона", font=f22, fill=(47, 107, 79))

    # зоны отсечки
    zones = [
        ("B", 3, 5, (40, 90, 170), (225, 233, 245)),
        ("C", 5, 10, (47, 107, 79), (226, 238, 230)),
        ("D", 10, 20, (185, 105, 40), (247, 235, 222)),
    ]
    for name, lo, hi, color, light in zones:
        d.rectangle((px(lo), py(0.03), px(hi), py(0.01)), fill=light, outline=color, width=3)
        for kk in (lo, hi):
            t_top = min(60 / (kk * kk - 1), 10000)
            y = py(t_top)
            while y < py(0.03):
                d.line((px(kk), y, px(kk), min(y + 14, py(0.03))), fill=color, width=3)
                y += 26
        label = {"B": "3–5", "C": "5–10", "D": "10–20"}[name]
        ctext(d, ((px(lo) + px(hi)) / 2, py(0.05)), f"{name}: {label}·Iном", f18, color)

    ctext(d, ((left + right) / 2, 840), "Отношение тока к номинальному I / Iном", f22)
    d.text((10, top - 45), "Время срабатывания, с", font=f22, fill=INK)
    d.text((left + 15, top + 10), "характеристики типов B, C и D", font=f18, fill=GREY)
    im.save(path)


def patch_docx():
    d = docx.Document(DOCX)
    target = None
    for par in d.paragraphs:
        if par._element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"):
            shapes = [sh for sh in d.inline_shapes if sh._inline.getparent() is not None]
            # берём абзац с графиком — второй по счёту рисунок
            target = par
    paragraphs = [p for p in d.paragraphs]
    drawing_pars = [p for p in paragraphs if p._element.findall(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing")]
    assert len(drawing_pars) == 2, len(drawing_pars)
    graph_par = drawing_pars[1]

    for run in list(graph_par.runs):
        if run._element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"):
            graph_par._element.remove(run._element)

    shape = graph_par.add_run().add_picture(PNG, width=Cm(9.5))
    sp = shape._inline.graphic.graphicData.pic.spPr
    ln = OxmlElement("a:ln")
    ln.set("w", "0")
    ln.append(OxmlElement("a:noFill"))
    sp.append(ln)
    graph_par.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for par in d.paragraphs:
        if par.text.strip().startswith("Рисунок 2 —"):
            for run in par.runs:
                run.text = ""
            par.runs[0].text = "Рисунок 2 — Время-токовые характеристики типов B, C и D"
        if par.text.strip().startswith("Время-токовая характеристика типа C"):
            for run in par.runs:
                run.text = ""
            par.runs[0].text = (
                "Характеристики типов B, C и D отличаются кратностью отсечки: 3–5, 5–10 и 10–20·Iном. "
                "При токах ниже отсечки работает тепловая защита, при больших токах срабатывает "
                "электромагнитный расцепитель."
            )

    out = os.environ.get("OUT_DOCX", DOCX)
    d.save(out)
    print("patched:", out)


if __name__ == "__main__":
    draw(PNG)
    patch_docx()
