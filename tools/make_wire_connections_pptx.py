# -*- coding: utf-8 -*-
"""Сборка презентации «Виды соединений жил проводов и кабелей» из подобранных иллюстраций."""

import os

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

WORK = os.path.join(os.environ["TEMP"], "eb-ppt", "trim")
OUT = (
    r"D:\Electrobaza\content\groups\1me-25\mdk-05-01\wire-connections\docs"
    r"\Презентация — Виды соединений жил проводов и кабелей.pptx"
)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

TITLE_COLOR = RGBColor(0x14, 0x23, 0x1B)
MUTED_COLOR = RGBColor(0x5A, 0x6B, 0x61)
ACCENT_COLOR = RGBColor(0x2F, 0x6B, 0x4F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "Segoe UI"

SLIDES = [
    ("01-overview.jpg", WHITE,
     "Обзор способов соединения жил",
     "Скрутка (так делать нельзя), опрессовка, винтовой клеммник, болтовое соединение, "
     "самозажимные клеммы — пять способов, которые встречаются в электромонтаже."),
    ("02-twist.png", WHITE,
     "Скрутка",
     "Скрутка многожильных и одножильных проводов: от простой до бандажной. "
     "По ПУЭ скрутка не допускается как самостоятельное соединение — только как подготовка к пайке или сварке."),
    ("03-solder.png", RGBColor(0x0E, 0x11, 0x12),
     "Пайка: физика процесса",
     "Диффузия олова в медь, этапы пайки, отличие холодной пайки от качественной. "
     "Припой ПОС-61, флюс, температура и чистота поверхности."),
    ("04-weld.png", WHITE,
     "Сварка жил",
     "Сварка угольным электродом: скрутка длиной 50 мм, дуга, аппарат 220/18–24 В. "
     "На конце скрутки образуется монолитный шарик расплавленного металла."),
    ("05-weld-photo.jpg", WHITE,
     "Сварное соединение",
     "Сварной шарик на торце скрутки: соединение не окисляется и не ослабевает. "
     "После остывания шарик зачищают, покрывают лаком и изолируют."),
    ("06-crimp-principle.jpg", WHITE,
     "Опрессовка: принцип",
     "Жила + гильза + жила = неразъёмное соединение. После обжатия место соединения "
     "закрывают термоусадочной трубкой."),
    ("07-crimp-ferrules.jpg", WHITE,
     "Опрессовка: гильзы",
     "Три гильзы после опрессовки: по два вдавливания по длине, медная и лужёная. "
     "Жила не должна проворачиваться и вытягиваться из гильзы."),
    ("08-box-sequence.jpg", WHITE,
     "Последовательность в коробке",
     "Три стадии: скрутка жил, установка гильз или зажимов, изоляция места соединения."),
    ("09-wago.jpg", WHITE,
     "Самозажимные клеммы",
     "Устройство самозажимной клеммы, монтаж в распределительной коробке и разрез: "
     "пружина прижимает жилу без винта и протяжки."),
    ("10-caps.png", WHITE,
     "Колпачковые соединители",
     "Соединитель Scotchlock: пружина с заострёнными концами, корпус, изолятор. "
     "Жилы зажимаются пружиной при завинчивании колпачка."),
    ("11-bolt-scheme.png", WHITE,
     "Болтовое соединение",
     "Порядок шайб и правильная разделка кольца: жила зажимается между шайбами, "
     "кольцо охватывает болт по направлению затяжки."),
    ("12-bolt-photo.jpg", WHITE,
     "Болтовое соединение шин",
     "Наконечники жил крепятся к медным шинам болтами с шайбами. "
     "Затяжку выполняют нормированным моментом."),
    ("13-piercing.png", WHITE,
     "Прокалывающие зажимы",
     "Ответвление от воздушной линии: фазный провод, самонесущая нейтральная жила, "
     "прокалывающий зажим и ответвление."),
]


def add_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_picture_contain(slide, path):
    with Image.open(path) as im:
        img_w, img_h = im.size
    aspect = img_w / img_h
    slide_aspect = SLIDE_W / SLIDE_H
    if aspect >= slide_aspect:
        width = SLIDE_W
        height = int(width / aspect)
        left = 0
        top = int((SLIDE_H - height) / 2)
    else:
        height = SLIDE_H
        width = int(height * aspect)
        left = int((SLIDE_W - width) / 2)
        top = 0
    slide.shapes.add_picture(path, Emu(left), Emu(top), width=Emu(width), height=Emu(height))


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def build_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, WHITE)

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(2.38), Inches(0.13), Inches(2.1))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT_COLOR
    bar.line.fill.background()
    bar.shadow.inherit = False
    style_el = bar._element.find(qn("p:style"))
    if style_el is not None:
        bar._element.remove(style_el)

    title_box = slide.shapes.add_textbox(Inches(1.4), Inches(2.18), Inches(11.0), Inches(1.7))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = "Виды соединений жил проводов и кабелей"
    run.font.name = FONT
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = TITLE_COLOR
    p.line_spacing = 1.05

    sub_box = slide.shapes.add_textbox(Inches(1.43), Inches(4.02), Inches(11.0), Inches(0.9))
    tf = sub_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = "Скрутка · пайка · сварка · опрессовка · зажимы"
    run.font.name = FONT
    run.font.size = Pt(22)
    run.font.color.rgb = MUTED_COLOR

    add_notes(
        slide,
        "Презентация к лекции «Виды соединений жил проводов и кабелей». "
        "Слайды с иллюстрациями идут в порядке разделов лекции.",
    )
    return slide


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    prs.core_properties.title = "Виды соединений жил проводов и кабелей"
    prs.core_properties.author = "Электробаза"

    build_title_slide(prs)

    for filename, bg, note_title, note_text in SLIDES:
        path = os.path.join(WORK, filename)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(slide, bg)
        add_picture_contain(slide, path)
        add_notes(slide, f"{note_title}. {note_text}")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    print("saved:", OUT)
    print("slides:", len(prs.slides))


if __name__ == "__main__":
    build()
