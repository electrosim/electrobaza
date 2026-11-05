# -*- coding: utf-8 -*-
"""Три лекции для 1ЭО-24, ОП.08: рисунки, Word-версии и HTML-фрагменты."""

import os
import re

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = r"D:\Electrobaza"
BASE = os.path.join(ROOT, "content", "groups", "1eo-24", "op-08")
FONT = "Times New Roman"

GREEN = (47, 107, 79)
GREEN_SOFT = (238, 244, 239)
INK = (20, 35, 27)
GREY = (90, 107, 97)

ARIAL = r"C:\Windows\Fonts\arial.ttf"
ARIAL_BOLD = r"C:\Windows\Fonts\arialbd.ttf"


# --------------------------------------------------------------------------- рисунки

def font(size, bold=False):
    return ImageFont.truetype(ARIAL_BOLD if bold else ARIAL, size)


def center_text(draw, xy, text, fnt, fill=INK):
    box = draw.textbbox((0, 0), text, font=fnt)
    w, h = box[2] - box[0], box[3] - box[1]
    draw.text((xy[0] - w / 2, xy[1] - h / 2 - box[1]), text, font=fnt, fill=fill)


def rounded_box(draw, x1, y1, x2, y2, lines, fnt, fill=GREEN_SOFT, outline=GREEN, radius=18, text_fill=INK):
    draw.rounded_rectangle((x1, y1, x2, y2), radius=radius, fill=fill, outline=outline, width=3)
    cy = (y1 + y2) / 2
    step = 34 if len(lines) > 1 else 0
    start = cy - step * (len(lines) - 1) / 2
    for i, line in enumerate(lines):
        center_text(draw, ((x1 + x2) / 2, start + i * step), line, fnt, text_fill)


def arrow(draw, x1, y1, x2, y2, fill=GREEN, width=4):
    draw.line((x1, y1, x2, y2), fill=fill, width=width)
    dx, dy = x2 - x1, y2 - y1
    length = max((dx * dx + dy * dy) ** 0.5, 1)
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    size = 16
    draw.polygon(
        [(x2, y2), (x2 - ux * size + px * size * 0.55, y2 - uy * size + py * size * 0.55),
         (x2 - ux * size - px * size * 0.55, y2 - uy * size - py * size * 0.55)],
        fill=fill,
    )


def draw_l1(path):
    W, H = 1600, 900
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    f25, f22, f19, f28 = font(25), font(22), font(19), font(28, True)

    d.text((60, 40), "Преобразование сведений", font=f28, fill=INK)
    boxes = [(70, 110, 470, 300, ["ДАННЫЕ", "факты, записи,", "показания приборов"]),
             (620, 110, 1020, 300, ["ИНФОРМАЦИЯ", "понятый смысл,", "снижение неопределённости"]),
             (1170, 110, 1570, 300, ["ЗНАНИЯ", "проверено практикой,", "основа решений"])]
    for x1, y1, x2, y2, lines in boxes:
        rounded_box(d, x1, y1, x2, y2, lines, f22)
    arrow(d, 480, 205, 610, 205)
    arrow(d, 1030, 205, 1160, 205)

    d.text((60, 380), "Этапы информационной технологии", font=f28, fill=INK)
    stages = ["СБОР", "ОБРАБОТКА", "ХРАНЕНИЕ", "ПЕРЕДАЧА", "ВЫВОД"]
    x = 60
    width, gap, y1, y2 = 270, 45, 450, 610
    for i, name in enumerate(stages):
        rounded_box(d, x, y1, x + width, y2, [name], f22)
        if i < len(stages) - 1:
            arrow(d, x + width + 6, (y1 + y2) / 2, x + width + gap - 8, (y1 + y2) / 2)
        x += width + gap

    d.text((60, 680), "Пример: размеры помещения и нормы освещённости → расчёт числа светильников → отчёт",
           font=f19, fill=GREY)
    img.save(path)


def draw_l2(path):
    W, H = 1600, 1000
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    f24, f20, f18, f27 = font(24), font(20), font(18), font(27, True)

    rounded_box(d, 560, 30, 1040, 120, ["ИНФОРМАЦИОННЫЕ СИСТЕМЫ"], f27)
    groups = [
        (20, "По масштабу", ["персональные", "групповые", "корпоративные", "глобальные"]),
        (410, "По сфере применения", ["производственные", "бухгалтерские", "учебные", "САПР, АСУ ТП"]),
        (800, "По автоматизации", ["ручные", "автоматизированные", "автоматические"]),
        (1190, "По обработке данных", ["поисковые", "управляющие", "СППР", "экспертные"]),
    ]
    for x, title, leaves in groups:
        rounded_box(d, x, 260, x + 370, 370, [title], f24)
        arrow(d, 800, 125, x + 185, 250)
        y = 430
        for leaf in leaves:
            rounded_box(d, x + 30, y, x + 340, y + 80, [leaf], f20,
                        fill="white", outline=(150, 170, 158), radius=14)
            d.line((x + 185, 375, x + 185, y), fill=(150, 170, 158), width=3)
            y += 120

    d.text((40, 930), "Обеспечивающие подсистемы: информационное, техническое, программное, математическое, организационное, правовое",
           font=f18, fill=GREY)
    img.save(path)


def draw_l3(path):
    W, H = 1600, 900
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    f22, f20 = font(22), font(20)

    left, right, top, bottom = 170, 1520, 80, 760
    d.line((left, bottom, right, bottom), fill=INK, width=3)
    d.line((left, top, left, bottom), fill=INK, width=3)

    def px(L):
        return left + L / 60 * (right - left)

    def py(pct):
        return bottom - pct / 6 * (bottom - top)

    for L in range(0, 61, 10):
        d.line((px(L), bottom, px(L), bottom + 10), fill=INK, width=2)
        center_text(d, (px(L), bottom + 34), str(L), f20)
    for pct in range(0, 7):
        d.line((left - 10, py(pct), left, py(pct)), fill=INK, width=2)
        d.text((left - 55, py(pct) - 12), f"{pct}%", font=f20, fill=INK)
    for L in range(10, 60, 10):
        d.line((px(L), top, px(L), bottom), fill=(226, 232, 227), width=2)
    for pct in range(1, 6):
        d.line((left, py(pct), right, py(pct)), fill=(226, 232, 227), width=2)

    # норма 5 %
    for L in range(0, 60, 2):
        d.line((px(L), py(5), px(L + 1), py(5)), fill=(200, 80, 60), width=3)
    d.text((right - 250, py(5) - 34), "норма 5 %", font=f22, fill=(200, 80, 60))

    data = {2.5: (47, 107, 79), 4.0: (40, 90, 170)}
    for S, color in data.items():
        points = []
        for i in range(0, 61):
            L = i
            pct = 2 * (3000 / 220) * L / (54 * S) / 220 * 100
            points.append((px(L), py(min(pct, 6))))
        d.line(points, fill=color, width=5, joint="curve")
        d.text((px(52), py(2 * (3000 / 220) * 52 / (54 * S) / 220 * 100) - 40),
               f"S = {S:g} мм²", font=f22, fill=color)

    center_text(d, ((left + right) / 2, 850), "Длина линии L, м", f22)
    d.text((20, top - 10), "Потеря напряжения", font=f22, fill=INK)
    d.text((20, top + 22), "ΔU, %", font=f22, fill=INK)
    d.text((left + 10, py(5) - 66), "P = 3 кВт, U = 220 В, медь", font=f20, fill=GREY)
    img.save(path)


# --------------------------------------------------------------------------- содержимое

LECTURES = [
    {
        "topic": "information-resources",
        "title": "Информационные ресурсы и информационные технологии",
        "description": "Данные, информация и знания; виды информационных ресурсов; этапы информационных технологий",
        "image": ("info-resources-scheme.png", "Схема преобразования данных в знания и этапы информационной технологии", 8.5),
        "blocks": [
            ("p", "Современное производство и учёба работают с большими объёмами сведений. Чтобы сведения "
                  "превратились в полезный результат, их нужно собрать, обработать и сохранить. Этим занимаются "
                  "информационные технологии, а источником сведений служат информационные ресурсы."),
            ("h2", "Данные, информация и знания"),
            ("p", "Данные — это зафиксированные факты и записи, которые сами по себе не имеют смысла: числа, "
                  "текст, показания приборов. Информация — сведения, которые получатель понял и которые уменьшили "
                  "его неопределённость. Знания — информация, проверенная практикой и пригодная для принятия решений."),
            ("p", "Одна и та же запись может быть данными для одного человека и информацией для другого: всё "
                  "зависит от подготовки и задачи."),
            ("p", "К информации предъявляют требования:"),
            ("ul", [
                "объективность — она не зависит от мнения и желания;",
                "достоверность — отражает реальное положение дел;",
                "полнота — достаточна для решения задачи;",
                "актуальность — соответствует текущему моменту;",
                "полезность — даёт практический эффект;",
                "понятность — изложена на доступном языке.",
            ]),
            ("h2", "Измерение информации"),
            ("p", "Количество информации измеряют в битах. Если сообщение выбирает один из N равновероятных "
                  "вариантов, то:"),
            ("formula", "N = 2^i^"),
            ("p", "где i — количество информации в битах. Сообщение о результате броска кубика содержит "
                  "i = log~2~ 6 ≈ 2,6 бита, а выбор одной из восьми ламп — 3 бита."),
            ("p", "Объём текста считают по числу символов:"),
            ("formula", "V = K · i"),
            ("p", "где K — количество символов, i — число бит на символ. Текст из 2000 символов в кодировке, "
                  "где символ занимает 8 бит, занимает 16 000 бит = 2000 байт ≈ 2 КБ."),
            ("fig", 0),
            ("h2", "Информационные ресурсы"),
            ("p", "Информационные ресурсы — совокупность данных, организованных для получения достоверной "
                  "информации: библиотеки, архивы, базы данных, справочные и правовые системы, интернет-сайты. "
                  "Закон «Об информации, информационных технологиях и о защите информации» определяет их как "
                  "отдельные документы и их массивы в информационных системах."),
            ("p", "Ресурсы различают по назначению (научно-технические, правовые, статистические, учебные), "
                  "по форме представления (текстовые, графические, мультимедийные) и по носителю (бумажные, "
                  "магнитные, оптические, облачные)."),
            ("p", "Главное отличие ресурса от простого набора данных — упорядоченность: ресурс организован так, "
                  "чтобы нужные сведения можно было быстро найти и использовать."),
            ("h2", "Информационные технологии"),
            ("p", "Информационная технология — совокупность методов и средств сбора, хранения, обработки, "
                  "передачи и представления информации. Технология отвечает на вопрос «как», а ресурс — «из чего»."),
            ("p", "Основные этапы:"),
            ("ol", [
                "сбор и ввод данных;",
                "обработка — расчёт, сортировка, поиск;",
                "хранение на носителе или в базе данных;",
                "передача по каналам связи;",
                "вывод и оформление результата.",
            ]),
            ("p", "Пример. Для расчёта освещения цеха нужны данные о размерах помещения и нормах освещённости "
                  "(ресурс). По формулам вычисляют световой поток и число светильников (обработка), результат "
                  "сохраняют и оформляют в отчёт (хранение и вывод)."),
            ("p", "Развитие технологий прошло этапы: ручной (перо, счёты), механический (арифмометр, пишущая "
                  "машинка), электронный (ЭВМ), сетевой (интернет, облака, мобильные устройства)."),
            ("h2", "Краткий вывод"),
            ("p", "Данные превращаются в информацию, а информация — в знания. Информационные ресурсы служат "
                  "источником сведений, информационные технологии — инструментом их обработки. К информации "
                  "предъявляют требования достоверности, полноты и актуальности, а её количество измеряют в битах."),
            ("h2", "Вопросы для самопроверки"),
            ("ol", [
                "Чем данные отличаются от информации?",
                "Какие требования предъявляют к информации?",
                "Сколько бит содержит сообщение о выборе одного из 16 вариантов?",
                "Что относят к информационным ресурсам?",
                "Назовите этапы информационной технологии.",
            ]),
        ],
    },
    {
        "topic": "information-systems",
        "title": "Информационные системы. Классификация информационных систем",
        "description": "Определение и состав информационной системы; классификация по масштабу, сфере, автоматизации и обработке данных",
        "image": ("is-classification.png", "Классификация информационных систем по основным признакам", 8.5),
        "blocks": [
            ("p", "Информационная система собирает, хранит и обрабатывает данные так, чтобы человек или программа "
                  "могли получить нужный результат. Такие системы окружают нас в учёбе и на производстве: "
                  "от электронного журнала до АСУ ТП подстанции."),
            ("h2", "Что такое информационная система"),
            ("p", "Закон «Об информации, информационных технологиях и о защите информации» определяет "
                  "информационную систему как совокупность содержащейся в базах данных информации "
                  "и обеспечивающих её обработку информационных технологий и технических средств."),
            ("p", "Признаки информационной системы:"),
            ("ul", [
                "база данных — организованное хранилище сведений;",
                "программное обеспечение — правила обработки данных;",
                "технические средства — компьютеры, серверы, датчики;",
                "персонал — пользователи и администраторы;",
                "регламенты — порядок работы и защиты информации.",
            ]),
            ("h2", "Состав информационной системы"),
            ("p", "Систему делят на обеспечивающие подсистемы:"),
            ("ul", [
                "информационное обеспечение — состав данных, форматы, справочники;",
                "техническое — оборудование и каналы связи;",
                "программное — операционные системы и прикладные программы;",
                "математическое — методы и алгоритмы расчёта;",
                "организационное — структура служб и инструкции;",
                "правовое — нормы и ответственность.",
            ]),
            ("p", "Функциональная часть — это подсистемы, которые решают конкретные задачи: учёт, планирование, "
                  "анализ, отчётность."),
            ("h2", "Классификация информационных систем"),
            ("p", "Системы классифицируют по нескольким признакам."),
            ("fig", 0),
            ("p", "По масштабу различают персональные, групповые (отдел, цех), корпоративные и глобальные системы."),
            ("p", "По сфере применения — производственные, бухгалтерские, кадровые, торговые, медицинские, "
                  "учебные, а также САПР (проектирование) и АСУ ТП (управление технологическим процессом)."),
            ("p", "По степени автоматизации — ручные, автоматизированные (решение принимает человек) "
                  "и автоматические (работают без участия человека)."),
            ("p", "По характеру обработки данных — информационно-поисковые, информационно-управляющие, системы "
                  "поддержки принятия решений и экспертные системы."),
            ("h2", "Примеры в профессиональной деятельности"),
            ("p", "Электромонтажник встречается с информационными системами ежедневно: наряд-заказы и журналы "
                  "работ, системы учёта электроэнергии, SCADA-системы подстанций, САПР для схем и чертежей, "
                  "складские и кадровые системы предприятия."),
            ("p", "Выбор класса системы зависит от задач: для небольшого участка достаточно персональной базы "
                  "данных, для энергообъекта нужна корпоративная АСУ ТП с резервированием и защитой информации."),
            ("h2", "Краткий вывод"),
            ("p", "Информационная система объединяет данные, программы, оборудование и людей. Её классифицируют "
                  "по масштабу, сфере применения, степени автоматизации и характеру обработки данных. Понимание "
                  "класса системы помогает правильно выбрать инструмент и организовать работу."),
            ("h2", "Вопросы для самопроверки"),
            ("ol", [
                "Какое определение информационной системы даёт закон?",
                "Назовите обеспечивающие подсистемы информационной системы.",
                "Чем автоматизированная система отличается от автоматической?",
                "Какие информационные системы применяют в электроэнергетике?",
                "По каким признакам классифицируют информационные системы?",
            ]),
        ],
    },
    {
        "topic": "engineering-calculations-automation",
        "title": "Автоматизация выполнения различных инженерных расчётов",
        "description": "Инструменты и этапы автоматизации расчётов; пример расчёта потери напряжения в таблице",
        "image": ("calc-graph.png", "График потери напряжения в линии в зависимости от длины", 9.5),
        "blocks": [
            ("p", "Инженерный расчёт состоит из повторяющихся операций: подставить данные, вычислить по формуле, "
                  "проверить результат, оформить. Автоматизация переносит эти операции в программу, поэтому "
                  "расчёт выполняется быстрее, точнее и без арифметических ошибок."),
            ("h2", "Зачем автоматизировать расчёт"),
            ("ul", [
                "скорость — серия вариантов считается за минуты;",
                "точность — исключены арифметические ошибки;",
                "повторяемость — одинаковый результат при тех же данных;",
                "наглядность — таблицы и графики;",
                "хранение — расчёт можно вернуть и проверить.",
            ]),
            ("h2", "Инструменты автоматизации"),
            ("p", "Для инженерных расчётов применяют табличные процессоры (Excel, LibreOffice Calc), математические "
                  "пакеты (MathCAD, MATLAB), языки программирования (Python с библиотекой NumPy) и специализированные "
                  "программы (САПР, «Электрик»)."),
            ("p", "Табличный процессор удобен, когда расчёт сводится к формулам: одна ячейка хранит исходные "
                  "данные, другая — результат. Абсолютные ссылки позволяют размножить формулу на весь столбец, "
                  "а диаграмма строится по готовой таблице."),
            ("h2", "Этапы автоматизации"),
            ("ol", [
                "постановка задачи — что дано и что нужно найти;",
                "математическая модель — формулы и ограничения;",
                "алгоритм — последовательность действий;",
                "реализация — таблица или программа;",
                "проверка — сравнение с ручным расчётом;",
                "анализ — выводы и оформление.",
            ]),
            ("h2", "Пример: расчёт потери напряжения"),
            ("p", "Ток нагрузки при однофазном включении:"),
            ("formula", "I = P / U"),
            ("p", "Потеря напряжения в двухпроводной линии:"),
            ("formula", "ΔU = 2 · I · L / (γ · S)"),
            ("p", "где L — длина линии, м; γ — удельная проводимость (для меди 54 м/(Ом·мм^2^)); S — сечение "
                  "жилы, мм^2^. Относительная потеря напряжения:"),
            ("formula", "ΔU% = ΔU / U · 100"),
            ("p", "Пример. P = 3 кВт, U = 220 В, L = 30 м, S = 2,5 мм^2^. Тогда I = 13,6 А, ΔU = 6,1 В, "
                  "ΔU% ≈ 2,8 % — в пределах нормы 5 %."),
            ("fig", 0),
            ("p", "В таблице достаточно задать P, U, L, S и γ, а формулы протянуть на десять вариантов. График "
                  "показывает, при какой длине линия выходит за допустимые 5 %: для сечения 2,5 мм^2^ это "
                  "около 54 м, для 4 мм^2^ — около 87 м."),
            ("h2", "Проверка результата"),
            ("p", "Автоматический расчёт всё равно проверяют: сверяют с ручным примером, следят за размерностями, "
                  "проверяют предельные случаи (нулевая длина, большое сечение) и оценивают разумность числа."),
            ("h2", "Краткий вывод"),
            ("p", "Автоматизация избавляет от рутинных вычислений и снижает число ошибок. Она требует "
                  "математической модели, корректных формул и проверки результата. Таблицы, пакеты и программы "
                  "позволяют быстро перебрать варианты и наглядно представить итог."),
            ("h2", "Вопросы для самопроверки"),
            ("ol", [
                "Какие преимущества даёт автоматизация расчётов?",
                "Назовите этапы автоматизации расчёта.",
                "Чем табличный процессор удобен для инженерных расчётов?",
                "Как проверить правильность автоматического расчёта?",
                "Какой инструмент выбрать для расчёта 50 вариантов по одной формуле?",
            ]),
        ],
    },
]


# --------------------------------------------------------------------------- Word

def _set_rfonts(rpr):
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


def style_font(style, size=None, bold=None, color=None):
    rpr = style.element.get_or_add_rPr()
    _set_rfonts(rpr)
    style.font.name = FONT
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
        col = rpr.find(qn("w:color"))
        if col is not None:
            for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
                if col.get(qn(attr)) is not None:
                    del col.attrib[qn(attr)]


def run_font(run, size=None, bold=None, italic=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    _set_rfonts(rpr)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def rich_docx(paragraph, text, size=None, bold=None, italic=None):
    for token in re.split(r"(~[^~]+~|\^[^\^]+\^|\*[^*]+\*)", text):
        if not token:
            continue
        if token.startswith("~") and token.endswith("~") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run_font(run, size, bold, italic)
            run.font.subscript = True
        elif token.startswith("^") and token.endswith("^") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run_font(run, size, bold, italic)
            run.font.superscript = True
        elif token.startswith("*") and token.endswith("*") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run_font(run, size, bold, True)
        else:
            run = paragraph.add_run(token)
            run_font(run, size, bold, italic)


def build_docx(lecture):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    normal = doc.styles["Normal"]
    style_font(normal, size=11)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.07

    h1 = doc.styles["Heading 1"]
    style_font(h1, size=11.5, bold=True, color=RGBColor(0, 0, 0))
    h1.paragraph_format.space_before = Pt(4)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich_docx(p, lecture["title"], size=14, bold=True)

    img_name, _, img_width = lecture["image"]

    for block in lecture["blocks"]:
        kind = block[0]
        if kind == "h2":
            p = doc.add_paragraph(style="Heading 1")
            rich_docx(p, block[1])
        elif kind == "p":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(0.8)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.07
            rich_docx(p, block[1])
        elif kind == "ul":
            for item in block[1]:
                p = doc.add_paragraph(style="List Bullet")
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05
                rich_docx(p, item)
        elif kind == "ol":
            for index, item in enumerate(block[1], start=1):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.left_indent = Cm(0.75)
                p.paragraph_format.first_line_indent = Cm(-0.75)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05
                rich_docx(p, f"{index}.\t{item}")
        elif kind == "formula":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            rich_docx(p, block[1])
        elif kind == "fig":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.keep_with_next = True
            shape = p.add_run().add_picture(os.path.join(BASE, lecture["topic"], "images", img_name), width=Cm(img_width))
            sp_pr = shape._inline.graphic.graphicData.pic.spPr
            line = OxmlElement("a:ln")
            line.set("w", "0")
            line.append(OxmlElement("a:noFill"))
            sp_pr.append(line)
            c = doc.add_paragraph()
            c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.paragraph_format.space_before = Pt(1)
            c.paragraph_format.space_after = Pt(4)
            rich_docx(c, "Рисунок 1 — " + lecture["image"][1], size=9, italic=True)

    out = os.path.join(BASE, lecture["topic"], "docs", f"Лекция — {lecture['title']}.docx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    return out


# --------------------------------------------------------------------------- HTML

def html_text(text):
    out = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    out = re.sub(r"~([^~]+)~", r"<sub>\1</sub>", out)
    out = re.sub(r"\^([^\^]+)\^", r"<sup>\1</sup>", out)
    out = re.sub(r"\*([^*]+)\*", r"<em>\1</em>", out)
    return out


def build_html(lecture):
    img_name, alt, _ = lecture["image"]
    parts = []
    open_section = False
    for block in lecture["blocks"]:
        kind = block[0]
        if kind == "h2":
            if open_section:
                parts.append("</section>")
            parts.append(f"<section>\n  <h2>{html_text(block[1])}</h2>")
            open_section = True
        elif kind == "p":
            parts.append(f"  <p>{html_text(block[1])}</p>")
        elif kind == "ul":
            parts.append("  <ul>")
            for item in block[1]:
                parts.append(f"    <li>{html_text(item)}</li>")
            parts.append("  </ul>")
        elif kind == "ol":
            parts.append("  <ol>")
            for item in block[1]:
                parts.append(f"    <li>{html_text(item)}</li>")
            parts.append("  </ol>")
        elif kind == "formula":
            parts.append(f'  <ul class="formulas">\n    <li>{html_text(block[1])}</li>\n  </ul>')
        elif kind == "fig":
            src = f"content/groups/1eo-24/op-08/{lecture['topic']}/images/{img_name}"
            parts.append(
                "  <figure>\n"
                f'    <img src="{src}" alt="{html_text(alt)}" loading="lazy" />\n'
                f"    <figcaption>Рисунок 1 — {html_text(alt)}</figcaption>\n"
                "  </figure>"
            )
    if open_section:
        parts.append("</section>")

    docx_name = f"Лекция — {lecture['title']}.docx"
    link = (
        f'<p><a href="content/groups/1eo-24/op-08/{lecture["topic"]}/docs/{docx_name}">'
        "Скачать лекцию в Word (.docx)</a></p>"
    )
    html = "\n".join(parts) + "\n\n" + link + "\n"
    out = os.path.join(BASE, lecture["topic"], "Лекция.html")
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(html)
    return out


def main():
    os.makedirs(os.path.join(BASE, "information-resources", "images"), exist_ok=True)
    draw_l1(os.path.join(BASE, "information-resources", "images", "info-resources-scheme.png"))
    os.makedirs(os.path.join(BASE, "information-systems", "images"), exist_ok=True)
    draw_l2(os.path.join(BASE, "information-systems", "images", "is-classification.png"))
    os.makedirs(os.path.join(BASE, "engineering-calculations-automation", "images"), exist_ok=True)
    draw_l3(os.path.join(BASE, "engineering-calculations-automation", "images", "calc-graph.png"))

    for lecture in LECTURES:
        print("docx:", build_docx(lecture))
        print("html:", build_html(lecture))


if __name__ == "__main__":
    main()
