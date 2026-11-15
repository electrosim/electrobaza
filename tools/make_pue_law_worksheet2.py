# -*- coding: utf-8 -*-
"""Двухлистовое задание: лист 1 — правовая база и понятия, лист 2 — две ситуации."""

import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\5-pue-i-gost\docs"
OUT = os.path.join(BASE, "Задание — Правовой анализ ПУЭ (2 листа).docx")
FONT = "Times New Roman"


def setup(doc):
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.4)
    s.top_margin = Cm(1.1)
    s.bottom_margin = Cm(1.0)
    normal = doc.styles["Normal"]
    rpr = normal.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)
    normal.font.name = FONT
    normal.font.size = Pt(9.5)
    normal.paragraph_format.space_after = Pt(1)
    normal.paragraph_format.line_spacing = 1.0


def para(doc, text, size=9.5, bold=False, italic=False, align=None, after=1, before=0):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    return p


def cell_para(cell, text, size=9, bold=False, after=1, first=False):
    if first and cell.paragraphs[0].text == "":
        p = cell.paragraphs[0]
    else:
        p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.bold = bold
    return p


def borders(table, sz=4, color="808080"):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)


def build():
    doc = Document()
    setup(doc)

    # ---------------- ЛИСТ 1 ----------------
    para(doc, "Правовая база: ПУЭ, ГОСТ и ответственность", size=14, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=1)
    para(doc, "Лист 1 из 2 · МДК 04.02 + ОП.12 · все ответы ищутся здесь", size=9, italic=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=3)

    para(doc, "Статус ПУЭ. ПУЭ — ведомственный нормативный акт, а не закон. Действует 6-е издание "
              "(изменения на 20.06.2003); 7-е издание вводилось отдельными главами и не прошло "
              "регистрацию в Минюсте. Обязательность для конкретного лица создают: технический регламент, "
              "проект, договор подряда, трудовые обязанности и надзор. ГОСТ применяется добровольно, "
              "но становится обязательным, если на него ссылается регламент или договор.",
         size=9, after=3)

    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t, 4, "A6A6A6")
    left, right = t.rows[0].cells
    left.width = Cm(9.1)
    right.width = Cm(9.1)
    cell_para(left, "Технические нормы", 10, bold=True, first=True)
    cell_para(left, "ПУЭ, п. 2.1.21 (гл. 2.1, 6-е изд.): соединение, ответвление и оконцевание жил "
                    "должны производиться при помощи опрессовки, сварки, пайки или сжимов (винтовых, "
                    "болтовых и т. п.).")
    cell_para(left, "ПУЭ, гл. 1.7: металлические части, которые могут оказаться под напряжением, "
                    "должны быть заземлены или занулены (изложение).")
    cell_para(left, "ПУЭ, гл. 3.1: защита сетей до 1 кВ от сверхтоков автоматическими выключателями "
                    "и предохранителями (изложение).")
    cell_para(left, "ПУЭ, п. 7.1.79: в групповых сетях, питающих штепсельные розетки, СЛЕДУЕТ "
                    "применять УЗО с током срабатывания не более 30 мА (рекомендация).")
    cell_para(left, "ПУЭ, п. 7.1.82: ОБЯЗАТЕЛЬНА установка УЗО не более 30 мА для розеточных сетей "
                    "вне помещений и в особо опасных помещениях.")
    cell_para(left, "ГОСТ Р 50571 — электроустановки зданий (гармонизирован с МЭК 60364); "
                    "ТР ТС 004/2011 и 020/2011 — обязательные требования к оборудованию и ЭМС.", after=2)

    cell_para(right, "Правовые нормы", 10, bold=True, first=True)
    cell_para(right, "ГК РФ, ст. 743: подрядчик выполняет работы по технической документации; "
                     "ст. 754, 755 — ответственность за качество и гарантии.")
    cell_para(right, "ГК РФ, ст. 1064: вред имуществу или личности возмещает причинитель вреда.")
    cell_para(right, "ГК РФ, ст. 1068: за вред, причинённый РАБОТНИКОМ при исполнении трудовых "
                     "обязанностей, отвечает работодатель. К подрядчику по ГПХ не применяется.")
    cell_para(right, "ГК РФ, ст. 1081 — регресс к причинителю; ст. 1085–1087 — возмещение вреда "
                     "здоровью; ст. 151 — моральный вред.")
    cell_para(right, "ТК РФ, ст. 192 — замечание, выговор, увольнение; ст. 241 — материальная "
                     "ответственность в пределах среднего месячного заработка; ст. 243 — случаи "
                     "полной ответственности.")
    cell_para(right, "КоАП РФ, ст. 9.11 — нарушение правил устройства и эксплуатации "
                     "электроустановок: штраф (граждане 1–2 тыс., должностные лица и ИП 2–4 тыс. руб.), "
                     "для организаций — вплоть до приостановления деятельности.")
    cell_para(right, "УК РФ, ст. 143 — нарушение требований охраны труда, повлёкшее по "
                     "неосторожности ТЯЖКИЙ вред здоровью или смерть. Субъект — лицо, обязанное соблюдать "
                     "требования охраны труда (руководитель, мастер).")
    cell_para(right, "УК РФ, ст. 168 — уничтожение имущества по неосторожности при КРУПНОМ ущербе; "
                     "ст. 219 — нарушение требований пожарной безопасности, повлёкшее тяжкий вред "
                     "или смерть.", after=2)

    para(doc, "Ключевые понятия", size=10, bold=True, before=5, after=1)
    para(doc, "Правонарушение — виновное противоправное деяние. Состав: объект (на что посягает), "
              "объективная сторона (само деяние, последствия, причинная связь), субъект (кто), "
              "субъективная сторона (вина: умысел или неосторожность).", size=9)
    para(doc, "Виды ответственности применяются независимо: административная не исключает гражданской, дисциплинарной и уголовной. Регресс — обратное требование того, кто возместил вред, к причинителю. Регресс работодателя к работнику ограничен нормами ТК о материальной ответственности (ст. 241, 243). Материальная "
              "ответственность — обязанность работника возместить ущерб работодателю (ТК, ст. 241, 243).",
         size=9, after=2)

    para(doc, "Как отвечать (алгоритм)", size=10, bold=True, before=3, after=1)
    para(doc, "1. Найдите норму ПУЭ: пункт, глава, издание.  2. Объясните, почему она обязательна "
              "именно для этого лица (регламент, проект, договор, трудовые обязанности).  "
              "3. Определите последствия: имущество или здоровье, тяжесть, размер.  4. Подберите "
              "статью по субъекту и последствиям; если данных не хватает — напишите, каких именно.",
         size=9, after=2)

    # ---------------- ЛИСТ 2 ----------------
    doc.add_page_break()
    para(doc, "Задание: две ситуации", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=1)
    para(doc, "Лист 2 из 2 · выберите один вариант · отвечать можно только по листу 1", size=9, italic=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=5)

    def variant(title, situation, questions):
        para(doc, title, size=11, bold=True, before=2, after=1)
        para(doc, situation, size=9.5, after=3)
        tb = doc.add_table(rows=0, cols=2)
        tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders(tb, 4, "A6A6A6")
        for q in questions:
            row = tb.add_row()
            row.cells[0].width = Cm(8.2)
            row.cells[1].width = Cm(9.8)
            cell_para(row.cells[0], q, 9, first=True)
            cell_para(row.cells[1], "", 9, first=True)
            row.height = Cm(1.3)

    variant(
        "Вариант 1",
        "Индивидуальный предприниматель по договору подряда смонтировал электропроводку в квартире. "
        "В договоре было условие: работы выполняются в соответствии с действующими ПУЭ и ГОСТ. "
        "Жилы в распределительной коробке он соединил скруткой. Через полгода место соединения "
        "нагрелось, начался пожар: повреждена мебель и отделка. Пострадавших нет; размер ущерба "
        "в акте не указан.",
        [
            "1. Какой пункт ПУЭ нарушен и в чём именно?",
            "2. Почему норма обязательна для ИП, если ПУЭ — не закон?",
            "3. Кто отвечает перед собственником? Применима ли ст. 1068 ГК?",
            "4. Есть ли состав ст. 9.11 КоАП? Кто субъект?",
            "5. Есть ли состав ст. 168 УК? Каких данных не хватает?",
            "6. Изменится ли квалификация, если ущерб окажется крупным?",
        ],
    )

    variant(
        "Вариант 2",
        "Работник ООО «Монтаж» по заданию мастера смонтировал щит в офисе по проекту, выполненному по ГОСТ Р 50571, но не установил УЗО 30 мА, "
        "предусмотренное проектом. Мастер подписал акт приёмки без проверки. Через месяц сотрудник "
        "офиса получил электротравму средней тяжести, прикоснувшись к повреждённому проводу.",
        [
            "1. Какие нормы нарушены: ПУЭ, проект, ГОСТ?",
            "2. Почему требование обязательно, если п. 7.1.79 говорит «следует»?",
            "3. Кто отвечает перед пострадавшим? Что такое регресс?",
            "4. Есть ли состав ст. 143 УК? Обоснуйте.",
            "5. Может ли ООО взыскать с монтажника всю сумму вреда? (ТК, ст. 241, 243)",
            "6. Изменится ли квалификация, если вред здоровью окажется тяжким?",
        ],
    )

    os.makedirs(BASE, exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
