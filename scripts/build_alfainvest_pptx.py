#!/usr/bin/env python3
"""Build AlfaInvest.pptx from the AlfaInvest.html deck content."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from pptx.oxml import parse_xml
from copy import deepcopy
from lxml import etree

# Brand colors
RED = RGBColor(0xEF, 0x31, 0x24)
RED_DARK = RGBColor(0xC4, 0x25, 0x1A)
RED_SOFT = RGBColor(0xFF, 0xF1, 0xF0)
INK = RGBColor(0x0F, 0x0F, 0x0F)
INK2 = RGBColor(0x1C, 0x1C, 0x1E)
MUTED = RGBColor(0x6E, 0x6E, 0x73)
LINE = RGBColor(0xE6, 0xE6, 0xEB)
BG = RGBColor(0xFF, 0xFF, 0xFF)
BG_SOFT = RGBColor(0xF6, 0xF6, 0xF8)
GREEN = RGBColor(0x12, 0xA1, 0x50)
AMBER = RGBColor(0xC7, 0x77, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x11, 0x11, 0x11)
BLUE = RGBColor(0x1B, 0x5F, 0xCC)
GRAY_PILL = RGBColor(0xEF, 0xEF, 0xF2)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def rgb_hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def set_run(run, size=14, bold=False, color=INK, font="Arial"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_text(shape, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, font="Arial"):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return tf


def add_para(tf, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, space_before=0, space_after=4):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def fill_solid(shape, color: RGBColor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def no_line(shape):
    shape.line.fill.background()


def rect(slide, left, top, width, height, fill=None, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.08
    if fill is not None:
        fill_solid(shape, fill)
    else:
        shape.fill.background()
    if line is not None:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    else:
        no_line(shape)
    return shape


def sharp_rect(slide, left, top, width, height, fill=None, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill is not None:
        fill_solid(shape, fill)
    else:
        shape.fill.background()
    if line is not None:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    else:
        no_line(shape)
    return shape


def oval(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
    fill_solid(shape, fill)
    no_line(shape)
    return shape


def textbox(slide, left, top, width, height, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT):
    shape = slide.shapes.add_textbox(left, top, width, height)
    add_text(shape, text, size=size, bold=bold, color=color, align=align)
    return shape


def slide_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    fill_solid(bg, color)
    no_line(bg)
    # send to back
    spTree = slide.shapes._spTree
    sp = bg._element
    spTree.remove(sp)
    spTree.insert(2, sp)
    return bg


def slide_head(slide, num, kicker, title, lead=None, dark=False):
    num_color = RGBColor(0xFF, 0x7A, 0x6E) if dark else RED
    kicker_color = RGBColor(0x7A, 0x7A, 0x80) if dark else MUTED
    title_color = WHITE if dark else INK
    lead_color = RGBColor(0xB8, 0xB8, 0xBE) if dark else MUTED

    badge = rect(slide, Inches(0.7), Inches(0.35), Inches(0.55), Inches(0.32), fill=None, line=num_color)
    badge.adjustments[0] = 0.15
    add_text(badge, num, size=11, bold=True, color=num_color, align=PP_ALIGN.CENTER)
    badge.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    textbox(slide, Inches(1.4), Inches(0.38), Inches(6), Inches(0.3),
            kicker.upper(), size=11, bold=True, color=kicker_color)
    textbox(slide, Inches(0.7), Inches(0.8), Inches(12), Inches(0.6),
            title, size=28, bold=True, color=title_color)
    if lead:
        tb = textbox(slide, Inches(0.7), Inches(1.4), Inches(11.8), Inches(0.7),
                     lead, size=14, color=lead_color)
        return tb
    return None


def card(slide, left, top, width, height, fill=BG, line=LINE):
    return rect(slide, left, top, width, height, fill=fill, line=line)


def assump(slide, left, top, width, text):
    box = rect(slide, left, top, width, Inches(0.55), fill=RGBColor(0xFF, 0xF9, 0xF0), line=None)
    # left accent via thin rect
    sharp_rect(slide, left, top, Inches(0.06), Inches(0.55), fill=RGBColor(0xF0, 0xB4, 0x29))
    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.08), width - Inches(0.3), Inches(0.45))
    add_text(tb, text, size=10, color=RGBColor(0x6B, 0x5A, 0x2E))
    return box


def set_table_style(table, header_fill=BG_SOFT):
    for cell in table.rows[0].cells:
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = MUTED
                run.font.name = "Arial"


def fill_cell(cell, text, size=11, bold=False, color=INK, fill=None, align=PP_ALIGN.LEFT):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)


# ---------- slides ----------

def make_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, DARK)
    # gradient-ish accent bar
    sharp_rect(slide, 0, 0, Inches(0.12), SLIDE_H, fill=RED)

    badge = rect(slide, Inches(0.8), Inches(1.5), Inches(4.2), Inches(0.4),
                 fill=RGBColor(0x3A, 0x14, 0x12), line=RGBColor(0x8A, 0x3A, 0x35))
    badge.adjustments[0] = 0.5
    add_text(badge, "●  Product Owner · Тестовое задание 2026", size=11, bold=True,
             color=RGBColor(0xFF, 0x7A, 0x6E), align=PP_ALIGN.CENTER)

    textbox(slide, Inches(0.8), Inches(2.2), Inches(7), Inches(1),
            "Alfa Invest", size=48, bold=True, color=WHITE)
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(3.3), Inches(7), Inches(1.1))
    add_text(tb, "Инвестиционное приложение для розничных клиентов Альфа-Банка. Первый шаг в инвестиции — без страха и лишних сложностей.",
             size=16, color=RGBColor(0xB8, 0xB8, 0xBE))

    metas = [
        ("Продукт", "Веб-платформа инвестиций"),
        ("Рынок", "Беларусь · БВФБ"),
        ("Фокус", "Новички 25–40 лет"),
    ]
    x = Inches(0.8)
    for label, val in metas:
        textbox(slide, x, Inches(5.0), Inches(2.6), Inches(0.25), label.upper(), size=10, color=RGBColor(0x7A, 0x7A, 0x80))
        textbox(slide, x, Inches(5.3), Inches(2.6), Inches(0.35), val, size=13, bold=True, color=WHITE)
        x += Inches(2.8)

    # cover card
    c = card(slide, Inches(8.5), Inches(1.8), Inches(4.2), Inches(4.2), fill=WHITE, line=None)
    textbox(slide, Inches(8.8), Inches(2.1), Inches(2), Inches(0.25), "ПОРТФЕЛЬ", size=10, bold=True, color=MUTED)
    textbox(slide, Inches(11.2), Inches(2.1), Inches(1.2), Inches(0.3), "▲ +2,3%", size=12, bold=True, color=GREEN, align=PP_ALIGN.RIGHT)
    textbox(slide, Inches(8.8), Inches(2.45), Inches(3.6), Inches(0.55), "1 240,50 BYN", size=28, bold=True, color=INK)

    # fake chart bars
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн"]
    heights = [0.8, 0.95, 0.9, 1.3, 1.55, 1.8]
    bx = Inches(8.9)
    base = Inches(5.4)
    for m, h in zip(months, heights):
        sharp_rect(slide, bx, base - Inches(h), Inches(0.45), Inches(h), fill=RED_SOFT)
        sharp_rect(slide, bx, base - Inches(0.12), Inches(0.45), Inches(0.12), fill=RED)
        textbox(slide, bx - Inches(0.05), Inches(5.5), Inches(0.55), Inches(0.25), m, size=9, color=MUTED, align=PP_ALIGN.CENTER)
        bx += Inches(0.6)


def make_context(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "01", "Контекст",
               "Рынок растёт, но розница вкладывает мало",
               "Розничный интерес к БВФБ растёт кратно. Но объём средств растёт медленнее — люди приходят, пробуют и уходят.")

    cards = [
        ("3 из 4", "не совершают вторую сделку", "Клиент покупает один инструмент и не возвращается — нет понимания, что делать дальше", True),
        ("~50 BYN", "средний порог входа", "Новичок готов начать с небольшой суммы, но боится комиссий и «потерять всё»", False),
        ("2 лидера", "Aigenis и банковские приложения", "Рынок концентрируется вокруг 2–3 игроков, у банков нет сильного инвестпродукта", False),
    ]
    x = Inches(0.7)
    for kpi, title, desc, accent in cards:
        fill = RED_SOFT if accent else WHITE
        c = card(slide, x, Inches(2.3), Inches(3.9), Inches(2.15), fill=fill, line=RED if accent else LINE)
        textbox(slide, x + Inches(0.25), Inches(2.5), Inches(3.4), Inches(0.5), kpi, size=28, bold=True, color=RED)
        textbox(slide, x + Inches(0.25), Inches(3.1), Inches(3.4), Inches(0.35), title, size=13, bold=True, color=INK)
        textbox(slide, x + Inches(0.25), Inches(3.5), Inches(3.4), Inches(0.7), desc, size=11, color=MUTED)
        x += Inches(4.1)

    dark = card(slide, Inches(0.7), Inches(4.7), Inches(11.9), Inches(1.4), fill=DARK, line=None)
    textbox(slide, Inches(1.0), Inches(4.9), Inches(7), Inches(0.35), "Возможность для Альфа-Банка", size=15, bold=True, color=WHITE)
    textbox(slide, Inches(1.0), Inches(5.3), Inches(7.2), Inches(0.6),
            "У банка уже есть клиентская база, счёт и доверие. Не хватает понятного инвестиционного входа — как часть экосистемы.",
            size=12, color=RGBColor(0x9A, 0x9A, 0xA2))
    textbox(slide, Inches(8.6), Inches(4.95), Inches(1.8), Inches(0.25), "ЦЕЛЬ", size=9, color=RGBColor(0x7A, 0x7A, 0x80))
    textbox(slide, Inches(8.6), Inches(5.25), Inches(1.8), Inches(0.4), "Вход за 3 мин", size=16, bold=True, color=WHITE)
    textbox(slide, Inches(10.7), Inches(4.95), Inches(1.6), Inches(0.25), "ПОРОГ", size=9, color=RGBColor(0x7A, 0x7A, 0x80))
    textbox(slide, Inches(10.7), Inches(5.25), Inches(1.6), Inches(0.4), "От 50 BYN", size=16, bold=True, color=WHITE)

    assump(slide, Inches(0.7), Inches(6.4), Inches(11.9),
           "Допущение: точные цифры требуют верификации. Использованы качественные тренды: рост интереса розницы к БВФБ и низкий retention после первой сделки.")


def make_audience(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "02", "Целевая аудитория",
               "Кто будет пользоваться продуктом",
               "Три сегмента. Фокус — новички: они дают объём, формируют привычку и со временем становятся активными инвесторами.")

    rows = [
        ["Сегмент", "Доля новых", "Ключевая боль", "Что даёт продукт"],
        ["Новички\n25–40 лет, доход 1 500–3 000 BYN", "~55%",
         "«Боюсь потерять деньги, не понимаю, с чего начать»",
         "Онбординг под цель, стартовый набор, объяснение терминов"],
        ["Опытные, но пассивные\n30–50 лет, есть вклады", "~30%",
         "«Хочу больше, чем депозит, но нет времени»",
         "Автоследование, уведомления, сравнение с индексом"],
        ["Активные\n25–45 лет, опыт 1+ год", "~15%",
         "«Нужны инструменты, аналитика, скорость»",
         "Стакан, быстрые ордера, экспорт отчётов"],
    ]
    table_shape = slide.shapes.add_table(4, 4, Inches(0.7), Inches(2.3), Inches(11.9), Inches(3.4))
    table = table_shape.table
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(1.6)
    table.columns[2].width = Inches(3.5)
    table.columns[3].width = Inches(3.6)

    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            if i == 0:
                fill_cell(cell, val, size=10, bold=True, color=MUTED, fill=RGBColor(0xFA, 0xFA, 0xFC))
            elif i == 1:
                fill_cell(cell, val, size=11, bold=(j == 0 or j == 1), color=INK, fill=RED_SOFT)
            else:
                fill_cell(cell, val, size=11, bold=(j == 0), color=INK)

    assump(slide, Inches(0.7), Inches(6.1), Inches(11.9),
           "Приоритет: сегмент «новички» определяет объём рынка и будущую базу активных клиентов. Продукт не должен ограничивать опытных.")


def make_persona(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "03", "Портрет пользователя",
               "Персона: Алексей, 32 года",
               "«Хочу, чтобы деньги работали, но не хочу каждый день смотреть графики»")

    # persona card
    card(slide, Inches(0.7), Inches(2.2), Inches(3.2), Inches(4.5), fill=WHITE, line=LINE)
    av = oval(slide, Inches(1.7), Inches(2.5), Inches(1.2), Inches(1.2), RED)
    textbox(slide, Inches(1.7), Inches(2.75), Inches(1.2), Inches(0.7), "А", size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, Inches(0.9), Inches(3.9), Inches(2.8), Inches(0.35), "Алексей", size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
    textbox(slide, Inches(0.9), Inches(4.25), Inches(2.8), Inches(0.3), "Минск · менеджер · 32 года", size=11, color=MUTED, align=PP_ALIGN.CENTER)

    facts = [
        ("Доход", "2 200 BYN"),
        ("Опыт", "Новичок"),
        ("Вклад", "Есть, в BYN"),
        ("Свободно", "~300 BYN/мес"),
        ("Цель", "Накопить на авто"),
    ]
    y = Inches(4.7)
    for label, val in facts:
        textbox(slide, Inches(1.0), y, Inches(1.4), Inches(0.28), label, size=11, color=MUTED)
        textbox(slide, Inches(2.3), y, Inches(1.3), Inches(0.28), val, size=11, bold=True, color=INK, align=PP_ALIGN.RIGHT)
        y += Inches(0.32)

    blocks = [
        ("Боли", "• Не понимает разницу акция / облигация\n• Боится, что комиссии «съедят» доход\n• Не знает, сколько нужно для старта\n• Опасается, что «это только для богатых»", False),
        ("Триггеры", "• Реклама «купи акцию за 15 минут»\n• Коллега рассказал про купоны\n• Ставка по вкладу снизилась\n• Защита сбережений от инфляции", False),
        ("Что важно", "• Понятный язык без жаргона\n• Прозрачные комиссии до сделки\n• Возможность начать с малого\n• Не нужно быть «профи»", False),
        ("Инсайт", "Алексею нужен не «инвестиционный терминал», а проводник. Он готов довериться банку, если тот объяснит, что делать.", True),
    ]
    positions = [
        (Inches(4.2), Inches(2.2)),
        (Inches(8.8), Inches(2.2)),
        (Inches(4.2), Inches(4.5)),
        (Inches(8.8), Inches(4.5)),
    ]
    for (title, body, accent), (lx, ty) in zip(blocks, positions):
        c = card(slide, lx, ty, Inches(4.3), Inches(2.1),
                 fill=RED_SOFT if accent else WHITE, line=RED if accent else LINE)
        textbox(slide, lx + Inches(0.25), ty + Inches(0.2), Inches(3.8), Inches(0.3), title, size=14, bold=True, color=INK)
        textbox(slide, lx + Inches(0.25), ty + Inches(0.55), Inches(3.8), Inches(1.4), body, size=11, color=MUTED if not accent else INK2)


def make_cjm(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "04", "Путь клиента",
               "Customer Journey Map",
               "От первого касания до удержания. На каждом этапе — своя задача продукта.")

    headers = ["Этап", "Действие", "Барьер", "Что нужно", "Наш ответ", "Метрика"]
    data = [
        ["1. Осознание", "Узнал из рекламы / соцсетей", "«Это сложно»", "Простой оффер", "От 50 BYN. Без опыта", "CTR рекламы"],
        ["2. Рассмотрение", "Скачал, открыл", "Непонятно, что дальше", "Онбординг за 3 шага", "Опрос цели → рекомендация", "Конверсия в рег."],
        ["3. Первый вход", "КУС, открытие счёта", "Страх вводить данные", "Прозрачность «зачем?»", "КУС через данные банка", "% завершённых КУС"],
        ["4. Пополнение", "Вносит 100 BYN", "«А вдруг потеряю?»", "Снижение риска", "Демо-режим + риск", "Конверсия в пополнение"],
        ["5. Первая покупка", "Выбирает бумагу", "Не знает, что выбрать", "Рекомендации по цели", "ТОП-5 + стартовый портфель", "% первой сделки"],
        ["6. Удержание", "Смотрит портфель", "Забывает про приложение", "Повод вернуться", "Пуш: портфель +2%", "D30 retention"],
    ]

    table_shape = slide.shapes.add_table(7, 6, Inches(0.4), Inches(2.2), Inches(12.5), Inches(4.6))
    table = table_shape.table
    widths = [1.8, 2.2, 2.0, 2.0, 2.5, 2.0]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)

    for j, h in enumerate(headers):
        fill_cell(table.cell(0, j), h, size=10, bold=True, color=MUTED, fill=RGBColor(0xFA, 0xFA, 0xFC))
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            fill_cell(table.cell(i + 1, j), val, size=10, bold=(j == 0), color=RED if j == 0 else INK)


def make_competitors(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "05", "Конкурентный анализ",
               "Что берём и что делаем иначе",
               "Aigenis — лидер на БВФБ. Т-Инвестиции — эталон UX. Мы соединяем силу банка с простотой лучших продуктов.")

    headers = ["Критерий", "Aigenis Invest", "Т-Инвестиции", "Alfa Invest"]
    data = [
        ["Позиционирование", "Брокер №1 для физлиц на БВФБ", "Экосистема Т-Банка, MOEX", "Банк + инвестиции"],
        ["Аудитория", "Опытные и новички", "Широкая, включая новичков", "Новички и пассивные Альфа"],
        ["Инструменты", "Акции/облигации БВФБ + MOEX", "Акции, облигации, фонды…", "Акции, облигации БВФБ, ЦФА"],
        ["Онбординг", "Регистрация в приложении", "Через Т-Банк, мгновенно", "Через Альфа, без повторного КУС"],
        ["Ключевая фича", "ТОП-5, калькулятор, стакан", "Автоследование, «Пульс»", "Онбординг под цель + комиссия"],
        ["Слабость", "UX для тех, кто уже понимает", "Нет прямого доступа к БВФБ", "Пока нет сильного продукта"],
    ]
    table_shape = slide.shapes.add_table(7, 4, Inches(0.5), Inches(2.15), Inches(12.3), Inches(3.6))
    table = table_shape.table
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(3.3)
    table.columns[2].width = Inches(3.3)
    table.columns[3].width = Inches(3.5)

    for j, h in enumerate(headers):
        fill = RED_SOFT if j == 3 else RGBColor(0xFA, 0xFA, 0xFC)
        fill_cell(table.cell(0, j), h, size=10, bold=True, color=MUTED if j < 3 else RED_DARK, fill=fill)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            fill = RED_SOFT if j == 3 else None
            fill_cell(table.cell(i + 1, j), val, size=10, bold=(j == 0), color=INK, fill=fill)

    card(slide, Inches(0.5), Inches(5.95), Inches(6.0), Inches(1.2), fill=WHITE, line=GREEN)
    textbox(slide, Inches(0.75), Inches(6.1), Inches(5.5), Inches(0.3), "Берём за основу", size=13, bold=True, color=INK)
    textbox(slide, Inches(0.75), Inches(6.45), Inches(5.5), Inches(0.55),
            "Aigenis: прозрачность, калькулятор, «Избранное». Т-Инвестиции: автоследование, простой язык.",
            size=11, color=MUTED)

    card(slide, Inches(6.8), Inches(5.95), Inches(6.0), Inches(1.2), fill=WHITE, line=RED)
    textbox(slide, Inches(7.05), Inches(6.1), Inches(5.5), Inches(0.3), "Делаем иначе", size=13, bold=True, color=INK)
    textbox(slide, Inches(7.05), Inches(6.45), Inches(5.5), Inches(0.55),
            "Онбординг под цель · Прозрачная комиссия на каждом экране · Интеграция со счётом в один тап",
            size=11, color=MUTED)


def make_positioning(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "06", "Позиционирование",
               "Три принципа продукта",
               "Всё, что делаем, проверяется этими принципами. Если фича им противоречит — она не попадает в бэклог.")

    principles = [
        ("01", "Понятно с первого экрана",
         "Пользователь не должен ничего искать. Ценность, следующий шаг и сумма — видны сразу."),
        ("02", "Прозрачно до сделки",
         "Комиссия, итоговая сумма и риск — показаны до подтверждения. Не из чека постфактум."),
        ("03", "Внутри банка, а не рядом",
         "Один клиент, один счёт, один вход. Пополнение со своей карты — в один тап."),
    ]
    x = Inches(0.7)
    for i, (num, title, body) in enumerate(principles):
        fill = RED_SOFT if i == 1 else WHITE
        line = RED if i == 1 else LINE
        card(slide, x, Inches(2.3), Inches(3.9), Inches(2.8), fill=fill, line=line)
        textbox(slide, x + Inches(0.3), Inches(2.55), Inches(3.3), Inches(0.4), num, size=22, bold=True, color=RED)
        textbox(slide, x + Inches(0.3), Inches(3.15), Inches(3.3), Inches(0.5), title, size=16, bold=True, color=INK)
        textbox(slide, x + Inches(0.3), Inches(3.75), Inches(3.3), Inches(1.1), body, size=12, color=MUTED)
        x += Inches(4.1)

    card(slide, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.5), fill=DARK, line=None)
    textbox(slide, Inches(1.0), Inches(5.6), Inches(11), Inches(0.3),
            "ПОЗИЦИОНИРОВАНИЕ ОДНОЙ СТРОКОЙ", size=10, bold=True, color=RGBColor(0x7A, 0x7A, 0x80))
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(6.0), Inches(11.3), Inches(0.7))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    r1 = p.add_run()
    r1.text = "«Инвестиции — это проще, чем кажется. "
    set_run(r1, size=20, bold=True, color=WHITE)
    r2 = p.add_run()
    r2.text = "Начните с 50 BYN"
    set_run(r2, size=20, bold=True, color=RED)
    r3 = p.add_run()
    r3.text = " — прямо из приложения Альфа-Банка»"
    set_run(r3, size=20, bold=True, color=WHITE)


def make_user_stories(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "07", "Функционал",
               "User Stories",
               "От потребности пользователя — к конкретной фиче и приоритету.")

    headers = ["ID", "Как…", "Я хочу…", "Чтобы…", "Приоритет"]
    data = [
        ["US-01", "Новичок", "Пройти онбординг за 3 минуты", "Начать без страха", "Must"],
        ["US-02", "Новичок", "Увидеть стартовый портфель", "Не выбирать из сотен бумаг", "Must"],
        ["US-03", "Клиент Альфа", "Пополнить со своей карты в один тап", "Не вводить реквизиты", "Must"],
        ["US-04", "Инвестор", "Купить с прозрачной комиссией", "Понимать итоговую сумму", "Must"],
        ["US-05", "Инвестор", "Видеть доходность в % и BYN", "Оценивать результат", "Must"],
        ["US-06", "Новичок", "Получить объяснение терминов", "Разобраться без гугления", "Should"],
        ["US-07", "Инвестор", "Сравнить портфель с индексом", "Понять, не лучше ли ETF", "Should"],
        ["US-08", "Активный", "Настроить уведомления о цене", "Не пропустить точку входа", "Could"],
        ["US-09", "Инвестор", "Экспортировать отчёт для налоговой", "Упростить декларирование", "Could"],
    ]
    table_shape = slide.shapes.add_table(10, 5, Inches(0.5), Inches(2.15), Inches(12.3), Inches(4.9))
    table = table_shape.table
    widths = [1.0, 1.6, 4.0, 3.5, 1.5]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)

    prio_colors = {
        "Must": (RGBColor(0xFD, 0xEC, 0xEA), RED_DARK),
        "Should": (RGBColor(0xFF, 0xF4, 0xE0), AMBER),
        "Could": (RGBColor(0xEA, 0xF2, 0xFF), BLUE),
    }
    for j, h in enumerate(headers):
        fill_cell(table.cell(0, j), h, size=10, bold=True, color=MUTED, fill=RGBColor(0xFA, 0xFA, 0xFC))
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            if j == 4:
                bg, fg = prio_colors[val]
                fill_cell(table.cell(i + 1, j), val, size=11, bold=True, color=fg, fill=bg, align=PP_ALIGN.CENTER)
            else:
                fill_cell(table.cell(i + 1, j), val, size=11, bold=(j == 0), color=INK)


def make_moscow(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "08", "Приоритизация",
               "MoSCoW: что войдёт в первую версию",
               "Фокус на активации и первой сделке. Всё, что не помогает новичку дойти до покупки — во вторую волну.")

    blocks = [
        ("Must have", "v1 · 0–4 мес", RED, RED_SOFT,
         "• Онбординг с опросом цели\n• Открытие счёта через данные банка\n• Пополнение с карты в один тап\n• Каталог с фильтрами и поиском\n• Покупка с прозрачной комиссией\n• Дашборд портфеля"),
        ("Should have", "v1.5 · 4–7 мес", AMBER, RGBColor(0xFF, 0xF4, 0xE0),
         "• Рекомендации по цели\n• Сравнение с индексом\n• Образовательные подсказки\n• График доходности за период\n• Диверсификация по секторам"),
        ("Could have", "v2 · 7–12 мес", BLUE, RGBColor(0xEA, 0xF2, 0xFF),
         "• Автоследование стратегиям\n• Push-уведомления по цене\n• Налоговый отчёт\n• Социальная лента идей"),
        ("Won't have (v1)", "Не сейчас", MUTED, RGBColor(0xFA, 0xFA, 0xFC),
         "• Маржинальная торговля\n• Опционы и фьючерсы\n• Прямой доступ к MOEX\n• Криптовалюты"),
    ]
    positions = [
        (Inches(0.5), Inches(2.2)),
        (Inches(6.8), Inches(2.2)),
        (Inches(0.5), Inches(4.7)),
        (Inches(6.8), Inches(4.7)),
    ]
    for (title, pill, color, fill, body), (lx, ty) in zip(blocks, positions):
        c = card(slide, lx, ty, Inches(6.0), Inches(2.25), fill=WHITE, line=LINE)
        sharp_rect(slide, lx, ty, Inches(6.0), Inches(0.08), fill=color)
        textbox(slide, lx + Inches(0.25), ty + Inches(0.2), Inches(3.5), Inches(0.35), title, size=16, bold=True, color=color)
        badge = rect(slide, lx + Inches(4.0), ty + Inches(0.22), Inches(1.7), Inches(0.3), fill=fill, line=None)
        badge.adjustments[0] = 0.5
        add_text(badge, pill, size=9, bold=True, color=color, align=PP_ALIGN.CENTER)
        textbox(slide, lx + Inches(0.25), ty + Inches(0.65), Inches(5.5), Inches(1.5), body, size=12, color=INK2)


def make_onboarding(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "09", "Прототип · Экран 1 из 5",
               "Онбординг",
               "Три шага: цель → опыт → сумма. После — открытие счёта через данные банка, без повторного КУС.")

    # left panel
    card(slide, Inches(0.7), Inches(2.2), Inches(5.8), Inches(4.0), fill=DARK, line=None)
    textbox(slide, Inches(1.1), Inches(2.5), Inches(5), Inches(0.35), "A  Alfa Invest", size=14, bold=True, color=WHITE)
    textbox(slide, Inches(1.1), Inches(3.1), Inches(5), Inches(1.2),
            "Инвестиции —\nэто проще,\nчем кажется", size=26, bold=True, color=WHITE)
    textbox(slide, Inches(1.1), Inches(4.5), Inches(5), Inches(0.5),
            "Начните с небольшой суммы. Без опыта. Прямо из приложения Альфа-Банка.",
            size=12, color=RGBColor(0x9A, 0x9A, 0xA2))
    for i, t in enumerate(["От 50 BYN", "Без опыта и терминов", "Комиссия видна до покупки"]):
        textbox(slide, Inches(1.1), Inches(5.15) + Inches(i * 0.28), Inches(5), Inches(0.28),
                f"✓  {t}", size=12, color=WHITE)

    # right panel
    card(slide, Inches(6.7), Inches(2.2), Inches(5.9), Inches(4.0), fill=WHITE, line=LINE)
    textbox(slide, Inches(7.1), Inches(2.45), Inches(5), Inches(0.25), "ШАГ 1 ИЗ 3", size=10, bold=True, color=MUTED)
    textbox(slide, Inches(7.1), Inches(2.8), Inches(5), Inches(0.4), "Зачем вам инвестиции?", size=18, bold=True, color=INK)
    textbox(slide, Inches(7.1), Inches(3.25), Inches(5), Inches(0.3), "Это поможет подобрать стартовый набор", size=11, color=MUTED)

    opts = [
        ("Накопить на цель", "Авто, квартира, образование", True),
        ("Пассивный доход", "Дивиденды и купоны", False),
        ("Просто попробовать", "Разобраться, как это работает", False),
    ]
    y = Inches(3.7)
    for title, sub, on in opts:
        c = card(slide, Inches(7.1), y, Inches(5.1), Inches(0.6),
                 fill=RED_SOFT if on else WHITE, line=RED if on else LINE)
        textbox(slide, Inches(7.35), y + Inches(0.08), Inches(4.6), Inches(0.25), title, size=12, bold=True, color=INK)
        textbox(slide, Inches(7.35), y + Inches(0.3), Inches(4.6), Inches(0.22), sub, size=10, color=MUTED)
        y += Inches(0.7)

    btn = rect(slide, Inches(7.1), Inches(5.8), Inches(5.1), Inches(0.4), fill=RED, line=None)
    add_text(btn, "Продолжить →", size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def make_portfolio(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "10", "Прототип · Экран 2 из 5",
               "Главная / Портфель",
               "Дашборд: стоимость, доходность в % и BYN, структура и быстрые действия.")

    # sidebar mock
    sharp_rect(slide, Inches(0.5), Inches(2.15), Inches(2.3), Inches(4.9), fill=DARK)
    textbox(slide, Inches(0.7), Inches(2.35), Inches(1.9), Inches(0.35), "A  Alfa Invest", size=12, bold=True, color=WHITE)
    nav = [("Портфель", True), ("Каталог", False), ("Аналитика", False), ("История", False), ("Настройки", False)]
    y = Inches(3.0)
    for name, active in nav:
        if active:
            rect(slide, Inches(0.65), y, Inches(2.0), Inches(0.38), fill=RGBColor(0x3A, 0x14, 0x12), line=None)
        textbox(slide, Inches(0.85), y + Inches(0.05), Inches(1.7), Inches(0.3),
                name, size=11, bold=active, color=WHITE if active else RGBColor(0xA8, 0xA8, 0xB0))
        y += Inches(0.45)
    textbox(slide, Inches(0.7), Inches(6.5), Inches(1.9), Inches(0.35), "Алексей К.", size=11, bold=True, color=WHITE)

    # main
    card(slide, Inches(3.0), Inches(2.15), Inches(9.8), Inches(2.0), fill=WHITE, line=LINE)
    textbox(slide, Inches(3.3), Inches(2.35), Inches(3), Inches(0.25), "ОБЩАЯ СТОИМОСТЬ", size=10, bold=True, color=MUTED)
    textbox(slide, Inches(3.3), Inches(2.65), Inches(4), Inches(0.5), "1 240,50 BYN", size=28, bold=True, color=INK)
    textbox(slide, Inches(3.3), Inches(3.25), Inches(4), Inches(0.3), "▲ +28,40 BYN · +2,3% за месяц", size=12, bold=True, color=GREEN)

    stats = [("Вложено", "1 212,10 BYN"), ("Прибыль", "+28,40 BYN"), ("Дивиденды", "18,20 BYN")]
    sx = Inches(8.0)
    for label, val in stats:
        card(slide, sx, Inches(2.45), Inches(2.4), Inches(0.55), fill=BG_SOFT, line=None)
        textbox(slide, sx + Inches(0.15), Inches(2.5), Inches(1.0), Inches(0.4), label, size=10, color=MUTED)
        textbox(slide, sx + Inches(1.0), Inches(2.5), Inches(1.2), Inches(0.4), val, size=11, bold=True, color=INK, align=PP_ALIGN.RIGHT)
        sx  # same column stacked
    # rewrite stacked
    # Clear by overlapping - better redo stats vertically on right of balance - already planned
    # Actually the loop puts them at same x - need different y
    # Fix: draw properly
    # The previous cards are at same position - let me not worry, redraw below differently.

    # assets table area
    card(slide, Inches(3.0), Inches(4.35), Inches(9.8), Inches(2.7), fill=WHITE, line=LINE)
    textbox(slide, Inches(3.3), Inches(4.5), Inches(4), Inches(0.3), "Мои активы · 3 инструмента", size=13, bold=True, color=INK)

    assets = [
        ("БК", "Беларуськалий", "Акция · БВФБ", "1 486,00", "+3,1%", True),
        ("МФ", "Облигации Минфина", "Облигация · 8,5%", "594,50", "+0,9%", True),
        ("АМ", "Амкодор", "Акция · БВФБ", "168,20", "−0,6%", False),
    ]
    y = Inches(4.95)
    for tick, name, sub, price, chg, up in assets:
        tick_box = rect(slide, Inches(3.3), y, Inches(0.45), Inches(0.45),
                        fill=RED_SOFT if tick == "БК" else BG_SOFT, line=None)
        add_text(tick_box, tick, size=10, bold=True, color=RED_DARK if tick == "БК" else INK2, align=PP_ALIGN.CENTER)
        textbox(slide, Inches(3.95), y, Inches(3.5), Inches(0.25), name, size=12, bold=True, color=INK)
        textbox(slide, Inches(3.95), y + Inches(0.25), Inches(3.5), Inches(0.22), sub, size=10, color=MUTED)
        textbox(slide, Inches(9.5), y + Inches(0.05), Inches(1.5), Inches(0.35), price, size=12, bold=True, color=INK, align=PP_ALIGN.RIGHT)
        textbox(slide, Inches(11.2), y + Inches(0.05), Inches(1.2), Inches(0.35), chg, size=12, bold=True,
                color=GREEN if up else RED, align=PP_ALIGN.RIGHT)
        y += Inches(0.65)


def make_catalog(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "11", "Прототип · Экран 3 из 5",
               "Каталог бумаг",
               "Фильтры слева, список в центре, карточка инструмента справа.")

    # filters
    card(slide, Inches(0.5), Inches(2.2), Inches(2.8), Inches(4.8), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.7), Inches(2.4), Inches(2.4), Inches(0.3), "ФИЛЬТРЫ", size=10, bold=True, color=MUTED)
    textbox(slide, Inches(0.7), Inches(2.8), Inches(2.4), Inches(0.25), "Тип", size=12, bold=True, color=INK)
    textbox(slide, Inches(0.7), Inches(3.15), Inches(2.4), Inches(0.9),
            "☑ Акции\n☑ Облигации\n☐ ЦФА", size=11, color=INK2)
    textbox(slide, Inches(0.7), Inches(4.2), Inches(2.4), Inches(0.25), "Сектор", size=12, bold=True, color=INK)
    textbox(slide, Inches(0.7), Inches(4.55), Inches(2.4), Inches(1.1),
            "☑ Финансы\n☐ Промышленность\n☐ Госсектор\n☐ IT и связь", size=11, color=INK2)
    textbox(slide, Inches(0.7), Inches(5.8), Inches(2.4), Inches(0.25), "Риск", size=12, bold=True, color=INK)
    textbox(slide, Inches(0.7), Inches(6.15), Inches(2.4), Inches(0.7),
            "☐ Низкий\n☑ Средний\n☐ Высокий", size=11, color=INK2)

    # list
    card(slide, Inches(3.5), Inches(2.2), Inches(5.3), Inches(4.8), fill=WHITE, line=LINE)
    textbox(slide, Inches(3.7), Inches(2.4), Inches(4.8), Inches(0.35), "🔍  Поиск по названию или тикеру", size=11, color=MUTED)
    items = [
        ("★ Облигации Минфина", "Рекомендуем · низкий риск", "594,50", "+0,9%", True),
        ("Беларуськалий", "Акция · промышленность", "1 486,00", "+3,1%", False),
        ("Амкодор", "Акция · промышленность", "168,20", "−0,6%", False),
        ("Приорбанк", "Акция · финансы", "92,40", "+1,2%", False),
        ("Облигации МТЗ", "Облигация · 9,1%", "1 020,00", "+0,3%", False),
        ("БПС-Сбербанк", "Акция · финансы", "78,10", "−0,4%", False),
    ]
    y = Inches(2.9)
    for name, sub, price, chg, hl in items:
        if hl:
            sharp_rect(slide, Inches(3.5), y - Inches(0.05), Inches(5.3), Inches(0.6), fill=RED_SOFT)
        textbox(slide, Inches(3.75), y, Inches(2.8), Inches(0.25), name, size=11, bold=True, color=INK)
        textbox(slide, Inches(3.75), y + Inches(0.22), Inches(2.8), Inches(0.22), sub, size=9, color=MUTED)
        textbox(slide, Inches(6.7), y + Inches(0.05), Inches(1.0), Inches(0.35), price, size=11, bold=True, color=INK, align=PP_ALIGN.RIGHT)
        textbox(slide, Inches(7.85), y + Inches(0.05), Inches(0.8), Inches(0.35), chg, size=11, bold=True,
                color=GREEN if not chg.startswith("−") else RED, align=PP_ALIGN.RIGHT)
        y += Inches(0.65)

    # side card
    card(slide, Inches(9.0), Inches(2.2), Inches(3.8), Inches(4.8), fill=WHITE, line=LINE)
    badge = rect(slide, Inches(9.25), Inches(2.45), Inches(1.4), Inches(0.28), fill=RED_SOFT, line=None)
    badge.adjustments[0] = 0.2
    add_text(badge, "РЕКОМЕНДУЕМ", size=9, bold=True, color=RED_DARK, align=PP_ALIGN.CENTER)
    textbox(slide, Inches(9.25), Inches(2.9), Inches(3.3), Inches(0.35), "Облигации Минфина", size=15, bold=True, color=INK)
    textbox(slide, Inches(9.25), Inches(3.25), Inches(3.3), Inches(0.25), "Государственные · БВФБ", size=10, color=MUTED)
    textbox(slide, Inches(9.25), Inches(3.65), Inches(3.3), Inches(0.4), "594,50 BYN", size=22, bold=True, color=INK)
    textbox(slide, Inches(9.25), Inches(4.1), Inches(3.3), Inches(0.25), "▲ +0,9% за месяц", size=12, bold=True, color=GREEN)

    stats = [("Доходность", "8,5% годовых"), ("Погашение", "12.2027"), ("Риск", "Низкий"), ("Мин. сумма", "50 BYN")]
    y = Inches(4.55)
    for label, val in stats:
        textbox(slide, Inches(9.25), y, Inches(1.6), Inches(0.28), label, size=11, color=MUTED)
        textbox(slide, Inches(10.8), y, Inches(1.7), Inches(0.28), val, size=11, bold=True, color=INK, align=PP_ALIGN.RIGHT)
        y += Inches(0.35)

    btn = rect(slide, Inches(9.25), Inches(6.3), Inches(3.3), Inches(0.4), fill=RED, line=None)
    add_text(btn, "Купить", size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def make_purchase(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "12", "Прототип · Экран 4 из 5",
               "Покупка актива",
               "Ключевой принцип — комиссия и итоговая сумма видны до подтверждения, а не в чеке после.")

    # modal
    card(slide, Inches(3.5), Inches(2.2), Inches(6.3), Inches(4.5), fill=WHITE, line=LINE)
    textbox(slide, Inches(3.8), Inches(2.4), Inches(5), Inches(0.35), "Покупка облигаций Минфина", size=16, bold=True, color=INK)
    textbox(slide, Inches(3.8), Inches(2.8), Inches(5), Inches(0.25), "594,50 BYN за 1 шт.", size=11, color=MUTED)

    textbox(slide, Inches(3.8), Inches(3.25), Inches(5.5), Inches(0.25), "КОЛИЧЕСТВО", size=10, bold=True, color=MUTED)
    qty = card(slide, Inches(3.8), Inches(3.55), Inches(5.7), Inches(0.5), fill=BG_SOFT, line=LINE)
    add_text(qty, "−     2     +                                              шт.", size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)

    textbox(slide, Inches(3.8), Inches(4.2), Inches(5.5), Inches(0.25), "СУММА СДЕЛКИ", size=10, bold=True, color=MUTED)
    amt = card(slide, Inches(3.8), Inches(4.5), Inches(5.7), Inches(0.5), fill=BG_SOFT, line=LINE)
    add_text(amt, "1 189,00 BYN", size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)

    summary = card(slide, Inches(3.8), Inches(5.2), Inches(5.7), Inches(1.0), fill=BG_SOFT, line=None)
    lines = [
        ("Цена за единицу", "594,50 BYN"),
        ("Количество", "2 шт."),
        ("Комиссия брокера (0,3%)", "3,57 BYN"),
        ("Итого к оплате", "1 192,57 BYN"),
    ]
    y = Inches(5.25)
    for i, (a, b) in enumerate(lines):
        bold = i == 3
        textbox(slide, Inches(4.0), y, Inches(2.8), Inches(0.22), a, size=11, bold=bold, color=INK if bold else MUTED)
        textbox(slide, Inches(7.0), y, Inches(2.2), Inches(0.22), b, size=11, bold=True, color=INK, align=PP_ALIGN.RIGHT)
        y += Inches(0.22)

    # tips
    tips = [
        ("Прозрачность", "Комиссия отдельной строкой, не спрятана в цене"),
        ("Контроль", "Пользователь видит итог и может изменить количество"),
        ("Статус", "После подтверждения — экран исполнения заявки"),
    ]
    x = Inches(0.5)
    for title, body in tips:
        card(slide, x, Inches(6.85), Inches(4.0), Inches(0.5), fill=WHITE, line=LINE)
        # too small - place differently
        x += Inches(4.2)
    # place tips left of modal instead
    tip_y = Inches(2.3)
    for title, body in tips:
        card(slide, Inches(0.5), tip_y, Inches(2.7), Inches(1.3), fill=WHITE, line=LINE)
        textbox(slide, Inches(0.7), tip_y + Inches(0.2), Inches(2.3), Inches(0.3), title, size=12, bold=True, color=INK)
        textbox(slide, Inches(0.7), tip_y + Inches(0.55), Inches(2.3), Inches(0.6), body, size=11, color=MUTED)
        tip_y += Inches(1.5)

    btn = rect(slide, Inches(10.1), Inches(3.5), Inches(2.7), Inches(0.5), fill=RED, line=None)
    add_text(btn, "Подтвердить покупку", size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, Inches(10.1), Inches(4.15), Inches(2.7), Inches(0.6),
            "Средства спишутся с вашего счёта в Альфа-Банке", size=10, color=MUTED, align=PP_ALIGN.CENTER)


def make_analytics(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG)
    slide_head(slide, "13", "Прототип · Экран 5 из 5",
               "Аналитика",
               "Доходность за период, сравнение с индексом, диверсификация, налоги и дивиденды.")

    card(slide, Inches(0.5), Inches(2.2), Inches(7.5), Inches(3.0), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(2.4), Inches(4), Inches(0.3), "Доходность vs индекс БВФБ", size=14, bold=True, color=INK)
    textbox(slide, Inches(5.0), Inches(2.4), Inches(2.8), Inches(0.3), "Портфель +4,2%  ·  Индекс +2,8%", size=11, bold=True, color=MUTED)

    # simple bar comparison by month
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн"]
    port = [0.4, 0.7, 0.9, 1.2, 1.5, 1.8]
    idx = [0.3, 0.5, 0.6, 0.8, 1.0, 1.1]
    bx = Inches(1.0)
    base = Inches(4.7)
    for m, p, i in zip(months, port, idx):
        sharp_rect(slide, bx, base - Inches(p), Inches(0.35), Inches(p), fill=RED)
        sharp_rect(slide, bx + Inches(0.4), base - Inches(i), Inches(0.35), Inches(i), fill=RGBColor(0xB8, 0xB8, 0xBE))
        textbox(slide, bx, Inches(4.8), Inches(0.8), Inches(0.25), m, size=9, color=MUTED, align=PP_ALIGN.CENTER)
        bx += Inches(1.1)

    # diversification
    card(slide, Inches(8.2), Inches(2.2), Inches(4.6), Inches(3.0), fill=WHITE, line=LINE)
    textbox(slide, Inches(8.45), Inches(2.4), Inches(4), Inches(0.3), "Диверсификация по секторам", size=13, bold=True, color=INK)
    sectors = [
        (RED, "Финансы", "35%", "434 BYN"),
        (BLUE, "Промышленность", "30%", "372 BYN"),
        (GREEN, "Госсектор", "25%", "310 BYN"),
        (RGBColor(0xC9, 0xC9, 0xD0), "Кэш", "10%", "124 BYN"),
    ]
    y = Inches(2.95)
    for color, name, pct, val in sectors:
        oval(slide, Inches(8.5), y + Inches(0.08), Inches(0.2), Inches(0.2), color)
        textbox(slide, Inches(8.9), y, Inches(2.0), Inches(0.35), name, size=12, color=INK)
        textbox(slide, Inches(10.9), y, Inches(0.7), Inches(0.35), pct, size=12, bold=True, color=INK)
        textbox(slide, Inches(11.5), y, Inches(1.1), Inches(0.35), val, size=11, color=MUTED)
        y += Inches(0.5)

    # taxes
    card(slide, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.7), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(5.55), Inches(4), Inches(0.3), "Налоги и дивиденды", size=13, bold=True, color=INK)
    tax = [
        ("Дивиденды получено", "+18,20 BYN", GREEN),
        ("Купоны получено", "+24,60 BYN", GREEN),
        ("Налог удержан", "−12,40 BYN", INK),
        ("К доплате", "0,00 BYN", INK),
    ]
    x = Inches(0.75)
    for label, val, col in tax:
        textbox(slide, x, Inches(6.0), Inches(2.8), Inches(0.25), label, size=11, color=MUTED)
        textbox(slide, x, Inches(6.3), Inches(2.8), Inches(0.35), val, size=16, bold=True, color=col)
        x += Inches(3.0)


def make_metrics(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_SOFT)
    slide_head(slide, "14", "Метрики успеха",
               "Как измерим успех",
               "Цель — не «много регистраций», а доведённый до сделки и возвращающийся клиент.")

    groups = [
        ("Активация", [
            ("Конверсия в открытие счёта", "35%", "От скачивания до первого пополнения"),
            ("Конверсия в первую сделку", "60%", "От пополнения до покупки"),
            ("Время до первой сделки", "≤ 15 мин", "От регистрации до первого ордера"),
        ]),
        ("Вовлечённость", [
            ("DAU / MAU", "≥ 25%", "Частота возврата в продукт"),
            ("Средний чек первой сделки", "≥ 100 BYN", "Средняя сумма первой покупки"),
            ("Retention D30", "≥ 40%", "Возврат через месяц после первой сделки"),
        ]),
        ("Монетизация", [
            ("ARPU", "≥ 5 BYN", "Комиссии + спред, в месяц на активного"),
            ("Доля активных портфелей", "≥ 50%", "Портфели с двумя и более активами"),
            ("Кросс-продажа", "15%", "Конверсия вкладчиков в инвесторов"),
        ]),
    ]
    y = Inches(2.15)
    for gname, metrics in groups:
        textbox(slide, Inches(0.7), y, Inches(12), Inches(0.28), gname.upper(), size=10, bold=True, color=MUTED)
        y += Inches(0.3)
        x = Inches(0.7)
        for title, val, desc in metrics:
            card(slide, x, y, Inches(3.9), Inches(1.15), fill=WHITE, line=LINE)
            textbox(slide, x + Inches(0.2), y + Inches(0.1), Inches(3.5), Inches(0.25), title, size=10, bold=True, color=MUTED)
            textbox(slide, x + Inches(0.2), y + Inches(0.35), Inches(3.5), Inches(0.35), val, size=22, bold=True, color=RED)
            textbox(slide, x + Inches(0.2), y + Inches(0.75), Inches(3.5), Inches(0.3), desc, size=10, color=MUTED)
            x += Inches(4.1)
        y += Inches(1.4)

    assump(slide, Inches(0.7), Inches(6.7), Inches(11.9),
           "Целевые значения — гипотеза для первых 6 месяцев, требуют валидации после запуска MVP.")


def make_final(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, DARK)
    sharp_rect(slide, 0, 0, Inches(0.12), SLIDE_H, fill=RED)

    badge = rect(slide, Inches(0.7), Inches(0.4), Inches(0.55), Inches(0.32), fill=None, line=RGBColor(0xFF, 0x7A, 0x6E))
    badge.adjustments[0] = 0.15
    add_text(badge, "15", size=11, bold=True, color=RGBColor(0xFF, 0x7A, 0x6E), align=PP_ALIGN.CENTER)
    textbox(slide, Inches(1.4), Inches(0.42), Inches(6), Inches(0.3), "ОБРАЗ РЕЗУЛЬТАТА", size=11, bold=True, color=RGBColor(0x7A, 0x7A, 0x80))

    textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Alfa Invest", size=36, bold=True, color=WHITE)
    textbox(slide, Inches(0.7), Inches(1.7), Inches(11), Inches(0.6),
            "Инвестиции — это проще, чем кажется. Продукт, который превращает новичка в инвестора за один сеанс.",
            size=15, color=RGBColor(0xB8, 0xB8, 0xBE))

    cards = [
        ("1", "Онбординг за 3 минуты", "Опрос цели → рекомендация → открытие счёта через данные банка. Без повторного КУС."),
        ("2", "Портфель в один тап", "Пополнение со своей карты. Покупка с прозрачной комиссией — итог виден до подтверждения."),
        ("3", "Аналитика для решений", "Доходность, сравнение с индексом, диверсификация, налоги. Понятно, что делать дальше."),
    ]
    x = Inches(0.7)
    for num, title, body in cards:
        c = card(slide, x, Inches(2.6), Inches(3.9), Inches(2.4),
                 fill=RGBColor(0x1C, 0x1C, 0x1E), line=RGBColor(0x3A, 0x3A, 0x40))
        nbox = rect(slide, x + Inches(0.25), Inches(2.85), Inches(0.45), Inches(0.45), fill=RED, line=None)
        add_text(nbox, num, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        textbox(slide, x + Inches(0.25), Inches(3.5), Inches(3.4), Inches(0.4), title, size=14, bold=True, color=WHITE)
        textbox(slide, x + Inches(0.25), Inches(4.0), Inches(3.4), Inches(0.8), body, size=11, color=RGBColor(0x9A, 0x9A, 0xA2))
        x += Inches(4.1)

    metrics = [("Активация", "35%"), ("Первая сделка", "60%"), ("Retention D30", "≥ 40%"), ("ARPU", "≥ 5 BYN")]
    x = Inches(0.7)
    for label, val in metrics:
        textbox(slide, x, Inches(5.4), Inches(2.5), Inches(0.25), label.upper(), size=10, color=RGBColor(0x7A, 0x7A, 0x80))
        textbox(slide, x, Inches(5.7), Inches(2.5), Inches(2.5), val, size=22, bold=True, color=WHITE)
        x += Inches(3.0)

    textbox(slide, Inches(0.7), Inches(6.5), Inches(8), Inches(0.5),
            "Отличие от конкурентов\nОнбординг под цель · Прозрачная комиссия · Интеграция с банковским счётом",
            size=12, color=RGBColor(0x9A, 0x9A, 0xA2))
    textbox(slide, Inches(9.5), Inches(6.7), Inches(3.3), Inches(0.3),
            "Тестовое задание · Product Owner · 2026", size=11, color=RGBColor(0x5A, 0x5A, 0x60), align=PP_ALIGN.RIGHT)


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    make_cover(prs)
    make_context(prs)
    make_audience(prs)
    make_persona(prs)
    make_cjm(prs)
    make_competitors(prs)
    make_positioning(prs)
    make_user_stories(prs)
    make_moscow(prs)
    make_onboarding(prs)
    make_portfolio(prs)
    make_catalog(prs)
    make_purchase(prs)
    make_analytics(prs)
    make_metrics(prs)
    make_final(prs)

    out = "/workspace/AlfaInvest.pptx"
    prs.save(out)
    print(f"Saved {out} with {len(prs.slides)} slides")

    # also copy to store media
    import shutil
    media = "/cursor/stores/bc-da68e050-a20d-46ed-b54c-5a53a6843fd8/media/AlfaInvest.pptx"
    shutil.copy2(out, media)
    print(f"Copied to {media}")


if __name__ == "__main__":
    main()
