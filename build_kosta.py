# -*- coding: utf-8 -*-
"""Презентация «Коста Хетагуров — поэт и художник».

Минимум текста, максимум картин и фотографий, БЕЗ zoom-эффектов,
с разнообразными переходами-«перелистываниями» между слайдами.
Совместимо с PowerPoint 2010.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from build_deck import (Timeline, attach, new_slide, text, overlay,
                        slide_number, TRANSITIONS, SLIDE_W, SLIDE_H, WHITE)

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets_kosta")

IVORY = RGBColor(0xF2, 0xEA, 0xD9)
INK = RGBColor(0x23, 0x26, 0x2B)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
RED = RGBColor(0xA6, 0x3A, 0x2B)
DARK = RGBColor(0x17, 0x13, 0x10)
GREY = RGBColor(0x8A, 0x85, 0x78)

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Дополнительные «перелистывающие» переходы (OOXML 2006)
TRANSITIONS.update({
    "push_r": '<p:transition spd="med" xmlns:p="%s"><p:push dir="r"/></p:transition>' % P_NS,
    "wipe_d": '<p:transition spd="med" xmlns:p="%s"><p:wipe dir="d"/></p:transition>' % P_NS,
    "wipe_l": '<p:transition spd="med" xmlns:p="%s"><p:wipe dir="l"/></p:transition>' % P_NS,
    "cover_r": '<p:transition spd="med" xmlns:p="%s"><p:cover dir="r"/></p:transition>' % P_NS,
    "comb": '<p:transition spd="med" xmlns:p="%s"><p:comb dir="horz"/></p:transition>' % P_NS,
    "strips": '<p:transition spd="med" xmlns:p="%s"><p:strips dir="ru"/></p:transition>' % P_NS,
    "crawl": '<p:transition spd="med" xmlns:p="%s"><p:crawl dir="l"/></p:transition>' % P_NS,
    "randombar": '<p:transition spd="med" xmlns:p="%s"><p:randomBar dir="vert"/></p:transition>' % P_NS,
    "wedge": '<p:transition spd="med" xmlns:p="%s"><p:wedge/></p:transition>' % P_NS,
})

TOTAL = 12


def pic(slide, name, x, y, w=None, h=None, frame=None, frame_w=Pt(2.5)):
    path = os.path.join(ASSETS, name)
    if w and h:
        p = slide.shapes.add_picture(path, x, y, w, h)
    elif h:
        p = slide.shapes.add_picture(path, x, y, height=h)
    else:
        p = slide.shapes.add_picture(path, x, y, width=w)
    if frame:
        p.line.color.rgb = frame
        p.line.width = frame_w
    p.shadow.inherit = False
    return p


def dark_bg(slide):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    shp.fill.solid(); shp.fill.fore_color.rgb = DARK
    shp.line.fill.background(); shp.shadow.inherit = False
    return shp


def bg_cover(slide, name):
    """Фон во весь слайд без искажения пропорций (cover)."""
    from PIL import Image
    with Image.open(os.path.join(ASSETS, name)) as im:
        iw, ih = im.size
    if iw * SLIDE_H > ih * SLIDE_W:
        h = SLIDE_H; w = int(SLIDE_H * iw / ih)
    else:
        w = SLIDE_W; h = int(SLIDE_W * ih / iw)
    p = slide.shapes.add_picture(os.path.join(ASSETS, name),
                                 Emu(int((SLIDE_W - w) / 2)),
                                 Emu(int((SLIDE_H - h) / 2)),
                                 Emu(w), Emu(h))
    p.shadow.inherit = False
    return p


def caption_bar(slide, lines, num=None):
    bar = overlay(slide, 0, Emu(SLIDE_H - Inches(1.05)), SLIDE_W, Inches(1.05),
                  alpha=62, rounded=False)
    paras = [{"text": lines[0], "size": 20, "bold": True, "color": WHITE,
              "align": PP_ALIGN.CENTER}]
    if len(lines) > 1:
        paras.append({"text": lines[1], "size": 14, "color": RGBColor(0xD8, 0xCE, 0xBE),
                      "align": PP_ALIGN.CENTER, "space_before": 4})
    box = text(slide, Inches(0.6), Emu(SLIDE_H - Inches(1.0)),
               Inches(12.13), Inches(0.95), paras, anchor=MSO_ANCHOR.MIDDLE)
    return bar, box


# --------------------------------------------------------------------------
def k01_title(prs):
    s = new_slide(prs)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = IVORY
    bg.line.fill.background(); bg.shadow.inherit = False
    img = pic(s, "hist_1.jpg", Emu(SLIDE_W - Inches(7.6)), Inches(1.55),
              w=Inches(7.2), frame=GOLD)
    name = text(s, Inches(0.7), Inches(1.6), Inches(6.0), Inches(2.2),
                [{"text": "Коста", "size": 66, "bold": True, "color": INK},
                 {"text": "Хетагуров", "size": 66, "bold": True, "color": INK}])
    line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.78), Inches(4.05),
                              Inches(2.2), Pt(4))
    line.fill.solid(); line.fill.fore_color.rgb = RED; line.line.fill.background()
    line.shadow.inherit = False
    years = text(s, Inches(0.7), Inches(4.3), Inches(6.0), Inches(0.6),
                 [{"text": "1859 — 1906", "size": 22, "bold": True, "color": RED}])
    role = text(s, Inches(0.7), Inches(5.0), Inches(6.2), Inches(1.4),
                [{"text": "поэт • художник • просветитель", "size": 17, "color": INK},
                 {"text": "основоположник осетинской литературы", "size": 14,
                  "color": GREY, "space_before": 4}])
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 300, "withEffect", 0),
              tl.wipe(img.shape_id, "right", 900, delay=100),
              tl.fly(name.shape_id, "from_left", 700, delay=300),
              tl.wipe(line.shape_id, "right", 500, delay=900),
              tl.fade(years.shape_id, 600, delay=1100),
              tl.fade(role.shape_id, 700, delay=1400)])
    attach(s, tl, "fade")
    return s


def k02_photo(prs):
    s = new_slide(prs)
    bg = dark_bg(s)
    photo = pic(s, "kosta_photo.jpg", Inches(1.3), Inches(1.1), h=Inches(5.3),
                frame=GOLD, frame_w=Pt(3))
    head = text(s, Inches(6.0), Inches(1.7), Inches(6.6), Inches(1.6),
                [{"text": "Коста Леванович", "size": 34, "bold": True, "color": WHITE},
                 {"text": "Хетагуров", "size": 34, "bold": True, "color": WHITE}])
    sub = text(s, Inches(6.0), Inches(3.4), Inches(6.6), Inches(2.2),
               [{"text": "осетинский поэт и художник,", "size": 17, "color": RGBColor(0xD8, 0xCE, 0xBE)},
                {"text": "основоположник осетинского", "size": 17, "color": RGBColor(0xD8, 0xCE, 0xBE)},
                {"text": "литературного языка", "size": 17, "color": RGBColor(0xD8, 0xCE, 0xBE)}])
    slide_number(s, 2, TOTAL)
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 300, "withEffect", 0),
              tl.fade(photo.shape_id, 900, delay=150),
              tl.fly(head.shape_id, "from_right", 700, delay=500),
              tl.fade(sub.shape_id, 800, delay=1100)])
    attach(s, tl, "push")
    return s


def k03_nar(prs):
    s = new_slide(prs)
    bg = bg_cover(s, "port_1.jpg")
    bar, box = caption_bar(s, ["Аул Нар — колыбель поэта",
                               "15 октября 1859, Мамисонское ущелье"])
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 400, "withEffect", 0),
              tl.fade(bar.shape_id, 500, delay=300),
              tl.fly(box.shape_id, "from_bottom", 700, delay=500)])
    attach(s, tl, "wipe")
    return s


def k04_study(prs):
    s = new_slide(prs)
    bg = dark_bg(s)
    draw = pic(s, "foto_2.jpg", Inches(7.2), Inches(1.5), w=Inches(5.4),
               frame=GOLD)
    head = text(s, Inches(0.9), Inches(1.5), Inches(6.0), Inches(0.9),
                [{"text": "Дорога к искусству", "size": 34, "bold": True, "color": WHITE}])
    items = text(s, Inches(0.95), Inches(2.7), Inches(6.0), Inches(3.4),
                 [{"text": "Владикавказская прогимназия", "size": 19, "color": WHITE,
                   "space_after": 14},
                  {"text": "Ставропольская гимназия", "size": 19, "color": WHITE,
                   "space_after": 14},
                  {"text": "Петербургская академия художеств, 1881", "size": 19,
                   "color": GOLD}])
    slide_number(s, 4, TOTAL)
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 300, "withEffect", 0),
              tl.fly(head.shape_id, "from_left", 700, delay=100),
              tl.wipe(draw.shape_id, "left", 900, delay=400),
              tl.fade(items.shape_id, 900, delay=900)])
    attach(s, tl, "cover")
    return s


def k05_fandyr(prs):
    s = new_slide(prs)
    bg = dark_bg(s)
    acc = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.9), Inches(2.0),
                             Inches(3.5), Pt(4))
    acc.fill.solid(); acc.fill.fore_color.rgb = GOLD; acc.line.fill.background()
    acc.shadow.inherit = False
    title = text(s, Inches(1.0), Inches(2.3), Inches(11.33), Inches(1.6),
                 [{"text": "«Ирон фæндыр»", "size": 58, "bold": True,
                   "color": GOLD, "align": PP_ALIGN.CENTER}])
    sub = text(s, Inches(1.5), Inches(4.0), Inches(10.33), Inches(0.7),
               [{"text": "1899 — первый сборник осетинской поэзии", "size": 20,
                 "color": WHITE, "align": PP_ALIGN.CENTER}])
    fact = text(s, Inches(1.5), Inches(4.8), Inches(10.33), Inches(0.7),
                [{"text": "стихотворение «Завещание» переведено на 108 языков",
                  "size": 15, "color": GREY, "align": PP_ALIGN.CENTER}])
    slide_number(s, 5, TOTAL)
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 300, "withEffect", 0),
              tl.wipe(acc.shape_id, "right", 600, delay=200),
              tl.wipe(title.shape_id, "down", 900, delay=400),
              tl.fade(sub.shape_id, 700, delay=1200),
              tl.fade(fact.shape_id, 800, delay=1700)])
    attach(s, tl, "split")
    return s


def gallery(prs, num, img_name, w, h, caption, transition, note=None):
    s = new_slide(prs)
    bg = dark_bg(s)
    x = Emu(int((SLIDE_W - w) / 2))
    y = Inches(0.55)
    p = pic(s, img_name, x, y, w=Emu(int(w)), h=Emu(int(h)), frame=GOLD,
            frame_w=Pt(3))
    lines = [caption]
    if note:
        lines.append(note)
    bar, box = caption_bar(s, lines)
    slide_number(s, num, TOTAL)
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 300, "withEffect", 0),
              tl.wipe(p.shape_id, "right", 1000, delay=150),
              tl.fade(bar.shape_id, 500, delay=800),
              tl.fade(box.shape_id, 700, delay=1000)])
    attach(s, tl, transition)
    return s


def k10_monument(prs):
    s = new_slide(prs)
    bg = bg_cover(s, "monument_bronze.jpg")
    bar, box = caption_bar(s, ["Владикавказ — памятник Коста",
                               "центральное место города, у гор и Терека"])
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 400, "withEffect", 0),
              tl.fade(bar.shape_id, 500, delay=300),
              tl.fly(box.shape_id, "from_bottom", 700, delay=500)])
    attach(s, tl, "push_r")
    return s


def k11_university(prs):
    s = new_slide(prs)
    bg = bg_cover(s, "hist_2.jpg")
    bar, box = caption_bar(s, ["Северо-Осетинский университет имени К. Л. Хетагурова",
                               "имя Коста носят улицы, театры и сёла по всему Кавказу"])
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 400, "withEffect", 0),
              tl.fade(bar.shape_id, 500, delay=300),
              tl.fly(box.shape_id, "from_bottom", 700, delay=500)])
    attach(s, tl, "comb")
    return s


def k12_final(prs):
    s = new_slide(prs)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = IVORY
    bg.line.fill.background(); bg.shadow.inherit = False
    img = pic(s, "draw_1.jpg", Inches(1.4), Inches(2.2), h=Inches(3.4),
              frame=RED)
    quote = text(s, Inches(5.6), Inches(2.4), Inches(6.8), Inches(2.0),
                 [{"text": "«Весь мир —", "size": 44, "bold": True, "color": RED},
                  {"text": "мой храм»", "size": 44, "bold": True, "color": RED}])
    sub = text(s, Inches(5.6), Inches(4.5), Inches(6.8), Inches(0.9),
               [{"text": "Коста Хетагуров, 1859—1906", "size": 16, "color": INK}])
    tl = Timeline()
    tl.group([tl.fade(bg.shape_id, 300, "withEffect", 0),
              tl.fade(img.shape_id, 800, delay=200),
              tl.wipe(quote.shape_id, "right", 900, delay=600),
              tl.fade(sub.shape_id, 700, delay=1400)])
    attach(s, tl, "fade")
    return s


def main():
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    prs.core_properties.title = "Коста Хетагуров — поэт и художник"
    prs.core_properties.author = "Arena.ai"

    k01_title(prs)
    k02_photo(prs)
    k03_nar(prs)
    k04_study(prs)
    k05_fandyr(prs)
    from PIL import Image
    specs = [
        (6, "paint_goryanka_crop.jpg", "«Горянка, идущая за водой»", "push_r",
         "живопись Коста Хетагурова"),
        (7, "paint_popova_crop.jpg", "Портрет А. Я. Поповой", "wipe_l",
         "той, что любила и была любима"),
        (8, "paint_master_crop.jpg", "«Мастерская художника»", "strips",
         "Коста за работой"),
        (9, "draw_4.jpg", "Образ Коста в графике", "wedge", None),
    ]
    for num, name, cap, tr, note in specs:
        with Image.open(os.path.join(ASSETS, name)) as im:
            iw, ih = im.size
        max_h = Inches(5.1)
        max_w = Inches(7.5)
        h = min(max_h, Emu(int(max_w * ih / iw)))
        w = Emu(int(h * iw / ih))
        gallery(prs, num, name, w, h, cap, tr, note)
    k10_monument(prs)
    k11_university(prs)
    k12_final(prs)

    out = os.path.join(HERE, "Kosta_Hetagurov.pptx")
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main()
