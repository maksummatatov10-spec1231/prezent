# -*- coding: utf-8 -*-
"""Генератор презентации «Космос: путешествие по Солнечной системе».

Совместимо с Microsoft PowerPoint 2010: используются только элементы
OOXML 2006 (p:timing, p:transition) и эффекты, проверенные в реальном
PowerPoint (пресеты Appear/Fly/Zoom/Wipe/Fade/Wheel/Spin/Grow + motion-твины
ppt_x/ppt_y/ppt_w/ppt_h).
"""
import os
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

# Широкоэкранный слайд 16:9
SLIDE_W, SLIDE_H = 12192000, 6858000

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xFF, 0xB3, 0x47)
CYAN = RGBColor(0x7F, 0xD4, 0xFF)
DIM = RGBColor(0xCF, 0xD8, 0xE3)

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


# --------------------------------------------------------------------------
# Движок таймлайна (анимации)
# --------------------------------------------------------------------------
class Timeline:
    """Собирает кликовые группы анимаций слайда и строит блок <p:timing>."""

    def __init__(self):
        self._id = 2          # 1 = tmRoot, 2 = mainSeq
        self.groups = []      # список групп: каждая — список строк (rows)

    def nid(self):
        self._id += 1
        return self._id

    def group(self, rows):
        self.groups.append(rows)

    # -- служебные обёртки -------------------------------------------------
    def _row(self, node, delay, preset, inner, dur):
        cid = self.nid()
        d = ' dur="%d"' % dur if dur else ""
        return ('<p:par><p:cTn id="%d"%s fill="hold" nodeType="%s" grpId="0" %s>'
                '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                '<p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
                % (cid, d, node, preset, inner))

    def _mid(self, row_xml, delay):
        return ('<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="%d"/>'
                '</p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
                % (self.nid(), delay, row_xml))

    # -- примитивы поведения ------------------------------------------------
    def set_vis(self, spid, val="visible", delay=0):
        return ('<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst></p:cTn>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl><p:attrNameLst>'
                '<p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                '</p:cBhvr><p:to><p:strVal val="%s"/></p:to></p:set>'
                % (self.nid(), delay, spid, val))

    def anim_eff(self, spid, dur, trans, filt):
        return ('<p:animEffect transition="%s" filter="%s"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '</p:cBhvr></p:animEffect>' % (trans, filt, self.nid(), dur, spid))

    def anim_axis(self, spid, attr, frm, to, dur):
        def val(v):
            if isinstance(v, (int, float)):
                return '<p:fltVal val="%s"/>' % v
            return '<p:strVal val="%s"/>' % v
        return ('<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
                '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/>'
                '</p:tgtEl><p:attrNameLst><p:attrName>%s</p:attrName>'
                '</p:attrNameLst></p:cBhvr><p:tavLst>'
                '<p:tav tm="0"><p:val>%s</p:val></p:tav>'
                '<p:tav tm="100000"><p:val>%s</p:val></p:tav>'
                '</p:tavLst></p:anim>' % (self.nid(), dur, spid, attr, val(frm), val(to)))

    def anim_rot(self, spid, dur, by_deg):
        return ('<p:animRot by="%d"><p:cBhvr><p:cTn id="%d" dur="%d" fill="hold"/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl><p:attrNameLst>'
                '<p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>'
                % (by_deg * 60000, self.nid(), dur, spid))

    # -- эффекты (row-строки) ------------------------------------------------
    def appear(self, spid, node="afterEffect", delay=0):
        return self._mid(self._row(node, 0, 'presetID="1" presetClass="entr" presetSubtype="0"',
                                   self.set_vis(spid), 1), delay)

    def fade(self, spid, dur=600, node="afterEffect", delay=0):
        inner = self.set_vis(spid) + self.anim_eff(spid, dur, "in", "fade")
        return self._mid(self._row(node, 0, 'presetID="9" presetClass="entr" presetSubtype="0"',
                                   inner, dur), delay)

    def dissolve(self, spid, dur=700, node="afterEffect", delay=0):
        inner = self.set_vis(spid) + self.anim_eff(spid, dur, "in", "dissolve")
        return self._mid(self._row(node, 0, 'presetID="10" presetClass="entr" presetSubtype="0"',
                                   inner, dur), delay)

    def wipe(self, spid, direction="right", dur=600, node="afterEffect", delay=0):
        sub = {"up": 1, "right": 2, "down": 4, "left": 8}[direction]
        inner = self.set_vis(spid) + self.anim_eff(spid, dur, "in", "wipe(%s)" % direction)
        return self._mid(self._row(node, 0, 'presetID="22" presetClass="entr" presetSubtype="%d"' % sub,
                                   inner, dur), delay)

    def wheel(self, spid, spokes=1, dur=700, node="afterEffect", delay=0):
        inner = self.set_vis(spid) + self.anim_eff(spid, dur, "in", "wheel(%d)" % spokes)
        return self._mid(self._row(node, 0, 'presetID="21" presetClass="entr" presetSubtype="%d"' % spokes,
                                   inner, dur), delay)

    def circle_in(self, spid, dur=700, node="afterEffect", delay=0):
        inner = self.set_vis(spid) + self.anim_eff(spid, dur, "in", "circle(in)")
        return self._mid(self._row(node, 0, 'presetID="18" presetClass="entr" presetSubtype="12"',
                                   inner, dur), delay)

    def fly(self, spid, direction="from_bottom", dur=600, node="afterEffect", delay=0):
        # direction: from_bottom / from_top / from_left / from_right
        if direction == "from_bottom":
            sub, ax = 4, ("ppt_y", "1+#ppt_h/2")
        elif direction == "from_top":
            sub, ax = 1, ("ppt_y", "0-#ppt_h/2")
        elif direction == "from_right":
            sub, ax = 2, ("ppt_x", "1+#ppt_w/2")
        else:
            sub, ax = 8, ("ppt_x", "0-#ppt_w/2")
        anims = ""
        for attr in ("ppt_x", "ppt_y"):
            frm = ax[1] if attr == ax[0] else "#%s" % attr
            anims += self.anim_axis(spid, attr, frm, "#%s" % attr, dur)
        inner = self.set_vis(spid) + anims
        return self._mid(self._row(node, 0, 'presetID="2" presetClass="entr" presetSubtype="%d"' % sub,
                                   inner, dur), delay)

    def zoom(self, spid, dur=700, node="afterEffect", delay=0):
        inner = (self.set_vis(spid)
                 + self.anim_axis(spid, "ppt_w", 0, "#ppt_w", dur)
                 + self.anim_axis(spid, "ppt_h", 0, "#ppt_h", dur)
                 + self.anim_axis(spid, "ppt_x", 0.5, "#ppt_x", dur)
                 + self.anim_axis(spid, "ppt_y", 0.5, "#ppt_y", dur))
        return self._mid(self._row(node, 0, 'presetID="23" presetClass="entr" presetSubtype="528"',
                                   inner, dur), delay)

    def spin(self, spid, deg=360, dur=1500, node="afterEffect", delay=0):
        return self._mid(self._row(node, 0, 'presetID="8" presetClass="emph" presetSubtype="0"',
                                   self.anim_rot(spid, dur, deg), dur), delay)

    def kenburns(self, spid, dur=10000, delay=0):
        inner = (self.anim_axis(spid, "ppt_w", "#ppt_w", "7/6*#ppt_w", dur)
                 + self.anim_axis(spid, "ppt_h", "#ppt_h", "7/6*#ppt_h", dur)
                 + self.anim_axis(spid, "ppt_x", "#ppt_x", "#ppt_x-#ppt_w/12", dur)
                 + self.anim_axis(spid, "ppt_y", "#ppt_y", "#ppt_y-#ppt_h/12", dur))
        return self._mid(self._row("withEffect", 0, '', inner, dur), delay)

    def exit_fade(self, spid, dur=500, node="afterEffect", delay=0):
        inner = (self.anim_eff(spid, dur, "out", "fade")
                 + self.set_vis(spid, "hidden", max(dur - 1, 0)))
        return self._mid(self._row(node, 0, 'presetID="10" presetClass="exit" presetSubtype="0"',
                                   inner, dur), delay)

    # -- сборка ---------------------------------------------------------------
    def xml(self):
        groups_xml = ""
        for g in self.groups:
            groups_xml += ('<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                           '</p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
                           % (self.nid(), "".join(g)))
        return ('<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
                'restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" '
                'nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
                '<p:childTnLst>%s</p:childTnLst></p:cTn><p:prevCondLst>'
                '<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
                '</p:prevCondLst><p:nextCondLst>'
                '<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
                '</p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
                % (P_NS, groups_xml))


TRANSITIONS = {
    "fade": '<p:transition spd="med" xmlns:p="%s"><p:fade/></p:transition>' % P_NS,
    "push": '<p:transition spd="med" xmlns:p="%s"><p:push dir="l"/></p:transition>' % P_NS,
    "wipe": '<p:transition spd="med" xmlns:p="%s"><p:wipe dir="r"/></p:transition>' % P_NS,
    "wheel": '<p:transition spd="med" xmlns:p="%s"><p:wheel spokes="4"/></p:transition>' % P_NS,
    "cover": '<p:transition spd="med" xmlns:p="%s"><p:cover dir="l"/></p:transition>' % P_NS,
    "split": '<p:transition spd="med" xmlns:p="%s"><p:split orient="horz"/></p:transition>' % P_NS,
    "zoom": '<p:transition spd="med" xmlns:p="%s"><p:zoom/></p:transition>' % P_NS,
}


def attach(slide, timeline, transition):
    sld = slide._element
    trans = etree.fromstring(TRANSITIONS[transition])
    timing = etree.fromstring(timeline.xml())
    sld.append(trans)
    sld.append(timing)


# --------------------------------------------------------------------------
# Конструктор слайдов
# --------------------------------------------------------------------------
def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def bg_picture(slide, name):
    return slide.shapes.add_picture(os.path.join(ASSETS, name), 0, 0, SLIDE_W, SLIDE_H)


def overlay(slide, x, y, w, h, alpha=45, rounded=True, color=RGBColor(0x05, 0x08, 0x14)):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    _set_alpha(shp, alpha)
    return shp


def _set_alpha(shape, alpha_pct):
    """Прозрачность заливки: alpha_pct = непрозрачность в %."""
    sp_pr = shape._element.spPr
    fill = sp_pr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
    clr = fill[0]
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    alpha = etree.SubElement(clr, '{%s}alpha' % a_ns)
    alpha.set('val', str(int(alpha_pct * 1000)))


def text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if p.get("space_after") is not None:
            para.space_after = Pt(p["space_after"])
        if p.get("space_before") is not None:
            para.space_before = Pt(p["space_before"])
        run = para.add_run()
        run.text = p["text"]
        f = run.font
        f.name = p.get("font", "Calibri")
        f.size = Pt(p.get("size", 18))
        f.bold = p.get("bold", False)
        f.italic = p.get("italic", False)
        f.color.rgb = p.get("color", WHITE)
    return box


def slide_number(slide, n, total):
    text(slide, Emu(SLIDE_W - Inches(1.2)), Emu(SLIDE_H - Inches(0.5)),
         Inches(1.0), Inches(0.35),
         [{"text": "%02d / %02d" % (n, total), "size": 11, "color": DIM,
           "align": PP_ALIGN.RIGHT}])


# --------------------------------------------------------------------------
# Контент
# --------------------------------------------------------------------------
TOTAL = 11


def s01_title(prs):
    s = new_slide(prs)
    bg = bg_picture(s, "cover.jpg")
    ov = overlay(s, Inches(4.4), Inches(2.1), Inches(8.5), Inches(3.4), alpha=38)
    line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.35), Inches(2.55),
                              Inches(2.6), Pt(5))
    line.fill.solid(); line.fill.fore_color.rgb = ORANGE; line.line.fill.background()
    line.shadow.inherit = False
    title = text(s, Inches(4.6), Inches(2.7), Inches(8.1), Inches(1.5),
                 [{"text": "КОСМОС", "size": 60, "bold": True, "align": PP_ALIGN.CENTER}])
    sub = text(s, Inches(4.6), Inches(4.15), Inches(8.1), Inches(0.9),
               [{"text": "Путешествие по Солнечной системе", "size": 24,
                 "color": CYAN, "align": PP_ALIGN.CENTER}])
    foot = text(s, Inches(4.6), Inches(4.95), Inches(8.1), Inches(0.5),
                [{"text": "интерактивная презентация • 2026", "size": 13,
                  "color": DIM, "align": PP_ALIGN.CENTER}])
    tl = Timeline()
    tl.group([tl.kenburns(bg.shape_id, 12000),
              tl.zoom(line.shape_id, 700, delay=200),
              tl.zoom(title.shape_id, 800, delay=300),
              tl.fly(sub.shape_id, "from_bottom", 700, delay=900),
              tl.fade(foot.shape_id, 800, delay=1400),
              tl.fade(ov.shape_id, 500, "withEffect", 0)])
    attach(s, tl, "fade")
    return s


def s02_toc(prs):
    s = new_slide(prs)
    bg = bg_picture(s, "section_nebula.jpg")
    ov = overlay(s, Inches(0.7), Inches(0.7), Inches(6.9), Inches(6.1), alpha=45)
    title = text(s, Inches(1.1), Inches(1.0), Inches(6.1), Inches(0.9),
                 [{"text": "Содержание", "size": 40, "bold": True}])
    items = ["Солнце — наша звезда", "Планеты земной группы",
             "Гиганты: Юпитер и Сатурн", "Ледяные миры на краю",
             "Удивительные факты о системе"]
    rows = []
    for i, it in enumerate(items):
        y = Inches(2.15 + i * 0.86)
        chip = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.15), y, Inches(0.62), Inches(0.62))
        chip.fill.solid(); chip.fill.fore_color.rgb = ORANGE; chip.line.fill.background()
        chip.shadow.inherit = False
        num = text(s, Inches(1.15), y + Inches(0.13), Inches(0.62), Inches(0.4),
                   [{"text": "0%d" % (i + 1), "size": 16, "bold": True,
                     "color": RGBColor(0x10, 0x0A, 0x00), "align": PP_ALIGN.CENTER}])
        lab = text(s, Inches(2.0), y + Inches(0.1), Inches(5.3), Inches(0.5),
                   [{"text": it, "size": 20, "color": WHITE}])
        rows.append((chip, num, lab))
    slide_number(s, 2, TOTAL)
    tl = Timeline()
    tl.group([tl.kenburns(bg.shape_id, 12000),
              tl.wipe(title.shape_id, "right", 700),
              tl.fade(ov.shape_id, 500, "withEffect", 0)])
    for chip, num, lab in rows:
        tl.group([tl.zoom(chip.shape_id, 500, "clickEffect"),
                  tl.fade(num.shape_id, 400, "withEffect", 50),
                  tl.fly(lab.shape_id, "from_left", 600, "withEffect", 100)])
    attach(s, tl, "push")
    return s


def planet_slide(prs, num, bg_name, title_txt, subtitle, facts, transition,
                 fly_dir="from_left"):
    s = new_slide(prs)
    bg = bg_picture(s, bg_name)
    panel = overlay(s, Inches(0.55), Inches(0.85), Inches(6.1), Inches(5.9), alpha=52)
    title = text(s, Inches(0.95), Inches(1.15), Inches(5.3), Inches(1.0),
                 [{"text": title_txt, "size": 36, "bold": True}])
    sub = text(s, Inches(0.95), Inches(2.05), Inches(5.3), Inches(0.6),
               [{"text": subtitle, "size": 16, "italic": True, "color": CYAN}])
    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.98), Inches(2.6),
                              Inches(1.7), Pt(4))
    rule.fill.solid(); rule.fill.fore_color.rgb = ORANGE; rule.line.fill.background()
    rule.shadow.inherit = False
    fact_boxes = []
    for i, f in enumerate(facts):
        fb = text(s, Inches(0.95), Inches(2.95 + i * 1.05), Inches(5.3), Inches(1.0),
                  [{"text": "●", "size": 13, "color": ORANGE}, ])
        fact_boxes.append(fb)
        fact_boxes.append(text(s, Inches(1.35), Inches(2.9 + i * 1.05), Inches(5.0), Inches(1.05),
                               [{"text": f, "size": 17, "color": WHITE}]))
    slide_number(s, num, TOTAL)
    tl = Timeline()
    tl.group([tl.kenburns(bg.shape_id, 12000),
              tl.fade(panel.shape_id, 500, "withEffect", 0),
              tl.fly(title.shape_id, fly_dir, 700, delay=100),
              tl.fade(sub.shape_id, 600, delay=600),
              tl.wipe(rule.shape_id, "right", 500, delay=900)])
    for i in range(0, len(fact_boxes), 2):
        b, t_ = fact_boxes[i], fact_boxes[i + 1]
        tl.group([tl.fade(b.shape_id, 400, "clickEffect"),
                  tl.fly(t_.shape_id, "from_left", 600, "withEffect", 80)])
    attach(s, tl, transition)
    return s


def s09_family(prs):
    s = new_slide(prs)
    bg = bg_picture(s, "section_nebula.jpg")
    ov = overlay(s, Inches(0.5), Inches(0.6), Inches(12.33), Inches(6.2), alpha=40)
    title = text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(0.9),
                 [{"text": "Семья Солнца: восемь планет", "size": 34, "bold": True,
                   "align": PP_ALIGN.CENTER}])
    planets = [("Солнце", 1.35, RGBColor(0xFD, 0xB8, 0x13)),
               ("Меркурий", 0.42, RGBColor(0x9E, 0x9E, 0x9E)),
               ("Венера", 0.62, RGBColor(0xE8, 0xC4, 0x6B)),
               ("Земля", 0.66, RGBColor(0x3B, 0x82, 0xC4)),
               ("Марс", 0.5, RGBColor(0xC1, 0x44, 0x0E)),
               ("Юпитер", 1.15, RGBColor(0xC8, 0x8B, 0x3A)),
               ("Сатурн", 1.0, RGBColor(0xE3, 0xC7, 0x8A)),
               ("Уран", 0.78, RGBColor(0x7F, 0xD4, 0xD4)),
               ("Нептун", 0.75, RGBColor(0x3E, 0x54, 0xE8))]
    xs = [0.75, 2.45, 3.55, 4.75, 5.95, 7.0, 8.75, 10.45, 11.75]
    base_y = 3.35
    shapes = []
    for (name, d, color), x in zip(planets, xs):
        size = Inches(d)
        y = Inches(base_y + (1.35 - d) / 2.0)
        ovl = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), y, size, size)
        ovl.fill.solid(); ovl.fill.fore_color.rgb = color; ovl.line.fill.background()
        ovl.shadow.inherit = False
        lab = text(s, Inches(x - 0.35), Inches(5.35), Inches(d + 0.7), Inches(0.4),
                   [{"text": name, "size": 12, "color": DIM, "align": PP_ALIGN.CENTER}])
        shapes.append((ovl, lab))
    slide_number(s, 9, TOTAL)
    tl = Timeline()
    tl.group([tl.kenburns(bg.shape_id, 12000),
              tl.fade(ov.shape_id, 500, "withEffect", 0),
              tl.fly(title.shape_id, "from_top", 700, delay=100)])
    rows = []
    for i, (ovl, lab) in enumerate(shapes):
        rows.append(tl.zoom(ovl.shape_id, 450, "withEffect", 600 + i * 300))
        rows.append(tl.fade(lab.shape_id, 400, "withEffect", 750 + i * 300))
    tl.group(rows)
    tl.group([tl.spin(shapes[6][0].shape_id, 360, 2500, "withEffect", 0)])
    attach(s, tl, "wheel")
    return s


def s10_facts(prs):
    s = new_slide(prs)
    bg = bg_picture(s, "milkyway.jpg")
    title = text(s, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.9),
                 [{"text": "Удивительные факты", "size": 38, "bold": True,
                   "align": PP_ALIGN.CENTER}])
    cards = [("01", "День на Венере длится дольше, чем её год"),
             ("02", "На Марсе закаты голубого цвета"),
             ("03", "Свет Солнца летит до Земли 8 минут 19 секунд"),
             ("04", "У планет Солнечной системы более 400 спутников")]
    boxes = []
    for i, (n, t_) in enumerate(cards):
        x = Inches(0.8 + (i % 2) * 6.15)
        y = Inches(1.8 + (i // 2) * 2.35)
        card = overlay(s, x, y, Inches(5.75), Inches(2.0), alpha=55)
        num = text(s, x + Inches(0.3), y + Inches(0.35), Inches(1.0), Inches(0.8),
                   [{"text": n, "size": 30, "bold": True, "color": ORANGE}])
        lab = text(s, x + Inches(1.35), y + Inches(0.45), Inches(4.1), Inches(1.3),
                   [{"text": t_, "size": 17, "color": WHITE}])
        boxes.append((card, num, lab))
    slide_number(s, 10, TOTAL)
    tl = Timeline()
    tl.group([tl.kenburns(bg.shape_id, 12000),
              tl.wipe(title.shape_id, "down", 700)])
    card_fx = [lambda sid: tl.zoom(sid, 700, "clickEffect"),
               lambda sid: tl.fly(sid, "from_bottom", 700, "clickEffect"),
               lambda sid: tl.wheel(sid, 2, 700, "clickEffect"),
               lambda sid: tl.circle_in(sid, 700, "clickEffect")]
    for i, (card, num, lab) in enumerate(boxes):
        tl.group([card_fx[i % len(card_fx)](card.shape_id),
                  tl.fade(num.shape_id, 500, "withEffect", 150),
                  tl.fly(lab.shape_id, "from_bottom", 600, "withEffect", 250)])
    attach(s, tl, "cover")
    return s


def s11_final(prs):
    s = new_slide(prs)
    bg = bg_picture(s, "astronaut.jpg")
    ov = overlay(s, Inches(3.1), Inches(1.5), Inches(7.1), Inches(3.2), alpha=42)
    star = text(s, Inches(6.25), Inches(1.7), Inches(0.8), Inches(0.8),
                [{"text": "★", "size": 34, "color": ORANGE, "align": PP_ALIGN.CENTER}])
    title = text(s, Inches(3.3), Inches(2.5), Inches(6.7), Inches(1.1),
                 [{"text": "Спасибо за внимание!", "size": 44, "bold": True,
                   "align": PP_ALIGN.CENTER}])
    sub = text(s, Inches(3.3), Inches(3.6), Inches(6.7), Inches(0.7),
               [{"text": "Космос ближе, чем кажется", "size": 20, "color": CYAN,
                 "align": PP_ALIGN.CENTER}])
    foot = text(s, Inches(3.3), Inches(4.25), Inches(6.7), Inches(0.5),
                [{"text": "F5 — начать показ • Esc — выйти", "size": 12, "color": DIM,
                  "align": PP_ALIGN.CENTER}])
    tl = Timeline()
    tl.group([tl.kenburns(bg.shape_id, 12000),
              tl.fade(ov.shape_id, 500, "withEffect", 0),
              tl.zoom(star.shape_id, 600, delay=200),
              tl.zoom(title.shape_id, 800, delay=500),
              tl.fly(sub.shape_id, "from_bottom", 700, delay=1100),
              tl.fade(foot.shape_id, 700, delay=1700),
              tl.spin(star.shape_id, 360, 3000, "withEffect", 2300)])
    attach(s, tl, "fade")
    return s


def main():
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    prs.core_properties.title = "Космос: путешествие по Солнечной системе"
    prs.core_properties.author = "Arena.ai"

    s01_title(prs)
    s02_toc(prs)
    planet_slide(prs, 3, "sun.jpg", "Солнце — наша звезда",
                 "жёлтый карлик, сердце Солнечной системы",
                 ["99,86 % всей массы Солнечной системы",
                  "Температура ядра — около 15 000 000 °C",
                  "Каждую секунду сгорает 4 млн тонн вещества"],
                 "fade", fly_dir="from_left")
    planet_slide(prs, 4, "earth.jpg", "Земля",
                 "третья планета от Солнца, наш дом",
                 ["Диаметр — 12 742 км",
                  "71 % поверхности покрыто океанами",
                  "Единственный известный мир с жизнью"],
                 "push", fly_dir="from_left")
    planet_slide(prs, 5, "mars.jpg", "Марс",
                 "красная планета, цель первых экспедиций",
                 ["Вулкан Олимп — высота 21 км",
                  "Сутки почти как земные: 24 ч 37 мин",
                  "Два спутника — Фобос и Деймос"],
                 "wipe", fly_dir="from_top")
    planet_slide(prs, 6, "jupiter.jpg", "Юпитер",
                 "гигант, больший всех планет вместе взятых",
                 ["В 11 раз шире Земли",
                  "Красное Пятно — шторм крупнее Земли",
                  "Почти 100 собственных спутников"],
                 "split", fly_dir="from_left")
    planet_slide(prs, 7, "saturn.jpg", "Сатурн",
                 "властелин колец",
                 ["Кольца из льда и камня — 282 000 км в ширину",
                  "Плотность меньше воды — он бы всплыл",
                  "Титан — спутник с плотной атмосферой"],
                 "wheel", fly_dir="from_left")
    planet_slide(prs, 8, "neptune.jpg", "Уран и Нептун",
                 "ледяные гиганты на краю системы",
                 ["Ветра Нептуна — до 2 100 км/ч",
                  "Год Нептуна = 165 земных лет",
                  "Уран вращается, лёжа «на боку»"],
                 "cover", fly_dir="from_right")
    s09_family(prs)
    s10_facts(prs)
    s11_final(prs)

    out = os.path.join(HERE, "Kosmos_Solnechnaya_sistema.pptx")
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main()
