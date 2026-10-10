# -*- coding: utf-8 -*-
"""Проверка собранного .pptx: целостность XML, ссылки анимаций на фигуры,
уникальность id таймлайна, порядок элементов слайда."""
import re
import sys
import zipfile
from lxml import etree

PATH = "Kosmos_Solnechnaya_sistema.pptx"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

errors = []
zf = zipfile.ZipFile(PATH)

# 1. Все XML-части валидны (well-formed)
for name in zf.namelist():
    if name.endswith(".xml") or name.endswith(".rels"):
        try:
            etree.fromstring(zf.read(name))
        except Exception as e:
            errors.append("%s: XML error: %s" % (name, e))

# 2. Проверки по слайдам
slide_names = sorted([n for n in zf.namelist()
                      if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                     key=lambda s: int(re.search(r"(\d+)", s).group(1)))
print("slides:", len(slide_names))

total_anims = 0
for name in slide_names:
    root = etree.fromstring(zf.read(name))
    # id всех фигур (cNvPr под sp / pic / grpSp)
    shape_ids = set()
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        if tag == "cNvPr":
            shape_ids.add(el.get("id"))
    # spid в тайминге
    spids = [el.get("spid") for el in root.iter()
             if el.tag == P + "spTgt"]
    for sp in spids:
        if sp not in shape_ids:
            errors.append("%s: animation targets missing shape id %s" % (name, sp))
    # уникальность cTn id
    ids = [el.get("id") for el in root.iter() if el.tag == P + "cTn"]
    dup = set(i for i in ids if ids.count(i) > 1)
    if dup:
        errors.append("%s: duplicate cTn ids %s" % (name, dup))
    # порядок: cSld ... transition ... timing последним
    kids = [c.tag for c in root]
    if P + "timing" in kids and kids[-1] != P + "timing":
        errors.append("%s: timing is not the last child" % name)
    if P + "transition" in kids and P + "timing" in kids:
        if kids.index(P + "transition") > kids.index(P + "timing"):
            errors.append("%s: transition after timing" % name)
    # кликовые группы
    groups = root.findall(".//" + P + "seq/" + P + "cTn/" + P + "childTnLst/" + P + "par")
    total_anims += len(spids)
    print("%s: shapes=%d anim_targets=%d click_groups=%d cTn_ids=%d"
          % (name, len(shape_ids), len(spids), len(groups), len(ids)))

print("total animation targets:", total_anims)

# 3. python-pptx открывает файл
from pptx import Presentation
prs = Presentation(PATH)
print("python-pptx opens OK, slides:", len(prs.slides))

if errors:
    print("\nERRORS:")
    for e in errors:
        print(" -", e)
    sys.exit(1)
print("\nALL CHECKS PASSED")
