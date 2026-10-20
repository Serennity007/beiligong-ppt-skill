# -*- coding: utf-8 -*-
"""步骤 A：扫描版式骨架（页眉横幅 / 统摄标题 / 页内小标题 / 底部条 / 页码 / 色条）

用途：改造任何一份 PPTX 之前，先摸清它"页眉写了什么、小标题在哪里、底部有没有绿条"，
      后面所有脚本都依赖这份清单。

用法：
    python ppt_chrome.py 源文件.pptx 输出.json [--txt 输出.txt] [--slide-h 7.5]
"""
import argparse
import json
import re
import sys

from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.path.insert(0, __file__.rsplit("\\", 1)[0].rsplit("/", 1)[0])
import pptkit as K

ap = argparse.ArgumentParser()
ap.add_argument("src")
ap.add_argument("out")
ap.add_argument("--txt", default=None, help="同时写一份人类可读的 txt")
ap.add_argument("--slide-h", type=float, default=7.5, help="页面高度（英寸），用于判定底部区")
ap.add_argument("--tag", default="nodeTag", help="节点标签形状名包含的串")
a = ap.parse_args()

# 版式判定阈值（13.333 x 7.5 的 16:9；其他尺寸按比例缩放）
SCALE = a.slide_h / 7.5
BANNER_W, BANNER_T = 12.0, 0.5 * SCALE
TITLE_T, TITLE_L, TITLE_W = 0.75 * SCALE, 5.0, 3.0
SUB_T0, SUB_T1, SUB_L, SUB_W = 0.88 * SCALE, 1.25 * SCALE, 1.6, 8.0
BOTTOM_T, BOTTOM_W = a.slide_h - 1.9 * SCALE, 8.0
ACC_W0, ACC_W1, ACC_H0, ACC_H1 = 0.15, 0.40, 0.5 * SCALE, 1.0 * SCALE


def classify(sh):
    g = K.geo(sh)
    txt = K.text_of(sh).strip()
    is_auto = sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
    if is_auto and g["w"] > BANNER_W and g["t"] < BANNER_T:
        return "banner"
    if is_auto and ACC_W0 < g["w"] < ACC_W1 and ACC_H0 < g["h"] < ACC_H1 and g["t"] < BANNER_T:
        return "accent"
    if g["t"] < TITLE_T and g["l"] < TITLE_L and g["w"] > TITLE_W and txt:
        return "title"
    if SUB_T0 <= g["t"] <= SUB_T1 and g["l"] < SUB_L and g["w"] > SUB_W and txt:
        return "subtitle"
    if g["t"] > BOTTOM_T and g["w"] > BOTTOM_W:
        return "bottom"
    if g["w"] < 1.2 and g["h"] < 0.5 and g["t"] > a.slide_h - 0.9 * SCALE and re.fullmatch(r"\d{1,3}", txt or ""):
        return "pageno"
    return None


prs = Presentation(a.src)
pages = []
for i, sl in enumerate(prs.slides, 1):
    rec = {"page": i, "bg": "FFFFFF", "roles": {}}
    try:
        f = sl.background.fill
        if f.type == MSO_FILL_TYPE.SOLID:
            rec["bg"] = str(f.fore_color.rgb).upper()
    except Exception:
        pass
    for sh, _, _ in K.iter_shapes(sl.shapes, i):
        role = classify(sh)
        if role is None:
            continue
        item = {
            "id": sh.shape_id, "name": sh.name, "fill": K.fill_hex(sh),
            "text": K.text_of(sh).strip()[:120], **K.geo(sh),
        }
        rec["roles"].setdefault(role, []).append(item)
    for sh, _, _ in K.iter_shapes(sl.shapes, i):
        nm = sh.name or ""
        if a.tag and a.tag in nm:
            rec["roles"].setdefault("nodetag", []).append(
                {"id": sh.shape_id, "name": nm, "text": K.text_of(sh).strip()[:80]})
    pages.append(rec)

json.dump({"src": a.src, "n": len(pages), "pages": pages},
          open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

if a.txt:
    L = []
    for r in pages:
        def g(role):
            return (r["roles"].get(role) or [{}])[0]
        L.append("P%-4d | 横幅:%-22s | 标题:%-26s | 小标题:%-24s | 底部:%-16s | 页码:%s" % (
            r["page"],
            g("banner").get("text", "")[:22],
            g("title").get("text", "")[:26],
            g("subtitle").get("text", "")[:24],
            g("bottom").get("text", "")[:16],
            g("pageno").get("text", "")))
    open(a.txt, "w", encoding="utf-8").write("\n".join(L))

from collections import Counter
c = Counter()
for r in pages:
    for k in r["roles"]:
        c[k] += 1
print("已扫描页数:", len(pages))
print("角色命中:", dict(c))
print("输出:", a.out)
