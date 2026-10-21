# -*- coding: utf-8 -*-
"""步骤 C：白底深字改造

三件事：
  1. 每一页的页面背景刷成纯白（只改形状不改背景，投影时仍是绿底）；
  2. 大面积浅绿/浅灰底纹去掉填充 —— 但同步把上面的字改成深色（只读性红线）；
  3. 深绿/深灰等主色加深一档（006633 -> 173404）。

用法：
    python ppt_white.py 源文件.pptx 输出.pptx [--config white_config.json] [--log 改动.json]

white_config.json 示例见 ../case/white_config.json
"""
import argparse
import json
import sys
from collections import Counter

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.util import Pt

sys.path.insert(0, __file__.rsplit("\\", 1)[0].rsplit("/", 1)[0])
import pptkit as K

ap = argparse.ArgumentParser()
ap.add_argument("src")
ap.add_argument("out")
ap.add_argument("--config", default=None)
ap.add_argument("--log", default=None)
ap.add_argument("--titles", default=None,
                help="ppt_titles.py 产出的方案 json，用于写入统摄标题 + 页内小标题")
a = ap.parse_args()

DEF = {
    "page_bg": "FFFFFF",
    "defill": ["E8F5E9", "F1F8F1", "EFF5EF", "F5F5F5", "FAFAFA"],
    "min_defill_w": 0.4,          # 小于这个宽度（英寸）的色块不动，多是色条/标记
    "min_defill_h": 0.28,
    "light_text_to": "173404",    # 浅底上的白字改成什么
    "light_text_to_body": "4A4A4A",
    "lum_threshold": 0.62,        # 底色亮度高于此值，认为其上的白字不可读
    "font_map": {
        "006633": "173404", "006C39": "173404", "044A2E": "122C03", "17865A": "173404",
        "595959": "2E2E2E", "6E6E6E": "3D3D3D", "222222": "141414", "404040": "262626",
        "4A4A4A": "2A2A2A", "9A9A9A": "555555", "BFBFBF": "666666", "C0C0C0": "666666",
        "CC3333": "B02020",
    },
    "fill_map": {"006633": "173404", "006C39": "173404"},
    "line_map": {},
    "font": "微软雅黑",
}
cfg = K.load_config(a.config, DEF)
DEFILL = set(x.upper() for x in cfg["defill"])
FMAP = {k.upper(): v.upper() for k, v in (cfg.get("fill_map") or {}).items()}
CMAP = {k.upper(): v.upper() for k, v in (cfg.get("font_map") or {}).items()}
LMAP = {k.upper(): v.upper() for k, v in (cfg.get("line_map") or {}).items()}
TH = float(cfg["lum_threshold"])

prs = Presentation(a.src)
log = []
stat = Counter()

# ---------- 1. 页面背景 ----------
for i, sl in enumerate(prs.slides, 1):
    try:
        f = sl.background.fill
        f.solid()
        f.fore_color.rgb = RGBColor.from_string(cfg["page_bg"].upper())
        stat["页面背景"] += 1
        log.append({"page": i, "what": "页面背景", "to": cfg["page_bg"]})
    except Exception as e:
        log.append({"page": i, "what": "页面背景失败", "err": str(e)})


def page_bg_fallback(sl):
    try:
        f = sl.background.fill
        if f.type == MSO_FILL_TYPE.SOLID:
            return str(f.fore_color.rgb).upper()
    except Exception:
        pass
    return cfg["page_bg"].upper()


# ---------- 2. 逐形状处理 ----------
for i, sl in enumerate(prs.slides, 1):
    bg = page_bg_fallback(sl)
    for sh, _, _ in K.iter_shapes(sl.shapes, i):
        g = K.geo(sh)
        fill = K.fill_hex(sh)

        # 2a. 去浅底
        removed = False
        if fill in DEFILL and g["w"] >= cfg["min_defill_w"] and g["h"] >= cfg["min_defill_h"]:
            K.set_fill(sh, None)
            removed = True
            stat["去浅底"] += 1
            log.append({"page": i, "id": sh.shape_id, "what": "去浅底", "from": fill})

        # 2b. 主色加深
        if fill in FMAP:
            K.set_fill(sh, FMAP[fill])
            stat["填充加深"] += 1

        # 2c. 线条
        try:
            lc = str(sh.line.color.rgb).upper()
            if lc in LMAP:
                sh.line.color.rgb = RGBColor.from_string(LMAP[lc])
        except Exception:
            pass

        if not getattr(sh, "has_text_frame", False):
            continue
        # 字所在"底"：自身填充优先，否则页面背景
        base = (bg if removed else (fill or bg))
        if fill is None and sh._element.getparent() is sl.shapes._spTree:
            # 仅检查文字之前的顶层形状：后面的形状不可能是文字背后的色带。
            cx, cy = g["l"] + g["w"] / 2, g["t"] + g["h"] / 2
            for behind in sl.shapes:
                if behind._element is sh._element:
                    break
                gb = K.geo(behind)
                if gb["l"] <= cx <= gb["l"] + gb["w"] and gb["t"] <= cy <= gb["t"] + gb["h"]:
                    color = K.fill_hex(behind)
                    if color:
                        base = color
                    elif behind.shape_type == 13:  # 图片底色不可静态判定
                        base = None
        light_bg = base is not None and K.luminance(base) > TH

        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                try:
                    c = str(r.font.color.rgb).upper()
                except Exception:
                    c = None
                if c in CMAP:
                    r.font.color.rgb = RGBColor.from_string(CMAP[c])
                    stat["字色加深"] += 1
                elif light_bg and c in ("FFFFFF", "FFFEFE", "F5F5F5", "E8F5E9"):
                    # 浅底白字：必改，否则整页"看不见字"
                    target = cfg["light_text_to_body"] if (g["h"] > 1.2 or g["w"] > 6) else cfg["light_text_to"]
                    r.font.color.rgb = RGBColor.from_string(target.upper())
                    stat["浅底白字改深"] += 1
                    log.append({"page": i, "id": sh.shape_id, "what": "浅底白字改深",
                                "from": c, "to": target, "text": r.text[:30]})
                K.set_ea_font(r, cfg.get("font", K.FONT_DEFAULT))

# ---------- 2.5 清编辑定位信息与导读（教师硬要求：B4/B5/F1） ----------
purge = cfg.get("purge") or {}
if purge.get("on"):
    import re as _re
    pats = [_re.compile("(?:" + p + r")(?:\s*[：:].*)?") for p in purge.get("patterns", [])]
    keep = set(purge.get("keep_pages", []))
    n_purge = 0
    for i, sl in enumerate(prs.slides, 1):
        if i in keep:
            continue
        for sh in list(sl.shapes):
            if not getattr(sh, "has_text_frame", False):
                continue
            txt = K.text_of(sh)
            if not txt.strip():
                continue
            for para in list(sh.text_frame.paragraphs):
                matched = next((pat for pat in pats if pat.fullmatch(para.text.strip())), None)
                if matched:
                    # 只删独立定位/备注段落，保留同框正文及其格式。
                    para._p.getparent().remove(para._p)
                    n_purge += 1
                    stat["删导读/页段"] += 1
                    log.append({"page": i, "what": "删导读或定位信息",
                                "text": para.text.strip()[:40], "hit": matched.pattern[:20]})
            if not K.text_of(sh).strip():
                K.remove_element(sh)
    print("删导读/页段/定位信息:", n_purge, "个框")

# ---------- 3. 写统摄标题 / 页内小标题（可选） ----------
if a.titles:
    from pptx.enum.shapes import MSO_SHAPE_TYPE as MST
    from pptx.util import Emu, Pt as _Pt

    plan = json.load(open(a.titles, encoding="utf-8"))["plan"]
    pm = {}
    for p in plan:
        if not p.get("title"):
            continue
        for i in range(p["from"], p["to"] + 1):
            pm[i] = p["title"]

    def set_one_line(sh, txt, size, bold, color):
        tf = sh.text_frame
        tf.word_wrap = True
        # 清空原有段落与 run，只留一行
        for k in range(len(tf.paragraphs) - 1, 0, -1):
            tf._txBody.remove(tf.paragraphs[k]._p)
        p0 = tf.paragraphs[0]
        for r in list(p0.runs):
            r._r.getparent().remove(r._r)
        r = p0.add_run()
        r.text = txt
        r.font.size = _Pt(size)
        r.font.bold = bold
        r.font.color.rgb = RGBColor.from_string(color.upper())
        K.set_ea_font(r, cfg.get("font", K.FONT_DEFAULT))

    n_title = 0
    for i, sl in enumerate(prs.slides, 1):
        want = pm.get(i)
        if not want:
            continue
        title_sh, sub_sh = None, None
        for sh in sl.shapes:
            if sh.shape_type == MST.GROUP or not getattr(sh, "has_text_frame", False):
                continue
            g = K.geo(sh)
            txt = K.text_of(sh).strip()
            if title_sh is None and g["t"] < 0.75 and g["l"] < 5.0 and g["w"] > 3.0 and txt:
                title_sh = sh
            if sub_sh is None and 0.88 <= g["t"] <= 1.25 and g["l"] < 1.6 and g["w"] > 8.0 and txt:
                sub_sh = sh
        if title_sh is None:
            continue
        old = K.text_of(title_sh).strip()
        if sub_sh is None and old and old != want:
            # 本页原标题降为页内小标题（后面 ppt_deco.py 会给它套色带）
            box = sl.shapes.add_textbox(Emu(int(0.68 * K.IN)), Emu(int(0.96 * K.IN)),
                                        Emu(int(12.0 * K.IN)), Emu(int(0.44 * K.IN)))
            set_one_line(box, old, 15, True, "4A4A4A")
            log.append({"page": i, "what": "原标题降为小标题", "text": old[:30]})
        if old != want:
            set_one_line(title_sh, want, 20, True, cfg["light_text_to"])
            n_title += 1
            log.append({"page": i, "what": "写统摄标题", "text": want[:30]})
    print("统摄标题写入:", n_title, "页")

prs.save(a.out)
if a.log:
    json.dump(log, open(a.log, "w", encoding="utf-8"), ensure_ascii=False)

print("保存:", a.out)
print("页数:", len(prs.slides._sldIdLst), "（包校验:", K.page_count_zip(a.out), "）")
print("改动统计:", dict(stat))
