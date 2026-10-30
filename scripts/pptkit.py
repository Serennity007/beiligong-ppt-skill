# -*- coding: utf-8 -*-
"""pptkit —— ppt制作 skill 的公共工具模块

所有 scripts/ 下的脚本都 `import pptkit`，复用这里的遍历、取色、改色、字体等能力。
"""
import json
import os
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu

IN = 914400.0          # 1 英寸 = 914400 EMU
FONT_DEFAULT = "微软雅黑"


def check_output_paths(inputs, outputs):
    """Reject input/output aliases before scripts read or write any artifacts."""
    source = {key: Path(value).expanduser().resolve() for key, value in inputs.items() if value}
    targets = {key: Path(value).expanduser().resolve() for key, value in outputs.items() if value}

    def same(left, right):
        return left == right or (left.exists() and right.exists() and os.path.samefile(left, right))

    previous = {}
    for name, target in targets.items():
        for input_name, path in source.items():
            if same(path, target):
                raise ValueError(f"输出 {name} 与输入 {input_name} 指向同一文件，请改用其他路径")
        for output_name, path in previous.items():
            if same(path, target):
                raise ValueError(f"输出 {name} 与输出 {output_name} 指向同一文件，请分别指定路径")
        previous[name] = target


# ---------------- 基础遍历 ----------------

def iter_shapes(shapes, page=0, _depth=0):
    """递归遍历形状，产出 (shape, page, depth)。组合（GROUP）会被展开。"""
    for sh in shapes:
        yield sh, page, _depth
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                yield from iter_shapes(sh.shapes, page, _depth + 1)
            except Exception:
                pass


def geo(sh):
    """返回形状的位置尺寸（英寸），缺失返回 None"""
    try:
        return {
            "l": (sh.left or 0) / IN, "t": (sh.top or 0) / IN,
            "w": (sh.width or 0) / IN, "h": (sh.height or 0) / IN,
        }
    except Exception:
        return {"l": 0.0, "t": 0.0, "w": 0.0, "h": 0.0}


def text_of(sh):
    try:
        if sh.has_text_frame:
            return sh.text_frame.text
    except Exception:
        pass
    return ""


# ---------------- 填充 / 颜色 ----------------

def fill_hex(sh):
    """实心填充色，字符串大写 HEX；非实心或无填充返回 None"""
    try:
        f = sh.fill
        if f.type == MSO_FILL_TYPE.SOLID:
            return str(f.fore_color.rgb).upper()
    except Exception:
        pass
    return None


def set_fill(sh, hexcolor):
    f = sh.fill
    if hexcolor is None:
        f.background()
    else:
        f.solid()
        f.fore_color.rgb = RGBColor.from_string(hexcolor.upper())


def luminance(hexcolor):
    """相对亮度 0~1，>0.62 视为浅底（白字不可读）"""
    h = hexcolor.upper()
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def set_ea_font(run, name=FONT_DEFAULT):
    """同时写 latin / ea / cs 三种字体的 typeface，保证中文字体真的生效"""
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", name)


def remove_element(sh):
    el = sh._element
    el.getparent().remove(el)


# ---------------- 配置 ----------------

def load_config(path, defaults=None):
    """读取 JSON 配置，缺失键用 defaults 兜底"""
    cfg = dict(defaults or {})
    if path and os.path.exists(path):
        cfg.update(json.load(open(path, "r", encoding="utf-8")))
    return cfg


def open_deck(path):
    prs = Presentation(path)
    return prs


def page_count_zip(path):
    """不依赖 python-pptx，直接从包里数 sldId，用于校验写入后页数没变"""
    import zipfile
    with zipfile.ZipFile(path) as z:
        xml = z.read("ppt/presentation.xml").decode("utf-8")
    return xml.count("<p:sldId ")


def round_half(v):
    return round(v * 2) / 2
