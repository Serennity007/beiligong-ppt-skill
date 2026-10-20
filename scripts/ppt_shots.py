# -*- coding: utf-8 -*-
"""把 PDF 渲染成 PNG 缩略图（跨平台，不依赖 PowerPoint）

导出 PDF 之后想快速翻一遍效果时用；也可以在 Linux/Mac 上给同事看。
用法：
    python ppt_shots.py 成品.pdf 输出目录 [--pages 1,2,3,10-20] [--dpi 110]
"""
import argparse
import os
import sys

ap = argparse.ArgumentParser()
ap.add_argument("pdf")
ap.add_argument("out")
ap.add_argument("--pages", default=None, help="如 1,2,5-9；不给则全渲染")
ap.add_argument("--dpi", type=int, default=110)
a = ap.parse_args()

try:
    import fitz  # pymupdf
except ImportError:
    raise SystemExit("需要 pymupdf：pip install pymupdf")

os.makedirs(a.out, exist_ok=True)
doc = fitz.open(a.pdf)
want = None
if a.pages:
    want = set()
    for part in a.pages.split(","):
        part = part.strip()
        if "-" in part:
            s, e = part.split("-")
            want.update(range(int(s), int(e) + 1))
        else:
            want.add(int(part))

mat = fitz.Matrix(a.dpi / 72.0, a.dpi / 72.0)
n = 0
for i, page in enumerate(doc, 1):
    if want and i not in want:
        continue
    pix = page.get_pixmap(matrix=mat)
    f = os.path.join(a.out, "p%03d.png" % i)
    pix.save(f)
    n += 1
print("渲染", n, "页 ->", a.out)
