# -*- coding: utf-8 -*-
"""导出 PDF（PowerPoint COM）并可选导出每页 PNG，用于人眼复核

用法：
    python ppt_export.py 成品.pptx --pdf 输出.pdf [--png 图片目录] [--png-pages 1,2,3]
"""
import argparse
import os
import sys

sys.path.insert(0, __file__.rsplit("\\", 1)[0].rsplit("/", 1)[0])
import pptkit as K

ap = argparse.ArgumentParser()
ap.add_argument("src")
ap.add_argument("--pdf", default=None)
ap.add_argument("--png", default=None, help="PNG 输出目录")
ap.add_argument("--png-pages", default=None, help="只导出这些页，如 1,2,3；不给则全导")
ap.add_argument("--png-w", type=int, default=1600)
a = ap.parse_args()

try:
    import win32com.client
except ImportError:
    raise SystemExit("需要 pywin32：pip install pywin32（且本机装有 PowerPoint）")

src = os.path.abspath(a.src)
app = win32com.client.DispatchEx("PowerPoint.Application")
app.DisplayAlerts = 1
deck = None
try:
    deck = app.Presentations.Open(src, ReadOnly=True, Untitled=False, WithWindow=False)
    print("页数:", deck.Slides.Count, flush=True)
    if a.pdf:
        pdf = os.path.abspath(a.pdf)
        os.makedirs(os.path.dirname(pdf) or ".", exist_ok=True)
        deck.SaveAs(pdf, 32)      # ppSaveAsPDF
        print("PDF:", pdf)
    if a.png:
        os.makedirs(a.png, exist_ok=True)
        pages = ([int(x) for x in a.png_pages.split(",")] if a.png_pages
                 else list(range(1, deck.Slides.Count + 1)))
        for n in pages:
            f = os.path.join(a.png, "p%03d.png" % n)
            deck.Slides(n).Export(f, "PNG", a.png_w)
        print("PNG:", len(pages), "页 ->", a.png)
finally:
    if deck is not None:
        deck.Close()
    app.Quit()

if a.pdf and os.path.exists(a.pdf):
    print("PDF 大小: %.1f MB" % (os.path.getsize(a.pdf) / 1048576))
