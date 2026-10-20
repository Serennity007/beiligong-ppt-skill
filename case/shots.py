# -*- coding: utf-8 -*-
import sys, os
import pymupdf

B = r"C:/Users/ztl/Desktop/智能体安全_党校授课_交付_20261021/教师意见_五章结构重构_20261017/v59_教师录音纠偏_20261020"
PDF = os.path.join(B, "智能体安全_技术趋势与风险防控_v60__20261020.pdf")
OUT = os.path.join(B, "_build", "preview")
os.makedirs(OUT, exist_ok=True)
doc = pymupdf.open(PDF)
pages = [int(x) for x in sys.argv[1:]] or [2, 4, 6, 7, 30]
for p in pages:
    pg = doc[p - 1]
    pix = pg.get_pixmap(dpi=110)
    f = os.path.join(OUT, f"p{p:03d}.png")
    pix.save(f)
    print(f)
