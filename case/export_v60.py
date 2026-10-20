# -*- coding: utf-8 -*-
from pathlib import Path
import sys, hashlib
import win32com.client

B = Path(__file__).resolve().parent.parent
SRC = B / "智能体安全_技术趋势与风险防控_v60__20261020.pptx"
OUT = B / "智能体安全_技术趋势与风险防控_v60__20261020.pdf"
sys.stdout.reconfigure(encoding="utf-8")
before = hashlib.sha256(SRC.read_bytes()).hexdigest()
app = win32com.client.DispatchEx("PowerPoint.Application")
app.DisplayAlerts = 1
deck = None
try:
    deck = app.Presentations.Open(str(SRC), ReadOnly=True, Untitled=False, WithWindow=False)
    print("slides:", deck.Slides.Count, flush=True)
    deck.SaveAs(str(OUT), 32)
    print("Exported ->", OUT, flush=True)
finally:
    if deck is not None:
        deck.Close()
    app.Quit()
assert hashlib.sha256(SRC.read_bytes()).hexdigest() == before, "PDF导出不得回写PPTX"
