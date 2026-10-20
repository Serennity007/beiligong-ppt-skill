# -*- coding: utf-8 -*-
"""步骤 E-2：用 PowerPoint 实测每个文字框，放不下的按比例回退到能放下的最大字号

为什么必须用 COM 实测：python-pptx 拿不到"文字实际占多大地方"，
只有 PowerPoint 的 TextFrame2.TextRange.BoundWidth/BoundHeight 才是真的墨迹边界。
不实测就放大，等于闭着眼睛排版。

原理：
  k 从 1.0 逐级往下试 [1.0, 0.94, 0.88, 0.82, 0.76, 0.70]，
  每级写入字号后量一次四边溢出，<= 容差即停；
  地板 kfloor = 原字号/目标字号 —— 保证绝不缩到比改造前还小。

用法（Windows + 已装 PowerPoint）：
    python ppt_fit.py 上一步.pptx --map size_map.json --out fit_result.json [--tol 3.0]
"""
import argparse
import json
import os
import sys
from collections import Counter

sys.path.insert(0, __file__.rsplit("\\", 1)[0].rsplit("/", 1)[0])
import pptkit as K

ap = argparse.ArgumentParser()
ap.add_argument("src")
ap.add_argument("--map", required=True)
ap.add_argument("--out", required=True)
ap.add_argument("--tol", type=float, default=3.0, help="允许溢出磅数")
ap.add_argument("--steps", default="1.0,0.94,0.88,0.82,0.76,0.70")
a = ap.parse_args()

STEPS = [float(x) for x in a.steps.split(",")]
recs = json.load(open(a.map, encoding="utf-8"))
mp = {(r["page"], r["id"]): r for r in recs}

try:
    import win32com.client
except ImportError:
    raise SystemExit("需要 pywin32：pip install pywin32（且本机装有 PowerPoint）")

app = win32com.client.DispatchEx("PowerPoint.Application")
app.DisplayAlerts = 1
deck = None
result = []
try:
    deck = app.Presentations.Open(os.path.abspath(a.src), ReadOnly=False,
                                  Untitled=False, WithWindow=False)
    total = deck.Slides.Count
    print("slides:", total, flush=True)

    def collect(shapes, page, out):
        for k in range(1, shapes.Count + 1):
            s = shapes.Item(k)
            if s.Type == 6:                       # msoGroup
                try:
                    collect(s.GroupItems, page, out)
                except Exception:
                    pass
                continue
            if not s.HasTextFrame:
                continue
            try:
                if not s.TextFrame2.TextRange.Text.strip():
                    continue
            except Exception:
                continue
            out.append((page, s))

    for n in range(1, total + 1):
        lst = []
        collect(deck.Slides(n).Shapes, n, lst)
        for page, s in lst:
            try:
                sid = int(s.Id)
            except Exception:
                continue
            rec = mp.get((page, sid))
            if rec is None:
                continue
            tr = s.TextFrame2.TextRange
            try:
                pcount = tr.Paragraphs.Count
            except Exception:
                pcount = 0
            pnew, porig = rec["new"], rec["orig"]
            use_par = (pcount == len(pnew) and pcount > 0)

            def apply(k):
                try:
                    if use_par:
                        for i in range(1, pcount + 1):
                            tr.Paragraphs.Item(i).Font.Size = round(pnew[i - 1] * k, 1)
                    else:
                        tr.Font.Size = round(max(pnew) * k, 1)
                except Exception:
                    pass

            def overflow():
                try:
                    bl, bt = float(tr.BoundLeft), float(tr.BoundTop)
                    bw, bh = float(tr.BoundWidth), float(tr.BoundHeight)
                    sl, st = float(s.Left), float(s.Top)
                    sw, shg = float(s.Width), float(s.Height)
                    return max(sl - bl, st - bt, (bl + bw) - (sl + sw), (bt + bh) - (st + shg))
                except Exception:
                    return 0.0

            kfloor = 1.0
            for o, nn in zip(porig, pnew):
                if o and nn:
                    kfloor = min(kfloor, o / nn)
            chosen, ov = kfloor, 0.0
            for k in STEPS:
                kk = max(k, kfloor)
                apply(kk)
                ov = overflow()
                chosen = kk
                if ov <= a.tol:
                    break
            result.append({"page": page, "id": sid, "name": rec["name"],
                           "k": round(chosen, 3), "over": round(ov, 1),
                           "new": pnew, "applied": [round(v * chosen, 1) for v in pnew]})
        if n % 25 == 0:
            print("...page", n, len(result), flush=True)
    print("测量完成", len(result), flush=True)
finally:
    if deck is not None:
        deck.Close()
    app.Quit()

json.dump(result, open(a.out, "w", encoding="utf-8"), ensure_ascii=False)
print("回退档位分布:", dict(Counter(r["k"] for r in result)))
bad = [r for r in result if r["over"] > a.tol]
print("仍越框(已退回原字号仍放不下):", len(bad))
for r in bad[:25]:
    print("  P", r["page"], r["name"], r["over"], r["applied"])
