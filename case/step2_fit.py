# -*- coding: utf-8 -*-
"""v60 步骤2：PowerPoint 实测每个文字框，放不下的按比例回退到能放下的最大字号"""
import json, os, sys
import win32com.client
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
B = r"C:/Users/ztl/Desktop/智能体安全_党校授课_交付_20261021/教师意见_五章结构重构_20261017/v59_教师录音纠偏_20261020"
SRC = os.path.join(B, "_build", "v60_step1.pptx")
OUTJSON = os.path.join(B, "_build", "fit_result.json")
TOL = 3.0
STEPS = [1.0, 0.94, 0.88, 0.82, 0.76, 0.70]

recs = json.load(open(os.path.join(B, "_build", "size_map.json"), encoding="utf-8"))
mp = {(r["page"], r["id"]): r for r in recs}

app = win32com.client.DispatchEx("PowerPoint.Application")
app.DisplayAlerts = 1
deck = None
result = []
try:
    deck = app.Presentations.Open(str(SRC), ReadOnly=False, Untitled=False, WithWindow=False)
    total = deck.Slides.Count
    print("slides:", total, flush=True)

    def collect(shapes, page, out):
        for k in range(1, shapes.Count + 1):
            s = shapes.Item(k)
            if s.Type == 6:
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
            pnew = rec["new"]
            porig = rec["orig"]
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
                    return max(sl - bl, st - bt,
                               (bl + bw) - (sl + sw), (bt + bh) - (st + shg))
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
                if ov <= TOL:
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

json.dump(result, open(OUTJSON, "w", encoding="utf-8"), ensure_ascii=False)
print("回退档位分布:", dict(Counter(r["k"] for r in result)))
bad = [r for r in result if r["over"] > TOL]
print("仍越框(已退回原字号):", len(bad))
for r in bad[:25]:
    print("  P", r["page"], r["name"], r["over"], r["applied"])
