# -*- coding: utf-8 -*-
import json

P = r"C:/Users/ztl/Desktop/智能体安全_党校授课_交付_20261021/教师意见_五章结构重构_20261017/v59_教师录音纠偏_20261020/_build/scan_v58.json"
d = json.load(open(P, encoding="utf-8"))

KEYS = ("chrome:title", "chrome:subtitle", "chrome:takeaway", "chrome:pageno",
        "v57:nodeTag", "supp:takeaway", "supp:subtitle", "v57:num", "v57:label", "v57:oneliner")

def flat(shapes, out):
    for s in shapes:
        out.append(s)
        if "children" in s:
            flat(s["children"], out)

rows = []
for sl in d["slides"]:
    fl = []
    flat(sl["shapes"], fl)
    rec = {"idx": sl["idx"]}
    for s in fl:
        n = s.get("name") or ""
        if n in KEYS and s.get("text"):
            rec.setdefault(n, []).append(s["text"].strip().replace("\n", " / ")[:120])
    rows.append(rec)

with open(r"C:/Users/ztl/Desktop/智能体安全_党校授课_交付_20261021/教师意见_五章结构重构_20261017/v59_教师录音纠偏_20261020/_build/chrome_dump.txt", "w", encoding="utf-8") as f:
    for r in rows:
        parts = [f"[{r['idx']:3d}]"]
        for k in KEYS:
            if k in r:
                parts.append(f"{k}={r[k]}")
        f.write(" | ".join(parts) + "\n")
print("done", len(rows))
