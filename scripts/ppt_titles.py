# -*- coding: utf-8 -*-
"""步骤 B：生成"统摄标题"方案

把连续若干页归为一个知识节点（nodeTag），给这组页配一个统一的页眉标题。
教师/客户口径的改写用 override 配置覆盖，其余沿用节点名。

用法：
    python ppt_titles.py chrome.json 输出.json [--override override.json] [--txt 输出.txt]

override.json 形如：
    {"智能体的定义与能力": "一、智能体的概念与技术构成",
     "风险总览": "二、风险版图：四层风险与防护控制"}
self_titled（不套统摄标题，用页内自己的标题）形如：
    {"self_titled": ["总览", "示例", "部分收束"]}
"""
import argparse
import json
import os
import sys

sys.path.insert(0, __file__.rsplit("\\", 1)[0].rsplit("/", 1)[0])
import pptkit as K

ap = argparse.ArgumentParser()
ap.add_argument("chrome")
ap.add_argument("out")
ap.add_argument("--override", default=None)
ap.add_argument("--txt", default=None)
ap.add_argument("--prefix", default="", help="统摄标题前缀，如 '智能体安全｜'")
ap.add_argument("--tag", default="nodeTag",
                help="节点标签形状名包含的串；找不到时自动退化为按页内标题分组")
a = ap.parse_args()
try:
    K.check_output_paths({'chrome': a.chrome, 'override': a.override}, {'out': a.out, 'txt': a.txt})
except (ValueError, OSError) as exc:
    ap.error(str(exc))

ov = {}
SELF = set()
cfg = K.load_config(a.override)
for k, v in cfg.items():
    if k == "self_titled":
        SELF = set(v)
    else:
        ov[k] = v

data = json.load(open(a.chrome, encoding="utf-8"))
pages = data["pages"]

# 1) 收集每页节点标签
tags = []
for r in pages:
    t = (r["roles"].get("nodetag") or [{}])[0].get("text", "")
    tags.append(t.strip())

if not any(tags):
    # 退化：用页内标题当标签（成品稿常已把节点标签消费掉）
    print("提示：未发现名为 *%s* 的节点标签形状，退化为按页内标题分组" % a.tag)
    tags = [(r["roles"].get("title") or [{}])[0].get("text", "").strip() for r in pages]

# 2) 连续归并成组
groups = []
for i, t in enumerate(tags, 1):
    if groups and groups[-1]["tag"] == t:
        groups[-1]["pages"].append(i)
    else:
        groups.append({"tag": t, "pages": [i]})

# 3) 定标题
plan = []
for grp in groups:
    tag = grp["tag"]
    sub = (pages[grp["pages"][0] - 1]["roles"].get("subtitle") or [{}])[0].get("text", "")
    if not tag or tag in SELF:
        title = ""          # 保留页内自有标题
        why = "自有标题"
    elif tag in ov:
        title = ov[tag]
        why = "override"
    else:
        title = a.prefix + tag
        why = "节点名"
    plan.append({
        "tag": tag, "title": title, "why": why,
        "from": grp["pages"][0], "to": grp["pages"][-1],
        "n": len(grp["pages"]), "first_subtitle": sub[:60],
    })

json.dump({"groups": len(plan), "plan": plan},
          open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

if a.txt:
    L = []
    for p in plan:
        L.append("P%03d-%03d (%2d页)  %-28s -> %s   [%s]"
                 % (p["from"], p["to"], p["n"], (p["tag"] or "(无)")[:28], p["title"] or "(保留页内标题)", p["why"]))
    open(a.txt, "w", encoding="utf-8").write("\n".join(L))

print("节点组:", len(plan))
print("已 override:", sum(1 for p in plan if p["why"] == "override"))
print("输出:", a.out)
