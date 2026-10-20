# -*- coding: utf-8 -*-
"""生成统摄标题方案（先出计划表，人工核对后再执行）"""
import json, re
from collections import OrderedDict

P = r"C:/Users/ztl/Desktop/智能体安全_党校授课_交付_20261021/教师意见_五章结构重构_20261017/v59_教师录音纠偏_20261020/_build/scan_v58.json"
d = json.load(open(P, encoding="utf-8"))

def flat(shapes, out):
    for s in shapes:
        out.append(s)
        if "children" in s:
            flat(s["children"], out)

# 人工确定的统摄标题映射（键为 nodeTag 的 "部分 · 节点" 原文）
OVERRIDE = {
    "一 · 定义与能力演进": "智能体的定义与能力演进",
    "一 · 政策部署与目标": "智能体的政策部署与目标",
    "一 · 运行机制与关键技术": "智能体的运行机制与关键技术",
    "一 · 多智能体协作与连接": "多智能体协作与连接",
    "一 · 工具调用与应用边界": "工具调用与应用边界",
    "二 · 风险总览": "四层风险总览",
    "二 · 基础设施层": "基础设施层风险与防护",
    "二 · 管理层": "管理层风险与防护",
    "二 · 协同层": "协同层风险与防护",
    "二 · 应用接口层": "应用接口层风险与防护",
    "二 · 防线落地": "四层防线的落地与验收",
    "三 · 回到任务链": "回到任务链看验证环节",
    "三 · 评测基准": "安全评测基准",
    "三 · 我国治理规则与合规要求": "我国治理规则与合规要求",
    "三 · 国际政策工具与标准": "国际政策工具与标准",
    "三 · 模型自身风险": "模型自身风险",
    "四 · 四类典型事件": "四类典型安全事件",
    "四 · 越界机理": "越界机理",
    "四 · 公开通报与近期事件": "公开通报与近期事件",
    "四 · 证据链与复盘方法": "证据链与复盘方法",
    "五 · 协作网络的架构与运行": "协作网络的架构与运行",
    "五 · 跨主体协作的信任缺口与瓶颈": "跨主体协作的信任缺口与瓶颈",
    "五 · 链上机制与链下控制": "链上机制与链下控制",
    "五 · 可信协作基础设施与研究方向": "可信协作基础设施与研究方向",
    "六 · 运行环境": "运行环境与安全要求",
    "六 · 四项安全能力": "四项安全能力",
    "六 · 失控事件与根因": "失控事件与根因",
    "六 · 单位行动与全篇收束": "单位行动与全篇收束",
}
# 单页节点：统摄标题直接用页面原标题
SELF_TITLED = {"总览", "示例", "部分收束"}

info = {}
for sl in d["slides"]:
    fl = []
    flat(sl["shapes"], fl)
    tag = None
    title = None
    for s in fl:
        if s.get("name") == "v57:nodeTag" and s.get("text"):
            tag = s["text"].strip()
        if s["type"] == "TEXT_BOX (17)" and s.get("text") and (s["t"] or 9) < 0.7 \
           and (s["l"] or 9) < 5 and (s["h"] or 0) < 1.0 and s.get("name") != "chrome:pageno":
            title = s["text"].strip()
    info[sl["idx"]] = (tag, title)

groups = OrderedDict()
for i in sorted(info):
    tag, title = info[i]
    if not tag:
        continue
    groups.setdefault(tag, []).append((i, title))

lines = []
for tag, pages in groups.items():
    m = re.match(r"^([一二三四五六])\s*·\s*(.+)$", tag)
    part, node = (m.group(1), m.group(2)) if m else ("", tag)
    if len(pages) >= 2 and node not in SELF_TITLED:
        gov = OVERRIDE.get(tag, node)
    else:
        gov = None  # 用页面原标题
    for idx, t in pages:
        lines.append((idx, tag, gov if gov else (t or ""), t or "", gov is None))

with open(r"C:/Users/ztl/Desktop/智能体安全_党校授课_交付_20261021/教师意见_五章结构重构_20261017/v59_教师录音纠偏_20261020/_build/title_plan.txt", "w", encoding="utf-8") as f:
    for idx, tag, gov, sub, selfd in lines:
        f.write(f"{idx}\t{tag}\t统摄={gov}\t小标题={'（无）' if selfd else sub}\n")

print("节点组数:", len(groups))
for tag, pages in groups.items():
    print(f"  {tag:28s} {len(pages):3d} 页  {pages[0][0]}–{pages[-1][0]}")
print("\n未覆盖标题的页:", [i for i in sorted(info) if not info[i][0] and 2 <= i <= 185])
